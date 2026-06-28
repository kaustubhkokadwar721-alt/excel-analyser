using System.Diagnostics;
using ExcelPerfDiag.Core.Abstractions;
using ExcelPerfDiag.Core.Model;

namespace ExcelPerfDiag.Core.Timing;

public sealed record TimingOptions(
    double BudgetSeconds = 150,
    int Iterations = 6,
    int MaxCopies = 2000,
    int TopK = 6);

/// <summary>
/// Layer 2 orchestration (pure — drives an <see cref="IExcelSession"/>): workbook
/// full/recalc/threading ratios, then batch-and-subtract per pattern on the worst K
/// sheets, bounded by a wall-clock budget with spill/full-column safety. Port of
/// the verified <c>dynamic_layer.py</c>.
/// </summary>
public static class TimingHarness
{
    public static void Run(
        IExcelSession session,
        DiagnosticReport report,
        IReadOnlyDictionary<string, double> suspicion,
        TimingOptions opt,
        Action<string, double>? progress = null)
    {
        var sw = Stopwatch.StartNew();
        var s = report.Summary;
        s.ExcelVersion = session.ExcelVersion;

        // ---- workbook-level ----
        session.SetMultiThreaded(true);
        s.MultiThreadMs = session.TimeFullCalc(opt.Iterations);
        s.FullCalcMs = s.MultiThreadMs;
        s.RecalcMs = session.TimeRecalc(opt.Iterations);
        s.VolatilityPct = s.FullCalcMs > 0 ? Math.Round(100 * s.RecalcMs / s.FullCalcMs, 1) : 0;

        try
        {
            session.SetMultiThreaded(false);
            s.SingleThreadMs = session.TimeFullCalc(Math.Max(3, opt.Iterations - 2));
            session.SetMultiThreaded(true);
        }
        catch { s.SingleThreadMs = s.MultiThreadMs; }
        s.MtEfficiency = s.MultiThreadMs > 0 ? Math.Round(s.SingleThreadMs / s.MultiThreadMs, 2) : 1;
        progress?.Invoke("workbook timing", sw.Elapsed.TotalSeconds);

        // ---- per-pattern timing on the worst K sheets ----
        var worst = suspicion.OrderByDescending(kv => kv.Value)
                             .Take(opt.TopK).Select(kv => kv.Key)
                             .ToHashSet(StringComparer.Ordinal);
        var todo = report.Patterns
            .Where(p => worst.Contains(p.Sheet))
            .OrderByDescending(p => (double)p.Occurrences * (p.IsVolatile ? 8 : 2))
            .ToList();

        var measured = 0;
        foreach (var p in todo)
        {
            if (sw.Elapsed.TotalSeconds > opt.BudgetSeconds)
            {
                s.Partial = true;
                s.Notes.Add($"Time budget {opt.BudgetSeconds}s hit after {measured} patterns; report is partial.");
                break;
            }
            try
            {
                var r1c1 = session.GetFormulaR1C1(p.Sheet, p.SampleCell);
                var copies = Math.Min(opt.MaxCopies, Math.Max(200, p.Occurrences));

                var spill = p.FuncClass is FunctionClass.DynamicArray or FunctionClass.Lambda
                            || session.HasSpill(p.Sheet, p.SampleCell);
                if (spill) copies = 1;
                if (p.CellsTouched >= 1_048_576 || p.FuncClass == FunctionClass.Array)
                    copies = Math.Min(copies, 100);

                var us = session.MeasureScratchUs(r1c1, copies, opt.Iterations, out var sd);

                // two-point linearity check on large patterns
                if (p.Occurrences >= 1000 && !spill)
                {
                    var usSmall = session.MeasureScratchUs(r1c1, Math.Max(100, copies / 4), opt.Iterations, out _);
                    if (us > 0 && Math.Abs(us - usSmall) / us > 0.5) p.Nonlinear = true;
                }

                p.UsPerOcc = Math.Round(us, 3);
                p.StdevUs = Math.Round(sd, 3);
                p.TotalMs = Math.Round(us * p.Occurrences / 1000.0, 2);
                p.Measured = true;
                measured++;
                progress?.Invoke($"measured {p.Sheet}!{p.SampleCell}", sw.Elapsed.TotalSeconds);
            }
            catch (Exception e)
            {
                p.FlagReason = $"measure error: {e.Message}";
            }
        }

        // ---- direct open/save timing ----
        try
        {
            session.TimeOpenSave(out var openMs, out var saveMs);
            s.OpenMs = openMs;
            s.SaveMs = saveMs;
        }
        catch { /* best-effort */ }
    }
}
