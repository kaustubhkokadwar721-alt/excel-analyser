using System.Diagnostics;
using System.Runtime.InteropServices;
using ExcelPerfDiag.Core.Abstractions;

namespace ExcelPerfDiag.Excel;

/// <summary>
/// Live Excel host via late-bound COM (no PIA dependency; any Excel 2016+). Runs
/// inside the crash-isolated worker. Operates on a temp COPY, saves/restores app
/// state, manages a scratch sheet, and assigns Excel to a Job Object so it can never
/// orphan. Port of the verified <c>dynamic_layer.py</c> timing harness.
/// </summary>
public sealed class ExcelComSession : IExcelSession
{
    private const int XlCalculationManual = -4135;

    private dynamic _app = null!;
    private dynamic? _wb;
    private dynamic? _scratch;
    private object? _savedCalc;
    private string _copyPath = "";
    private string _tmpDir = "";
    private readonly JobObject _job = new();
    private bool _disposed;

    public string ExcelVersion { get; private set; } = "";

    public void Open(string path)
    {
        _tmpDir = Path.Combine(Path.GetTempPath(), "perfdiag_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_tmpDir);
        _copyPath = Path.Combine(_tmpDir, Path.GetFileName(path));
        File.Copy(path, _copyPath, overwrite: true);

        var type = Type.GetTypeFromProgID("Excel.Application")
                   ?? throw new InvalidOperationException("Excel is not installed (no Excel.Application ProgID).");
        _app = Activator.CreateInstance(type)!;
        _app.Visible = false;
        _app.DisplayAlerts = false;
        _app.ScreenUpdating = false;
        _app.EnableEvents = false;
        ExcelVersion = (string)_app.Version;

        _wb = _app.Workbooks.Open(_copyPath, 0); // UpdateLinks:=0

        // Assign the spawned Excel to the kill-on-close job — no orphan, ever.
        try
        {
            var hwnd = (int)_app.Hwnd;
            _ = JobObject.GetWindowThreadProcessId((IntPtr)hwnd, out var pid);
            _job.AssignProcess((int)pid);
        }
        catch { /* job assignment is best-effort */ }

        try { _savedCalc = _app.Calculation; _app.Calculation = XlCalculationManual; }
        catch { /* Calculation can only be set with a workbook open; ignore */ }
    }

    public double TimeFullCalc(int iters)
    {
        var samples = new List<double>();
        for (var i = 0; i < iters; i++)
        {
            var sw = Stopwatch.StartNew();
            _app.CalculateFull();
            sw.Stop();
            if (i > 0) samples.Add(sw.Elapsed.TotalSeconds);
        }
        return Mean(samples) * 1000.0;
    }

    public double TimeRecalc(int iters)
    {
        var samples = new List<double>();
        for (var i = 0; i < iters; i++)
        {
            var sw = Stopwatch.StartNew();
            _app.Calculate();
            sw.Stop();
            if (i > 0) samples.Add(sw.Elapsed.TotalSeconds);
        }
        return Mean(samples) * 1000.0;
    }

    public void SetMultiThreaded(bool enabled)
    {
        try { _app.MultiThreadedCalculation.Enabled = enabled; } catch { /* legacy Excel */ }
    }

    public bool HasSpill(string sheet, string a1)
    {
        try { return (bool)_wb!.Worksheets[sheet].Range[a1].HasSpill; }
        catch { return false; }
    }

    public string GetFormulaR1C1(string sheet, string a1)
        => (string)_wb!.Worksheets[sheet].Range[a1].FormulaR1C1;

    public double MeasureScratchUs(string r1c1, int copies, int iters, out double stdevUs)
    {
        var sc = GetScratch();
        try { sc.UsedRange.ClearContents(); } catch { /* clear spill residue */ }

        dynamic rng = sc.Range[sc.Cells[1, 1], sc.Cells[copies, 1]];
        rng.ClearContents();
        rng.FormulaR1C1 = r1c1;

        var patt = new List<double>();
        for (var i = 0; i < iters; i++)
        {
            var sw = Stopwatch.StartNew();
            rng.Calculate();
            sw.Stop();
            if (i > 0) patt.Add(sw.Elapsed.TotalSeconds);
        }

        rng.ClearContents();
        rng.FormulaR1C1 = "=1";
        double over = MeanCalc(rng, iters);   // explicit double: dynamic-arg call is otherwise typed dynamic
        rng.ClearContents();

        var perCell = patt.Select(p => Math.Max(p - over, 0.0) / copies * 1_000_000.0).ToList();
        stdevUs = Pstdev(perCell);
        return perCell.Count > 0 ? perCell.Average() : 0.0;
    }

    public void TimeOpenSave(out double openMs, out double saveMs)
    {
        openMs = 0; saveMs = 0;
        try
        {
            if (_scratch is not null)
            {
                _app.DisplayAlerts = false;
                _scratch.Delete();
                _scratch = null;
            }
            var ts = Stopwatch.StartNew();
            _wb!.Save();
            saveMs = ts.Elapsed.TotalMilliseconds;

            _wb.Close(false);
            _wb = null;

            var to = Stopwatch.StartNew();
            _wb = _app.Workbooks.Open(_copyPath, 0);
            openMs = to.Elapsed.TotalMilliseconds;
        }
        catch { /* best-effort */ }
    }

    private dynamic GetScratch()
    {
        if (_scratch is null)
        {
            try { _scratch = _wb!.Worksheets["_perf_scratch"]; } catch { _scratch = null; }
            if (_scratch is null)
            {
                _scratch = _wb!.Worksheets.Add();
                _scratch.Name = "_perf_scratch";
            }
        }
        return _scratch;
    }

    private static double MeanCalc(dynamic rng, int iters)
    {
        var samples = new List<double>();
        for (var i = 0; i < iters; i++)
        {
            var sw = Stopwatch.StartNew();
            rng.Calculate();
            sw.Stop();
            if (i > 0) samples.Add(sw.Elapsed.TotalSeconds);
        }
        return Mean(samples);
    }

    private static double Mean(IReadOnlyCollection<double> xs) => xs.Count > 0 ? xs.Average() : 0.0;

    private static double Pstdev(IReadOnlyCollection<double> xs)
    {
        if (xs.Count == 0) return 0;
        var m = xs.Average();
        return Math.Sqrt(xs.Sum(x => (x - m) * (x - m)) / xs.Count);
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        try { if (_savedCalc is not null) _app.Calculation = _savedCalc; } catch { }
        try { _wb?.Close(false); } catch { }
        try { _app?.Quit(); } catch { }
        try { if (_wb is not null) Marshal.FinalReleaseComObject(_wb); } catch { }
        try { if (_app is not null) Marshal.FinalReleaseComObject(_app); } catch { }
        _wb = null; _app = null!;
        GC.Collect();
        GC.WaitForPendingFinalizers();

        _job.Dispose(); // kills the Excel process if anything is still alive
        try { Directory.Delete(_tmpDir, recursive: true); } catch { }
    }
}
