namespace ExcelPerfDiag.Core.Abstractions;

/// <summary>
/// Abstracts a live Excel host so the timing harness never touches COM directly
/// (keeps Core platform-agnostic and unit-testable with a fake). Implemented by
/// ExcelPerfDiag.Excel via late-bound COM, hosted inside the crash-isolated worker.
/// </summary>
public interface IExcelSession : IDisposable
{
    string ExcelVersion { get; }

    /// <summary>Open a working copy of the file and enter the safety envelope (manual calc, etc.).</summary>
    void Open(string path);

    double TimeFullCalc(int iters);     // CalculateFull, averaged ms (warm-up dropped)
    double TimeRecalc(int iters);       // Application.Calculate, averaged ms
    void SetMultiThreaded(bool enabled);

    bool HasSpill(string sheet, string a1);
    string GetFormulaR1C1(string sheet, string a1);

    /// <summary>Batch-and-subtract: returns measured microseconds per occurrence.</summary>
    double MeasureScratchUs(string r1c1, int copies, int iters, out double stdevUs);

    void TimeOpenSave(out double openMs, out double saveMs);
}
