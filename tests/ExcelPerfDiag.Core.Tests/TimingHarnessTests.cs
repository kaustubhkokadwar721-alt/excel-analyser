using ExcelPerfDiag.Core.Abstractions;
using ExcelPerfDiag.Core.Model;
using ExcelPerfDiag.Core.Timing;
using Xunit;

namespace ExcelPerfDiag.Core.Tests;

/// <summary>Deterministic fake session — exercises the harness logic without Excel.</summary>
internal sealed class FakeSession : IExcelSession
{
    private bool _mt = true;
    public string ExcelVersion => "fake-16";
    public int MeasureCalls;
    public int LastCopies;

    public void Open(string path) { }
    public double TimeFullCalc(int iters) => _mt ? 10.0 : 20.0;   // multi=10ms, single=20ms -> eff 2.0
    public double TimeRecalc(int iters) => 5.0;                    // vol% = 5/10 = 50
    public void SetMultiThreaded(bool enabled) => _mt = enabled;
    public bool HasSpill(string sheet, string a1) => false;
    public string GetFormulaR1C1(string sheet, string a1) => "=R1C1";
    public double MeasureScratchUs(string r1c1, int copies, int iters, out double stdevUs)
    {
        MeasureCalls++; LastCopies = copies; stdevUs = 0.1;
        return 2.0; // 2 µs/occ
    }
    public void TimeOpenSave(out double openMs, out double saveMs) { openMs = 100; saveMs = 50; }
    public void Dispose() { }
}

public class TimingHarnessTests
{
    private static DiagnosticReport ReportWith(params PatternCost[] patterns)
    {
        var r = new DiagnosticReport();
        r.Patterns.AddRange(patterns);
        return r;
    }

    [Fact]
    public void FillsWorkbookRatios()
    {
        var r = ReportWith(new PatternCost { Sheet = "S", R1C1 = "=R1C1", SampleCell = "A1", Occurrences = 100 });
        var susp = new Dictionary<string, double> { ["S"] = 1 };

        TimingHarness.Run(new FakeSession(), r, susp, new TimingOptions(Iterations: 3, TopK: 6));

        Assert.Equal(10.0, r.Summary.MultiThreadMs);
        Assert.Equal(20.0, r.Summary.SingleThreadMs);
        Assert.Equal(2.0, r.Summary.MtEfficiency);
        Assert.Equal(50.0, r.Summary.VolatilityPct);
        Assert.Equal(100, r.Summary.OpenMs);
    }

    [Fact]
    public void MeasuresPatterns_AndComputesTotalMs()
    {
        var p = new PatternCost { Sheet = "S", R1C1 = "=R1C1", SampleCell = "A1", Occurrences = 500 };
        var r = ReportWith(p);
        TimingHarness.Run(new FakeSession(), r, new Dictionary<string, double> { ["S"] = 1 },
            new TimingOptions(Iterations: 3));

        Assert.True(p.Measured);
        Assert.Equal(2.0, p.UsPerOcc);
        Assert.Equal(1.0, p.TotalMs, 3);   // 2 µs * 500 / 1000 = 1.0 ms
    }

    [Fact]
    public void OnlyTimesWorstKSheets()
    {
        var fake = new FakeSession();
        var r = ReportWith(
            new PatternCost { Sheet = "Hot", R1C1 = "=a", SampleCell = "A1", Occurrences = 10 },
            new PatternCost { Sheet = "Cold", R1C1 = "=b", SampleCell = "A1", Occurrences = 10 });
        var susp = new Dictionary<string, double> { ["Hot"] = 100, ["Cold"] = 1 };

        TimingHarness.Run(fake, r, susp, new TimingOptions(TopK: 1));

        Assert.True(r.Patterns.Single(p => p.Sheet == "Hot").Measured);
        Assert.False(r.Patterns.Single(p => p.Sheet == "Cold").Measured);
    }

    [Fact]
    public void DynamicArray_Pattern_UsesSingleCopy()
    {
        var fake = new FakeSession();
        var p = new PatternCost { Sheet = "S", R1C1 = "=FILTER", SampleCell = "A1", Occurrences = 5000, FuncClass = FunctionClass.DynamicArray };
        TimingHarness.Run(fake, ReportWith(p), new Dictionary<string, double> { ["S"] = 1 }, new TimingOptions());
        Assert.Equal(1, fake.LastCopies);   // spill-safety: never mass-replicate
    }
}
