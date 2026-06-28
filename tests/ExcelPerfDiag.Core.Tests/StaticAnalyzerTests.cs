using ExcelPerfDiag.Core.Analysis;
using ExcelPerfDiag.Core.Model;
using ExcelPerfDiag.Fixtures;
using Xunit;

namespace ExcelPerfDiag.Core.Tests;

public sealed class StaticAnalyzerTests : IDisposable
{
    private readonly string _path;
    private readonly StaticResult _r;

    public StaticAnalyzerTests()
    {
        _path = FixtureBuilder.BuildSmall();
        _r = StaticAnalyzer.Analyze(_path);
    }

    public void Dispose()
    {
        try { File.Delete(_path); } catch { /* best effort */ }
    }

    private PatternCost? Find(string sheet, FunctionClass cls) =>
        _r.Patterns.FirstOrDefault(p => p.Sheet == sheet && p.FuncClass == cls);

    [Fact]
    public void CopiedFormulas_CollapseToOnePattern()
    {
        var vlookup = Find("Lookup_Legacy", FunctionClass.Lookup);
        Assert.NotNull(vlookup);
        Assert.Equal(30, vlookup!.Occurrences);   // 30 copied rows -> one pattern
    }

    [Fact]
    public void LegacyVsModern_Classified()
    {
        Assert.NotNull(Find("Lookup_Legacy", FunctionClass.Lookup));
        Assert.NotNull(Find("Lookup_Modern", FunctionClass.Lookup));   // XLOOKUP
        Assert.NotNull(Find("Agg_Legacy", FunctionClass.Array));       // SUMPRODUCT
        Assert.NotNull(Find("Agg_Modern", FunctionClass.Cheap));       // SUMIFS
    }

    [Fact]
    public void Volatile_And_SingleThreaded_Flagged()
    {
        var vols = _r.Patterns.Where(p => p.Sheet == "Volatile" && p.IsVolatile).ToList();
        Assert.NotEmpty(vols);
        var indirect = _r.Patterns.First(p => p.Sheet == "Volatile" && p.Funcs.Contains("INDIRECT"));
        Assert.True(indirect.IsSingleThreaded);
    }

    [Fact]
    public void FullColumn_Pattern_Saturates_CellsTouched()
    {
        var full = Find("FullColumn", FunctionClass.Cheap);   // SUMIF
        Assert.NotNull(full);
        Assert.True(full!.CellsTouched >= 1_048_576);
    }

    [Fact]
    public void Suspicion_RanksRiskySheetsAboveCheap()
    {
        var volatileSusp = _r.Suspicion["Volatile"];
        var fullColSusp = _r.Suspicion["FullColumn"];
        var aggModernSusp = _r.Suspicion["Agg_Modern"];
        Assert.True(volatileSusp > aggModernSusp);
        Assert.True(fullColSusp > aggModernSusp);
    }

    [Fact]
    public void Structure_FindsSheets_AndStyleCount()
    {
        Assert.Equal(7, _r.Structure.Count);            // 7 worksheets
        Assert.True(_r.StyleCount >= 1);
        Assert.Equal(0, _r.ExternalLinkCount);
    }
}
