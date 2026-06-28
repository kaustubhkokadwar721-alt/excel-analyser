using ExcelPerfDiag.Core.Model;
using Xunit;

namespace ExcelPerfDiag.Core.Tests;

public class ModelTests
{
    [Theory]
    [InlineData(FunctionClass.DynamicArray, "dynamic-array")]
    [InlineData(FunctionClass.Volatile, "volatile")]
    [InlineData(FunctionClass.Cheap, "cheap")]
    [InlineData(FunctionClass.Lookup, "lookup")]
    public void FunctionClass_WireRoundTrip(FunctionClass c, string wire)
    {
        Assert.Equal(wire, c.ToWire());
        Assert.Equal(c, FunctionClassExtensions.FromWire(wire));
    }

    [Fact]
    public void PatternCost_Defaults_And_Mutation()
    {
        var p = new PatternCost { Sheet = "Data", R1C1 = "=RC[-1]*2", FuncClass = FunctionClass.Cheap };
        Assert.Equal(1, p.CellsTouched);
        Assert.False(p.Measured);

        p.UsPerOcc = 5.0;
        p.Occurrences = 100;
        p.TotalMs = p.UsPerOcc * p.Occurrences / 1000.0;
        p.Measured = true;

        Assert.Equal(0.5, p.TotalMs, 3);
        Assert.True(p.Measured);
    }

    [Fact]
    public void DiagnosticReport_StartsEmpty()
    {
        var r = new DiagnosticReport();
        Assert.Empty(r.Patterns);
        Assert.Empty(r.Actions);
        Assert.False(r.Summary.Partial);
        Assert.Null(r.Deltas);
    }
}
