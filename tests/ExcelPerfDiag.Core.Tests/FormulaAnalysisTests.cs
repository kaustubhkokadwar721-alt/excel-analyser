using ExcelPerfDiag.Core.Model;
using ExcelPerfDiag.Core.Analysis;
using Xunit;

namespace ExcelPerfDiag.Core.Tests;

public class FormulaAnalysisTests
{
    [Theory]
    [InlineData("=VLOOKUP(A2,Data!$A:$H,3,FALSE)", FunctionClass.Lookup)]
    [InlineData("=_xlfn.XLOOKUP(A2,Data!$A:$A,Data!$C:$C,\"na\")", FunctionClass.Lookup)]
    [InlineData("=SUMPRODUCT((Data!$B$2:$B$100=A2)*Data!$C$2:$C$100)", FunctionClass.Array)]
    [InlineData("=SUMIFS(Data!$C$2:$C$100,Data!$B$2:$B$100,A2)", FunctionClass.Cheap)]
    [InlineData("=NOW()", FunctionClass.Volatile)]
    [InlineData("=INDIRECT(\"Data!C\"&ROW())", FunctionClass.Volatile)]
    [InlineData("=SUM(OFFSET(Data!$C$1,ROW()-1,0,5,1))", FunctionClass.Volatile)]
    [InlineData("=_xlfn._xlws.FILTER(Data!$A$2:$H$100,Data!$C$2:$C$100=\"N\")", FunctionClass.DynamicArray)]
    [InlineData("=_xlfn.REDUCE(0,SEQUENCE(10),_xlfn.LAMBDA(a,b,a+b))", FunctionClass.Lambda)]
    [InlineData("=A2*B2", FunctionClass.Cheap)]
    public void Classify_Cases(string formula, FunctionClass expected)
        => Assert.Equal(expected, FormulaAnalysis.Classify(formula));

    [Fact]
    public void Indirect_Is_Volatile_And_SingleThreaded()
    {
        const string f = "=INDIRECT(\"Data!C\"&ROW())";
        Assert.True(FormulaAnalysis.IsVolatile(f));
        Assert.True(FormulaAnalysis.IsSingleThreaded(f));
    }

    [Theory]
    [InlineData("=SUMIF(Data!C:C,\"N\",Data!F:F)", true)]      // sheet-qualified full column
    [InlineData("=VLOOKUP(A2,Data!$A:$H,3,FALSE)", true)]       // full-column lookup array
    [InlineData("=A:A", true)]
    [InlineData("=SUMIFS(Data!$C$2:$C$10,Data!$B$2:$B$10,A2)", false)] // bounded
    public void HasFullColumn_Cases(string formula, bool expected)
        => Assert.Equal(expected, FormulaAnalysis.HasFullColumn(formula));

    [Fact]
    public void CellsTouched_FullColumn_Saturates_Bucket()
    {
        var n = FormulaAnalysis.CellsTouched("=VLOOKUP(A2,Data!$A:$H,3,FALSE)");
        Assert.True(n >= 1_048_576);
        Assert.Equal(99, FormulaAnalysis.CellsBucket(n));
    }

    [Fact]
    public void CellsTouched_BoundedRange_AreaSum()
    {
        // $C$2:$C$11 = 10 cells
        Assert.Equal(10, FormulaAnalysis.CellsTouched("=SUM(Data!$C$2:$C$11)"));
    }

    [Fact]
    public void ColToNum_Works()
    {
        Assert.Equal(1, FormulaAnalysis.ColToNum("A"));
        Assert.Equal(27, FormulaAnalysis.ColToNum("AA"));
        Assert.Equal(3, FormulaAnalysis.ColToNum("C"));
    }

    [Fact]
    public void NormalizeR1C1_CopiedRelative_Collapses()
    {
        // Same formula shape copied down a column must produce identical R1C1.
        var a = FormulaAnalysis.NormalizeR1C1("=VLOOKUP(A2,Data!$A:$H,3,FALSE)", 2, 2);
        var b = FormulaAnalysis.NormalizeR1C1("=VLOOKUP(A3,Data!$A:$H,3,FALSE)", 3, 2);
        Assert.Equal(a, b);
        Assert.Contains("R[0]C[-1]", a);          // A2 relative to B2
        Assert.Contains("Data!$A:$H", a);          // sheet-qualified ref left fixed
    }

    [Fact]
    public void GroupingKey_CollapsesCopies_ButSplitsByCellsBucket()
    {
        var k1 = FormulaAnalysis.GroupingKey("=A2*2", 2, 2);
        var k2 = FormulaAnalysis.GroupingKey("=A3*2", 3, 2);
        Assert.Equal(k1, k2);  // copied relative -> one group

        // Same R1C1 shape but very different referenced-range size -> different bucket.
        var small = FormulaAnalysis.GroupingKey("=SUM(B2:B11)", 2, 1);
        var big = FormulaAnalysis.GroupingKey("=SUM(B2:B100002)", 2, 1);
        Assert.NotEqual(small.Bucket, big.Bucket);
    }
}
