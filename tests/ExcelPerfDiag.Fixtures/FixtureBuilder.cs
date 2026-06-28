using ClosedXML.Excel;

namespace ExcelPerfDiag.Fixtures;

/// <summary>
/// Generates small, deterministic test workbooks with known legacy-vs-modern
/// formulas on shared data. Used by static-layer tests (no Excel) and as the
/// seed for integration fixtures. C# port of make_small_sample.py.
/// </summary>
public static class FixtureBuilder
{
    /// <summary>Build a small fixture workbook; returns the path written.</summary>
    public static string BuildSmall(string? path = null)
    {
        path ??= Path.Combine(Path.GetTempPath(), $"perfdiag_fixture_{Guid.NewGuid():N}.xlsx");
        const int n = 50;   // data rows
        const int m = 30;   // formula rows
        var regions = new[] { "North", "South", "East", "West" };
        var rnd = new Random(1);

        using var wb = new XLWorkbook();

        var d = wb.AddWorksheet("Data");
        d.Cell(1, 1).Value = "ID"; d.Cell(1, 2).Value = "Region";
        d.Cell(1, 3).Value = "Qty"; d.Cell(1, 4).Value = "Price";
        for (var i = 1; i <= n; i++)
        {
            d.Cell(i + 1, 1).Value = i;
            d.Cell(i + 1, 2).Value = regions[rnd.Next(regions.Length)];
            d.Cell(i + 1, 3).Value = rnd.Next(1, 500);
            d.Cell(i + 1, 4).Value = Math.Round(rnd.NextDouble() * 500, 2);
        }
        var last = n + 1;

        var ll = wb.AddWorksheet("Lookup_Legacy");
        for (var r = 2; r <= m + 1; r++)
        {
            ll.Cell(r, 1).Value = rnd.Next(1, n);
            ll.Cell(r, 2).FormulaA1 = $"VLOOKUP(A{r},Data!$A$2:$D${last},2,FALSE)";
            ll.Cell(r, 3).FormulaA1 = $"INDEX(Data!$C$2:$C${last},MATCH(A{r},Data!$A$2:$A${last},0))";
        }

        var lm = wb.AddWorksheet("Lookup_Modern");
        for (var r = 2; r <= m + 1; r++)
        {
            lm.Cell(r, 1).Value = rnd.Next(1, n);
            lm.Cell(r, 2).FormulaA1 = $"XLOOKUP(A{r},Data!$A$2:$A${last},Data!$B$2:$B${last},\"na\")";
        }

        var al = wb.AddWorksheet("Agg_Legacy");
        for (var i = 0; i < regions.Length; i++)
        {
            al.Cell(i + 2, 1).Value = regions[i];
            al.Cell(i + 2, 2).FormulaA1 = $"SUMPRODUCT((Data!$B$2:$B${last}=A{i + 2})*Data!$C$2:$C${last})";
        }

        var am = wb.AddWorksheet("Agg_Modern");
        for (var i = 0; i < regions.Length; i++)
        {
            am.Cell(i + 2, 1).Value = regions[i];
            am.Cell(i + 2, 2).FormulaA1 = $"SUMIFS(Data!$C$2:$C${last},Data!$B$2:$B${last},A{i + 2})";
        }

        var vol = wb.AddWorksheet("Volatile");
        for (var r = 2; r <= 21; r++)
        {
            vol.Cell(r, 1).FormulaA1 = "NOW()";
            vol.Cell(r, 2).FormulaA1 = "RAND()";
            vol.Cell(r, 3).FormulaA1 = "INDIRECT(\"Data!C\"&ROW())";
        }

        var fc = wb.AddWorksheet("FullColumn");
        for (var r = 2; r <= 11; r++)
            fc.Cell(r, 1).FormulaA1 = "SUMIF(Data!C:C,\"North\",Data!D:D)";

        wb.SaveAs(path);
        return path;
    }
}
