using ClosedXML.Excel;
using ExcelPerfDiag.Core.Model;

namespace ExcelPerfDiag.Core.Analysis;

/// <summary>Result of Layer 1 (static) analysis — no Excel required.</summary>
public sealed record StaticResult(
    List<PatternCost> Patterns,
    List<StructureFinding> Structure,
    Dictionary<string, double> Suspicion,
    List<string> DefinedNames,
    int StyleCount,
    int ExternalLinkCount);

/// <summary>
/// Layer 1 — static analyzer. Reads an .xlsx via ClosedXML, groups formulas into
/// cost patterns by (R1C1, cells-bucket), computes structure, and ranks each sheet's
/// suspicion (which sheets Layer 2 should deep-time). Port of <c>static_layer.py</c>.
/// </summary>
public static class StaticAnalyzer
{
    private sealed class Group
    {
        public int Occ;
        public string Sample = "";
        public string Addr = "";
        public string R1C1 = "";
        public long Cells = 1;
        public bool Vol;
        public bool St;
        public FunctionClass Cls;
        public string[] Funcs = System.Array.Empty<string>();
    }

    public static StaticResult Analyze(string path)
    {
        // Read raw-zip parts first, before ClosedXML opens (and may lock) the file.
        var styleCount = StructureAnalysis.StyleCount(path);
        var linkCount = StructureAnalysis.ExternalLinkCount(path);

        using var wb = new XLWorkbook(path);

        var patterns = new List<PatternCost>();
        var structure = new List<StructureFinding>();
        var suspicion = new Dictionary<string, double>(StringComparer.Ordinal);

        foreach (var ws in wb.Worksheets)
        {
            var formulaCells = new Dictionary<(int, int), string>();
            var groups = new Dictionary<(string, string, int), Group>();

            foreach (var cell in ws.CellsUsed(c => c.HasFormula))
            {
                string ftext;
                try { ftext = "=" + cell.FormulaA1; }
                catch { continue; }
                if (string.IsNullOrWhiteSpace(cell.FormulaA1)) continue;

                int r = cell.Address.RowNumber, c = cell.Address.ColumnNumber;
                formulaCells[(r, c)] = ftext;

                var (r1c1, bucket) = FormulaAnalysis.GroupingKey(ftext, r, c);
                var key = (ws.Name, r1c1, bucket);
                if (!groups.TryGetValue(key, out var g))
                {
                    g = new Group
                    {
                        Sample = ftext,
                        Addr = cell.Address.ToString()!,
                        R1C1 = r1c1,
                        Cells = FormulaAnalysis.CellsTouched(ftext),
                        Cls = FormulaAnalysis.Classify(ftext),
                        Vol = FormulaAnalysis.IsVolatile(ftext),
                        St = FormulaAnalysis.IsSingleThreaded(ftext),
                        Funcs = FormulaAnalysis.Functions(ftext).Distinct().OrderBy(x => x, StringComparer.Ordinal).ToArray(),
                    };
                    groups[key] = g;
                }
                g.Occ++;
            }

            double sheetSusp = 0;
            foreach (var (key, g) in groups)
            {
                patterns.Add(new PatternCost
                {
                    Sheet = key.Item1,
                    R1C1 = g.R1C1,
                    SampleA1 = g.Sample,
                    SampleCell = g.Addr,
                    FuncClass = g.Cls,
                    Occurrences = g.Occ,
                    CellsTouched = g.Cells,
                    IsVolatile = g.Vol,
                    IsSingleThreaded = g.St,
                    Funcs = g.Funcs,
                });

                double w = g.Occ * ClassWeight(g.Cls);
                if (g.Vol) w *= 3;
                if (g.St) w *= 2;
                if (g.Cells >= 1_048_576) w *= 2;
                sheetSusp += w;
            }
            suspicion[ws.Name] = sheetSusp;

            structure.Add(AnalyzeSheetStructure(ws, formulaCells));
        }

        var definedNames = wb.DefinedNames.Select(n => n.Name).ToList();
        return new StaticResult(patterns, structure, suspicion, definedNames, styleCount, linkCount);
    }

    private static StructureFinding AnalyzeSheetStructure(IXLWorksheet ws, Dictionary<(int, int), string> formulaCells)
    {
        long populated = ws.CellsUsed().Count();
        var last = ws.LastCellUsed(XLCellsUsedOptions.All);
        long maxRow = last?.Address.RowNumber ?? 1;
        long maxCol = last?.Address.ColumnNumber ?? 1;
        long dimCells = maxRow * maxCol;
        double waste = dimCells > 0 ? (1 - (double)populated / dimCells) : 0.0;

        int cfRules = 0;
        bool cfFull = false;
        try
        {
            foreach (var cf in ws.ConditionalFormats)
            {
                cfRules++;
                var addr = cf.Range?.RangeAddress;
                if (addr is not null && addr.LastAddress.RowNumber >= 1_048_576) cfFull = true;
            }
        }
        catch { /* CF parsing is best-effort */ }

        return new StructureFinding
        {
            Sheet = ws.Name,
            Dimension = last is null ? "A1" : $"A1:{last.Address}",
            PopulatedCells = populated,
            DimensionCells = dimCells,
            UsedRangeWastePct = Math.Round(waste * 100, 1),
            CfRuleCount = cfRules,
            CfFullColumn = cfFull,
            MaxChainDepth = StructureAnalysis.MaxChainDepth(formulaCells),
        };
    }

    private static double ClassWeight(FunctionClass c) => c switch
    {
        FunctionClass.Volatile => 8,
        FunctionClass.Array => 6,
        FunctionClass.Lookup => 4,
        FunctionClass.DynamicArray => 3,
        FunctionClass.Lambda => 3,
        FunctionClass.Let => 1,
        FunctionClass.Cheap => 1,
        _ => 2,
    };

    public static double FileSizeMb(string path) =>
        Math.Round(new FileInfo(path).Length / (1024.0 * 1024.0), 3);
}
