using System.IO.Compression;
using System.Text.RegularExpressions;

namespace ExcelPerfDiag.Core.Analysis;

/// <summary>
/// Structural (non-formula) analysis: dependency-chain depth, styles.xml cell-format
/// count, external links. Port of the verified Python <c>structure.py</c>.
/// </summary>
public static partial class StructureAnalysis
{
    // Same-sheet A1 cell ref (cross-sheet refs, preceded by '!', are ignored as leaves).
    [GeneratedRegex(@"(?<![A-Za-z0-9_$!'])\$?([A-Za-z]{1,3})\$?(\d+)(?![A-Za-z0-9_(])")]
    private static partial Regex SelfRefRe();

    // Optional namespace prefix: openpyxl writes <cellXfs ...>, ClosedXML writes <x:cellXfs ...>.
    [GeneratedRegex("<(?:\\w+:)?cellXfs count=\"(\\d+)\"")]
    private static partial Regex CellXfsRe();

    [GeneratedRegex("<(?:\\w+:)?xf ")]
    private static partial Regex XfRe();

    public static List<(int Row, int Col)> SelfRefs(string formula)
    {
        var f = FormulaAnalysis.Clean(formula);
        var result = new List<(int, int)>();
        foreach (Match m in SelfRefRe().Matches(f))
        {
            // exclude sheet-qualified refs (char before the match is '!')
            if (m.Index > 0 && f[m.Index - 1] == '!') continue;
            result.Add((int.Parse(m.Groups[2].Value), FormulaAnalysis.ColToNum(m.Groups[1].Value)));
        }
        return result;
    }

    /// <summary>
    /// Longest dependency chain inside one sheet. Cross-sheet/constant refs are depth-0
    /// leaves; cycles are broken (iterative DFS with memo) to stay safe on iterative-calc files.
    /// </summary>
    public static int MaxChainDepth(IReadOnlyDictionary<(int Row, int Col), string> formulaCells)
    {
        var graph = new Dictionary<(int, int), List<(int, int)>>(formulaCells.Count);
        foreach (var ((r, c), f) in formulaCells)
        {
            var deps = new List<(int, int)>();
            foreach (var dep in SelfRefs(f))
                if (formulaCells.ContainsKey(dep)) deps.Add(dep);
            graph[(r, c)] = deps;
        }

        const int InProg = -1;
        var depth = new Dictionary<(int, int), int>(graph.Count);
        var best = 0;
        var stack = new Stack<((int, int) Node, bool Processed)>();

        foreach (var node in graph.Keys)
        {
            stack.Push((node, false));
            while (stack.Count > 0)
            {
                var (n, processed) = stack.Pop();
                if (processed)
                {
                    var d = 0;
                    foreach (var dep in graph[n])
                        if (depth.TryGetValue(dep, out var dd) && dd != InProg)
                            d = Math.Max(d, dd + 1);
                    depth[n] = d;
                    best = Math.Max(best, d);
                    continue;
                }
                if (depth.TryGetValue(n, out var cur) && cur != InProg) continue;
                depth[n] = InProg;
                stack.Push((n, true));
                foreach (var dep in graph[n])
                {
                    if (depth.TryGetValue(dep, out var ds) && ds == InProg) continue; // cycle
                    if (!depth.ContainsKey(dep)) stack.Push((dep, false));
                }
            }
        }
        return best;
    }

    /// <summary>cellXfs count from xl/styles.xml — the "64,000 ceiling" metric.</summary>
    public static int StyleCount(string path)
    {
        try
        {
            using var zip = ZipFile.OpenRead(path);
            var entry = zip.GetEntry("xl/styles.xml");
            if (entry is null) return 0;
            using var reader = new StreamReader(entry.Open());
            var xml = reader.ReadToEnd();
            var m = CellXfsRe().Match(xml);
            if (m.Success) return int.Parse(m.Groups[1].Value);
            return XfRe().Matches(xml).Count;
        }
        catch { return 0; }
    }

    public static int ExternalLinkCount(string path)
    {
        try
        {
            using var zip = ZipFile.OpenRead(path);
            return zip.Entries.Count(e =>
                e.FullName.StartsWith("xl/externalLinks/externalLink", StringComparison.Ordinal) &&
                e.FullName.EndsWith(".xml", StringComparison.Ordinal));
        }
        catch { return 0; }
    }
}
