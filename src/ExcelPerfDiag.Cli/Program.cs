using System.Text.Json;
using ExcelPerfDiag.Core.Analysis;
using ExcelPerfDiag.Core.Model;

string? file = null, jsonOut = null;
var staticOnly = false;
var top = 15;
for (var i = 0; i < args.Length; i++)
{
    var a = args[i];
    if (a is "--static-only") staticOnly = true;
    else if (a is "--json") jsonOut = ++i < args.Length ? args[i] : null;
    else if (a is "--top") top = ++i < args.Length && int.TryParse(args[i], out var t) ? t : top;
    else if (!a.StartsWith('-')) file = a;
}

if (file is null)
{
    Console.WriteLine("usage: ExcelPerfDiag.Cli <file.xlsx> [--static-only] [--json <out.json>] [--top N]");
    return 2;
}
if (!File.Exists(file))
{
    Console.Error.WriteLine($"ERROR: file not found: {file}");
    return 2;
}

Console.WriteLine($"[1/2] Layer 1 - static analysis of {Path.GetFileName(file)} ...");
var r = StaticAnalyzer.Analyze(file);
Console.WriteLine($"      {r.Patterns.Count} patterns across {r.Structure.Count} sheets");

if (!staticOnly)
    Console.WriteLine("[2/2] Layer 2 - dynamic timing not yet wired (P2). Showing static triage only.");

static double Weight(PatternCost p)
{
    var w = p.FuncClass switch
    {
        FunctionClass.Volatile => 8.0, FunctionClass.Array => 6, FunctionClass.Lookup => 4,
        FunctionClass.DynamicArray => 3, FunctionClass.Lambda => 3, _ => 1,
    };
    return p.Occurrences * w * (p.CellsTouched >= 1_048_576 ? 2 : 1);
}

Console.WriteLine("\n-- worst sheets (static suspicion) --");
foreach (var (s, v) in r.Suspicion.OrderByDescending(kv => kv.Value).Take(6))
    Console.WriteLine($"   {s,-18} {v,12:N0}");

Console.WriteLine("\n-- top suspect patterns --");
Console.WriteLine($"   {"sheet",-16} {"class",-13} {"occ",6} {"cells",10}  flags  formula");
foreach (var p in r.Patterns.OrderByDescending(Weight).Take(top))
{
    var flags = (p.IsVolatile ? "V" : " ") + (p.IsSingleThreaded ? "S" : " ") + (p.CellsTouched >= 1_048_576 ? "F" : " ");
    var f = p.SampleA1.Length > 46 ? p.SampleA1[..46] : p.SampleA1;
    Console.WriteLine($"   {p.Sheet,-16} {p.FuncClass.ToWire(),-13} {p.Occurrences,6} {p.CellsTouched,10}  {flags}   {f}");
}

Console.WriteLine("\n-- structure --");
Console.WriteLine($"   {"sheet",-18} {"waste%",7} {"CF",4} {"chain",6}");
foreach (var s in r.Structure.OrderByDescending(s => s.UsedRangeWastePct))
    Console.WriteLine($"   {s.Sheet,-18} {s.UsedRangeWastePct,7:N1} {s.CfRuleCount,4} {s.MaxChainDepth,6}");
Console.WriteLine($"   styles={r.StyleCount}  external-links={r.ExternalLinkCount}  defined-names={r.DefinedNames.Count}  size={StaticAnalyzer.FileSizeMb(file):N2}MB");

if (jsonOut is not null)
{
    var projection = new
    {
        file = Path.GetFullPath(file),
        patterns = r.Patterns.Select(p => new
        {
            sheet = p.Sheet, r1c1 = p.R1C1, cls = p.FuncClass.ToWire(),
            occ = p.Occurrences, cells = p.CellsTouched,
            vol = p.IsVolatile, st = p.IsSingleThreaded,
        }),
        structure = r.Structure.Select(s => new
        {
            sheet = s.Sheet, wastePct = s.UsedRangeWastePct,
            cf = s.CfRuleCount, chain = s.MaxChainDepth,
        }),
        styleCount = r.StyleCount, externalLinks = r.ExternalLinkCount,
        definedNames = r.DefinedNames.Count,
    };
    File.WriteAllText(jsonOut, JsonSerializer.Serialize(projection, new JsonSerializerOptions { WriteIndented = true }));
    Console.WriteLine($"\nstatic JSON written: {jsonOut}");
}

return 0;
