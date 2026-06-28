using System.Text.Json;
using System.Text.Json.Serialization;
using ExcelPerfDiag.Core.Analysis;
using ExcelPerfDiag.Core.Model;
using ExcelPerfDiag.Core.Timing;
using ExcelPerfDiag.Excel;

// Crash-isolated worker: drives Excel via COM, streams progress on stderr, emits the
// final DiagnosticReport as a single JSON line on stdout. If Excel hard-crashes, this
// process exits non-zero and the parent (CLI/UI) reports gracefully — the Job Object
// guarantees the spawned Excel never orphans.

string? file = null;
double budget = 150;
int copies = 2000, iters = 6, topk = 6;
for (var i = 0; i < args.Length; i++)
{
    switch (args[i])
    {
        case "--budget" when i + 1 < args.Length: budget = double.Parse(args[++i]); break;
        case "--copies" when i + 1 < args.Length: copies = int.Parse(args[++i]); break;
        case "--iters" when i + 1 < args.Length: iters = int.Parse(args[++i]); break;
        case "--topk" when i + 1 < args.Length: topk = int.Parse(args[++i]); break;
        default: if (!args[i].StartsWith('-')) file = args[i]; break;
    }
}
if (file is null || !File.Exists(file))
{
    Console.Error.WriteLine("worker: usage: <file.xlsx> [--budget s] [--copies n] [--iters n] [--topk n]");
    return 2;
}

void Progress(string msg, double t) => Console.Error.WriteLine($"[{t,5:F1}s] {msg}");

Progress("static analysis", 0);
var sr = StaticAnalyzer.Analyze(file);

var report = new DiagnosticReport();
report.Patterns.AddRange(sr.Patterns);
report.Structure.AddRange(sr.Structure);
var s = report.Summary;
s.File = Path.GetFullPath(file);
s.Machine = Environment.MachineName;
s.Timestamp = DateTime.Now.ToString("s");
s.FileSizeMb = StaticAnalyzer.FileSizeMb(file);
s.StyleCount = sr.StyleCount;
s.ExternalLinkCount = sr.ExternalLinkCount;
s.DefinedNameCount = sr.DefinedNames.Count;
s.UsedRangeWastePct = sr.Structure.Count > 0 ? sr.Structure.Max(x => x.UsedRangeWastePct) : 0;
s.MaxDependencyDepth = sr.Structure.Count > 0 ? sr.Structure.Max(x => x.MaxChainDepth) : 0;

try
{
    using var session = new ExcelComSession();
    Progress("opening excel (temp copy)", 0);
    session.Open(file);
    TimingHarness.Run(session, report, sr.Suspicion, new TimingOptions(budget, iters, copies, topk), Progress);
    Progress("dynamic timing complete", 0);
}
catch (Exception e)
{
    s.Notes.Add("dynamic timing failed: " + e.Message);
    Progress("dynamic error: " + e.Message, 0);
}

var json = JsonSerializer.Serialize(report, new JsonSerializerOptions
{
    Converters = { new JsonStringEnumConverter() },
});
Console.Out.WriteLine(json);
return 0;
