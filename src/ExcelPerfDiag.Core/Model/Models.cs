namespace ExcelPerfDiag.Core.Model;

/// <summary>
/// One row of the Formula Cost sheet: a unique (R1C1, cells-touched bucket) pattern.
/// Static analysis fills the identity/classification; dynamic timing fills the cost.
/// </summary>
public sealed class PatternCost
{
    public required string Sheet { get; init; }
    public required string R1C1 { get; init; }
    public string SampleA1 { get; init; } = "";
    public string SampleCell { get; init; } = "";   // e.g. "B2"
    public FunctionClass FuncClass { get; init; }
    public int Occurrences { get; set; }
    public long CellsTouched { get; init; } = 1;     // referenced-range size per occurrence
    public bool IsVolatile { get; init; }
    public bool IsSingleThreaded { get; init; }
    public IReadOnlyList<string> Funcs { get; init; } = System.Array.Empty<string>();

    // --- filled by Layer 2 (dynamic timing) ---
    public double UsPerOcc { get; set; }
    public double StdevUs { get; set; }
    public double TotalMs { get; set; }
    public bool Nonlinear { get; set; }
    public bool Measured { get; set; }

    // --- ranking helpers (filled by report) ---
    public double PctOfSheet { get; set; }
    public double CumPct { get; set; }
    public string FlagReason { get; set; } = "";
}

public sealed class SheetTiming
{
    public required string Sheet { get; init; }
    public double FullCalcMs { get; set; }
    public double OverheadMs { get; set; }
    public bool Measured { get; set; }
}

public sealed class StructureFinding
{
    public required string Sheet { get; init; }
    public string Dimension { get; set; } = "";
    public long PopulatedCells { get; set; }
    public long DimensionCells { get; set; }
    public double UsedRangeWastePct { get; set; }
    public int CfRuleCount { get; set; }
    public bool CfFullColumn { get; set; }
    public int MaxChainDepth { get; set; }
    public string Note { get; set; } = "";
}

public sealed class PowerQueryTiming
{
    public required string Name { get; init; }
    public double RefreshSeconds { get; set; }
    public bool Ok { get; set; } = true;
    public string Note { get; set; } = "";
}

/// <summary>One evidence-bound, ROI-ranked recommendation (Actions sheet).</summary>
public sealed class ActionItem
{
    public required string Anchor { get; init; }       // evidence row reference
    public required string MeasuredCost { get; init; }
    public required string Why { get; init; }
    public required string Fix { get; init; }
    public string EstGain { get; init; } = "";
    public string Effort { get; init; } = "medium";    // low | medium | high
    public string Confidence { get; init; } = "medium";
    public double Roi { get; set; }
}

public sealed class Summary
{
    public string File { get; set; } = "";
    public string ExcelVersion { get; set; } = "";
    public string Machine { get; set; } = "";
    public string Timestamp { get; set; } = "";
    public double FullCalcMs { get; set; }
    public double RecalcMs { get; set; }
    public double VolatilityPct { get; set; }
    public double MtEfficiency { get; set; }
    public double SingleThreadMs { get; set; }
    public double MultiThreadMs { get; set; }
    public int MaxDependencyDepth { get; set; }
    public double FileSizeMb { get; set; }
    public double UsedRangeWastePct { get; set; }
    public double OpenMs { get; set; }
    public double SaveMs { get; set; }
    public int StyleCount { get; set; }
    public int ExternalLinkCount { get; set; }
    public int DefinedNameCount { get; set; }
    public double PqTotalRefreshSeconds { get; set; }
    public bool Partial { get; set; }
    public List<string> Notes { get; } = new();
}

/// <summary>The complete diagnostic — what every face (CLI, app, exporters) consumes.</summary>
public sealed class DiagnosticReport
{
    public Summary Summary { get; init; } = new();
    public List<PatternCost> Patterns { get; init; } = new();
    public List<SheetTiming> SheetTimings { get; init; } = new();
    public List<StructureFinding> Structure { get; init; } = new();
    public List<PowerQueryTiming> PowerQueries { get; init; } = new();
    public List<ActionItem> Actions { get; init; } = new();
    public RegressionDelta? Deltas { get; set; }
}

/// <summary>Before/after deltas vs the previous run (spec §13).</summary>
public sealed class RegressionDelta
{
    public string? PreviousTimestamp { get; set; }
    public Dictionary<string, MetricDelta> SummaryDeltas { get; } = new();
    public Dictionary<string, MetricDelta> PatternDeltas { get; } = new();
}

public readonly record struct MetricDelta(double Previous, double Now, double Delta, double? Pct);
