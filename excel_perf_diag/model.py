"""Data model: the structured results every layer fills in and the report reads."""

from __future__ import annotations

from dataclasses import dataclass, field, asdict
from typing import Optional


# Function-class taxonomy (the "encoding type" analogue from VertiPaq).
CLASS_CHEAP = "cheap"
CLASS_LOOKUP = "lookup"
CLASS_ARRAY = "array"
CLASS_VOLATILE = "volatile"
CLASS_DYNAMIC = "dynamic-array"
CLASS_LAMBDA = "lambda"
CLASS_LET = "let"
CLASS_OTHER = "other"


@dataclass
class PatternCost:
    """One row of the Formula Cost sheet: a unique (R1C1, cells-bucket) pattern."""
    sheet: str
    r1c1: str
    sample_a1: str                 # an example A1 formula, for readability
    sample_cell: str = ""          # address of one occurrence, e.g. "B2"
    func_class: str = ""
    occurrences: int = 0
    cells_touched: int = 1         # referenced-range size per occurrence (estimate)
    is_volatile: bool = False
    is_single_threaded: bool = False
    funcs: tuple = ()              # function names seen in the pattern
    # --- filled by Layer 2 (dynamic timing) ---
    us_per_occ: float = 0.0        # measured microseconds per occurrence
    stdev_us: float = 0.0          # run-to-run variability (trust signal)
    total_ms: float = 0.0          # us_per_occ * occurrences / 1000
    nonlinear: bool = False        # two-point sampling found non-linear cost
    measured: bool = False
    # --- ranking helpers (filled by report) ---
    pct_of_sheet: float = 0.0
    cum_pct: float = 0.0
    flag_reason: str = ""

    def to_row(self):
        return asdict(self)


@dataclass
class SheetTiming:
    sheet: str
    full_calc_ms: float = 0.0
    overhead_ms: float = 0.0
    measured: bool = False


@dataclass
class StructureFinding:
    sheet: str
    dimension: str = ""
    populated_cells: int = 0
    dimension_cells: int = 0
    used_range_waste_pct: float = 0.0
    cf_rule_count: int = 0
    cf_full_column: bool = False
    max_chain_depth: int = 0
    note: str = ""


@dataclass
class PowerQueryTiming:
    name: str
    refresh_s: float = 0.0
    ok: bool = True
    note: str = ""


@dataclass
class Action:
    anchor: str                    # evidence row reference, e.g. "Volatile!INDIRECT pattern"
    measured_cost: str             # human-readable measured value
    why: str
    fix: str
    est_gain: str
    effort: str                    # low | medium | high
    confidence: str                # low | medium | high
    roi: float = 0.0               # (gain_score * conf) / effort, for sorting


@dataclass
class Summary:
    file: str = ""
    excel_version: str = ""
    machine: str = ""
    timestamp: str = ""
    full_calc_ms: float = 0.0
    recalc_ms: float = 0.0
    volatility_pct: float = 0.0
    mt_efficiency: float = 0.0
    single_thread_ms: float = 0.0
    multi_thread_ms: float = 0.0
    max_dependency_depth: int = 0
    file_size_mb: float = 0.0
    used_range_waste_pct: float = 0.0
    open_ms: float = 0.0
    save_ms: float = 0.0
    style_count: int = 0
    external_link_count: int = 0
    defined_name_count: int = 0
    pq_total_refresh_s: float = 0.0
    partial: bool = False          # safety budget tripped → incomplete
    notes: list = field(default_factory=list)


@dataclass
class Report:
    summary: Summary = field(default_factory=Summary)
    patterns: list = field(default_factory=list)        # list[PatternCost]
    sheet_timings: list = field(default_factory=list)   # list[SheetTiming]
    structure: list = field(default_factory=list)       # list[StructureFinding]
    power_queries: list = field(default_factory=list)   # list[PowerQueryTiming]
    actions: list = field(default_factory=list)         # list[Action]
    deltas: dict = field(default_factory=dict)          # regression deltas
