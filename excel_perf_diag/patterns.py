"""Formula parsing: R1C1 normalisation, function classification, grouping key.

The grouping unit (the VertiPaq "column" analogue) is
    (R1C1 pattern, cells-touched bucket)
not R1C1 alone — copied formulas pointing at different-sized ranges have the
same relative shape but different cost (spec §9.1).
"""

from __future__ import annotations

import re
from .model import (
    CLASS_CHEAP, CLASS_LOOKUP, CLASS_ARRAY, CLASS_VOLATILE,
    CLASS_DYNAMIC, CLASS_LAMBDA, CLASS_LET, CLASS_OTHER,
)

# Excel stores modern functions with these internal prefixes; strip for analysis.
_PREFIX_RE = re.compile(r"_xlfn\.(_xlws\.)?|_xlpm\.", re.IGNORECASE)

VOLATILE = {"OFFSET", "INDIRECT", "NOW", "TODAY", "RAND", "RANDBETWEEN",
            "RANDARRAY", "CELL", "INFO"}
# Single-threaded / multi-thread-blocking constructs.
SINGLE_THREADED = {"INDIRECT", "GETPIVOTDATA", "CELL", "INFO"}
LOOKUP = {"VLOOKUP", "HLOOKUP", "LOOKUP", "MATCH", "INDEX", "XLOOKUP", "XMATCH"}
ARRAY_LIKE = {"SUMPRODUCT", "MMULT", "TRANSPOSE", "DSUM", "DCOUNT", "DGET", "DAVERAGE"}
DYNAMIC = {"FILTER", "SORT", "SORTBY", "UNIQUE", "SEQUENCE", "RANDARRAY"}
LAMBDA_FN = {"LAMBDA", "MAP", "REDUCE", "SCAN", "BYROW", "BYCOL", "MAKEARRAY"}
CHEAP = {"SUM", "SUMIF", "SUMIFS", "COUNT", "COUNTIF", "COUNTIFS", "AVERAGE",
         "AVERAGEIF", "AVERAGEIFS", "MAX", "MIN", "MAXIFS", "MINIFS", "IF",
         "IFERROR", "ROUND", "ABS", "AND", "OR", "CONCAT", "TEXT"}

_FUNC_RE = re.compile(r"([A-Za-z_][A-Za-z0-9_.]*)\s*\(")
# Full-column ref (A:A or $A:$H), including sheet-qualified (Data!$A:$H).
_FULLCOL_RE = re.compile(r"(?<![A-Za-z0-9_$])\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}(?![A-Za-z0-9_])")
_A1_CELL_RE = re.compile(r"\$?([A-Za-z]{1,3})\$?(\d+)")
# A1 reference, optionally a range, optionally sheet-qualified.
_REF_RE = re.compile(
    r"(?:'?[A-Za-z0-9_ ]+'?!)?\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?"
)


def clean(formula: str) -> str:
    """Strip leading '=' and Excel internal prefixes."""
    f = formula or ""
    if f.startswith("="):
        f = f[1:]
    return _PREFIX_RE.sub("", f)


def functions(formula: str) -> list[str]:
    """All function names used, uppercased, prefixes removed."""
    return [m.group(1).upper() for m in _FUNC_RE.finditer(clean(formula))]


def classify(formula: str) -> str:
    """Single dominant function class (worst-wins ordering)."""
    fns = set(functions(formula))
    if fns & VOLATILE:
        return CLASS_VOLATILE
    if fns & ARRAY_LIKE:
        return CLASS_ARRAY
    if fns & LAMBDA_FN:
        return CLASS_LAMBDA
    if fns & DYNAMIC:
        return CLASS_DYNAMIC
    if fns & LOOKUP:
        return CLASS_LOOKUP
    if "LET" in fns and not (fns - {"LET"}) & (VOLATILE | ARRAY_LIKE | LOOKUP):
        return CLASS_LET
    if fns & CHEAP:
        return CLASS_CHEAP
    return CLASS_OTHER if fns else CLASS_CHEAP


def is_volatile(formula: str) -> bool:
    return bool(set(functions(formula)) & VOLATILE)


def is_single_threaded(formula: str) -> bool:
    return bool(set(functions(formula)) & SINGLE_THREADED)


def has_full_column(formula: str) -> bool:
    return bool(_FULLCOL_RE.search(clean(formula)))


def _col_to_num(col: str) -> int:
    n = 0
    for ch in col.upper():
        n = n * 26 + (ord(ch) - 64)
    return n


def cells_touched(formula: str) -> int:
    """Estimate referenced-range size summed across the formula's ranges.

    Full-column refs are charged a large fixed cost; bounded ranges by area.
    Coarse but monotonic — good enough for ranking and bucketing.
    """
    f = clean(formula)
    total = 0
    if has_full_column(f):
        total += 1_048_576 * (1 + f.count(":") )  # whole-column penalty
    for m in re.finditer(
        r"\$?([A-Za-z]{1,3})\$?(\d+):\$?([A-Za-z]{1,3})\$?(\d+)", f
    ):
        c1, r1, c2, r2 = m.group(1), int(m.group(2)), m.group(3), int(m.group(4))
        w = abs(_col_to_num(c2) - _col_to_num(c1)) + 1
        h = abs(r2 - r1) + 1
        total += w * h
    return max(total, 1)


def _cells_bucket(n: int) -> int:
    """Bucket cells-touched to log-ish bands so similar-cost cells group."""
    if n <= 1:
        return 0
    if n >= 1_048_576:
        return 99
    import math
    return int(math.log10(n) * 2)  # 2 buckets per order of magnitude


def normalize_r1c1(formula: str, row: int, col: int) -> str:
    """Convert an A1 formula at (row, col) to an R1C1-style normalised form.

    Copied/relative formulas collapse to one pattern; absolute refs stay fixed.
    A lightweight converter (handles the common cases that matter for grouping).
    """
    f = clean(formula)

    def repl(m):
        token = m.group(0)
        # skip sheet-qualified refs entirely (treat as fixed external anchor)
        if "!" in token:
            return token
        cm = _A1_CELL_RE.fullmatch(token)
        if not cm:
            return token
        col_part, row_part = m.group(0), None
        # rebuild from the regex groups
        full = m.group(0)
        abs_col = full.lstrip().startswith("$") or ("$" + cm.group(1)) in full
        # Determine absolute markers precisely:
        col_abs = bool(re.match(r"\$[A-Za-z]", full))
        row_abs = bool(re.search(r"[A-Za-z]\$\d", full))
        tcol = _col_to_num(cm.group(1))
        trow = int(cm.group(2))
        c = f"C{tcol}" if col_abs else f"C[{tcol - col}]"
        r = f"R{trow}" if row_abs else f"R[{trow - row}]"
        return r + c

    # Only convert standalone cell tokens (not ranges handled token-by-token).
    out = re.sub(r"\$?[A-Za-z]{1,3}\$?\d+", repl, f)
    return out


def grouping_key(formula: str, row: int, col: int) -> tuple[str, int]:
    """The cost-attribution unit: (R1C1 pattern, cells-touched bucket)."""
    return (normalize_r1c1(formula, row, col), _cells_bucket(cells_touched(formula)))
