"""Layer 1 structural analysis: used-range waste, CF, styles, links, names, chain depth.

All static — no Excel. Reads what openpyxl exposes plus a couple of raw
zip/xml peeks (styles.xml, externalLinks) that openpyxl summarises poorly.
"""

from __future__ import annotations

import re
import zipfile
from .model import StructureFinding
from . import patterns as P


def style_count(path: str) -> int:
    """cellXfs count from xl/styles.xml — the '64,000 ceiling' metric."""
    try:
        with zipfile.ZipFile(path) as z:
            xml = z.read("xl/styles.xml").decode("utf-8", "ignore")
        m = re.search(r"<cellXfs count=\"(\d+)\"", xml)
        if m:
            return int(m.group(1))
        return xml.count("<xf ")
    except Exception:
        return 0


def external_link_count(path: str) -> int:
    try:
        with zipfile.ZipFile(path) as z:
            return sum(
                1 for n in z.namelist()
                if n.startswith("xl/externalLinks/externalLink") and n.endswith(".xml")
            )
    except Exception:
        return 0


def _is_full_column_range(sqref: str) -> bool:
    # e.g. 'A1:A1048576' or '$A:$A'
    return "1048576" in sqref or bool(re.search(r"[A-Za-z]+:[A-Za-z]+(\s|$)", sqref))


def _self_refs(formula: str, row: int, col: int):
    """Same-sheet A1 cell refs in a formula (cross-sheet refs ignored as leaves)."""
    f = P.clean(formula)
    out = []
    for m in re.finditer(r"(?<![A-Za-z0-9_$!'])\$?([A-Za-z]{1,3})\$?(\d+)(?![A-Za-z0-9_(])", f):
        # crude exclusion of sheet-qualified refs (preceded by '!')
        start = m.start()
        if start > 0 and f[start - 1] == "!":
            continue
        out.append((int(m.group(2)), P._col_to_num(m.group(1))))
    return out


def max_chain_depth(formula_cells: dict) -> int:
    """Longest dependency chain inside one sheet.

    formula_cells: {(row,col): formula_str}. Iterative longest-path with memo;
    cross-sheet/external refs and refs to constants count as depth-0 leaves.
    Cycles are broken (return current best) to stay safe on iterative-calc files.
    """
    graph = {}
    for (r, c), f in formula_cells.items():
        deps = [(rr, cc) for (rr, cc) in _self_refs(f, r, c) if (rr, cc) in formula_cells]
        graph[(r, c)] = deps
    depth = {}
    INPROG = -1
    best = 0
    for node in graph:
        # iterative DFS
        stack = [(node, False)]
        while stack:
            n, processed = stack.pop()
            if processed:
                d = 0
                for dep in graph.get(n, ()):  # dep already resolved
                    d = max(d, depth.get(dep, 0) + 1)
                depth[n] = d
                best = max(best, d)
                continue
            if n in depth and depth[n] != INPROG:
                continue
            depth[n] = INPROG
            stack.append((n, True))
            for dep in graph.get(n, ()):
                if depth.get(dep, 0) == INPROG:
                    continue  # cycle: skip
                if dep not in depth:
                    stack.append((dep, False))
    return best


def analyze_sheet_structure(ws, formula_cells: dict) -> StructureFinding:
    populated = len(ws._cells)
    max_row = ws.max_row or 1
    max_col = ws.max_column or 1
    dim_cells = max_row * max_col
    waste = (1 - populated / dim_cells) if dim_cells else 0.0

    cf_rules = 0
    cf_full = False
    try:
        for rng in ws.conditional_formatting:
            sqref = str(rng.sqref)
            rules = ws.conditional_formatting[rng]
            cf_rules += len(rules)
            if _is_full_column_range(sqref):
                cf_full = True
    except Exception:
        pass

    return StructureFinding(
        sheet=ws.title,
        dimension=ws.dimensions,
        populated_cells=populated,
        dimension_cells=dim_cells,
        used_range_waste_pct=round(waste * 100, 1),
        cf_rule_count=cf_rules,
        cf_full_column=cf_full,
        max_chain_depth=max_chain_depth(formula_cells),
    )
