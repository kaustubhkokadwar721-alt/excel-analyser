"""Recommendation engine — evidence-bound actions from measured output (spec §12).

Every Action is anchored to a measured row. Ranked by
    ROI = (gain_score * confidence) / effort
Nothing generic or unanchored is emitted.
"""

from __future__ import annotations

from .model import Action

_EFFORT = {"low": 1.0, "medium": 2.0, "high": 4.0}
_CONF = {"low": 0.4, "medium": 0.7, "high": 1.0}


def _roi(gain_ms, effort, conf):
    return round((gain_ms * _CONF[conf]) / _EFFORT[effort], 2)


def build(report):
    actions = []
    s = report.summary
    measured = [p for p in report.patterns if p.measured and p.total_ms > 0]
    total_measured = sum(p.total_ms for p in measured) or 1.0

    # 1) Volatile high-cost patterns → de-volatilise
    for p in sorted([p for p in measured if p.is_volatile], key=lambda x: -x.total_ms)[:8]:
        fns = set(p.funcs)
        if "OFFSET" in fns:
            fix = "Replace OFFSET with INDEX (non-volatile, same result)."
        elif "INDIRECT" in fns:
            fix = "Replace INDIRECT with CHOOSE / structured references / direct refs."
        elif fns & {"NOW", "TODAY"}:
            fix = "Compute the timestamp once in a single helper cell; reference it."
        elif fns & {"RAND", "RANDBETWEEN"}:
            fix = "Freeze random values (paste-as-values) once generated."
        else:
            fix = "Remove the volatile function; it recalculates on every change."
        gain = p.total_ms  # whole pattern recalcs needlessly
        actions.append(Action(
            anchor=f"{p.sheet}!{p.sample_cell} ({p.func_class})",
            measured_cost=f"{p.total_ms:.0f} ms total, {p.us_per_occ:.1f} µs×{p.occurrences}",
            why="Volatile: recalculates on every workbook change, dragging dependents.",
            fix=fix, est_gain=f"~{gain:.0f} ms recalc/edit",
            effort="medium", confidence="high",
            roi=_roi(gain, "medium", "high")))

    # 2) Full-column references → bound the range
    for p in sorted([p for p in measured if p.cells_touched >= 1_048_576],
                    key=lambda x: -x.total_ms)[:6]:
        gain = p.total_ms * 0.7
        actions.append(Action(
            anchor=f"{p.sheet}!{p.sample_cell} ({p.func_class})",
            measured_cost=f"{p.total_ms:.0f} ms total; whole-column scan",
            why="Full-column reference (A:A) scans far more cells than the data uses.",
            fix="Bound the range to the used data, or convert the source to a Table.",
            est_gain=f"~{gain:.0f} ms", effort="low", confidence="high",
            roi=_roi(gain, "low", "high")))

    # 3) Single-threaded blockers when MT efficiency is poor
    if s.mt_efficiency and s.mt_efficiency < 1.5:
        blockers = sorted([p for p in measured if p.is_single_threaded],
                          key=lambda x: -x.total_ms)[:5]
        for p in blockers:
            gain = p.total_ms * 0.5
            actions.append(Action(
                anchor=f"{p.sheet}!{p.sample_cell} ({p.func_class})",
                measured_cost=f"MT efficiency {s.mt_efficiency}× (cores idle)",
                why="Single-threaded function blocks Excel from using other cores.",
                fix="Replace INDIRECT/GETPIVOTDATA/CELL/INFO with multi-thread-safe equivalents.",
                est_gain=f"~{gain:.0f} ms via parallelism", effort="medium",
                confidence="medium", roi=_roi(gain, "medium", "medium")))

    # 4) Expensive lookups (measured costly + lookup class)
    lk = sorted([p for p in measured if p.func_class == "lookup"], key=lambda x: -x.total_ms)
    for p in lk[:5]:
        if p.total_ms < total_measured * 0.05:
            continue
        gain = p.total_ms * 0.6
        actions.append(Action(
            anchor=f"{p.sheet}!{p.sample_cell} (lookup)",
            measured_cost=f"{p.total_ms:.0f} ms total, {p.us_per_occ:.1f} µs each",
            why="Exact-match lookup over a large range scans row-by-row.",
            fix="Use XLOOKUP / sorted approximate match / a helper key; avoid full-column lookup arrays.",
            est_gain=f"~{gain:.0f} ms", effort="medium", confidence="medium",
            roi=_roi(gain, "medium", "medium")))

    # 5) Used-range bloat (structure)
    for st in sorted(report.structure, key=lambda x: -x.used_range_waste_pct):
        if st.used_range_waste_pct >= 95 and st.dimension_cells > 100_000:
            actions.append(Action(
                anchor=f"{st.sheet} (used range {st.dimension})",
                measured_cost=f"{st.used_range_waste_pct:.0f}% of used range is empty",
                why="Ghost cells extend the used range, inflating memory, size and calc.",
                fix="Delete the empty rows/columns below/right of real data, save, reopen.",
                est_gain="smaller file, faster open", effort="low", confidence="high",
                roi=_roi(40, "low", "high")))
            break

    # 6) Full-column conditional formatting
    for st in report.structure:
        if st.cf_full_column:
            actions.append(Action(
                anchor=f"{st.sheet} (CF, {st.cf_rule_count} rules)",
                measured_cost="full-column CF re-evaluated on scroll/select",
                why="Conditional formatting over a whole column re-evaluates constantly.",
                fix="Limit conditional-format ranges to the used data only.",
                est_gain="smoother scrolling/selection", effort="low", confidence="medium",
                roi=_roi(25, "low", "medium")))
            break

    # 7) styles.xml bloat
    if s.style_count and s.style_count > 5000:
        actions.append(Action(
            anchor=f"styles.xml ({s.style_count} cell formats)",
            measured_cost=f"{s.style_count} distinct cell formats",
            why="Large style table slows save and inflates file (64,000 ceiling).",
            fix="Clean excess cell formatting (Inquire) or rebuild styles; often 70–90% reducible.",
            est_gain="faster save, smaller file", effort="low", confidence="medium",
            roi=_roi(20, "low", "medium")))

    # 8) Deep dependency chains
    for st in sorted(report.structure, key=lambda x: -x.max_chain_depth)[:1]:
        if st.max_chain_depth >= 25:
            actions.append(Action(
                anchor=f"{st.sheet} (chain depth {st.max_chain_depth})",
                measured_cost=f"longest serial chain = {st.max_chain_depth}",
                why="Deep chains serialise calc even with multiple cores available.",
                fix="Break the chain / pre-compute / paste-values intermediate stages.",
                est_gain="better core utilisation", effort="medium", confidence="medium",
                roi=_roi(30, "medium", "medium")))

    actions.sort(key=lambda a: -a.roi)
    report.actions = actions
    return actions
