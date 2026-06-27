"""Orchestrator / CLI.

  python -m excel_perf_diag <file.xlsx> [options]

Options:
  --static-only       Layer 1 only (no Excel); fast triage.
  --budget N          Layer 2 time budget in seconds (default 150).
  --copies N          Max scratch copies per pattern (default 2000).
  --iters N           Calc iterations per measurement (default 6).
  --topk N            Deep-time only the N worst sheets (default 6).
  --out PATH          Report path (default <file>.perfreport.xlsx).
"""

from __future__ import annotations

import argparse
import os
import socket
import sys
from datetime import datetime

from .model import Report, Summary
from . import static_layer as SL
from . import structure as S
from . import recommend, regression, report as RW


def _summarize_static(rep: Report, path, structure, defined_names):
    s = rep.summary
    s.file = os.path.abspath(path)
    s.machine = socket.gethostname()
    s.file_size_mb = SL.file_size_mb(path)
    s.style_count = S.style_count(path)
    s.external_link_count = S.external_link_count(path)
    s.defined_name_count = len(defined_names)
    s.used_range_waste_pct = max((st.used_range_waste_pct for st in structure), default=0.0)
    s.max_dependency_depth = max((st.max_chain_depth for st in structure), default=0)


def run(path, static_only=False, budget=150, copies=2000, iters=6, topk=6, out=None):
    if not os.path.exists(path):
        print(f"ERROR: file not found: {path}")
        return 2
    out = out or (os.path.splitext(path)[0] + ".perfreport.xlsx")

    print(f"[1/4] Layer 1 — static analysis of {os.path.basename(path)} ...")
    patterns, structure, suspicion, defined_names = SL.analyze(path)
    print(f"      {len(patterns)} patterns across {len(structure)} sheets.")

    rep = Report()
    rep.summary = Summary()
    rep.patterns = patterns
    rep.structure = structure
    _summarize_static(rep, path, structure, defined_names)
    rep.summary.timestamp = datetime.now().isoformat(timespec="seconds")

    if not static_only:
        print(f"[2/4] Layer 2 - dynamic timing (budget {budget}s, copies<={copies}, top{topk} sheets) ...")
        from .dynamic_layer import DynamicTimer
        pq = []
        rep.power_queries = pq
        timer = DynamicTimer(budget_s=budget, iterations=iters,
                             max_copies=copies, top_k=topk)
        try:
            timer.run(path, patterns, suspicion, rep.summary, pq_results=pq)
        except Exception as e:
            rep.summary.notes.append(f"Layer 2 failed: {e}")
            print(f"      Layer 2 error: {e}")
    else:
        print("[2/4] Layer 2 — skipped (--static-only).")

    print("[3/4] Regression diff + recommendations ...")
    prev = regression.load_previous(path)
    regression.compute_deltas(rep, prev)
    recommend.build(rep)

    print("[4/4] Writing report ...")
    RW.write(rep, out)
    regression.save(rep, path)

    measured = sum(1 for p in patterns if p.measured)
    print(f"\nDone. {measured} patterns timed. Report: {out}")
    if rep.summary.partial:
        print("NOTE: partial report — time budget was hit.")
    return 0


def main(argv=None):
    ap = argparse.ArgumentParser(prog="excel_perf_diag")
    ap.add_argument("file")
    ap.add_argument("--static-only", action="store_true")
    ap.add_argument("--budget", type=int, default=150)
    ap.add_argument("--copies", type=int, default=2000)
    ap.add_argument("--iters", type=int, default=6)
    ap.add_argument("--topk", type=int, default=6)
    ap.add_argument("--out", default=None)
    a = ap.parse_args(argv)
    return run(a.file, a.static_only, a.budget, a.copies, a.iters, a.topk, a.out)


if __name__ == "__main__":
    sys.exit(main())
