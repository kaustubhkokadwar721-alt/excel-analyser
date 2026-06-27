"""Before/after regression (spec §13): fingerprint each run, diff vs the last.

Stores a JSON sidecar next to the analyzed file. On the next run, computes
per-pattern and summary deltas so a fix can be *proven*, not eyeballed.
"""

from __future__ import annotations

import json
import os
from datetime import datetime


def _path(src_path):
    return os.path.splitext(src_path)[0] + ".perffingerprint.json"


def _fingerprint(report):
    return {
        "timestamp": report.summary.timestamp,
        "summary": {
            "full_calc_ms": report.summary.full_calc_ms,
            "recalc_ms": report.summary.recalc_ms,
            "volatility_pct": report.summary.volatility_pct,
            "mt_efficiency": report.summary.mt_efficiency,
            "open_ms": report.summary.open_ms,
            "save_ms": report.summary.save_ms,
            "pq_total_refresh_s": report.summary.pq_total_refresh_s,
        },
        "patterns": {
            f"{p.sheet}|{p.r1c1}": round(p.total_ms, 2)
            for p in report.patterns if p.measured
        },
    }


def load_previous(src_path):
    p = _path(src_path)
    if os.path.exists(p):
        try:
            with open(p, encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return None
    return None


def compute_deltas(report, previous):
    """Fill report.deltas with summary + per-pattern changes vs previous run."""
    if not previous:
        return
    deltas = {"summary": {}, "patterns": {}, "prev_timestamp": previous.get("timestamp")}
    cur = _fingerprint(report)
    for k, v in cur["summary"].items():
        pv = previous.get("summary", {}).get(k)
        if pv is not None:
            deltas["summary"][k] = {"prev": pv, "now": v, "delta": round(v - pv, 2),
                                    "pct": (round(100 * (v - pv) / pv, 1) if pv else None)}
    prev_pat = previous.get("patterns", {})
    for key, now in cur["patterns"].items():
        if key in prev_pat:
            pv = prev_pat[key]
            deltas["patterns"][key] = {"prev": pv, "now": now, "delta": round(now - pv, 2),
                                       "pct": (round(100 * (now - pv) / pv, 1) if pv else None)}
    report.deltas = deltas


def save(report, src_path):
    if not report.summary.timestamp:
        report.summary.timestamp = datetime.now().isoformat(timespec="seconds")
    try:
        with open(_path(src_path), "w", encoding="utf-8") as f:
            json.dump(_fingerprint(report), f, indent=2)
    except Exception:
        pass
