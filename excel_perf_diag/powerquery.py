"""Power Query / connection refresh timing — a SEPARATE animal from recalc.

Timed per query with BackgroundQuery:=False so each completes before the next,
never lumped into calc time (spec §2.4). Called inside the Layer 2 COM session.
"""

from __future__ import annotations

import time
from .model import PowerQueryTiming


def time_queries(app, wb, summary, results, budget_check=None):
    """Append PowerQueryTiming rows to `results`; set summary.pq_total_refresh_s."""
    total = 0.0
    # Names from the modern Queries collection (Excel 2016+), if present.
    try:
        nq = wb.Queries.Count
    except Exception:
        nq = 0

    try:
        conns = wb.Connections
        ccount = conns.Count
    except Exception:
        ccount = 0

    if ccount == 0:
        if nq:
            for i in range(1, nq + 1):
                try:
                    results.append(PowerQueryTiming(
                        name=str(wb.Queries(i).Name), refresh_s=0.0, ok=False,
                        note="query present but no refreshable connection"))
                except Exception:
                    pass
        return

    for i in range(1, ccount + 1):
        if budget_check and budget_check():
            summary.notes.append("PQ timing skipped (budget).")
            break
        try:
            conn = wb.Connections(i)
            name = str(conn.Name)
        except Exception:
            continue
        # Force foreground refresh where the connection type supports it.
        for attr in ("OLEDBConnection", "ODBCConnection"):
            try:
                getattr(conn, attr).BackgroundQuery = False
            except Exception:
                pass
        try:
            t = time.perf_counter()
            conn.Refresh()
            dt = time.perf_counter() - t
            total += dt
            results.append(PowerQueryTiming(name=name, refresh_s=round(dt, 3), ok=True))
        except Exception as e:
            results.append(PowerQueryTiming(name=name, refresh_s=0.0, ok=False,
                                            note=str(e)[:80]))

    summary.pq_total_refresh_s = round(total, 3)
