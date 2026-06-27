"""Layer 2 — dynamic timing harness via win32com (the only source of real cost).

Implements the FastExcel method: group identical formulas (done in Layer 1),
batch-calculate many copies single-threaded, subtract empty-sheet overhead,
average several runs. Plus workbook full/recalc, single-vs-multi-thread ratio,
and direct open/save timing.

Safety envelope (spec §7) is non-negotiable and enforced here:
  - operate on a TEMP COPY, never the live file
  - dedicated, hidden Excel instance
  - state saved/restored in a finally block on every exit path
  - global time budget → partial report
  - sample only the worst K sheets; cap copies per pattern
  - Esc/kill safe; idempotent scratch sheet
"""

from __future__ import annotations

import os
import shutil
import statistics
import tempfile
import time
from contextlib import contextmanager

import pythoncom
import win32com.client as win32

XL_MANUAL = -4135
XL_AUTOMATIC = -4105
SCRATCH = "_perf_scratch"


def _mean(xs):
    xs = [x for x in xs if x is not None]
    return sum(xs) / len(xs) if xs else 0.0


class DynamicTimer:
    def __init__(self, budget_s=150, iterations=6, max_copies=2000,
                 top_k=6, log=print):
        self.budget_s = budget_s
        self.iterations = iterations          # incl. one discarded warm-up
        self.max_copies = max_copies
        self.top_k = top_k
        self.log = log
        self._t0 = 0.0

    def _over_budget(self):
        return (time.perf_counter() - self._t0) > self.budget_s

    # ---------------------------------------------------------------- timing
    def _time_calc(self, com_obj, iters=None):
        """Time .Calculate() on a Range/Worksheet; drop warm-up, return samples."""
        iters = iters or self.iterations
        samples = []
        for _ in range(iters):
            t = time.perf_counter()
            com_obj.Calculate()
            samples.append(time.perf_counter() - t)
        return samples[1:] if len(samples) > 1 else samples

    def _time_full(self, app, iters=None):
        iters = iters or self.iterations
        samples = []
        for _ in range(iters):
            t = time.perf_counter()
            app.CalculateFull()
            samples.append(time.perf_counter() - t)
        return samples[1:] if len(samples) > 1 else samples

    # ---------------------------------------------------------------- scratch
    def _fill_scratch(self, scratch, n, formula_r1c1):
        rng = scratch.Range(scratch.Cells(1, 1), scratch.Cells(n, 1))
        rng.ClearContents()
        rng.FormulaR1C1 = formula_r1c1
        return rng

    def _measure_pattern(self, scratch, src_formula_r1c1, n):
        """Batch-and-subtract → (us_per_occ, stdev_us)."""
        rng = self._fill_scratch(scratch, n, src_formula_r1c1)
        patt = self._time_calc(rng)
        over_rng = self._fill_scratch(scratch, n, "1")
        over = self._time_calc(over_rng)
        per_cell = [(max(p - _mean(over), 0.0)) / n * 1e6 for p in patt]
        rng.ClearContents()
        return _mean(per_cell), (statistics.pstdev(per_cell) if len(per_cell) > 1 else 0.0)

    # ---------------------------------------------------------------- run
    def run(self, src_path, patterns, suspicion, summary, pq_results=None):
        """Mutate `patterns` in place with measured cost; fill `summary`. Returns nothing."""
        from . import powerquery
        self._t0 = time.perf_counter()
        tmpdir = tempfile.mkdtemp(prefix="perfdiag_")
        copy = os.path.join(tmpdir, os.path.basename(src_path))
        shutil.copy2(src_path, copy)

        pythoncom.CoInitialize()
        app = None
        wb = None
        saved = {}
        try:
            app = win32.DispatchEx("Excel.Application")
            app.Visible = False
            app.DisplayAlerts = False
            app.ScreenUpdating = False
            app.EnableEvents = False
            try:
                summary.excel_version = str(app.Version)
            except Exception:
                pass

            t_open0 = time.perf_counter()
            wb = app.Workbooks.Open(copy, UpdateLinks=0, ReadOnly=False)
            summary.open_ms = (time.perf_counter() - t_open0) * 1000

            # Calculation can only be set once a workbook is open.
            try:
                saved["calc"] = app.Calculation
                app.Calculation = XL_MANUAL
            except Exception:
                pass

            # ---- workbook-level: multi-thread full calc
            try:
                app.MultiThreadedCalculation.Enabled = True
            except Exception:
                pass
            multi = self._time_full(app)
            summary.multi_thread_ms = _mean(multi) * 1000
            summary.full_calc_ms = summary.multi_thread_ms

            # ---- recalc (volatiles + dependents only)
            recalc = self._time_calc(app, iters=self.iterations)
            summary.recalc_ms = _mean(recalc) * 1000
            summary.volatility_pct = round(
                100 * summary.recalc_ms / summary.full_calc_ms, 1
            ) if summary.full_calc_ms else 0.0

            # ---- single-thread full calc → MT efficiency
            try:
                app.MultiThreadedCalculation.Enabled = False
                single = self._time_full(app, iters=max(3, self.iterations - 2))
                summary.single_thread_ms = _mean(single) * 1000
                app.MultiThreadedCalculation.Enabled = True
            except Exception:
                summary.single_thread_ms = summary.full_calc_ms
            summary.mt_efficiency = round(
                summary.single_thread_ms / summary.multi_thread_ms, 2
            ) if summary.multi_thread_ms else 1.0

            # ---- scratch sheet
            try:
                scratch = wb.Sheets(SCRATCH)
            except Exception:
                scratch = wb.Sheets.Add()
                scratch.Name = SCRATCH

            # ---- per-pattern timing on worst K sheets, ranked by suspicion
            worst = [s for s, _ in sorted(suspicion.items(), key=lambda x: -x[1])][: self.top_k]
            worst_set = set(worst)
            todo = [p for p in patterns if p.sheet in worst_set]
            todo.sort(key=lambda p: -(p.occurrences * (8 if p.is_volatile else 2)))

            measured_n = 0
            for p in todo:
                if self._over_budget():
                    summary.partial = True
                    summary.notes.append(
                        f"Time budget {self.budget_s}s hit after {measured_n} patterns; "
                        f"report is partial."
                    )
                    break
                try:
                    src = wb.Sheets(p.sheet).Range(p.sample_cell)
                    r1c1 = src.FormulaR1C1
                    big = min(self.max_copies, max(200, p.occurrences))
                    us, sd = self._measure_pattern(scratch, r1c1, big)
                    # two-point linearity check on large patterns
                    if p.occurrences >= 1000:
                        us_small, _ = self._measure_pattern(scratch, r1c1, max(100, big // 4))
                        if us > 0 and abs(us - us_small) / max(us, 1e-9) > 0.5:
                            p.nonlinear = True
                    p.us_per_occ = round(us, 3)
                    p.stdev_us = round(sd, 3)
                    p.total_ms = round(us * p.occurrences / 1000, 2)
                    p.measured = True
                    measured_n += 1
                except Exception as e:
                    p.flag_reason = f"measure error: {e}"

            self.log(f"  measured {measured_n} patterns "
                     f"({time.perf_counter() - self._t0:.1f}s elapsed)")

            # ---- Power Query refresh timing (separate from recalc)
            if pq_results is not None:
                try:
                    powerquery.time_queries(app, wb, summary, pq_results,
                                            budget_check=self._over_budget)
                except Exception as e:
                    summary.notes.append(f"PQ timing error: {e}")

            # ---- direct save timing
            try:
                if SCRATCH in [s.Name for s in wb.Sheets]:
                    app.DisplayAlerts = False
                    wb.Sheets(SCRATCH).Delete()
                t_save = time.perf_counter()
                wb.Save()
                summary.save_ms = (time.perf_counter() - t_save) * 1000
            except Exception:
                pass

            # ---- direct reopen timing
            try:
                wb.Close(SaveChanges=False)
                wb = None
                t_re = time.perf_counter()
                wb = app.Workbooks.Open(copy, UpdateLinks=0)
                summary.open_ms = (time.perf_counter() - t_re) * 1000
            except Exception:
                pass

        finally:
            try:
                if "calc" in saved and app is not None:
                    app.Calculation = saved["calc"]
            except Exception:
                pass
            try:
                if wb is not None:
                    wb.Close(SaveChanges=False)
            except Exception:
                pass
            try:
                if app is not None:
                    app.Quit()
            except Exception:
                pass
            app = None
            wb = None
            pythoncom.CoUninitialize()
            try:
                shutil.rmtree(tmpdir, ignore_errors=True)
            except Exception:
                pass
