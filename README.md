# Excel Workbook Performance Diagnostic Tool

Diagnoses **what specifically** slows an Excel workbook — per formula pattern, per
sheet, per structure — with **measured, file-specific figures**, not estimates.
Built on the FastExcel batch-and-subtract timing method and the VertiPaq Analyzer
philosophy: *object → cost → reason, ranked*. See
[Excel_Performance_Diagnostic_Tool_Spec_v2.md](Excel_Performance_Diagnostic_Tool_Spec_v2.md)
for the full design.

There are **two interchangeable implementations** of the same engine:

| | Python engine | VBA add-in (`.xlam`) |
|---|---|---|
| Where | runs Excel out-of-process via COM | runs in-process inside Excel |
| Needs | Python + `pywin32` + Excel | only Excel (portable, no install) |
| Best for | development, batch runs, validation, deep structure | locked-down rotating PCs, one-click on the open file |
| Output | separate `*.perfreport.xlsx` (4 sheets) | new report workbook (4 sheets) |

Both produce the same four sheets: **Summary** (health strip + Top-5 Pareto),
**Formula Cost** (per-pattern µs, data-bar Pareto, cumulative %), **Structure**
(used-range waste, links, names), **Actions** (ROI-ranked, evidence-anchored fixes).

---

## 1. Python engine

### Install
```
pip install -r requirements.txt
```

### Run
```
python -m excel_perf_diag <file.xlsx>                 # full Layer 1 + Layer 2
python -m excel_perf_diag <file.xlsx> --static-only   # Layer 1 only, no Excel
```
Options: `--budget 150` (Layer-2 wall-clock seconds), `--copies 2000`
(max scratch copies per pattern), `--iters 6` (calc passes), `--topk 6`
(deep-time only the N worst sheets), `--out PATH`.

Output: `<file>.perfreport.xlsx`. A run fingerprint (`<file>.perffingerprint.json`)
is stored so the **next run shows before/after deltas** — prove a fix worked.

### Package layout
```
excel_perf_diag/
  static_layer.py   Layer 1: pattern grouping + suspicion ranking (openpyxl)
  patterns.py       R1C1 normalise, classify, (R1C1, cells-bucket) grouping key
  structure.py      used-range waste, CF, styles, links, names, chain depth
  dynamic_layer.py  Layer 2: COM timing harness + safety envelope
  powerquery.py     per-query refresh timing (BackgroundQuery:=False)
  recommend.py      evidence-bound ROI recommendation engine
  regression.py     fingerprint store + before/after diff
  report.py         four-sheet report writer + presentation
  cli.py            orchestrator
```

---

## 2. VBA add-in (`.xlam`)

### Build
With Excel installed and **File ▸ Options ▸ Trust Center ▸ Macro Settings ▸
Trust access to the VBA project object model** enabled, run from `xlam/`:
```
powershell -File build_xlam.ps1     # imports the 4 .bas modules + injects the ribbon
```
This imports `modTimer/modCollect/modReport/modRibbon.bas`, saves `PerfDiag.xlam`,
then runs `inject_ribbon.py` to add the **Perf Diagnostic** ribbon tab (the custom
UI is an OOXML package part the VBE import cannot add). The pre-built
`xlam/PerfDiag.xlam` (ribbon included) ships in this folder.

### Install & use
1. Drop `PerfDiag.xlam` into your add-ins folder (e.g. the existing `_MySetup`
   convention) or **File ▸ Options ▸ Add-ins ▸ Manage Excel Add-ins ▸ Browse**.
2. Open the workbook to diagnose and click **Perf Diagnostic ▸ Diagnose Workbook**
   on the ribbon (or run the `RunPerfDiagnostic` macro via Alt+F8). It diagnoses the
   **active workbook** and opens a report workbook.

The ribbon tab has **Diagnose Workbook** (runs the diagnostic) and **About**.
Callbacks are in `modRibbon.bas`; the XML is `customUI14.xml`.

### Safety
Operates with a state save/restore envelope: manual calc, screen updating off,
events off, all restored on every exit path. A scratch sheet is added to the
active workbook and removed on completion. Deep timing is bounded by a wall-clock
budget (partial report if hit), spill formulas are never mass-replicated, and
full-column/array patterns use few copies — so the tool will not hang or OOM the
weak hardware it runs on.

### Modules
`modTimer` (QueryPerformanceCounter timing + full/recalc), `modCollect`
(FormulaR1C1 grouping, classification, batch-and-subtract), `modReport`
(orchestration, safety, four-sheet output, recommendations). Set
`PerfDiagDebug = True` to write phase logs to `%TEMP%\perfdiag_log.txt`.

---

## 3. Sample / test workbooks
- `make_sample_workbook.py` → `PerfTest_Sample.xlsx`: full stress workbook with
  **legacy vs modern formulas on identical data** (VLOOKUP↔XLOOKUP,
  SUMPRODUCT↔SUMIFS, volatiles, full-column, dynamic arrays, LAMBDA, structured
  refs, deep chain, CF, styles bloat, ghost cells). `python make_sample_workbook.py [N_ROWS]`.
- `make_small_sample.py` → `PerfTest_Small.xlsx`: fast fixture for quick `.xlam` checks.

These double as **self-validating fixtures**: matched legacy/modern pairs have a
known expected ranking (modern native beats legacy array/exact-match/volatile),
so the engine's output can be checked without a paid reference tool.

---

## 4. Known limits
- Per-formula timing is noisy in absolute terms — reliable for **ranking**, not billing (stdev is reported).
- `.xlsb`/`.xlsm`: the Python static layer reads `.xlsx`; `.xlsb` needs `pyxlsb` or an Excel convert.
- Always logged: Excel version (dynamic arrays, structured-ref behaviour, and multi-threading are version-dependent).
