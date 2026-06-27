# Excel Workbook Performance Diagnostic Tool — Build Specification v2

*Research synthesis, architecture, and build specification (revised)*

**Prepared for:** Kaustubh Kokadwar
**Target build environment:** Claude Code (Excel CLI available)
**Original spec date:** 27 June 2026
**This revision:** integrates de-risked sequencing, on-weak-hardware safety, modern-Excel cost model, before/after regression loop, presentation layer, and an evidence-bound recommendation engine.

> **Status of this document.** v2 supersedes the original `.docx` brief. The original thesis is unchanged and correct: **measure cost, attribute it to specific objects, explain why each cost is what it is.** v2 fixes sequencing risk, adds the safety envelope the tool needs to run on the same weak hardware it diagnoses, modernises the cost model, and specifies how output is presented and turned into ranked actions.

---

## Contents

1. The problem we are solving
2. What the research established (unchanged — retained from v1)
3. The design template: VertiPaq Analyzer (principle, not copy)
4. Form factor decision
5. Architecture: two layers
6. **Revised sequencing — de-risk the core first**
7. **Safety envelope — do not crash the patient**
8. **Modern-Excel cost model — beyond the legacy drivers**
9. **The measurement details that must be right** (grouping, sampling, ratios)
10. **Output design — three diagnostic sheets + Actions sheet**
11. **Presentation layer — make conclusions jump out**
12. **Recommendation engine — actions from output**
13. **Before/after regression mode — closing the loop**
14. Build brief for Claude Code (modules, signatures)
15. Validation plan
16. Caveats and known risks
17. Decisions locked
18. The sample test workbook

---

## 1. The problem we are solving

**Constraint.** The firm runs on weak, ageing hardware (11th-gen i3-class laptops, 4 cores, 8 GB RAM). Most work is Excel-based. Small files are fine; once data volume, formula count, or Power Query load grows, performance collapses — slow recalculation, slow open/save, slow refresh.

**Requirement.** A tool that diagnoses *what specifically* slows a given workbook — data, formats, formatting, tables, formulas, conditional formatting, Power Query — and reports **exact, measured, file-specific figures**, not ballparks.

**Key insight (load-bearing).** Counting formulas is close to useless. Each formula carries a different compute cost depending on volatility, range size touched, function class, data structure, and the table/column it lives in. The whole design follows: **measure cost, attribute it to specific objects, explain why.**

> **What "good" looks like**
> - Never outputs a verdict ("your file is slow"). Outputs an X-ray.
> - Every output row = one concrete object + one cost number + one reason.
> - User diagnoses themselves, fast, because the tool handed them ranked evidence.
> - Exact figures validated against a reference before they are trusted.
> - **v2 addition:** the X-ray also yields a ranked, evidence-bound action list — each action cites the measured row it came from. This is still not a verdict; it is a prioritised worklist derived from evidence.

---

## 2. What the research established

*(Retained verbatim in intent from v1 — the actionable distillation.)*

### 2.1 Why per-formula cost can only be measured by timing
A single formula calculates in microseconds — below any clock's noise floor. Authoritative method (FastExcel):
1. **Group identical formulas** by their R1C1 form.
2. **Calculate many copies at once** single-threaded via `Range.Calculate`, divide total by cell count.
3. **Subtract empty-sheet overhead** — time an empty worksheet the same way and subtract.
4. **Average several runs** — OS scheduling/CPU caching make single readings noisy.

*Consequence: implement batch-and-subtract. Absolute numbers never perfectly repeatable; reliable for ranking, which is what matters.*

### 2.2 What drives compute cost (ranked)

| # | Driver | Why expensive |
|---|---|---|
| 1 | Volatile functions (OFFSET, INDIRECT, NOW, TODAY, RAND, CELL, INFO) | Recalc on every change anywhere; drag all dependents. INDIRECT triply bad: volatile + single-threaded + costly ref resolution. |
| 2 | Array formulas / SUMPRODUCT over large ranges | Force many calculations; often single-threaded; cost scales with cells touched. |
| 3 | Full-column / full-row references (A:A) | Excel scans far more used cells than needed. |
| 4 | Exact-match lookups on large unsorted tables | VLOOKUP/INDEX-MATCH exact scan row-by-row. |
| 5 | VBA/Automation UDFs, GETPIVOTDATA, INDIRECT | Single-threaded — block other cores. Acute on i3. |
| 6 | External workbook links (esp. closed/network) | Add seconds to open; fragile. |
| 7 | Conditional formatting over entire columns | Re-evaluated on every recalc/scroll/selection. |
| 8 | Used-range bloat ("ghost" cells) | Stray formatting extends used range to row 1,048,576. |
| 9 | styles.xml bloat | 64,000-combination ceiling; slows save; often 70–90% reducible. |
| 10 | What-if Data Tables | Recalc whole workbook once per row/column value; single-threaded. |

### 2.3 Two ratios that tell you the nature of the problem
- **Volatility % = recalc time ÷ full-calc time.** High = saturated with volatiles poisoning smart recalc. Fix: OFFSET→INDEX, INDIRECT→CHOOSE/structured refs.
- **Multi-thread efficiency = single-thread time ÷ multi-thread full-calc time.** Low = single-threaded functions/UDFs blocking core use — damaging on few-core i3. Fix: hunt UDFs, INDIRECT, GETPIVOTDATA.

### 2.4 Power Query is a separate animal
PQ refresh is **not** part of Excel's calc chain. Time independently — per query, `BackgroundQuery:=False` — never lumped into recalc. Built-in Query Diagnostics (Exclusive Duration per step) is the cross-check.

### 2.5 Existing tools — and the gap
Spreadsheet Inquire = structural counts only, enterprise-licence, may be admin-disabled, no calc/PQ timing. Check Performance (M365) = cosmetic only. FastExcel Profiler = gold-standard method (we borrow the method, not the product). **Nothing gives exact, file-specific, cost-ranked diagnosis across data + formatting + formulas + Power Query in one place.** Building justified.

---

## 3. The design template: VertiPaq Analyzer (principle, not copy)

VertiPaq Analyzer (DAX Studio, SQLBI) connects to a live model and asks the engine how much memory each object consumes, then ranks it. **Principle worth stealing:** never says "your model is slow"; says *this table → this column → this property → costs this many bytes, ranked.* 80% of cost sits in 20% of objects → ranked Pareto is the core deliverable.

**Why adapt not copy:** VertiPaq queries a live engine over DMVs. Excel exposes no external interface to its calc engine — real timing must happen *inside* Excel's object model. Form factor differs; philosophy carries.

### 3.1 Concept mapping — VertiPaq → Excel

| VertiPaq | Excel equivalent | Measures |
|---|---|---|
| Table | Sheet | Unit of grouping |
| Column | Unique formula pattern (R1C1-normalised) **+ cells-touched bucket** | Unit of cost attribution |
| Cardinality | Cells touched (referenced range × occurrences) | Primary cost driver |
| Encoding (Value vs Hash) | Function class (cheap vs lookup/array/volatile) | Why the cost is what it is |
| Dictionary Size | Volatility tax (cost × needless recalc frequency) | Hidden recurring cost |
| Data Size | Total ms (µs/formula × count) | Actual measured weight |
| Hierarchy Size (unused) | Conditional-format rule count over used range | Structural overhead nobody asked for |
| Relationships / missing keys | External links / broken named ranges | Structural risk |
| 80% in 20% | Pareto of Total ms by pattern | Where to spend fix-time |

---

## 4. Form factor decision

**Decision: a VBA add-in (`.xlam`), not a standalone program.**
1. **Real timing requires being inside Excel.** `Range.Calculate` / `Worksheet.Calculate` / empty-sheet-overhead all need the in-process object model — i.e. VBA (or an XLL via Excel-DNA, overkill here).
2. **Portability across locked-down, rotating PCs.** `.xlam` drops into the existing `_MySetup` folder convention — no install, runs anywhere Excel runs. Same pattern as the existing LAMBDA add-in.
3. **Output-as-sheet is correct for the job.** Diagnostic is kept, compared before/after, pasted into working papers.

**Python keeps a role** as the optional, separate Layer 1 static pass — run outside Excel to triage a file before committing to opening it on weak hardware.

---

## 5. Architecture: two layers

- **Layer 1 — Static analyzer (Python).** Reads `.xlsx` as ZIP + XML (or openpyxl). Fast, no Excel. Produces ranked suspect list + structural figures in seconds.
- **Layer 2 — Dynamic timer (VBA inside Excel).** Batch-and-subtract per-pattern timing, the two ratios, PQ timing. The only source of real cost.

**Split rationale:** cheap static triage finds suspects; expensive dynamic timing measures them. Static ranks structure; only timing reveals real cost.

---

## 6. Revised sequencing — de-risk the core first

> **v2 change.** v1 shipped Layer 1 first as the "quick win." But Layer 1 is structural counting — the exact thing the thesis calls "close to useless" on its own — and the **risky, novel** part is the Layer 2 timing harness. Building the easy/weak piece first means you only learn whether timing works at the very end.

**New order:**

1. **Spike the timing harness (Layer 2 core) first.** Minimum viable: `MicroTimer`, full-calc timer, single-sheet batch-and-subtract on **one known workbook**. Goal: prove the numbers are real and rank correctly before building anything around them. ~1–2 days of risk retired up front.
2. **Power Query timing module.** Safer than micro-timing, and for this firm ("data volume grows → collapse") often the *actual* bottleneck. High value, low noise. Build early.
3. **Layer 1 static analyzer.** Now framed explicitly as **triage that Layer 2 validates**, not a standalone verdict. Fast win, feeds the sampler in step 4.
4. **Full Layer 2:** per-sheet timing, the two ratios, sampling + linearity, open/save + memory probes.
5. **Output writer + presentation layer + Actions sheet.**
6. **Before/after regression mode.**
7. **Ribbon + one-click entry point.**

**Principle:** build the thing most likely to fail first. If timing can't be made reliable, everything else is decoration — find out on day 1, not day 30.

---

## 7. Safety envelope — do not crash the patient

> **v2 addition.** The diagnostic runs on the *same* i3/8 GB machine, against an *already-slow* file. Layer 2 mutates the workbook (scratch ranges, `EnableCalculation` toggling, `CalculateFull`, N copies, ≥5 iterations). Unbounded, the tool itself hangs or OOMs Excel. The safety envelope is **non-negotiable**.

**Hard rules:**
1. **Operate on a temp copy, never the live file.** Save-As a working copy to a temp path; all mutation happens there; original is never touched.
2. **Always restore state in an error handler.** `Calculation`, `ScreenUpdating`, `EnableEvents`, `DisplayAlerts` saved on entry, restored on every exit path (including error). A `Finally`-style cleanup block.
3. **Time budget.** Global wall-clock cap (e.g. 120 s default, configurable). On breach → stop, emit a **partial report** flagged as incomplete, restore state.
4. **Sample only the worst K sheets.** Rank sheets by Layer 1 suspicion / full-calc time; deep-time only the top K (default 5). Don't micro-time everything.
5. **Cap N copies.** Per-pattern batch size capped (e.g. 2,000–4,000). Huge patterns measured by sampling + linearity check (§9), not by copying all occurrences.
6. **Memory guard.** Before allocating scratch ranges, check available memory / used-cell count; back off if near limit. Never build a 500k-cell scratch on 8 GB.
7. **Kill switch.** `Esc`-interruptible (`Application.EnableCancelKey`); clean abort restores state and writes partial output.
8. **Idempotent output.** Re-running clears prior tool-generated sheets cleanly; never corrupts the workbook.

---

## 8. Modern-Excel cost model — beyond the legacy drivers

> **v2 addition.** v1's driver list is the *legacy* model. The firm runs modern Excel and already uses a **LAMBDA add-in** — the classifier must understand current functions or it will misrank real files.

Additional drivers and classifier rules:

| Modern construct | Cost behaviour | Classifier action |
|---|---|---|
| **Dynamic arrays / spill** (FILTER, SORT, UNIQUE, SEQUENCE, RANDARRAY) | Different dirty/recalc behaviour; spill range size = cost; 365/2021 only | Tag `dynamic-array`; cost ∝ spill size; note version dependence |
| **LAMBDA + helpers** (MAP, REDUCE, SCAN, BYROW, BYCOL, MAKEARRAY) | **Not** a VBA UDF — not single-thread-blocking the same way, but can be deep/iterative and costly | Tag `lambda`; classify by iteration breadth, *not* as a single-thread blocker |
| **LET** | Caches intermediate results — often *cheaper*; presence is usually good | Tag `let`; do not flag as cost unless body is expensive |
| **XLOOKUP / XMATCH** | Binary-search mode much faster than exact VLOOKUP; default exact | Compare against VLOOKUP equivalents; recommend where legacy exact-match dominates |
| **SUMIFS/COUNTIFS/AVERAGEIFS/MAXIFS** | Native, multi-threaded, cheap vs SUMPRODUCT/array | Tag `cheap`; recommend as replacement target |

**Dependency-chain depth (new driver).** 100k formulas in a 50-deep chain serialise even when multi-threaded; 100k flat parallelise. Build a precedent graph (evaluate pycel / vinci1it2000-formulas first) and report **max chain depth + longest path**. Explains "MTC efficiency low even with no UDFs."

**Memory dimension (new).** On 8 GB, paging/thrash slows everything and won't show in calc timing. Track footprint: file size + used-cell count + PQ buffer + style-table size → memory-risk band.

**Iterative calc / circular refs (new).** If iterative calc is enabled, recalc cost explodes — detect and flag.

---

## 9. The measurement details that must be right

**9.1 Grouping key — R1C1 alone is wrong.**
Copied formulas pointing at different table columns share an R1C1 form but touch different-sized ranges = different cost. Grouping purely by R1C1 merges different-cost formulas.
> **Fix:** cluster key = `(R1C1 pattern, cells-touched bucket)`. Measure cells-touched per occurrence; sub-bucket when it varies materially.

**9.2 Sampling + linearity check for huge patterns.**
A pattern with 500k occurrences must not be copied 500k times. Measure at **two sizes** (e.g. 1k and 4k copies), extrapolate, and **verify linearity**. Lookups grow with table size → non-linear; a single-point measure lies. Two-point catches it.

**9.3 Volatility ratio needs a defined trigger.**
`recalc ÷ full-calc` is only meaningful if recalc actually fires something. After a full calc the sheet is clean; define a **representative edit** (dirty one input cell) before timing recalc, or use `CalcSeqCountRef`-style probes to confirm which formulas re-fire.

**9.4 Trust signal.**
Report **stdev across runs** per pattern. User sees which numbers are solid vs noisy. Aligns with "rank, don't bill."

**9.5 Range.Calculate caveat.**
`Range.Calculate` is single-threaded and ignores some dependency behaviour — perfect for *comparing* formula cost, wrong for *real-world* recalc. Use full-calc/recalc for workbook-level figures, batch-and-subtract for per-pattern comparison.

**9.6 Direct open/save + CF timing (the invisible complaints).**
- **Open/save:** the complaint list says "slow open/save," which ≠ calc slowness. Measure directly: close + reopen the temp copy, time it; save the copy, time it; attribute to external links / used-range bloat / styles / volatile-on-open.
- **Conditional formatting:** CF cost is dynamic (re-eval on scroll/select). Measure by timing a forced screen refresh / selection change with CF enabled vs disabled on the copy — turn the structural guess into a number.

---

## 10. Output design — three diagnostic sheets + Actions sheet

Mirrors VertiPaq's Summary / Columns / Relationships split, **plus a fourth Actions sheet** (v2). Every flag carries its reason inline — a red cell with no explanation is decorative, not diagnostic.

### 10.1 Sheet 1 — Summary
One screen: "is this file sick, and in what way."

| Metric | Meaning |
|---|---|
| Full-calc ms | Worst-case: every formula recalculated |
| Recalc ms | Best-case: only dirty + volatile + dependents |
| Volatility % | recalc ÷ full-calc. High = volatile-saturated |
| Multi-thread efficiency | single ÷ multi. Low = single-threading bottleneck |
| Max dependency depth | Longest serial chain (new) |
| File size / used-range waste % | Bloat indicators |
| Open ms / Save ms | Measured directly (new) |
| Memory-risk band | Footprint vs 8 GB (new) |
| PQ total refresh (s) | Sum of per-query times (timed separately) |

### 10.2 Sheet 2 — Formula Cost (the Pareto, the main one)
One row per `(pattern, cells-touched bucket)` per sheet, **sorted by Total ms descending.**

| Column | Content |
|---|---|
| Sheet | Where the pattern lives |
| Pattern (R1C1) | Normalised formula form |
| Function class | cheap / lookup / array / volatile / dynamic-array / lambda |
| Occurrences | Cells sharing this pattern |
| Cells touched / occurrence | Referenced range size |
| µs / occurrence | Measured (batch-and-subtract) |
| **Stdev µs** | Trust signal (new) |
| Total ms | µs × occurrences — the ranking key |
| Cumulative % | Pareto cutoff (new) |
| % of sheet time | Share of sheet calc cost |
| Volatile? / Single-threaded? | Flags |
| Flag + reason | e.g. "INDIRECT — volatile + single-threaded + ref-resolution cost" |

### 10.3 Sheet 3 — Structure
Non-formula cost and risk surface: used range vs actual data (waste %); CF rule count + range size (flag full-column); style count vs 64k ceiling; external link count; named-range count with unused flagged; max dependency depth; PQ list with per-query refresh times.

### 10.4 Sheet 4 — Actions (new, §12)
Ranked, evidence-bound worklist. Each row cites the Formula-Cost / Structure row it derives from.

---

## 11. Presentation layer — make conclusions jump out

> **v2 addition.** Reconciles with the anti-verdict rule: never opinion, always measured number → color. Every coloured cell carries its reason in the adjacent cell.

**Summary sheet — top to bottom, one screen:**
```
┌─ HEALTH STRIP ─────────────────────────────────┐
│ Volatility ███░░ 62% ⚠   Threading ██░░░ 31% ✗  │  ← band color driven by
│ Bloat ████░ waste 71% ⚠  PQ ██████ 18.4s   ✗   │    measured threshold,
│ Memory ███░░ 5.8/8 GB ⚠  Open ████ 22s     ✗   │    not opinion
└─────────────────────────────────────────────────┘
TOP 5 COSTS  (= 80% of calc time)
 1. Sheet3  INDIRECT pattern     4,200 ms  ▮▮▮▮▮▮▮▮
 2. Sheet1  SUMPRODUCT A:A       1,800 ms  ▮▮▮▮
 3. ...
[Full 9,400ms | Recalc 6,100ms | Vol 65% | MTC 31% | Depth 47]
```
- **Health strip:** 4–6 dimensions, each a color band tied to a measured threshold (number drives color, not judgement).
- **"Top 5 = 80%" callout** pulled from the Pareto to the top. User reads 5 rows and is done.

**Formula Cost sheet — the Pareto, made visual:**
- **Data bars** (conditional format) on `Total ms` → the 80/20 is visible without reading numbers.
- **Cumulative %** column → the Pareto cutoff line ("rows above 80%").
- **Outline grouping** → collapse to per-sheet rollup, expand to patterns.
- **Icon flags** (⚡ volatile, 🔒 single-thread, ⬚ spill) not just booleans.
- Pre-applied: freeze header, autofilter, sorted by Total ms desc.
- **Stdev column** as trust signal.

**Universal rule:** every red cell → reason in the next cell. No naked colour.

---

## 12. Recommendation engine — actions from output

> **v2 addition.** Still not a verdict: a prioritised worklist, each row **anchored to a measured evidence row**. Generic, unanchored advice is forbidden.

**Sheet 4 — Actions**, ranked by measured ROI:

| Anchor (evidence row) | Measured cost | Why | Fix | Est. gain | Effort | Confidence | **ROI** |
|---|---|---|---|---|---|---|---|

`ROI = (Est. gain × Confidence) ÷ Effort`, sorted descending.

**Signal → action rules** (each evidence-bound; "Est. gain" derived from a measured component already captured — never fabricated):

| Measured signal | Fix | Est. gain source |
|---|---|---|
| Volatile pattern, high Total ms | OFFSET→INDEX, INDIRECT→CHOOSE/structured ref | volatility-tax portion of recalc |
| Full-col ref, cells-touched ≫ data rows | Bound to used range / convert to Table | wasted-cell ratio |
| Exact-match lookup, large unsorted, high µs | Sort + approx, or XLOOKUP, or helper key | scan-reduction estimate |
| MTC low + UDF/INDIRECT/GETPIVOTDATA | Remove single-thread blockers | core-utilisation recovery |
| Used-range waste % high | Delete ghost rows, save, reopen | size + open-time delta |
| styles near 64k | Clean excess formatting | save time + file size |
| CF full-column | Limit CF to used range | measured CF refresh delta (§9.6) |
| PQ slow + folding broken | Push filters early / fix folding / disable unused load | per-query refresh time |
| Deep dependency chain | Break chain / paste-values intermediate | serialisation relief |

**Honesty rule:** "Est. gain" is labelled an estimate, derived from a measured component. Keeps philosophy intact — hand ranked evidence + ranked actions, user decides.

---

## 13. Before/after regression mode — closing the loop

> **v2 addition. The payoff loop the user actually wants.**

Store a **run fingerprint** in a hidden sheet: pattern list + per-pattern timings + Excel version + machine id + timestamp. On re-run, **diff vs the last fingerprint** and add a delta column:

```
OFFSET→INDEX on Sheet3:  1,200 ms → 300 ms   (−75%)  ✓ improved
Full-column CF removed:    refresh 18s → 6s   (−67%)  ✓ improved
```

This **proves** a fix worked instead of making the user eyeball two reports. Diagnose → act → prove → repeat.

---

## 14. Build brief for Claude Code

Recommended environment: **Claude Code** (Excel CLI). Multi-file, write-run-verify work; can inject and run macros against a live workbook to confirm timing is real.

### 14.1 VBA module structure (`.xlam`)

| Module | Responsibility |
|---|---|
| `modTimer` | MicroTimer via QueryPerformanceCounter/Frequency (~µs resolution) |
| `modSafety` | **(new)** temp-copy management, state save/restore, time budget, kill switch, memory guard |
| `modCalc` | FullCalcTimer, RecalcTimer, SheetTimer, RangeTimer; EnableCalculation toggling |
| `modPatterns` | Group cells by `(R1C1, cells-touched bucket)`; classify function class (incl. dynamic-array/lambda/let); flag volatile/single-threaded |
| `modMeasure` | Batch-and-subtract per-pattern timing; sampling + linearity check; overhead sheet; N-run averaging + stdev |
| `modDepend` | **(new)** precedent graph; max chain depth + longest path |
| `modIO` | **(new)** open/save timing; CF-on/off refresh timing; memory footprint |
| `modPQ` | Enumerate + time queries individually (`BackgroundQuery:=False`) |
| `modStructure` | Used-range waste, conditional formats, styles, links, names |
| `modRecommend` | **(new)** signal→action rules; ROI scoring |
| `modRegression` | **(new)** fingerprint store + diff vs last run |
| `modOutput` | Write Summary / Formula Cost / Structure / Actions; data bars, cumulative %, outline grouping, inline reasons |
| `modRibbon` | Ribbon button(s) / one-click "Diagnose this workbook" entry point |

### 14.2 Key function signatures (starting point)
```vba
Function MicroTimer() As Double
Function TimeRange(rng As Range, iterations As Long) As Double
Function FullCalcMs(wb As Workbook) As Double
Function SheetFullCalcMs(ws As Worksheet) As Double     ' toggles EnableCalculation
Function NormalizeR1C1(f As String) As String
Function ClassifyFormula(f As String) As String         ' cheap|lookup|array|volatile|dynamic|lambda|let
Function MeasurePatternUs(pattern As String, n As Long, ByRef stdevOut As Double) As Double  ' batch-and-subtract + variance
Function MeasurePatternSampled(pattern As String) As Double   ' two-point + linearity check
Function MaxChainDepth(ws As Worksheet) As Long
Sub TimeOpenSave(srcPath As String, ByRef openMs As Double, ByRef saveMs As Double)
Sub TimeAllQueries(wb As Workbook)                      ' BackgroundQuery:=False
Function ScoreActions() As Variant                       ' ROI-ranked recommendations
Sub WriteReport(wb As Workbook)                          ' four sheets + presentation
Sub WithSafety()                                         ' temp copy + state save/restore + budget
```

---

## 15. Validation plan
1. Keep one genuinely slow workbook **plus the synthetic stress file (§18)** for testing on real i3 hardware.
2. **Self-validating fixtures** (since FastExcel is paid and may be unavailable): the sample workbook (§18) has sheets of *known relative cost* — pure arithmetic vs pure INDIRECT vs big SUMPRODUCT — so ranking correctness is checkable without FastExcel. Where FastExcel *is* available, cross-check and reconcile threading/overhead before trusting output.
3. Confirm Layer 1 suspect ranking actually predicts Layer 2 measured cost — if a "scary" pattern measures cheap, tune the classifier.
4. Record the Excel version in every report (structured-ref behaviour, dynamic arrays, multi-threading are all version-dependent).

---

## 16. Caveats and known risks
- **Per-formula timing is inherently noisy.** Use for ranking, not billing. Report stdev.
- **Range.Calculate is single-threaded** and ignores some dependency behaviour — comparison tool, not real-world recalc.
- **Static analysis cannot measure time.** Ranks suspects by structure; only timing reveals real cost.
- **Version dependence is real.** Structured-ref dirty-flagging ~1000× slower pre-2016; dynamic arrays only 365/2021; multi-threading absent in legacy `.xls`. Always log version.
- **Inquire is enterprise-only** and may be disabled — replicate its useful signals in Layer 1.
- **Open-source calc engines implement a subset** — use for structure/dependency extraction, not as a timing substitute.
- **`.xlsb` / `.xlsm`:** heavy real files are often binary `.xlsb` — openpyxl/ZIP-XML **cannot read it**. Detect format; use `pyxlsb` for read or convert via Excel first. (new)
- **The tool runs on the weak hardware it diagnoses** — without the §7 safety envelope it can hang/OOM Excel. (new)

---

## 17. Decisions locked

| Question | Decision |
|---|---|
| Count or measure? | **Measure.** Counting misleads. |
| Form factor? | **VBA `.xlam`** add-in (portable, no install, in-process timing). |
| Copy VertiPaq? | **No.** Adapt its principle (object + cost + reason, ranked). |
| Architecture? | **Two layers:** Python static triage + VBA dynamic timing. PQ timed separately. |
| Output? | **Four sheets:** Summary, Formula Cost, Structure, **Actions**. |
| Sequencing? | **Spike timing first**, then PQ, then Layer 1, then full Layer 2, then output, then regression. (revised) |
| Safety? | **Temp copy + state restore + time budget + sampling + kill switch.** (new) |
| Cost model? | **Legacy + modern** (dynamic arrays, LAMBDA, LET, chain depth, memory). (new) |
| Where to build? | **Claude Code** (Excel CLI) for write-run-verify; keep a real slow workbook + the synthetic file. |
| First deliverable? | **The Layer 2 timing spike** — retire the core risk first. (revised) |

---

## 18. The sample test workbook

A synthetic workbook is generated to exercise every driver and to put **legacy vs modern formulas on the same data** so the tool can measure the difference. See `make_sample_workbook.py`; output `PerfTest_Sample.xlsx`.

**Design:** one shared `Data` table feeds matched legacy/modern sheet pairs so cost comparisons are apples-to-apples.

| Sheet | Tests | Expected ranking signal |
|---|---|---|
| `README` | What each sheet tests + expected verdict | — |
| `Data` | Shared base dataset (N rows × typed columns) + one Excel Table | source of truth |
| `Lookup_Legacy` | VLOOKUP exact, INDEX/MATCH exact, nested IFERROR(VLOOKUP) | expensive (exact scan) |
| `Lookup_Modern` | XLOOKUP, FILTER, XMATCH — same results | cheaper (binary / native) |
| `Agg_Legacy` | SUMPRODUCT, CSE array SUM(IF), DSUM | expensive (array) |
| `Agg_Modern` | SUMIFS, COUNTIFS, AVERAGEIFS, MAXIFS | cheap (native, multi-thread) |
| `Volatile` | OFFSET, INDIRECT, NOW, TODAY, RAND, CELL, INFO | **poison** — high volatility % |
| `FullColumn` | `A:A` refs vs bounded equivalents | full-column waste |
| `DynamicArrays` | FILTER, SORT, UNIQUE, SEQUENCE, spill | dynamic-array class |
| `Lambda` | LAMBDA + MAP/REDUCE/SCAN, named LAMBDA | lambda class (not UDF-blocking) |
| `Tables_Structured` | Excel Table `[@col]` structured refs vs plain `A2*B2` | structured-ref behaviour |
| `DepChain` | Deep chain (each cell refs previous) | max chain depth driver |
| `CondFormat` | Full-column CF + heavy rules | dynamic CF cost |
| `Formats_Bloat` | Many distinct cell formats | styles.xml stress |
| `GhostCells` | Formatting far down/right | used-range bloat |

The matched pairs (`*_Legacy` vs `*_Modern`) are the core of the test: identical inputs, different formula technology, so the diagnostic's per-pattern numbers can be checked against the *known* expectation that modern native functions beat legacy array/exact-match/volatile equivalents.
