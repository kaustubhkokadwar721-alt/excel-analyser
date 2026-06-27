"""
excel_perf_diag — Excel Workbook Performance Diagnostic Tool.

Two-layer engine per the build spec:
  Layer 1 (static)  : openpyxl / zip+xml structural triage, no Excel needed.
  Layer 2 (dynamic) : win32com timing harness (batch-and-subtract), real cost.

Produces a four-sheet report (Summary / Formula Cost / Structure / Actions)
with an evidence-bound recommendation engine and before/after regression.
"""

__version__ = "1.0.0"
