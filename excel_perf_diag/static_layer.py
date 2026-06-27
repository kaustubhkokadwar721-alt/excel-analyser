"""Layer 1 — static analyzer. Groups formulas into cost patterns, no Excel.

Output: list[PatternCost] (unmeasured), list[StructureFinding], and a per-sheet
suspicion score used to pick which sheets Layer 2 will deep-time.
"""

from __future__ import annotations

import os
from collections import defaultdict
from openpyxl import load_workbook
from openpyxl.worksheet.formula import ArrayFormula

from .model import PatternCost
from . import patterns as P
from . import structure as S


def _formula_text(cell):
    v = cell.value
    if isinstance(v, ArrayFormula):
        return v.text, True
    if isinstance(v, str) and v.startswith("="):
        return v, False
    return None, False


def analyze(path: str):
    """Return (patterns, structure_findings, sheet_suspicion, defined_names)."""
    wb = load_workbook(path, data_only=False, read_only=False)

    patterns: dict[tuple, PatternCost] = {}
    structure = []
    suspicion: dict[str, float] = {}

    for ws in wb.worksheets:
        formula_cells: dict[tuple[int, int], str] = {}
        groups: dict[tuple, dict] = defaultdict(
            lambda: {"occ": 0, "sample": None, "rc": None, "cells": 1,
                     "vol": False, "st": False, "cls": None, "funcs": ()}
        )
        for cell in ws._cells.values():
            ftext, is_array = _formula_text(cell)
            if not ftext:
                continue
            r, c = cell.row, cell.column
            formula_cells[(r, c)] = ftext
            key = (ws.title,) + P.grouping_key(ftext, r, c)
            g = groups[key]
            g["occ"] += 1
            if g["sample"] is None:
                g["sample"] = ftext
                g["addr"] = cell.coordinate
                g["rc"] = P.normalize_r1c1(ftext, r, c)
                g["cells"] = P.cells_touched(ftext)
                g["cls"] = P.classify(ftext)
                g["vol"] = P.is_volatile(ftext) or is_array and False
                g["st"] = P.is_single_threaded(ftext)
                g["funcs"] = tuple(sorted(set(P.functions(ftext))))
                if is_array:
                    g["cls"] = "array"

        # materialise PatternCost rows + accumulate suspicion
        sheet_susp = 0.0
        for key, g in groups.items():
            sheet, r1c1, bucket = key[0], key[1], key[2]
            pc = PatternCost(
                sheet=sheet, r1c1=g["rc"] or r1c1, sample_a1=g["sample"],
                sample_cell=g.get("addr", ""),
                func_class=g["cls"], occurrences=g["occ"],
                cells_touched=g["cells"], is_volatile=g["vol"],
                is_single_threaded=g["st"], funcs=g["funcs"],
            )
            patterns[key] = pc
            # static suspicion weight (predicts, does not measure)
            w = g["occ"] * _class_weight(g["cls"])
            if g["vol"]:
                w *= 3
            if g["st"]:
                w *= 2
            if g["cells"] >= 1_048_576:
                w *= 2
            sheet_susp += w
        suspicion[ws.title] = sheet_susp

        structure.append(S.analyze_sheet_structure(ws, formula_cells))

    defined_names = list(wb.defined_names.keys()) if hasattr(wb.defined_names, "keys") else []
    wb.close()
    return list(patterns.values()), structure, suspicion, defined_names


def _class_weight(cls: str) -> float:
    return {
        "volatile": 8, "array": 6, "lookup": 4, "dynamic-array": 3,
        "lambda": 3, "let": 1, "cheap": 1, "other": 2,
    }.get(cls, 1)


def file_size_mb(path: str) -> float:
    return round(os.path.getsize(path) / (1024 * 1024), 3)
