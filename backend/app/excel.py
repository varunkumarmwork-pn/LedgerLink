"""The client's workbook as an Excel file: the saved file with the CA's edits applied.

Used for "Download Excel", as the starting point of a roll-forward, and to recheck figures
after edits. Formulas stay live; Excel recalculates them when the file is opened.
"""
import json
import os
import tempfile
from functools import lru_cache
from io import BytesIO
from pathlib import Path

import openpyxl
from openpyxl.workbook import Workbook

from app.learn import LearnResult, learn_workbook


def effective_workbook(path: str | Path, edits: dict) -> Workbook:
    """The saved workbook (formulas, not values) with each CA edit written into its cell."""
    wb = openpyxl.load_workbook(path)
    for sheet, cells in edits.items():
        for address, edit in cells.items():
            wb[sheet][address].value = edit.get("formula") or edit.get("value")
    return wb


def workbook_bytes(path: str | Path, edits: dict) -> bytes:
    wb = effective_workbook(path, edits)
    wb.calculation.fullCalcOnLoad = True  # openpyxl cannot calculate; ask Excel to on open
    buffer = BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


def learn_effective(path: str | Path, edits: dict) -> LearnResult:
    """Learn the workbook as the CA has left it (edits included). Cached: most review clicks
    change a decision, not the cells, so the same workbook is often learnt twice in a row."""
    return _learn_cached(str(path), json.dumps(edits, sort_keys=True))


@lru_cache(maxsize=16)
def _learn_cached(path: str, edits_json: str) -> LearnResult:
    edits = json.loads(edits_json)
    if not edits:
        return learn_workbook(path)
    handle, temp = tempfile.mkstemp(suffix=".xlsx")
    os.close(handle)
    try:
        effective_workbook(path, edits).save(temp)
        return learn_workbook(temp)
    finally:
        os.unlink(temp)
