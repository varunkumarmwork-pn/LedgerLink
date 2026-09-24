import json
from datetime import datetime
from pathlib import Path

import pytest

from app.learn import learn_workbook
from app.workbook_view import PLUM, RED, STATUS_BG, STATUS_INK, apply_edits, edited_cells, workbook_view

VENUS = Path(__file__).resolve().parents[2] / "samples" / "Venus_Agencies_2025-26.xlsx"
NOW = datetime(2026, 9, 23, 11, 30)


@pytest.fixture(scope="module")
def learnt():
    result = learn_workbook(VENUS)
    # round-trip through JSON, as stored in the database
    return json.loads(json.dumps(result.mapping())), json.loads(json.dumps(result.report()))


@pytest.fixture(scope="module")
def view(learnt):
    mapping, report = learnt
    return workbook_view(VENUS, mapping, report)


def sheet(view, name):
    return next(s for s in view["snapshot"]["sheets"].values() if s["name"] == name)


def cell(view, sheet_name, row, col):
    return sheet(view, sheet_name)["cellData"][row - 1][col - 1]


def style(view, c):
    return view["snapshot"]["styles"][c["s"]]


# ---------------------------------------------------------------- review colours and traces

@pytest.mark.parametrize("sheet_name, address, status", [
    ("Balance Sheet", "D21", "r"),   # current year links row 74, last year row 76
    ("P&L Notes", "C26", "r"),       # insurance typed in instead of TB ledgers
    ("TB", "A15", "r"),              # GODS A/C not used
    ("TB", "C15", "r"),
    ("BS Schedules", "C55", "g"),    # =TB!B48
    ("Balance Sheet", "D18", "g"),   # ='BS Schedules'!C58
    ("BS Schedules", "D55", None),   # last year's typed figure
    ("TB", "A48", None),             # Bank of Baroda is used
])
def test_statuses(view, sheet_name, address, status):
    assert view["cells"][sheet_name][address]["status"] == status


@pytest.mark.parametrize("sheet_name, address, trace", [
    ("BS Schedules", "C55", ["Bank of Baroda", "TB: Bank of Baroda"]),
    ("Balance Sheet", "D18", ["Cash & Bank Balances", "BS Schedules: Total Cash & Bank"]),
    ("P&L Notes", "C26", ["Insurance Paid",
                          "typed in (3540+6080); equals TB Shop Insurance A/c, Vehicle Insurance, not linked"]),
    ("TB", "A15", ["GODS A/C", "not used anywhere in the statements"]),
    ("TB", "B48", ["Bank of Baroda", "used in 'BS Schedules'!C55"]),
    ("BS Schedules", "D55", ["Bank of Baroda", "last year's figure, typed in"]),
])
def test_plain_words_trace(view, sheet_name, address, trace):
    assert view["cells"][sheet_name][address]["trace"] == trace


# ---------------------------------------------------------------- Univer snapshot

def test_snapshot_keeps_sheets_in_order(view):
    snapshot = view["snapshot"]
    names = [snapshot["sheets"][sid]["name"] for sid in snapshot["sheetOrder"]]
    assert names == ["Balance Sheet", "Profit & Loss", "BS Schedules", "P&L Notes", "PPE", "TB", "Accounting policies"]


def test_snapshot_has_live_formulas_with_saved_values(view):
    d18 = cell(view, "Balance Sheet", 18, 4)
    assert (d18["f"], d18["v"], d18["t"]) == ("='BS Schedules'!C58", 2221798.16, 2)
    assert cell(view, "Balance Sheet", 18, 2)["v"] == "Cash & Bank Balances"


def test_snapshot_shows_review_colours_and_indian_grouping(view):
    d21 = style(view, cell(view, "Balance Sheet", 21, 4))
    assert d21["bg"] == {"rgb": STATUS_BG[RED]} and d21["cl"] == {"rgb": STATUS_INK[RED]}
    assert d21["n"]["pattern"].startswith('[>=10000000]#","##","##","##0')
    assert d21["ff"] == "Times New Roman"  # client's own font kept


def test_snapshot_layout(view):
    bs = sheet(view, "Balance Sheet")
    assert bs["showGridlines"] == 0
    assert bs["columnData"][1]["w"] == 261  # column B, 36.6 characters wide
    assert len(bs["columnData"]) <= bs["columnCount"]  # not all 16,384 columns
    assert len(bs["mergeData"]) == 2  # the two merged title rows
    assert len(json.dumps(view)) < 300_000


# ---------------------------------------------------------------- CA edits

def test_edit_is_saved_marked_plum_and_laid_over_the_file(learnt):
    mapping, report = learnt
    change = {"sheet": "P&L Notes", "cell": "C26", "formula": "=TB!B79+TB!B82", "value": None}
    edited = apply_edits(VENUS, mapping, [change], NOW)
    assert edited["edits"] == {"P&L Notes": {"C26": {"formula": "=TB!B79+TB!B82", "value": None,
                                                     "edited_at": "2026-09-23T11:30:00"}}}
    assert "edits" not in mapping  # stored mapping not changed in place

    [result] = edited_cells(VENUS, edited, report, [change])
    assert (result["status"], result["bg"], result["ink"]) == (PLUM, STATUS_BG[PLUM], STATUS_INK[PLUM])
    assert result["trace"] == ["Insurance Paid", "edited by CA on 23 Sep 2026: =TB!B79+TB!B82"]

    view = workbook_view(VENUS, edited, report)
    c26 = cell(view, "P&L Notes", 26, 3)
    assert c26["f"] == "=TB!B79+TB!B82"
    assert style(view, c26)["bg"] == {"rgb": STATUS_BG[PLUM]}


def test_edit_on_an_empty_cell_and_typed_value(learnt):
    mapping, report = learnt
    changes = [{"sheet": "PPE", "cell": "c4", "value": 150000, "formula": None},
               {"sheet": "PPE", "cell": "J20", "value": "note", "formula": None}]
    edited = apply_edits(VENUS, mapping, changes, NOW)
    assert set(edited["edits"]["PPE"]) == {"C4", "J20"}
    [c4, _] = edited_cells(VENUS, edited, report, changes)
    assert c4["trace"] == ["Car – Opening Balance", "edited by CA on 23 Sep 2026: 1,50,000.00"]
    view = workbook_view(VENUS, edited, report)
    assert cell(view, "PPE", 4, 3)["v"] == 150000
    assert cell(view, "PPE", 20, 10)["v"] == "note"


def test_changing_a_cell_back_removes_the_edit(learnt):
    mapping, report = learnt
    edited = apply_edits(VENUS, mapping, [{"sheet": "P&L Notes", "cell": "C26", "formula": "=9620"}], NOW)
    undone = apply_edits(VENUS, edited, [{"sheet": "P&L Notes", "cell": "C26", "formula": "=3540+6080"}], NOW)
    assert undone["edits"] == {}
    [result] = edited_cells(VENUS, undone, report, [{"sheet": "P&L Notes", "cell": "C26"}])
    assert result["status"] == RED and result["bg"] == STATUS_BG[RED]  # back to the engine's colour


def test_edit_on_unknown_sheet_is_refused(learnt):
    mapping, _ = learnt
    with pytest.raises(ValueError, match="No sheet"):
        apply_edits(VENUS, mapping, [{"sheet": "Nope", "cell": "A1", "value": 1}], NOW)
