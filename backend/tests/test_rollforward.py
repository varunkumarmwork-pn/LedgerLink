"""ROLL-FORWARD on Venus with a FY 2026-27 TB that has shifted rows, a renamed ledger, two new
ledgers and a dropped one (see tb_samples.py)."""
import json
from datetime import datetime

import openpyxl
import pytest

from app.learn import learn_workbook
from app.review import apply_action, build_review
from app.rollforward import MISSING, RENAMED, SAME, roll_forward, shift_text
from tb_samples import VENUS, make_modified_venus_tb

NOW = datetime(2026, 9, 24, 10, 0)


@pytest.fixture(scope="module")
def rolled(tmp_path_factory):
    folder = tmp_path_factory.mktemp("rf")
    new_tb = make_modified_venus_tb(folder / "TB 2026-27.xlsx")
    mapping = json.loads(json.dumps(learn_workbook(VENUS).mapping()))
    out = folder / "Venus FY 2026-27.xlsx"
    facts = json.loads(json.dumps(roll_forward(VENUS, mapping, new_tb, out)))
    result = learn_workbook(out)
    stored = {**json.loads(json.dumps(result.mapping())), "rollforward": facts}
    report = json.loads(json.dumps(result.report()))
    return {"out": out, "facts": facts, "stored": stored, "report": report, "wb": openpyxl.load_workbook(out),
            "tb": new_tb, "folder": folder}


def value(rolled, sheet, cell):
    return rolled["wb"][sheet][cell].value


# ---------------------------------------------------------------- ledgers matched by name

def test_ledgers_matched_by_name(rolled):
    matches = rolled["facts"]["matches"]
    renamed = matches["Indirect Expenses > Bank Chargers"]  # found by the AI layer: amber, with a reason
    assert (renamed["new"], renamed["how"], renamed["source"]) == ("Indirect Expenses > Bank Charges", RENAMED,
                                                                   "similarity")
    assert renamed["confidence"] == 0.98  # name 96% alike, same group, same balance
    assert "name 96% alike" in renamed["reason"] and "98% confidence" in renamed["reason"]
    assert matches["Indirect Expenses > Tally Software Services"]["how"] == MISSING
    assert matches["Current Assets > Bank Accounts > Bank of Baroda"]["how"] == SAME
    assert rolled["facts"]["new_ledgers"] == ["Capital Account > Interest on SB Account",
                                              "Current Assets > Bank Accounts > HDFC Bank"]


@pytest.mark.parametrize("sheet, cell, formula", [
    ("BS Schedules", "C55", "=TB!B49"),            # Bank of Baroda: row 48 -> 49
    ("BS Schedules", "C56", "=TB!B50"),            # Union Bank of India
    ("BS Schedules", "C28", "=TB!C25-TB!B25"),     # Sundry Creditors, Cr less Dr
    ("BS Schedules", "C8", "=TB!C14"),             # above the first new ledger: unchanged row
    ("P&L Notes", "C5", "=TB!C53"),                # Sales Accounts group
    ("P&L Notes", "C16", "=TB!B52+TB!C70"),        # Inventory Difference + Closing Stock
    ("P&L Notes", "C37", "=TB!B73"),               # Bank Chargers renamed Bank Charges
    ("P&L Notes", "C27", "=0"),                    # Tally Software Services dropped
])
def test_tb_formulas_rewritten_to_new_rows(rolled, sheet, cell, formula):
    assert value(rolled, sheet, cell) == formula


def test_tb_sheet_replaced_under_the_same_name(rolled):
    tb = rolled["wb"]["TB"]  # formulas elsewhere still say TB!
    assert tb["A49"].value == "Bank of Baroda" and tb["B49"].value == 1029943.56
    assert tb["A16"].value == "Interest on SB Account"
    assert tb["A7"].value == "1-Apr-26 to 31-Mar-27"
    assert tb["A16"].font.i  # Tally formatting kept (ledger = italic)


# ---------------------------------------------------------------- prior year and headings

@pytest.mark.parametrize("sheet, cell, expected", [
    ("BS Schedules", "D55", 1029943.56),     # last year's Bank of Baroda, typed
    ("BS Schedules", "D5", 12525518.09),     # last year's opening capital
    ("BS Schedules", "D18", 13674167.85),    # last year's closing capital
    ("Balance Sheet", "E11", 17328403.28),   # was a formula; now last year's figure
    ("Balance Sheet", "E21", 1714.68),       # this year's D21 value, not D76's
    ("P&L Notes", "D26", 9620),              # insurance =3540+6080 -> its value
])
def test_current_year_becomes_prior_year_as_numbers(rolled, sheet, cell, expected):
    assert value(rolled, sheet, cell) == expected


@pytest.mark.parametrize("sheet, cell, text", [
    ("BS Schedules", "C4", "FY 2026-27 (₹)"),
    ("BS Schedules", "D4", "FY 2025-26 (₹)"),
    ("Balance Sheet", "B2", "Balance Sheet AS AT 31.03.2027"),
    ("PPE", "D3", "Addition upto 3-10-2026"),
])
def test_headings_move_on_a_year(rolled, sheet, cell, text):
    assert value(rolled, sheet, cell) == text


def test_shift_text_leaves_other_numbers_alone():
    assert shift_text("Income Tax Act, 1961; Membership No.024012; FY 2019-20", 2025) == \
        "Income Tax Act, 1961; Membership No.024012; FY 2019-20"


# ---------------------------------------------------------------- carried forward

def test_openings_carried_forward(rolled):
    assert value(rolled, "BS Schedules", "C5") == 13674167.85  # As per last Balance Sheet = last closing
    assert value(rolled, "PPE", "C4") == 127262.85  # Car: opening = last year's closing WDV
    assert (value(rolled, "PPE", "D5"), value(rolled, "PPE", "E5")) == (0, 0)  # additions reset
    assert value(rolled, "PPE", "G4").startswith("=")  # depreciation formula kept
    assert rolled["facts"]["last_closing_stock"] == {"cell": "'P&L Notes'!C16", "amount": 4021447.26}


# ---------------------------------------------------------------- review and checks

@pytest.fixture(scope="module")
def review(rolled):
    return build_review(rolled["out"], rolled["stored"], rolled["stored"], rolled["report"])


def test_checks(review):
    checks = {c["id"]: c for c in review["checks"]}
    assert checks["tb"]["ok"] and checks["capital"]["ok"] and checks["stock"]["ok"]
    # this test TB is last year's TB with changes, so the balance sheet cannot tally
    assert not checks["bs"]["ok"] and checks["bs"]["label"] == "Balance sheet difference ₹12,28,973.33"
    assert not checks["unused"]["ok"]


def test_review_items_and_colours(review):
    items = {i["id"]: i for i in review["items"]}
    renamed = items["ledger:Indirect Expenses > Bank Chargers"]
    assert renamed["tone"] == "amber" and renamed["title"] == "Bank Charges"
    assert items["ledger:Indirect Expenses > Tally Software Services"]["tone"] == "red"
    assert "New in this year's TB." in items["unused:Current Assets > Bank Accounts > HDFC Bank"]["text"]
    assert items["link:'P&L Notes'!C26"]["change"]["formula"] == "=TB!B81+TB!B83"  # insurance, new rows
    assert "manual:'BS Schedules'!C11" in items  # LIC income typed last year
    assert "ppe:PPE" in items and "check:bs" in items

    colours = review["overrides"]
    assert colours["'P&L Notes'!C37"]["status"] == "a"  # renamed ledger
    assert colours["'P&L Notes'!C27"]["status"] == "r"  # missing ledger
    assert colours["'BS Schedules'!C5"] == {"status": "g", "reason": "carried forward from last year's closing "
                                                                     "('BS Schedules'!C18)"}


def test_approve_and_change_update_the_mapping(rolled, review):
    items = {i["id"]: i for i in review["items"]}
    stored = rolled["stored"]

    # Change: map the dropped ledger to another one; the formula follows and next year repeats it
    stored = apply_action(rolled["out"], stored, items["ledger:Indirect Expenses > Tally Software Services"],
                          "change", {"ledger": "Indirect Expenses > Telephone Charges"}, NOW)
    assert stored["aliases"] == {"Indirect Expenses > Tally Software Services": "Indirect Expenses > Telephone Charges"}
    assert stored["edits"]["P&L Notes"]["C27"]["formula"] == "=TB!B82"

    # Approve a link suggestion: the typed amount becomes a TB formula
    stored = apply_action(rolled["out"], stored, items["link:'P&L Notes'!C26"], "approve", {}, NOW)
    assert stored["edits"]["P&L Notes"]["C26"]["formula"] == "=TB!B81+TB!B83"

    # Leave a ledger unused on purpose: remembered for next year
    stored = apply_action(rolled["out"], stored, items["unused:Capital Account > GODS A/C"], "approve", {}, NOW)
    assert "Capital Account > GODS A/C" in stored["ignored_ledgers"]

    after = build_review(rolled["out"], stored, rolled["stored"], rolled["report"])
    open_ids = {i["id"] for i in after["items"]}
    assert not open_ids & {"ledger:Indirect Expenses > Tally Software Services", "link:'P&L Notes'!C26",
                           "unused:Capital Account > GODS A/C"}
    assert after["overrides"]["TB!A15"]["status"] is None  # left unused: no colour


def test_wrong_year_tb_is_refused(rolled):
    mapping = json.loads(json.dumps(learn_workbook(VENUS).mapping()))
    with pytest.raises(ValueError, match="for FY 2025-26, but the next year is FY 2026-27"):
        roll_forward(VENUS, mapping, VENUS, rolled["folder"] / "wrong.xlsx")
