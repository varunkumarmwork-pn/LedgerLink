import json
from pathlib import Path

import pytest

from app.learn import INTERNAL, MANUAL, PRIOR_YEAR, TB_LINKED, TBLink, learn_workbook

VENUS = Path(__file__).resolve().parents[2] / "samples" / "Venus_Agencies_2025-26.xlsx"


@pytest.fixture(scope="module")
def venus():
    return learn_workbook(VENUS)


def value(result, cell):
    return round(result.cells[cell].value, 2)


# ---------------------------------------------------------------- expected figures (CLAUDE.md)

def test_expected_figures_recalculated_to_the_paisa(venus):
    assert value(venus, "'Balance Sheet'!D11") == 17328403.28  # Total Capital & Liabilities
    assert value(venus, "'Balance Sheet'!D22") == 17328403.28  # Total Assets
    assert value(venus, "'Profit & Loss'!D12") == 1669525.88  # Profit
    assert value(venus, "'Balance Sheet'!D6") == 13674167.85  # Owners' funds
    assert value(venus, "'BS Schedules'!C58") == 2221798.16  # Cash & bank
    assert value(venus, "PPE!G11") == 35425.45  # Depreciation


def test_every_formula_recalculates_like_excel(venus):
    assert venus.warnings == []  # nothing fell back to Excel's saved value
    assert not any("saved value" in i["message"] for i in venus.inconsistencies)


def test_financial_year_from_tb_period(venus):
    assert venus.financial_year == "2025-26"


# ---------------------------------------------------------------- the three known issues

def test_gods_ac_is_unused(venus):
    gods = [u for u in venus.unused_ledgers if u["name"] == "GODS A/C"]
    assert gods == [{"ledger": "Capital Account > GODS A/C", "name": "GODS A/C", "group_path": "Capital Account",
                     "dr": 0.0, "cr": 3384.0, "typed_in": []}]


def test_insurance_is_manual_and_matches_tb_ledgers(venus):
    cell = venus.cells["'P&L Notes'!C26"]
    assert (cell.category, cell.formula, cell.label) == (MANUAL, "=3540+6080", "Insurance Paid")
    manual = next(m for m in venus.manual_cells if m["cell"] == "'P&L Notes'!C26")
    assert manual["matches_ledgers"] == ["Shop Insurance A/c", "Vehicle Insurance"]
    unused = {u["name"]: u["typed_in"] for u in venus.unused_ledgers}
    assert unused["Shop Insurance A/c"] == ["'P&L Notes'!C26"]
    assert unused["Vehicle Insurance"] == ["'P&L Notes'!C26"]


def test_balance_sheet_row_21_inconsistency(venus):
    # Only row 21 is flagged; PPE row 14 (closing vs opening column, same row) is fine.
    assert len(venus.inconsistencies) == 1
    issue = venus.inconsistencies[0]
    assert issue["cell"] == "'Balance Sheet'!D21"
    assert issue["label"] == "Other Current Assets"
    assert "'BS Schedules'!C74" in issue["message"] and "'BS Schedules'!D76" in issue["message"]


def test_full_unused_ledger_list(venus):
    # Everything no formula reads. Fixed assets and depreciation come from the typed PPE
    # sheet; capital from last year's closing; Advance Income Tax is typed into Drawings.
    assert [u["name"] for u in venus.unused_ledgers] == [
        "Advance Income Tax", "GODS A/C", "T L Venkatesh",
        "Car Accunnt", "Computer  A/C", "Furnitures & Fixtures", "Sanghavi Opera Flat (50% Share)",
        "Scooter", "Shop Building", "Weighing Scale A/C",
        "Depreciation", "Shop Insurance A/c", "Vehicle Insurance", "Profit & Loss A/c",
    ]
    unused = {u["name"]: u["typed_in"] for u in venus.unused_ledgers}
    assert unused["Advance Income Tax"] == ["'BS Schedules'!C14"]  # =360000+...+195000


def test_group_reference_covers_its_ledgers(venus):
    # =TB!B63-TB!C63 reads the Direct Expenses group, so Round Off (under it) is used.
    unused = {u["name"] for u in venus.unused_ledgers}
    assert {"Round Off", "Cooli", "Gst Sales @ 5%", "Discount"}.isdisjoint(unused)


# ---------------------------------------------------------------- TB references by ledger name

def test_tb_references_resolve_to_ledger_names(venus):
    assert venus.cells["'BS Schedules'!C55"].tb_links == [
        TBLink("Current Assets > Bank Accounts > Bank of Baroda", "Dr")]
    assert venus.cells["'BS Schedules'!C28"].template == (
        "={TB:Current Liabilities > Sundry Creditors|Cr}-{TB:Current Liabilities > Sundry Creditors|Dr}")
    assert venus.cells["'P&L Notes'!C5"].template == "={TB:Sales Accounts|Cr}"  # a group
    assert venus.cells["'P&L Notes'!C16"].template == (
        "={TB:Current Assets > Inventory Difference Account|Dr}+{TB:Closing Stock|Cr}")


def test_mapping_is_keyed_by_ledger_name_and_saves_as_json(venus):
    mapping = venus.mapping()
    json.dumps(mapping)  # must be storable
    assert mapping["ledgers"]["Current Assets > Bank Accounts > Bank of Baroda"]["used_in"] == [
        {"cell": "'BS Schedules'!C55", "side": "Dr"}]
    assert mapping["ledgers"]["Capital Account > GODS A/C"]["used_in"] == []
    # manual inputs are kept with their formula and value
    drawings = mapping["cells"]["'BS Schedules'!C14"]
    assert (drawings["category"], drawings["formula"], drawings["value"]) == (
        MANUAL, "=360000+45120+363+63073+195000", 663556.0)


# ---------------------------------------------------------------- classification and graph

@pytest.mark.parametrize("cell, category", [
    ("'BS Schedules'!C55", TB_LINKED),    # =TB!B48
    ("'Balance Sheet'!D18", INTERNAL),    # ='BS Schedules'!C58
    ("'BS Schedules'!C58", INTERNAL),     # =SUM(C55:C57)
    ("'BS Schedules'!C5", MANUAL),        # typed "As per last Balance Sheet"
    ("'BS Schedules'!C10", MANUAL),       # =(5101.88)
    ("PPE!F4", MANUAL),                   # depreciation rate
    ("'Balance Sheet'!E21", PRIOR_YEAR),  # ='BS Schedules'!D76
    ("'BS Schedules'!D55", PRIOR_YEAR),   # typed last-year figure
])
def test_cell_categories(venus, cell, category):
    assert venus.cells[cell].category == category


def test_note_numbers_and_text_are_not_amounts(venus):
    assert "'Balance Sheet'!C6" not in venus.cells  # note number 1
    assert "'Profit & Loss'!B16" not in venus.cells  # formula returning text


def test_dependency_graph(venus):
    g = venus.precedents
    assert g["'Balance Sheet'!D11"] == [f"'Balance Sheet'!D{r}" for r in range(6, 11)]
    assert g["'Balance Sheet'!D6"] == ["'BS Schedules'!C18"]
    assert g["'BS Schedules'!C7"] == ["'Profit & Loss'!D12"]  # profit -> capital note
    assert g["'P&L Notes'!C32"] == ["PPE!G11"]  # depreciation -> P&L note
