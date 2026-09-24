"""Roll-forward, review, finalise and download through the API (Workbook screen)."""
from io import BytesIO

import openpyxl
import pytest

from tb_samples import VENUS, make_modified_venus_tb

CLIENT = {"name": "Venus Agencies", "entity_type": "Proprietorship", "financial_year": "2025-26"}


def learnt(client) -> int:
    client_id = client.post("/api/clients", json=CLIENT).json()["id"]
    client.post(f"/api/clients/{client_id}/learn", files={"file": (VENUS.name, VENUS.read_bytes())})
    return client_id


def rolled_forward(client, tmp_path) -> tuple[int, dict]:
    client_id = learnt(client)
    tb = make_modified_venus_tb(tmp_path / "TB 2026-27.xlsx")
    job_id = client.post(f"/api/clients/{client_id}/rollforward", files={"file": (tb.name, tb.read_bytes())}).json()["job_id"]
    return client_id, client.get(f"/api/jobs/{job_id}").json()


def act(client, client_id, item, action, **payload):
    return client.post(f"/api/clients/{client_id}/workbook/review", json={"item": item, "action": action, **payload})


def test_learned_workbook_has_a_real_review(client, storage):
    client_id = learnt(client)
    body = client.get(f"/api/clients/{client_id}/workbook").json()
    assert (body["status"], body["source"], body["next_financial_year"]) == ("learned", "learn", "2026-27")
    ids = {i["id"] for i in body["review"]}
    assert {"unused:Capital Account > GODS A/C", "issue:'Balance Sheet'!D21", "link:'P&L Notes'!C26",
            "stock:'P&L Notes'!C16", "check:capital"} <= ids
    checks = {c["id"]: c for c in body["checks"]}
    assert checks["bs"] == {"id": "bs", "ok": True, "label": "Balance sheet tallies, difference ₹0.00", "detail": ""}


def test_roll_forward_job_builds_a_draft_for_next_year(client, storage, tmp_path):
    client_id, job = rolled_forward(client, tmp_path)
    assert job["status"] == "done", job["error"]
    assert job["action"] == "Rolling forward" and job["titles"][1] == "Match ledgers to last year"
    assert job["details"][1] == "71 same, 1 renamed or moved, 1 missing, 2 new."
    assert ["'P&L Notes'!C37", "=TB!B73", "Bank Charges"] in job["log"]
    assert "AI matching: rapidfuzz + sentence-transformers" in job["details"][3]

    body = client.get(f"/api/clients/{client_id}/workbook").json()
    assert (body["financial_year"], body["status"], body["source"]) == ("2026-27", "draft", "roll-forward")
    assert body["cells"]["P&L Notes"]["C37"]["status"] == "a"
    assert body["cells"]["P&L Notes"]["C27"]["status"] == "r"
    assert body["cells"]["BS Schedules"]["C5"]["reason"].startswith("carried forward")
    hdfc = next(i for i in body["review"] if i["id"] == "unused:Current Assets > Bank Accounts > HDFC Bank")
    assert hdfc["tone"] == "amber" and "confidence" in hdfc["text"]  # AI suggestion, waiting for the CA
    listed = client.get("/api/clients").json()[0]
    assert listed["financial_year"] == "2026-27" and listed["review_status"].endswith("to check")


def test_roll_forward_needs_a_workbook_and_the_right_year(client, storage, tmp_path):
    client_id = client.post("/api/clients", json=CLIENT).json()["id"]
    tb = make_modified_venus_tb(tmp_path / "tb.xlsx")
    assert client.post(f"/api/clients/{client_id}/rollforward", files={"file": ("tb.xlsx", tb.read_bytes())}).status_code == 409

    client_id = learnt(client)
    job_id = client.post(f"/api/clients/{client_id}/rollforward",
                         files={"file": (VENUS.name, VENUS.read_bytes())}).json()["job_id"]  # same year's TB
    job = client.get(f"/api/jobs/{job_id}").json()
    assert job["status"] == "failed" and "next year is FY 2026-27" in job["error"]


def test_approve_and_change_through_the_api(client, storage, tmp_path):
    client_id, _ = rolled_forward(client, tmp_path)
    body = act(client, client_id, "ledger:Indirect Expenses > Tally Software Services", "change",
               ledger="Indirect Expenses > Telephone Charges").json()
    assert body["cells"]["P&L Notes"]["C27"]["status"] == "p"  # changed by the CA
    body = act(client, client_id, "ledger:Indirect Expenses > Bank Chargers", "approve").json()
    assert body["cells"]["P&L Notes"]["C37"]["status"] == "g"
    ids = {i["id"] for i in body["review"]}
    assert "ledger:Indirect Expenses > Bank Chargers" not in ids

    bad = act(client, client_id, "ledger:Indirect Expenses > Bank Chargers", "approve")
    assert bad.status_code == 404  # already cleared
    bad = act(client, client_id, "unused:Capital Account > GODS A/C", "change", cell="Nowhere!A1")
    assert bad.status_code == 422


def test_finalise_locks_the_year_and_download_has_live_formulas(client, storage):
    client_id = learnt(client)
    assert client.post(f"/api/clients/{client_id}/workbook/finalise").status_code == 409  # items open

    body = act(client, client_id, "link:'P&L Notes'!C26", "approve").json()  # insurance -> =TB!B79+TB!B82
    for item in body["review"]:
        action = next(a["action"] for a in item["actions"] if a["action"] in ("approve", "dismiss"))
        body = act(client, client_id, item["id"], action).json()
    assert body["review"] == []

    body = client.post(f"/api/clients/{client_id}/workbook/finalise").json()
    assert body["status"] == "finalised" and body["finalised_at"]
    assert client.get("/api/clients").json()[0]["review_status"] == "Finalised"
    locked = client.put(f"/api/clients/{client_id}/workbook/cells",
                        json={"changes": [{"sheet": "PPE", "cell": "C4", "value": 1}]})
    assert locked.status_code == 409

    download = client.get(f"/api/clients/{client_id}/workbook/download")
    assert download.headers["content-disposition"] == 'attachment; filename="Venus Agencies FY 2025-26.xlsx"'
    wb = openpyxl.load_workbook(BytesIO(download.content))
    assert wb.sheetnames == ["Balance Sheet", "Profit & Loss", "BS Schedules", "P&L Notes", "PPE", "TB",
                             "Accounting policies"]
    assert wb["BS Schedules"]["C55"].value == "=TB!B48"  # live formula, not a pasted value
    assert wb["P&L Notes"]["C26"].value == "=TB!B79+TB!B82"  # CA's approved link is in the file
    assert wb["Balance Sheet"]["D6"].font.name == "Times New Roman"  # client's own format
    assert wb.calculation.fullCalcOnLoad  # Excel recalculates on open
