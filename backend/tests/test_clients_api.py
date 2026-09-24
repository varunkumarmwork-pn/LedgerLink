import json
from datetime import date
from pathlib import Path

import pytest


VENUS = {
    "name": "Venus Agencies",
    "entity_type": "Proprietorship",
    "financial_year": "2025-26",
    "tb_updated_on": "2026-09-12",
    "review_status": "3 to check",
}


def test_create_list_delete_client(client):
    created = client.post("/api/clients", json=VENUS)
    assert created.status_code == 201
    body = created.json()
    assert body == {**VENUS, "id": body["id"]}

    assert client.get("/api/clients").json() == [body]

    assert client.delete(f"/api/clients/{body['id']}").status_code == 204
    assert client.get("/api/clients").json() == []


def test_new_client_defaults(client):
    body = client.post("/api/clients", json={"name": "Hebbal Hardware Stores",
                                             "entity_type": "Proprietorship", "financial_year": "2024-25"}).json()
    assert body["tb_updated_on"] is None
    assert body["review_status"] == "Not started"


@pytest.mark.parametrize("field, bad", [("entity_type", "Trust"), ("financial_year", "2025"), ("name", "")])
def test_rejects_bad_input(client, field, bad):
    assert client.post("/api/clients", json={**VENUS, field: bad}).status_code == 422


def test_delete_missing_client(client):
    assert client.delete("/api/clients/999").status_code == 404


# ---------------------------------------------------------------- LEARN MODE upload

SAMPLE = Path(__file__).resolve().parents[2] / "samples" / "Venus_Agencies_2025-26.xlsx"


def upload(client, client_id, name=SAMPLE.name, content=None):
    content = SAMPLE.read_bytes() if content is None else content
    return client.post(f"/api/clients/{client_id}/learn", files={"file": (name, content)})


def test_learn_runs_as_a_job_and_stores_workbook(client, storage):
    client_id = client.post("/api/clients", json=VENUS).json()["id"]
    response = upload(client, client_id)
    assert response.status_code == 202

    # TestClient runs the background job before returning, so it has finished here.
    job = client.get(f"/api/jobs/{response.json()['job_id']}").json()
    assert job["status"] == "done", job["error"]
    assert (job["client_id"], job["client_name"], job["filename"]) == (client_id, "Venus Agencies", SAMPLE.name)
    assert job["step"] == 4
    assert job["details"][0] == 'Sheet "TB". Debit and credit both ₹10,12,52,474.34.'
    assert job["details"][3] == "14 TB ledgers not used, 41 typed amounts, 1 inconsistency."
    assert ["'BS Schedules'!C55", "=TB!B48", "Bank of Baroda"] in job["log"]

    saved = list((storage / "clients" / str(client_id)).iterdir())
    assert len(saved) == 1 and saved[0].name.endswith(SAMPLE.name)
    assert saved[0].read_bytes() == SAMPLE.read_bytes()  # original kept unchanged

    listed = client.get("/api/clients").json()[0]
    # 14 unused ledgers, row 21, 2 typed amounts to link, closing stock, opening capital 0.42
    assert listed["review_status"] == "19 to check"
    assert listed["tb_updated_on"] == date.today().isoformat()


def test_job_events_stream(client, storage):
    client_id = client.post("/api/clients", json=VENUS).json()["id"]
    job_id = upload(client, client_id).json()["job_id"]
    with client.stream("GET", f"/api/jobs/{job_id}/events") as response:
        assert response.headers["content-type"].startswith("text/event-stream")
        events = [json.loads(line[len("data: "):]) for line in response.iter_lines() if line.startswith("data: ")]
    assert events[-1]["status"] == "done"  # stream ends once the job has ended


def test_learn_bad_workbook_fails_the_job(client, storage):
    client_id = client.post("/api/clients", json=VENUS).json()["id"]
    job_id = upload(client, client_id, name="tb.xlsx", content=b"not excel").json()["job_id"]
    job = client.get(f"/api/jobs/{job_id}").json()
    assert job["status"] == "failed" and job["error"].startswith("Could not read this workbook.")
    assert list((storage / "clients" / str(client_id)).iterdir()) == []  # upload not kept
    assert client.get("/api/clients").json()[0]["review_status"] == "Workbook not read"


def test_learn_rejects_non_excel_file(client, storage):
    client_id = client.post("/api/clients", json=VENUS).json()["id"]
    assert upload(client, client_id, name="tb.pdf", content=b"x").status_code == 422


def test_learn_unknown_client(client, storage):
    assert upload(client, 999, content=b"").status_code == 404


def test_unknown_job(client):
    assert client.get("/api/jobs/nope").status_code == 404
    assert client.get("/api/jobs/nope/events").status_code == 404


# ---------------------------------------------------------------- workbook screen

def learnt_client(client):
    client_id = client.post("/api/clients", json=VENUS).json()["id"]
    upload(client, client_id)
    return client_id


def test_get_workbook(client, storage):
    client_id = learnt_client(client)
    body = client.get(f"/api/clients/{client_id}/workbook").json()
    assert body["financial_year"] == "2025-26"
    assert len(body["snapshot"]["sheetOrder"]) == 7
    assert body["cells"]["Balance Sheet"]["D21"]["status"] == "r"


def test_save_cell_changes(client, storage):
    client_id = learnt_client(client)
    changes = {"changes": [{"sheet": "P&L Notes", "cell": "C26", "formula": "=TB!B79+TB!B82"}]}
    saved = client.put(f"/api/clients/{client_id}/workbook/cells", json=changes)
    assert saved.status_code == 200
    [cell] = saved.json()["cells"]
    assert (cell["sheet"], cell["cell"], cell["status"], cell["bg"]) == ("P&L Notes", "C26", "p", "#EEE6F5")

    body = client.get(f"/api/clients/{client_id}/workbook").json()  # edit survives a reload
    assert body["cells"]["P&L Notes"]["C26"]["status"] == "p"


@pytest.mark.parametrize("change", [
    {"sheet": "Nope", "cell": "A1", "value": 1},
    {"sheet": "PPE", "cell": "not-a-cell", "value": 1},
    {"sheet": "PPE", "cell": "A1", "formula": "SUM(A2)"},  # formulas start with =
])
def test_save_rejects_bad_changes(client, storage, change):
    client_id = learnt_client(client)
    assert client.put(f"/api/clients/{client_id}/workbook/cells", json={"changes": [change]}).status_code == 422


def test_workbook_not_learnt_yet(client):
    client_id = client.post("/api/clients", json=VENUS).json()["id"]
    assert client.get(f"/api/clients/{client_id}/workbook").status_code == 404
