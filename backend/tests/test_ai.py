"""AI mapping layer (Phase 7): client's own mapping first, then the firm dictionary and
similarity, then Ollama for the words. Every match is amber, with a reason and a confidence."""
import json
from collections import Counter
from datetime import datetime

import pytest

from app import ai
from app.learn import learn_workbook
from app.review import apply_action, build_review
from app.rollforward import roll_forward
from tb_samples import VENUS, make_modified_venus_tb

NOW = datetime(2026, 9, 24, 10, 0)


def ledger(name, group, amount=1000.0, key=None):
    return {"key": key or f"{group} > {name}", "name": name, "group": group, "amount": amount}


def cell(key, label, ledgers=(), kind="add", value=0.0):
    return {"cell": key, "label": label, "via": label, "kind": kind, "value": value, "ledgers": list(ledgers)}


# ---------------------------------------------------------------- similarity

def test_spelling_and_meaning():
    assert ai.spelling("Bank Chargers", "Bank Charges") > 0.9
    assert ai.norm("Salary  A/c") == "salary"
    assert ai.meaning.available()  # sentence-transformers, from the local cache only
    # meaning finds what spelling cannot
    assert ai.meaning.similarity("Staff Salary", "Wages paid") > ai.spelling("Staff Salary", "Wages paid") - 0.2
    assert ai.similarity("Staff Salary", "Wages paid") > ai.similarity("Staff Salary", "Rent")


# ---------------------------------------------------------------- 1. the client's own past mapping

def test_client_decision_comes_first():
    old = ledger("Tally Software Services", "Indirect Expenses", 6500)
    candidates = [ledger("Telephone Charges", "Indirect Expenses", 31786.74), ledger("Tally Subscription", "Indirect Expenses", 6500)]
    by_name = ai.match_ledger(old, candidates)
    assert by_name["new"] == "Indirect Expenses > Tally Subscription" and by_name["source"] == "similarity"

    history = {"aliases": {old["key"]: "Indirect Expenses > Telephone Charges"}}
    by_history = ai.match_ledger(old, candidates, history)  # the CA's choice beats similarity
    assert by_history["new"] == "Indirect Expenses > Telephone Charges"
    assert (by_history["source"], by_history["confidence"]) == ("client", 0.97)


def test_client_history_from_earlier_workbooks():
    mapping = json.loads(json.dumps(learn_workbook(VENUS).mapping()))
    history = ai.client_history([mapping])
    assert history["usage"]["bank of baroda"] == Counter({"Bank of Baroda": 1})
    found = ai.suggest_target(ledger("Bank of Baroda", "Current Assets > Bank Accounts"),
                              [cell("X!C5", "Bank of Baroda"), cell("X!C9", "Rent")], history)
    assert found["cell"] == "X!C5" and found["source"] == "client"


# ---------------------------------------------------------------- 2. firm dictionary + similarity

def test_firm_dictionary_from_other_clients():
    other = json.loads(json.dumps(learn_workbook(VENUS).mapping()))
    dictionary = ai.firm_dictionary([other])
    assert dictionary["bank of baroda"] == Counter({("Bank of Baroda", "Current Assets"): 1})

    dictionary = {"staff welfare": Counter({("Employee Benefit Expenses", "Indirect Expenses"): 4})}
    candidates = [cell("X!C5", "Employee Benefit Expenses"), cell("X!C9", "Rent")]
    found = ai.suggest_target(ledger("Staff Welfare Exps", "Indirect Expenses"), candidates, None, dictionary)
    assert found["cell"] == "X!C5" and found["source"] == "firm"
    assert "other clients put “staff welfare” in “Employee Benefit Expenses” (4×)" in found["evidence"]


def test_dictionary_only_counts_the_same_kind_of_group():
    # a loan from "T.L.Venkatesh (H.U.F)" must not pull the proprietor's capital "T L Venkatesh"
    dictionary = {"t l venkatesh h u f": Counter({("Unsecured Borrowings", "Loans (Liability)"): 1})}
    candidates = [cell("X!C23", "Unsecured Borrowings", [{"name": "T.L.Venkatesh (H.U.F )", "group": "Loans (Liability)"}])]
    assert ai.suggest_target(ledger("T L Venkatesh", "Capital Account"), candidates, None, dictionary) is None


def test_nothing_below_the_threshold():
    assert ai.suggest_target(ledger("Suspense", ""), [cell("X!C5", "Rent"), cell("X!C9", "Audit Fees")]) is None
    assert ai.match_ledger(ledger("Tally Software Services", "Indirect Expenses", 6500),
                           [ledger("HDFC Bank", "Current Assets > Bank Accounts", 50000)]) is None


# ---------------------------------------------------------------- 3. the reason: Ollama or template

def test_reason_from_template_when_ollama_is_not_running(monkeypatch):
    monkeypatch.setattr(ai, "_ollama_state", {"checked": True, "up": False})
    found = {"confidence": 0.81, "source": "similarity", "evidence": ["name 86% alike “Total Cash & Bank”"]}
    assert ai.reason("HDFC Bank", "add to “Total Cash & Bank”", found) == \
        "add to “Total Cash & Bank”: name 86% alike “Total Cash & Bank” (from name similarity, 81% confidence)."


def test_reason_from_local_ollama(monkeypatch):
    monkeypatch.setattr(ai, "_ollama_state", {"checked": True, "up": True})
    sent = {}

    def fake_post(url, body):
        sent.update(url=url, body=body)
        return {"response": "HDFC Bank sits with the other bank accounts, so it belongs in Cash & Bank."}

    monkeypatch.setattr(ai, "_post", fake_post)
    found = {"confidence": 0.81, "source": "similarity", "evidence": ["same group"]}
    text = ai.reason("HDFC Bank", "add to “Total Cash & Bank”", found)
    assert text == "HDFC Bank sits with the other bank accounts, so it belongs in Cash & Bank. (81% confidence)"
    assert sent["url"].startswith("http://127.0.0.1:11434") and sent["body"]["stream"] is False


def test_only_a_model_on_this_machine_is_ever_called():
    assert ai._local("http://127.0.0.1:11434") and ai._local("http://localhost:11434")
    assert not ai._local("https://api.example.com")
    assert ai._post("https://api.example.com/v1/chat", {"x": 1}) is None  # refused before any network call


# ---------------------------------------------------------------- in the roll-forward review

@pytest.fixture(scope="module")
def draft(tmp_path_factory):
    folder = tmp_path_factory.mktemp("ai")
    base = json.loads(json.dumps(learn_workbook(VENUS).mapping()))
    out = folder / "draft.xlsx"
    facts = json.loads(json.dumps(roll_forward(VENUS, base, make_modified_venus_tb(folder / "tb.xlsx"), out)))
    result = learn_workbook(out)
    mapping = {**json.loads(json.dumps(result.mapping())), "rollforward": facts}
    report = json.loads(json.dumps(result.report()))
    covered = {s["cell"] for s in facts["suggestions"]}  # cells that already have a "link to TB" item
    targets = ai.suggest_for_unused(mapping, report, ai.client_history([base]), {}, covered)
    mapping["ai"] = {"targets": json.loads(json.dumps(targets))}
    return {"out": out, "mapping": mapping, "report": report, "targets": targets}


def test_new_ledgers_get_ai_suggestions(draft):
    hdfc = draft["targets"]["Current Assets > Bank Accounts > HDFC Bank"]
    assert (hdfc["cell"], hdfc["kind"], hdfc["source"]) == ("'BS Schedules'!C58", "add", "similarity")
    assert hdfc["confidence"] >= 0.8 and "Bank of Baroda" in hdfc["reason"]
    sb = draft["targets"]["Capital Account > Interest on SB Account"]
    assert (sb["cell"], sb["kind"]) == ("'BS Schedules'!C10", "replace")  # typed SB interest = the ledger
    assert "typed amount ₹5,101.88 equals the ledger" in sb["reason"]
    assert "Capital Account > T L Venkatesh" not in draft["targets"]  # not confident: stays red


def test_ai_suggestions_are_amber_and_never_auto_approved(draft):
    review = build_review(draft["out"], draft["mapping"], draft["mapping"], draft["report"])
    items = {i["id"]: i for i in review["items"]}
    hdfc = items["unused:Current Assets > Bank Accounts > HDFC Bank"]
    assert hdfc["tone"] == "amber" and hdfc["title"] == "HDFC Bank: AI suggestion"
    assert [a["label"] for a in hdfc["actions"]] == ["Approve", "Change", "Leave unused"]
    assert review["overrides"]["TB!A51"]["status"] == "a"
    assert "edits" not in draft["mapping"]  # nothing written until the CA approves

    approved = apply_action(draft["out"], draft["mapping"], hdfc, "approve", {}, NOW)
    assert approved["edits"]["BS Schedules"]["C58"]["formula"] == "=SUM(C55:C57)+TB!B51"
    assert approved["decisions"][hdfc["id"]]["ai_confidence"] == hdfc["change"]["suggestion"]["confidence"]

    sb = items["unused:Capital Account > Interest on SB Account"]
    approved = apply_action(draft["out"], approved, sb, "approve", {}, NOW)
    assert approved["edits"]["BS Schedules"]["C10"]["formula"] == "=TB!C16"  # replaces the typed amount

    left = apply_action(draft["out"], draft["mapping"], hdfc, "dismiss", {}, NOW)  # "Leave unused"
    assert "Current Assets > Bank Accounts > HDFC Bank" in left["ignored_ledgers"]
