"""AI mapping layer (CLAUDE.md Phase 7). Everything runs on this machine; nothing is sent out.

It answers two questions, always as a SUGGESTION (amber, with a reason and a confidence);
the CA approves or changes it, nothing is applied as final on its own:
  match_ledger    last year's ledger is not in the new TB: which new ledger is it now?
  suggest_target  a ledger no formula uses: which cell should take it?

Evidence is tried in this order:
  1. the client's own past mapping    CA decisions (aliases) and where a ledger was used before
  2. the firm-wide dictionary         where other clients' workbooks use a similarly named ledger,
     plus similarity                  rapidfuzz (spelling) and sentence-transformers (meaning)
  3. a local Ollama model, if one is running, words the reason; otherwise a template does.
"""
import os
import re
import threading
from collections import Counter
from difflib import SequenceMatcher
from urllib.parse import urlparse

from app.tally_tb import format_inr

MODEL = "sentence-transformers/all-MiniLM-L6-v2"  # ~90 MB, downloaded once, then used offline
THRESHOLD = 0.6  # below this confidence nothing is suggested
OLLAMA_URL = os.environ.get("LEDGERLINK_OLLAMA_URL", "http://127.0.0.1:11434")
OLLAMA_MODEL = os.environ.get("LEDGERLINK_OLLAMA_MODEL", "qwen2.5")
FILLER = re.compile(r"\b(a c|ac|account|accounts|acct|ledger)\b")

try:
    from rapidfuzz import fuzz
except ImportError:  # rapidfuzz not installed: fall back to the standard library
    fuzz = None


# ---------------------------------------------------------------- similarity

def norm(name: str) -> str:
    """"Salary  A/c" -> "salary"."""
    text = re.sub(r"[^a-z0-9]+", " ", (name or "").lower())
    return re.sub(r"\s+", " ", FILLER.sub(" ", text)).strip()


def spelling(a: str, b: str) -> float:
    a, b = norm(a), norm(b)
    if not a or not b:
        return 0.0
    if fuzz is not None:
        return fuzz.WRatio(a, b) / 100
    return SequenceMatcher(None, a, b).ratio()


class _Meaning:
    """sentence-transformers, loaded once and only from the local cache (never downloads)."""

    def __init__(self):
        self.model, self.tried, self.vectors = None, False, {}
        self.lock = threading.Lock()

    def available(self) -> bool:
        with self.lock:
            if not self.tried:
                self.tried = True
                if os.environ.get("LEDGERLINK_EMBEDDINGS", "1") != "0":
                    try:
                        from sentence_transformers import SentenceTransformer
                        self.model = SentenceTransformer(MODEL, device="cpu", local_files_only=True)
                    except Exception:  # not installed, or model not downloaded: spelling only
                        self.model = None
            return self.model is not None

    def similarity(self, a: str, b: str) -> float:
        if not self.available():
            return 0.0
        va, vb = self._vector(norm(a)), self._vector(norm(b))
        return max(0.0, float(va @ vb))

    def _vector(self, text: str):
        with self.lock:
            if text not in self.vectors:
                self.vectors[text] = self.model.encode(text or " ", normalize_embeddings=True)
            return self.vectors[text]


meaning = _Meaning()


def similarity(a: str, b: str) -> float:
    """0..1: the better of spelling (rapidfuzz) and meaning (sentence-transformers)."""
    return max(spelling(a, b), meaning.similarity(a, b))


# ---------------------------------------------------------------- 1. client's own past mapping

def client_history(mappings: list[dict]) -> dict:
    """From the client's earlier workbooks: CA's ledger choices, and the cell labels each
    ledger name was used in."""
    aliases, usage = {}, {}
    for mapping in mappings:
        aliases.update(mapping.get("aliases", {}))
        for used, label, _ in _ledger_labels(mapping):
            usage.setdefault(norm(used), Counter())[label] += 1
    return {"aliases": aliases, "usage": usage}


# ---------------------------------------------------------------- 2. firm-wide dictionary

def firm_dictionary(mappings: list[dict]) -> dict[str, Counter]:
    """{ledger name (normalised): Counter((cell label, main Tally group))} across the firm's
    other learned clients. The main group keeps a person's capital and a loan from the same
    person apart."""
    dictionary: dict[str, Counter] = {}
    for mapping in mappings:
        for used, label, group in _ledger_labels(mapping):
            dictionary.setdefault(norm(used), Counter())[(label, group)] += 1
    return dictionary


def _ledger_labels(mapping: dict):
    """(ledger name, label of a cell it feeds, main group) for every TB link in a stored mapping."""
    cells = mapping.get("cells", {})
    for ledger in mapping.get("ledgers", {}).values():
        for use in ledger.get("used_in", []):
            label = cells.get(use["cell"], {}).get("label")
            if label:
                yield ledger["name"], label, _main_group(ledger.get("group_path", ""))


def _main_group(group_path: str) -> str:
    return (group_path or "").split(" > ")[0]


# ---------------------------------------------------------------- match a missing ledger

def match_ledger(old: dict, candidates: list[dict], history: dict | None = None) -> dict | None:
    """old / candidates: {"key", "name", "group", "amount"}. Best new ledger for an old one."""
    history = history or {"aliases": {}}
    alias = history["aliases"].get(old["key"])
    for cand in candidates:
        if alias and cand["key"] == alias:
            return _found(cand["key"], 0.97, "client", [f"you mapped “{old['name']}” to “{cand['name']}” before"])

    best = None
    for cand in candidates:
        name = similarity(old["name"], cand["name"])
        group = 1.0 if cand["group"] == old["group"] else 0.5 * similarity(old["group"], cand["group"])
        amount = _closeness(old["amount"], cand["amount"])
        score = 0.6 * name + 0.25 * group + 0.15 * amount
        if best is None or score > best[0]:
            evidence = [f"name {name:.0%} alike (“{old['name']}” / “{cand['name']}”)"]
            evidence.append("same group" if group == 1.0 else f"group {cand['group'] or 'top level'}")
            if amount > 0.9:
                evidence.append(f"balance close (₹{format_inr(old['amount'])} last year, ₹{format_inr(cand['amount'])} now)")
            best = (score, cand, evidence)
    if best and best[0] >= THRESHOLD:
        return _found(best[1]["key"], best[0], "similarity", best[2])
    return None


def _closeness(a: float, b: float) -> float:
    a, b = abs(a), abs(b)
    return 1.0 if a == b else max(0.0, 1 - abs(a - b) / max(a, b, 1.0))


def _found(key, confidence, source, evidence) -> dict:
    return {"new": key, "confidence": round(min(confidence, 0.99), 2), "source": source, "evidence": evidence}


# ---------------------------------------------------------------- a place for an unused ledger

def target_candidates(mapping: dict) -> list[dict]:
    """Cells an unused ledger could go to, from a learnt mapping:
    add     a TB-linked cell this year; the ledger is added to its note total when it has one
    replace a typed amount this year; the TB link would replace it"""
    cells, ledgers, fy = mapping["cells"], mapping["ledgers"], mapping.get("financial_year")
    totals = _note_totals(mapping)
    result = []
    for key, cell in cells.items():
        if cell["financial_year"] not in (fy, None) or not cell["label"]:
            continue
        if cell["category"] == "tb-linked":
            linked = [ledgers[l["ledger"]] for l in cell["links"] if l["ledger"] in ledgers]
            total = totals.get(key)
            result.append({"cell": total or key, "label": cells[total]["label"] if total else cell["label"],
                           "via": cell["label"], "kind": "add", "value": cell["value"],
                           "ledgers": [{"name": l["name"], "group": l["group_path"]} for l in linked]})
        elif cell["category"] == "manual" and cell["value"] and "rate" not in cell["label"].lower():
            result.append({"cell": key, "label": cell["label"], "via": cell["label"], "kind": "replace",
                           "value": cell["value"], "ledgers": []})
    return result


def _note_totals(mapping: dict) -> dict[str, str]:
    """TB-linked cell -> the nearest "Total ..." cell below it on the same sheet that adds it up."""
    cells, graph = mapping["cells"], mapping.get("precedents", {})
    totals: dict[str, tuple[int, str]] = {}
    for key, reads in graph.items():
        if not ((cells.get(key) or {}).get("label") or "").strip().lower().startswith("total"):
            continue
        sheet, row = key.rsplit("!", 1)[0], _row(key)
        for member in reads:
            if member.rsplit("!", 1)[0] == sheet and _row(member) < row:
                distance = row - _row(member)
                if member not in totals or distance < totals[member][0]:
                    totals[member] = (distance, key)
    return {member: key for member, (_, key) in totals.items()}


def _row(key: str) -> int:
    return int(re.search(r"(\d+)$", key)[1])


def suggest_target(ledger: dict, candidates: list[dict], history: dict | None = None,
                   dictionary: dict | None = None) -> dict | None:
    """ledger: {"key", "name", "group", "amount"}. The cell it most likely belongs in."""
    history, dictionary = history or {"usage": {}}, dictionary or {}
    known = history["usage"].get(norm(ledger["name"]), Counter())
    similar_names = sorted(((similarity(ledger["name"], name), name) for name in dictionary), reverse=True)[:3]

    best = None
    for cand in candidates:
        options = []
        for label, _ in known.most_common(3):  # 1. the client's own past mapping
            s = similarity(label, cand["label"])
            options.append((0.55 + 0.4 * s, "client", [f"used in “{label}” in an earlier year"]))
        for s_name, name in similar_names:  # 2. the firm-wide dictionary
            if s_name < 0.75:
                continue
            for (label, group), count in dictionary[name].most_common(2):
                if group != _main_group(ledger["group"]):
                    continue  # same name in another kind of group (e.g. capital vs a loan)
                s = similarity(label, cand["label"])
                options.append((0.5 * s_name + 0.45 * s, "firm",
                                [f"other clients put “{name}” in “{label}” ({count}×)"]))
        # 2. similarity with this workbook: the cell's own label and the ledgers it already reads
        # ledgers already there count only from the same main group: the proprietor's capital
        # "T L Venkatesh" must not follow a loan from "T.L.Venkatesh (H.U.F)"
        main = _main_group(ledger["group"])
        same_kind = [l for l in cand["ledgers"] if _main_group(l["group"]) == main]
        names = [cand["label"], cand["via"]] + [l["name"] for l in same_kind]
        name_sim, closest = max((similarity(ledger["name"], n), n) for n in names)
        groups = [l["group"] for l in cand["ledgers"] if l["group"]]
        if not ledger["group"]:
            group = 0.0  # a top-level line (e.g. Profit & Loss A/c) shares no group with anything
        elif ledger["group"] in groups:
            group = 1.0
        else:
            group = max((0.8 * similarity(ledger["group"], g) for g in groups), default=0.0)
        amount = 0.3 if cand["kind"] == "replace" and abs(cand["value"] - ledger["amount"]) < 0.005 else 0.0
        evidence = [f"name {name_sim:.0%} alike “{closest}”"]
        if group == 1.0:
            siblings = ", ".join(l["name"] for l in cand["ledgers"] if l["group"] == ledger["group"])
            evidence.append(f"same group ({ledger['group']}) as {siblings}")
        elif group:
            evidence.append("similar group")
        if amount:
            evidence.append(f"typed amount ₹{format_inr(cand['value'])} equals the ledger")
        options.append((0.65 * name_sim + 0.25 * group + amount, "similarity", evidence))

        for score, source, why in options:
            if best is None or score > best[0]:
                best = (score, source, why, cand)
    if best is None or best[0] < THRESHOLD:
        return None
    score, source, why, cand = best
    return {"cell": cand["cell"], "label": cand["label"], "kind": cand["kind"], "confidence": round(min(score, 0.99), 2),
            "source": source, "evidence": why}


# ---------------------------------------------------------------- 3. the reason in plain words

SOURCES = {"client": "your own mapping", "firm": "the firm's other clients", "similarity": "name similarity"}


def reason(subject: str, suggestion: str, found: dict) -> str:
    """One plain sentence. From the local Ollama model when it runs, else from a template."""
    facts = "; ".join(found["evidence"])
    text = _ask_ollama(subject, suggestion, facts, found["confidence"]) if ollama_running() else None
    return text or f"{suggestion}: {facts} (from {SOURCES[found['source']]}, {found['confidence']:.0%} confidence)."


_ollama_state = {"checked": False, "up": False}


def ollama_running() -> bool:
    """Is a local Ollama answering? Checked once per process. Only localhost is ever used."""
    if not _ollama_state["checked"]:
        _ollama_state["checked"] = True
        _ollama_state["up"] = _local(OLLAMA_URL) and _get(f"{OLLAMA_URL}/api/tags") is not None
    return _ollama_state["up"]


def _ask_ollama(subject: str, suggestion: str, facts: str, confidence: float) -> str | None:
    prompt = (f"You help a chartered accountant review a mapping of trial balance ledgers. In one short sentence "
              f"of plain English, say why this is suggested. Do not add facts.\nLedger: {subject}\n"
              f"Suggestion: {suggestion}\nFacts: {facts}\nConfidence: {confidence:.0%}")
    answer = _post(f"{OLLAMA_URL}/api/generate", {"model": OLLAMA_MODEL, "prompt": prompt, "stream": False,
                                                  "options": {"temperature": 0}})
    text = (answer or {}).get("response", "").strip()
    return f"{text} ({confidence:.0%} confidence)" if text else None


def _local(url: str) -> bool:
    """Client data must never leave the machine: only talk to a model on this computer."""
    return urlparse(url).hostname in ("127.0.0.1", "localhost", "::1")


def _get(url: str):
    try:
        import httpx
        response = httpx.get(url, timeout=1.0)
        return response.json() if response.status_code == 200 else None
    except Exception:
        return None


def _post(url: str, body: dict):
    if not _local(url):
        return None
    try:
        import httpx
        response = httpx.post(url, json=body, timeout=30.0)
        return response.json() if response.status_code == 200 else None
    except Exception:
        return None


# ---------------------------------------------------------------- for the review

def suggest_for_unused(mapping: dict, report: dict, history: dict, dictionary: dict,
                       covered: set[str] = frozenset()) -> dict[str, dict]:
    """{ledger key: suggestion with its reason} for every TB ledger no formula uses.
    covered: cells that already have an exact "link to TB" item, so they are left to it."""
    candidates = [c for c in target_candidates(mapping) if "opening" not in c["label"].lower()]
    suggestions = {}
    for unused in report["unused_ledgers"]:
        if any(cell in covered for cell in unused["typed_in"]):
            continue
        ledger = {"key": unused["ledger"], "name": unused["name"], "group": unused["group_path"],
                  "amount": unused["dr"] or unused["cr"]}
        found = suggest_target(ledger, candidates, history, dictionary)
        if found:
            action = "replace the typed amount in" if found["kind"] == "replace" else "add to"
            found["reason"] = reason(unused["name"], f"{action} “{found['label']}”", found)
            suggestions[unused["ledger"]] = found
    return suggestions


def describe() -> str:
    """Which parts of the AI layer are working, for the progress log."""
    parts = ["rapidfuzz" if fuzz is not None else "difflib"]
    if meaning.available():
        parts.append("sentence-transformers")
    words = f"Ollama ({OLLAMA_MODEL})" if ollama_running() else "template reasons (Ollama not running)"
    return f"AI matching: {' + '.join(parts)}; {words}."
