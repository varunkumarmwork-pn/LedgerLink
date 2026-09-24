"""The Review panel: items the CA must clear, the checks, and the colour of each cell.

Everything is worked out afresh from the workbook as the CA has left it, so edits and
decisions show at once. Decisions (Approve / Change) are kept in the mapping:
  decisions        {item id: {"action", "at", ...}}   cleared items
  ignored_ledgers  [ledger key]                         "leave unused"; carried to next year
  aliases          {last year's key: this year's key}   CA's choice of ledger; next year repeats it
The year can be finalised only when no item is open.
"""
import re
from datetime import datetime
from pathlib import Path

import openpyxl

from app.formula import cell_key
from app.learn import PRIOR_YEAR, TB_LINKED
from app.rollforward import MISSING, MOVED, OPENING_LABEL, RENAMED, link_suggestions, render_template
from app.tally_tb import TrialBalance, format_inr, parse_trial_balance
from app.workbook_view import apply_edits

TOTAL_LIABILITIES = re.compile(r"^total\b.*(liabilit|equity)", re.I)
TOTAL_ASSETS = re.compile(r"^total\s+assets", re.I)
BALANCE_SHEET = re.compile(r"balance\s*sheet|^bs$", re.I)


def build_review(path: str | Path, stored: dict, live_mapping: dict, live_report: dict) -> dict:
    """{"items": open items, "checks": [...], "overrides": {cell: {"status", "reason"}}}

    stored: the mapping kept for this workbook (roll-forward facts, decisions, edits).
    live_mapping / live_report: learn mode run on the workbook with the CA's edits.
    """
    tb = parse_trial_balance(path)
    ctx = _Context(tb, stored, live_mapping, live_report)
    items = (_ledger_items(ctx) + _unused_items(ctx) + _issue_items(ctx) + _link_items(ctx)
             + _stock_items(ctx) + _manual_items(ctx) + _ppe_items(ctx))
    checks = _checks(ctx)
    for check in checks:
        if not check["ok"] and check["item"]:
            items.append(_item(f"check:{check['id']}", "red", check["label"], "Check", None,
                               check["detail"], [("Accept", "approve")], reason=check["label"]))

    decisions, edits = stored.get("decisions", {}), stored.get("edits", {})
    overrides = {c["cell"]: {"status": "g", "reason": f"carried forward from last year's closing ({c['from']})"}
                 for c in (ctx.facts or {}).get("carried", [])}
    for key in ctx.ignored & set(ctx.rows):  # left unused on purpose, this year or an earlier one
        for cell in _tb_row_cells(ctx, key):
            overrides[cell] = {"status": None, "reason": "left unused by CA"}
    open_items = []
    for item in items:
        edited = item.get("change", {}).get("kind") == "value" and _is_edited(item["cell"], edits)
        if item["id"] in decisions or edited:
            left_unused = item["id"].startswith("unused:")
            for cell in item["cells"]:
                overrides[cell] = {"status": None if left_unused else "g",
                                   "reason": "left unused by CA" if left_unused else "approved by CA"}
        else:
            open_items.append(item)
            for cell in item["cells"]:
                overrides[cell] = {"status": "r" if item["tone"] == "red" else "a", "reason": item["reason"]}
    return {"items": open_items, "checks": checks, "overrides": overrides}


class _Context:
    def __init__(self, tb: TrialBalance, stored: dict, live_mapping: dict, live_report: dict):
        self.tb, self.stored, self.mapping, self.report = tb, stored, live_mapping, live_report
        self.facts = stored.get("rollforward")
        self.fy = live_mapping.get("financial_year")
        self.rows = {r.key: r for r in tb.rows}
        self.ignored = set(stored.get("ignored_ledgers", []))
        self.aliases = stored.get("aliases", {})
        self.ai_targets = stored.get("ai", {}).get("targets", {})  # AI suggestions for unused ledgers

    def resolve(self, old_key: str) -> str | None:
        """This year's ledger for last year's key (CA's choice first, then the automatic match)."""
        if old_key in self.aliases:
            return self.aliases[old_key]
        if self.facts:
            return self.facts["matches"].get(old_key, {}).get("new")
        return old_key


def _item(item_id, tone, title, ref, cell, text, actions, cells=None, change=None, reason=None) -> dict:
    return {"id": item_id, "tone": tone, "title": title, "ref": ref, "cell": cell, "text": text,
            "actions": [{"label": label, "action": action} for label, action in actions],
            "cells": cells if cells is not None else ([cell] if cell else []),
            "change": change or {}, "reason": reason or title}


# ---------------------------------------------------------------- items

def _ledger_items(ctx: _Context) -> list[dict]:
    """Roll-forward: ledgers found under a new name or group (amber) or not found (red)."""
    if not ctx.facts:
        return []
    items = []
    for old, match in ctx.facts["matches"].items():
        cells = [cell for cell, template in ctx.facts["templates"].items() if f"{{TB:{old}|" in template]
        if match["how"] not in (RENAMED, MOVED, MISSING) or not cells:
            continue
        old_name = old.split(" > ")[-1]
        refs = ", ".join(_where(c) for c in cells)
        if match["how"] == MISSING:
            items.append(_item(f"ledger:{old}", "red", f"{old_name} not in this year's TB", _short(cells[0]), cells[0],
                               f"Used last year in {refs}; those cells now read 0. Map it to a ledger in the "
                               f"FY {ctx.fy} TB, or accept 0.",
                               [("Map to ledger", "change"), ("Accept 0", "approve")], cells,
                               {"kind": "ledger", "old": old}, f"{old_name} is not in this year's TB"))
        else:
            new_name = match["new"].split(" > ")[-1]
            if match["how"] == RENAMED:
                how = f"renamed from “{old_name}”"
                why = match.get("reason") or f"{match['score']:.0%} alike"
                text = f"Looks {how}. {why} Linked in {refs}."
            else:
                how = f"moved from {' > '.join(old.split(' > ')[:-1]) or 'the top level'}"
                text = f"Looks {how}. Linked in {refs}."
            items.append(_item(f"ledger:{old}", "amber", new_name, _short(cells[0]), cells[0], text,
                               [("Approve", "approve"), ("Change", "change")], cells,
                               {"kind": "ledger", "old": old}, f"{new_name}: {how}"))
    return items


def _unused_items(ctx: _Context) -> list[dict]:
    new = set((ctx.facts or {}).get("new_ledgers", []))
    unmapped = set(ctx.stored.get("build", {}).get("unmapped", []))  # BUILD MODE: no default line
    items = []
    for u in ctx.report["unused_ledgers"]:
        key = u["ledger"]
        if key in ctx.ignored or key not in ctx.rows:
            continue
        amount = f"₹{format_inr(u['dr'])} debit" if u["dr"] else f"₹{format_inr(u['cr'])} credit"
        where = f" under {u['group_path']}" if u["group_path"] else ""
        text = f"{amount}{where} does not reach any schedule."
        if key in new:
            text += " New in this year's TB."
        if key in unmapped:
            text += " Its Tally group has no line in the default format."
        if u["typed_in"]:
            text += f" The amount is typed into {', '.join(_where(c) for c in u['typed_in'])}."
        tb_cells = _tb_row_cells(ctx, key)
        change = {"kind": "cell", "ledger": key, "options": _target_cells(ctx, key)}
        suggestion = ctx.ai_targets.get(key)
        if suggestion:  # AI suggestion: amber, never applied until the CA approves it
            change["suggestion"] = suggestion
            items.append(_item(f"unused:{key}", "amber", f"{u['name']}: AI suggestion", f"TB row {ctx.rows[key].row}",
                               tb_cells[0], f"{text} Suggested: {suggestion['reason']}",
                               [("Approve", "approve"), ("Change", "change"), ("Leave unused", "dismiss")], tb_cells,
                               change, f"AI suggestion ({suggestion['confidence']:.0%}): {suggestion['label']}"))
        else:
            items.append(_item(f"unused:{key}", "red", f"{u['name']} not used", f"TB row {ctx.rows[key].row}",
                               tb_cells[0], text, [("Map ledger", "change"), ("Leave unused", "dismiss")], tb_cells,
                               change, f"{u['name']} is not used in any schedule"))
    return items


def _tb_row_cells(ctx: _Context, key: str) -> list[str]:
    """Name, Dr and Cr cells of a ledger's row in the TB sheet."""
    row = ctx.rows[key].row
    return [cell_key(ctx.tb.sheet_name, row, _col(c)) for c in (ctx.tb.name_column, ctx.tb.debit_column,
                                                                ctx.tb.credit_column)]


def _target_cells(ctx: _Context, ledger: str) -> list[dict]:
    """Cells a ledger could be added to: this year's TB-linked cells in the same main group."""
    group = ledger.split(" > ")[0]
    linked = {k: c for k, c in ctx.mapping["cells"].items() if c["category"] == TB_LINKED}
    same = {k: c for k, c in linked.items() if any(l["ledger"].split(" > ")[0] == group for l in c["links"])}
    return [{"cell": k, "label": f"{_sheet(k)} {_short(k)} · {c['label']}"} for k, c in (same or linked).items()]


def _issue_items(ctx: _Context) -> list[dict]:
    """Current year and last year read different lines (e.g. Venus Balance Sheet row 21)."""
    issues = list(ctx.report["inconsistencies"])
    if ctx.facts:
        issues += [dict(i, message="Last year: " + i["message"]) for i in ctx.facts.get("last_year_issues", [])]
    return [_item(f"issue:{i['cell']}", "red", i["label"] or _short(i["cell"]), _short(i["cell"]), i["cell"],
                  i["message"], [("Keep as is", "approve"), ("Change", "change")], change={"kind": "value"},
                  reason="this year and last year read different lines")
            for i in issues]


def _link_items(ctx: _Context) -> list[dict]:
    """Typed amounts that equal TB ledgers: offer the linked formula."""
    if ctx.facts:
        suggestions = ctx.facts.get("suggestions", [])
    else:
        suggestions = link_suggestions(ctx.report["manual_cells"], ctx.tb)
    items = []
    for s in suggestions:
        formula, missing = render_template(s["template"], ctx.tb, ctx.tb.sheet_name, ctx.resolve)
        if missing:
            continue
        names = ", ".join(k.split(" > ")[-1] for k in s["ledgers"])
        typed = s["formula"] or format_inr(s["value"])
        items.append(_item(f"link:{s['cell']}", "amber", s["label"], _short(s["cell"]), s["cell"],
                           f"Typed as {typed}, which equals TB {names}. Link it so it updates with the TB: {formula}",
                           [("Link to TB", "approve"), ("Keep typed", "dismiss")],
                           change={"kind": "link", "formula": formula}, reason=f"typed; equals TB {names}"))
    return items


def _stock_items(ctx: _Context) -> list[dict]:
    """Closing stock taken from the TB together with other ledgers (Venus: + Inventory Difference)."""
    memo = {r.key for r in ctx.tb.rows if r.memo}
    items = []
    for key, cell in ctx.mapping["cells"].items():
        ledgers = list(dict.fromkeys(l["ledger"] for l in cell["links"]))
        others = [l for l in ledgers if l not in memo and l.split(" > ")[-1].lower() != "opening stock"]
        if cell["category"] != TB_LINKED or not any(l in memo for l in ledgers) or not others:
            continue  # closing stock alone, or with opening stock (changes in inventories), is normal
        parts = [f"{k.split(' > ')[-1]} ₹{format_inr(_amount(ctx.rows[k]))}" for k in ledgers if k in ctx.rows]
        items.append(_item(f"stock:{key}", "amber", cell["label"] or _short(key), _short(key), key,
                           f"{' plus '.join(parts)} from the TB. Confirm the difference belongs in stock.",
                           [("Approve", "approve"), ("Change", "change")], change={"kind": "value"},
                           reason="closing stock plus other ledgers; confirm"))
    return items


def _manual_items(ctx: _Context) -> list[dict]:
    """Roll-forward: amounts typed last year, kept as they were; the CA confirms or changes them."""
    if not ctx.facts:
        return []
    linked = {s["cell"] for s in ctx.facts.get("suggestions", [])}
    return [_item(f"manual:{m['cell']}", "red", m["label"] or _short(m["cell"]), _short(m["cell"]), m["cell"],
                  f"Typed last year: ₹{format_inr(m['value'])}{' (' + m['formula'] + ')' if m['formula'] else ''}. "
                  f"Enter the FY {ctx.fy} figure, or approve to keep it.",
                  [("Approve", "approve"), ("Change", "change")], change={"kind": "value"},
                  reason="typed last year; enter this year's figure")
            for m in ctx.facts.get("manual", []) if m["cell"] not in linked]


def _ppe_items(ctx: _Context) -> list[dict]:
    if not ctx.facts:
        return []
    return [_item(f"ppe:{p['sheet']}", "red", f"{p['sheet']}: additions for FY {ctx.fy}", p["sheet"], p["reset"][0],
                  "Opening WDV is carried forward from last year's closing. Additions and deletions were reset "
                  "to 0: enter this year's from the fixed asset ledgers, then mark done.",
                  [("Done", "approve")], p["reset"], reason="enter this year's additions and deletions")
            for p in ctx.facts.get("ppe", []) if p["reset"]]


# ---------------------------------------------------------------- checks

def _checks(ctx: _Context) -> list[dict]:
    """CLAUDE.md: TB Dr = Cr, BS difference = 0, every TB ledger used, opening capital = last
    closing; plus opening stock = last year's closing stock."""
    tb = ctx.tb
    checks = [_check("tb", tb.total_dr == tb.total_cr, "Trial balance tallies",
                     f"Debit ₹{format_inr(tb.total_dr)} and credit ₹{format_inr(tb.total_cr)} do not match.",
                     "Trial balance does not tally")]

    liabilities, assets = _find_total(ctx, TOTAL_LIABILITIES), _find_total(ctx, TOTAL_ASSETS)
    if liabilities is not None and assets is not None:
        diff = liabilities - assets
        checks.append(_check("bs", abs(diff) < 0.005, f"Balance sheet tallies, difference ₹{format_inr(abs(diff))}",
                             f"Total capital & liabilities ₹{format_inr(liabilities)} but total assets "
                             f"₹{format_inr(assets)}.", f"Balance sheet difference ₹{format_inr(abs(diff))}"))

    unused = [u for u in ctx.report["unused_ledgers"] if u["ledger"] not in ctx.ignored]
    word = "ledger" if len(unused) == 1 else "ledgers"
    checks.append(_check("unused", not unused, "Every TB ledger is used",
                         f"{len(unused)} {word} from the TB {'is' if len(unused) == 1 else 'are'} not used.",
                         f"{len(unused)} {word} from the TB {'is' if len(unused) == 1 else 'are'} not used",
                         item=False))

    capital = _opening_capital(ctx)
    if capital:
        opening, last_closing = capital
        diff = opening - last_closing
        checks.append(_check("capital", abs(diff) < 0.005, "Opening capital equals last year's closing",
                             f"Opening ₹{format_inr(opening)}, last year's closing ₹{format_inr(last_closing)}.",
                             f"Opening capital differs from last year's closing by ₹{format_inr(abs(diff))}"))

    stock = _opening_stock(ctx)
    if stock:
        opening, last_closing = stock
        diff = opening - last_closing
        checks.append(_check("stock", abs(diff) < 0.005, "Opening stock equals last year's closing stock",
                             f"TB opening stock ₹{format_inr(opening)}, last year's closing ₹{format_inr(last_closing)}.",
                             f"Opening stock differs from last year's closing by ₹{format_inr(abs(diff))}"))
    return checks


def _check(check_id, ok, ok_label, detail, bad_label, item=True) -> dict:
    """item=False when the check already has its own review items (unused ledgers)."""
    return {"id": check_id, "ok": bool(ok), "label": ok_label if ok else bad_label,
            "detail": "" if ok else detail, "item": item}


def _find_total(ctx: _Context, label: re.Pattern) -> float | None:
    """This year's figure on the Balance Sheet row whose label matches (e.g. "Total Assets")."""
    for key, cell in ctx.mapping["cells"].items():
        if BALANCE_SHEET.search(_sheet(key)) and label.search(cell["label"] or "") and cell["financial_year"] == ctx.fy:
            return cell["value"]
    return None


def _opening_capital(ctx: _Context) -> tuple[float, float] | None:
    """("As per last Balance Sheet" this year, last year's closing on the row it leads to)."""
    cells, graph = ctx.mapping["cells"], ctx.mapping.get("precedents", {})
    for key, cell in cells.items():
        if cell["financial_year"] != ctx.fy or not OPENING_LABEL.search(cell["label"] or ""):
            continue
        closing = next((k for k in graph if key in graph[k] and _sheet(k) == _sheet(key)
                        and _column(k) == _column(key) and _row(k) > _row(key)), None)
        prior = _prior_cell(ctx, closing) if closing else None
        if prior in cells:
            return cell["value"], cells[prior]["value"]
    return None


def _opening_stock(ctx: _Context) -> tuple[float, float] | None:
    row = next((r for r in ctx.tb.rows if r.name.strip().lower() == "opening stock"), None)
    if row is None:
        return None
    last = (ctx.facts or {}).get("last_closing_stock")
    if last:
        return float(row.dr), last["amount"]
    memo = {r.key for r in ctx.tb.rows if r.memo}
    for key, cell in ctx.mapping["cells"].items():
        if cell["category"] == TB_LINKED and any(l["ledger"] in memo for l in cell["links"]):
            prior = _prior_cell(ctx, key)
            if prior in ctx.mapping["cells"] and ctx.mapping["cells"][prior]["category"] == PRIOR_YEAR:
                return float(row.dr), ctx.mapping["cells"][prior]["value"]
    return None


def _prior_cell(ctx: _Context, key: str) -> str | None:
    """The prior-year cell on the same row as a current-year cell."""
    years = ctx.mapping.get("year_columns", {}).get(_sheet(key), {})
    this_year = int(ctx.fy[:4])
    current = sorted(int(c) for c, y in years.items() if y == this_year)
    prior = sorted(int(c) for c, y in years.items() if y == this_year - 1)
    pairs = dict(zip(current, prior))
    col = pairs.get(_column(key))
    return cell_key(_sheet(key), _row(key), col) if col else None


# ---------------------------------------------------------------- Approve / Change

def apply_action(path: str | Path, stored: dict, item: dict, action: str, payload: dict, now: datetime) -> dict:
    """Record the CA's decision on a review item; returns the new stored mapping."""
    change = item["change"]
    decision = {"action": action, "at": now.isoformat(timespec="seconds")}
    mapping = dict(stored)

    if action == "approve" and change.get("kind") == "link":
        mapping = apply_edits(path, mapping, [_edit(item["cell"], formula=change["formula"])], now)
    elif action == "approve" and change.get("suggestion"):
        found = change["suggestion"]
        mapping = _add_ledger(path, mapping, change["ledger"], found["cell"], now, replace=found["kind"] == "replace")
        decision.update({"cell": found["cell"], "ai_confidence": found["confidence"]})
    elif action in ("approve", "dismiss") and item["id"].startswith("unused:"):
        mapping["ignored_ledgers"] = sorted(set(mapping.get("ignored_ledgers", [])) | {change["ledger"]})
    elif action == "change":
        mapping = _apply_change(path, mapping, item, payload, now)
        decision.update({k: v for k, v in payload.items() if k in ("ledger", "cell", "formula", "value")})
    elif action not in ("approve", "dismiss"):
        raise ValueError(f"Unknown action {action!r}.")
    mapping["decisions"] = {**mapping.get("decisions", {}), item["id"]: decision}
    return mapping


def _apply_change(path, mapping: dict, item: dict, payload: dict, now: datetime) -> dict:
    kind = item["change"].get("kind")
    tb = parse_trial_balance(path)
    rows = {r.key: r for r in tb.rows}
    if kind == "ledger":
        new_key = payload.get("ledger")
        if new_key not in rows:
            raise ValueError("Choose a ledger from this year's TB.")
        old = item["change"]["old"]
        facts = mapping["rollforward"]
        aliases = {**mapping.get("aliases", {}), old: new_key}
        resolve = lambda k: aliases.get(k) or facts["matches"].get(k, {}).get("new")  # noqa: E731
        changes = [_edit(cell, formula=render_template(facts["templates"][cell], tb, tb.sheet_name, resolve)[0])
                   for cell in item["cells"] if cell in facts["templates"]]
        mapping = apply_edits(path, mapping, changes, now)
        mapping["aliases"] = aliases
        return mapping
    if kind == "cell":
        target = payload.get("cell")
        if target not in {o["cell"] for o in item["change"]["options"]}:
            raise ValueError("Choose one of the offered cells.")
        return _add_ledger(path, mapping, item["change"]["ledger"], target, now)
    if kind == "value":
        formula, value = payload.get("formula"), payload.get("value")
        if formula is None and value is None:
            raise ValueError("Enter a value or a formula.")
        return apply_edits(path, mapping, [_edit(item["cell"], formula=formula, value=value)], now)
    raise ValueError("This item cannot be changed here; edit the cells in the sheet instead.")


def _add_ledger(path, mapping: dict, ledger: str, target: str, now: datetime, replace: bool = False) -> dict:
    """Link a TB ledger into a cell: added to what the cell has, or replacing a typed amount."""
    tb = parse_trial_balance(path)
    row = next(r for r in tb.rows if r.key == ledger)
    column = tb.debit_column if row.dr >= row.cr else tb.credit_column
    sheet = tb.sheet_name if tb.sheet_name.replace("_", "").isalnum() else f"'{tb.sheet_name}'"
    ref = f"{sheet}!{column}{row.row}"
    if replace:
        return apply_edits(path, mapping, [_edit(target, formula=f"={ref}")], now)
    current = _current_content(path, mapping, target)
    base = current if isinstance(current, str) and current.startswith("=") else f"={current or 0}"
    return apply_edits(path, mapping, [_edit(target, formula=f"{base}+{ref}")], now)


def _current_content(path, mapping: dict, key: str):
    sheet, address = _split(key)
    edit = mapping.get("edits", {}).get(sheet, {}).get(address)
    if edit:
        return edit.get("formula") or edit.get("value")
    return openpyxl.load_workbook(path)[sheet][address].value


def _edit(key: str, formula=None, value=None) -> dict:
    sheet, address = _split(key)
    return {"sheet": sheet, "cell": address, "formula": formula, "value": value}


def _is_edited(key: str | None, edits: dict) -> bool:
    if not key:
        return False
    sheet, address = _split(key)
    return address in edits.get(sheet, {})


# ---------------------------------------------------------------- helpers

def _amount(row) -> float:
    return float(row.dr or row.cr)


def _split(key: str) -> tuple[str, str]:
    sheet, address = key.rsplit("!", 1)
    return (sheet[1:-1].replace("''", "'") if sheet.startswith("'") else sheet), address


def _sheet(key: str) -> str:
    return _split(key)[0]


def _short(key: str) -> str:
    return _split(key)[1]


def _where(key: str) -> str:
    """"'BS Schedules'!C14" -> "BS Schedules C14"."""
    sheet, address = _split(key)
    return f"{sheet} {address}"


def _row(key: str) -> int:
    return int(re.search(r"\d+$", key)[0])


def _column(key: str) -> int:
    return _col(re.search(r"!\$?([A-Z]+)", key)[1])


def _col(letters: str) -> int:
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n
