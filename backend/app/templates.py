"""Default statement formats for BUILD MODE and the default mapping from Tally groups.

Two formats (CLAUDE.md):
  Non-corporate  ICAI Guidance Note on Financial Statements of Non-Corporate Entities
  Company        Schedule III to the Companies Act 2013, Division I

Each statement is a list of rows: ("heading", text), ("line", line id), ("total", text, [line ids]).
A ledger goes to a line by the first Tally group on its path (most specific first) found in
GROUP_LINES; ledgers under expense groups are refined by their name (EXPENSE_NAMES). A ledger
whose groups are not known (e.g. Suspense A/c, a primary group the client created) is left
unmapped and shows red in review.
"""
import re

COMPANY_TYPES = {"Private limited company"}

# Line id -> (label, side). side: "liability"/"income" take Cr - Dr; "asset"/"expense" take Dr - Cr.
LINES = {
    # Balance Sheet
    "capital": ("Owners' Capital Account", "liability"),
    "share_capital": ("Share Capital", "liability"),
    "reserves": ("Reserves and Surplus", "liability"),
    "long_term_borrowings": ("Long-term Borrowings", "liability"),
    "short_term_borrowings": ("Short-term Borrowings", "liability"),
    "trade_payables": ("Trade Payables", "liability"),
    "other_current_liabilities": ("Other Current Liabilities", "liability"),
    "short_term_provisions": ("Short-term Provisions", "liability"),
    "ppe": ("Property, Plant and Equipment", "asset"),
    "investments": ("Non-current Investments", "asset"),
    "inventories": ("Inventories", "asset"),
    "trade_receivables": ("Trade Receivables", "asset"),
    "cash_bank": ("Cash and Bank Balances", "asset"),
    "short_term_loans_advances": ("Short-term Loans and Advances", "asset"),
    "other_current_assets": ("Other Current Assets", "asset"),
    # Statement of Profit and Loss
    "revenue": ("Revenue from Operations", "income"),
    "other_income": ("Other Income", "income"),
    "purchases": ("Purchases of Stock-in-Trade", "expense"),
    "direct_expenses": ("Direct Expenses", "expense"),
    "inventory_change": ("Changes in Inventories of Stock-in-Trade", "expense"),
    "employee": ("Employee Benefits Expense", "expense"),
    "finance": ("Finance Costs", "expense"),
    "depreciation": ("Depreciation and Amortisation Expense", "expense"),
    "other_expenses": ("Other Expenses", "expense"),
    "tax": ("Tax Expense", "expense"),
}

CAPITAL_LABELS = {  # the owners' funds line reads differently by entity
    "Proprietorship": "Proprietor's Capital Account",
    "Partnership firm": "Partners' Capital Accounts",
    "LLP": "Partners' Contribution",
}

NON_CORPORATE = {
    "name": "Non-corporate (ICAI Guidance Note)",
    "balance_sheet": [
        ("heading", "I. OWNERS' FUNDS AND LIABILITIES"),
        ("heading", "(1) Owners' Funds"),
        ("line", "capital"), ("line", "reserves"),
        ("heading", "(2) Non-current Liabilities"),
        ("line", "long_term_borrowings"),
        ("heading", "(3) Current Liabilities"),
        ("line", "short_term_borrowings"), ("line", "trade_payables"),
        ("line", "other_current_liabilities"), ("line", "short_term_provisions"),
        ("total", "Total Owners' Funds and Liabilities", ["capital", "reserves", "long_term_borrowings",
                                                          "short_term_borrowings", "trade_payables",
                                                          "other_current_liabilities", "short_term_provisions"]),
        ("heading", "II. ASSETS"),
        ("heading", "(1) Non-current Assets"),
        ("line", "ppe"), ("line", "investments"),
        ("heading", "(2) Current Assets"),
        ("line", "inventories"), ("line", "trade_receivables"), ("line", "cash_bank"),
        ("line", "short_term_loans_advances"), ("line", "other_current_assets"),
        ("total", "Total Assets", ["ppe", "investments", "inventories", "trade_receivables", "cash_bank",
                                   "short_term_loans_advances", "other_current_assets"]),
    ],
    "profit_and_loss": [
        ("line", "revenue"), ("line", "other_income"),
        ("total", "Total Income", ["revenue", "other_income"]),
        ("heading", "Expenses"),
        ("line", "purchases"), ("line", "direct_expenses"), ("line", "inventory_change"), ("line", "employee"),
        ("line", "finance"), ("line", "depreciation"), ("line", "other_expenses"),
        ("total", "Total Expenses", ["purchases", "direct_expenses", "inventory_change", "employee", "finance",
                                     "depreciation", "other_expenses"]),
    ],
    "policies": [
        ("Basis of preparation", "The financial statements are prepared under the historical cost convention on the "
         "accrual basis of accounting, in accordance with the Accounting Standards issued by the Institute of "
         "Chartered Accountants of India as applicable to non-corporate entities, and the Guidance Note on "
         "Financial Statements of Non-Corporate Entities."),
        ("Use of estimates", "Preparing the financial statements requires estimates and assumptions that affect "
         "reported amounts. Actual results may differ; revisions are recognised in the period they are made."),
        ("Property, plant and equipment", "Stated at cost of acquisition less accumulated depreciation. Cost "
         "includes expenses directly attributable to bringing the asset to working condition."),
        ("Depreciation", "Provided on the written down value method at the rates prescribed under the Income-tax "
         "Act, 1961; half the rate for assets put to use for less than 180 days in the year."),
        ("Inventories", "Valued at the lower of cost and net realisable value."),
        ("Revenue recognition", "Sales are recognised when the significant risks and rewards of ownership pass to "
         "the buyer, net of GST and trade discounts."),
        ("Investments", "Long-term investments are stated at cost less any permanent diminution in value."),
    ],
}

COMPANY = {
    "name": "Company (Schedule III, Division I)",
    "balance_sheet": [
        ("heading", "I. EQUITY AND LIABILITIES"),
        ("heading", "(1) Shareholders' Funds"),
        ("line", "share_capital"), ("line", "reserves"),
        ("heading", "(2) Non-current Liabilities"),
        ("line", "long_term_borrowings"),
        ("heading", "(3) Current Liabilities"),
        ("line", "short_term_borrowings"), ("line", "trade_payables"),
        ("line", "other_current_liabilities"), ("line", "short_term_provisions"),
        ("total", "Total Equity and Liabilities", ["share_capital", "reserves", "long_term_borrowings",
                                                   "short_term_borrowings", "trade_payables",
                                                   "other_current_liabilities", "short_term_provisions"]),
        ("heading", "II. ASSETS"),
        ("heading", "(1) Non-current Assets"),
        ("line", "ppe"), ("line", "investments"),
        ("heading", "(2) Current Assets"),
        ("line", "inventories"), ("line", "trade_receivables"), ("line", "cash_bank"),
        ("line", "short_term_loans_advances"), ("line", "other_current_assets"),
        ("total", "Total Assets", ["ppe", "investments", "inventories", "trade_receivables", "cash_bank",
                                   "short_term_loans_advances", "other_current_assets"]),
    ],
    "profit_and_loss": [
        ("line", "revenue"), ("line", "other_income"),
        ("total", "Total Income", ["revenue", "other_income"]),
        ("heading", "Expenses"),
        ("line", "purchases"), ("line", "inventory_change"), ("line", "employee"), ("line", "finance"),
        ("line", "depreciation"), ("line", "other_expenses"),
        ("total", "Total Expenses", ["purchases", "inventory_change", "employee", "finance", "depreciation",
                                     "other_expenses"]),
    ],
    "policies": [
        ("Basis of preparation", "The financial statements are prepared under the historical cost convention on the "
         "accrual basis, in accordance with the Accounting Standards notified under section 133 of the Companies "
         "Act, 2013 and presented as required by Division I of Schedule III to the Act."),
        ("Use of estimates", "Preparing the financial statements requires estimates and assumptions that affect "
         "reported amounts. Actual results may differ; revisions are recognised in the period they are made."),
        ("Property, plant and equipment", "Stated at cost less accumulated depreciation and impairment, if any."),
        ("Depreciation", "Provided on the written down value method over the useful lives prescribed in Schedule II "
         "to the Companies Act, 2013."),
        ("Inventories", "Valued at the lower of cost and net realisable value."),
        ("Revenue recognition", "Revenue is recognised when the significant risks and rewards of ownership pass to "
         "the buyer, net of GST and trade discounts, as required by AS 9."),
        ("Taxes on income", "Current tax is provided on taxable income at the applicable rates. Deferred tax is "
         "recognised on timing differences as required by AS 22."),
    ],
}

# Tally's predefined groups (normalised names) -> line. Company-only differences in COMPANY_GROUPS.
GROUP_LINES = {
    "capital account": "capital",
    "reserves surplus": "reserves",
    "profit loss a c": "reserves",  # Tally's P&L A/c: profit of earlier years
    "secured loans": "long_term_borrowings",
    "unsecured loans": "long_term_borrowings",
    "bank od a c": "short_term_borrowings",
    "bank occ a c": "short_term_borrowings",
    "loans liability": "long_term_borrowings",
    "sundry creditors": "trade_payables",
    "duties taxes": "other_current_liabilities",
    "provisions": "short_term_provisions",
    "current liabilities": "other_current_liabilities",
    "fixed assets": "ppe",
    "investments": "investments",
    "opening stock": "inventory_change",
    "closing stock": "inventories",
    "stock in hand": "inventories",
    "sundry debtors": "trade_receivables",
    "bank accounts": "cash_bank",
    "cash in hand": "cash_bank",
    "deposits asset": "other_current_assets",
    "loans advances asset": "short_term_loans_advances",
    "current assets": "other_current_assets",
    "misc expenses asset": "other_current_assets",
    "sales accounts": "revenue",
    "direct incomes": "revenue",
    "indirect incomes": "other_income",
    "purchase accounts": "purchases",
    "direct expenses": "direct_expenses",
    "indirect expenses": "other_expenses",
}
COMPANY_GROUPS = {"capital account": "share_capital", "direct expenses": "other_expenses"}

# Within expense groups, a ledger's name can say where it belongs.
EXPENSE_NAMES = [
    (re.compile(r"salar|wage|bonus|staff|provident|\bp\.?f\b|\besi\b|gratuity|remuneration", re.I), "employee"),
    (re.compile(r"interest|bank charge|bank charger|loan processing", re.I), "finance"),
    (re.compile(r"depreciation|amorti", re.I), "depreciation"),
    (re.compile(r"income tax|corporation tax|tax provision|provision for tax", re.I), "tax"),
]
EXPENSE_LINES = {"direct_expenses", "other_expenses"}


def template_for(entity_type: str) -> dict:
    return COMPANY if entity_type in COMPANY_TYPES else NON_CORPORATE


def is_company(entity_type: str) -> bool:
    return entity_type in COMPANY_TYPES


def line_for(names_most_specific_first: list[str], entity_type: str) -> str | None:
    """The statement line for a TB row, from its own name and its groups; None = unmapped."""
    groups = {**GROUP_LINES, **(COMPANY_GROUPS if is_company(entity_type) else {})}
    for name in names_most_specific_first:
        line = groups.get(_norm(name))
        if line is None:
            continue
        if line in EXPENSE_LINES:
            for pattern, special in EXPENSE_NAMES:
                if pattern.search(names_most_specific_first[0]):
                    return special
        return line
    return None


def _norm(name: str) -> str:
    return re.sub(r"[^a-z0-9]+", " ", name.lower()).strip()
