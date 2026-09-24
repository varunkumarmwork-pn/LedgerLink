import pytest

from app.formula import FormulaError, Ref, cell_key, constants, evaluate, parse_ref

CELLS = {"A1": 2, "A2": 3, "B1:B4": [1, 2, "text", None], "C1": "Car"}


def lookup(text):
    return CELLS.get(text)


@pytest.mark.parametrize("formula, expected", [
    ("=1+2*3", 7),
    ("=10-4-3", 3),
    ("=-2^2", 4),  # Excel applies the minus sign first
    ("=(5101.88)", 5101.88),
    ("=SUM(B1:B4)+A1", 5),  # SUM skips text and blanks
    ("=SUM(ROUND(A1/3,2),A2)", 3.67),
    ("=ROUND(2.5,0)", 3),  # halves round away from zero, as in Excel
    ("=ABS(-A2)", 3),
    ("=A2*10%", 0.3),
    ("=Z9+1", 1),  # blank cell counts as 0
])
def test_evaluate(formula, expected):
    assert evaluate(formula, lookup) == pytest.approx(expected)


def test_single_reference_to_text_returns_text():
    assert evaluate("=C1", lookup) == "Car"


@pytest.mark.parametrize("formula", ["=A1/0", "=C1+1", "=VLOOKUP(A1,B1:B4,1)"])
def test_evaluate_errors(formula):
    with pytest.raises(FormulaError):
        evaluate(formula, lookup)


def test_parse_ref():
    assert parse_ref("'BS Schedules'!$C$74", "X") == Ref("BS Schedules", 3, 74, 3, 74)
    assert parse_ref("C7:C11", "BS") == Ref("BS", 3, 7, 3, 11)
    assert parse_ref("'It''s'!A1", "X").sheet == "It's"
    assert parse_ref("MyName", "X") is None  # defined name
    assert parse_ref("[1]Other!A1", "X") is None  # another file


def test_cell_key_quotes_like_excel():
    assert cell_key("PPE", 11, 7) == "PPE!G11"
    assert cell_key("BS Schedules", 74, 3) == "'BS Schedules'!C74"


def test_constants():
    assert constants("=3540+6080") == [3540, 6080]
    assert constants("=TB!B48") == []
