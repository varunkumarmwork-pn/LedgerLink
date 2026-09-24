"""Excel formula helpers: find cell references and recalculate simple formulas.

Built on openpyxl's formula tokenizer. The calculator covers what finished statements
use (+ - * / ^ %, brackets, SUM, ROUND, ABS); anything else raises FormulaError so the
caller can fall back to Excel's saved value.
"""
from dataclasses import dataclass
from decimal import ROUND_HALF_UP, Decimal
from typing import Callable

from openpyxl.formula.tokenizer import Token, Tokenizer
from openpyxl.utils.cell import get_column_letter, range_boundaries


class FormulaError(Exception):
    pass


@dataclass(frozen=True)
class Ref:
    """A reference to one cell or a rectangular range. Rows/cols are None for whole columns/rows."""
    sheet: str
    min_col: int | None
    min_row: int | None
    max_col: int | None
    max_row: int | None

    @property
    def is_single_cell(self) -> bool:
        return self.min_col == self.max_col and self.min_row == self.max_row and self.min_row is not None

    def cells(self, sheet_max_row: int, sheet_max_col: int) -> list[tuple[int, int]]:
        """(row, col) of every cell covered; whole-column/row refs stop at the sheet's used area."""
        rows = range(self.min_row or 1, (self.max_row or sheet_max_row) + 1)
        cols = range(self.min_col or 1, (self.max_col or sheet_max_col) + 1)
        return [(r, c) for r in rows for c in cols]


def cell_key(sheet: str, row: int, col: int) -> str:
    """Excel-style address, e.g. "'BS Schedules'!C74" or "PPE!G11"."""
    needs_quotes = not sheet.replace("_", "").isalnum()
    name = "'" + sheet.replace("'", "''") + "'" if needs_quotes else sheet
    return f"{name}!{get_column_letter(col)}{row}"


def tokens(formula: str) -> list[Token]:
    return [t for t in Tokenizer(formula).items if t.type != Token.WSPACE]


def parse_ref(text: str, current_sheet: str) -> Ref | None:
    """Turn a reference token like 'BS Schedules'!$C$74 into a Ref. None if we cannot follow it
    (defined names, links to other files, 3-D references, #REF!)."""
    sheet, address = current_sheet, text
    if "!" in text:
        sheet, address = text.rsplit("!", 1)
        if sheet.startswith("'") and sheet.endswith("'"):
            sheet = sheet[1:-1].replace("''", "'")
        if sheet.startswith("[") or ":" in sheet:
            return None
    try:
        min_col, min_row, max_col, max_row = range_boundaries(address.replace("$", ""))
    except (ValueError, TypeError):
        return None
    return Ref(sheet, min_col, min_row, max_col, max_row)


def references(formula: str, current_sheet: str) -> list[tuple[str, Ref | None]]:
    """Every reference token in a formula, with its parsed Ref (None if not followable)."""
    return [(t.value, parse_ref(t.value, current_sheet))
            for t in tokens(formula) if t.type == Token.OPERAND and t.subtype == Token.RANGE]


def constants(formula: str) -> list[float]:
    """Numbers typed inside a formula, e.g. =3540+6080 -> [3540, 6080]."""
    return [float(t.value) for t in tokens(formula) if t.type == Token.OPERAND and t.subtype == Token.NUMBER]


# ---------------------------------------------------------------- calculator

PRECEDENCE = {"=": 1, "<>": 1, "<": 1, ">": 1, "<=": 1, ">=": 1, "&": 2, "+": 3, "-": 3, "*": 4, "/": 4, "^": 5}


def evaluate(formula: str, lookup: Callable[[str], object]):
    """Recalculate a formula. `lookup(reference_text)` returns a cell value, or a list of
    values for a range."""
    parser = _Parser(tokens(formula), lookup)
    value = parser.expression()
    if parser.peek() is not None:
        raise FormulaError(f"cannot read formula past '{parser.peek().value}'")
    return value


class _Parser:
    """Precedence-climbing parser that evaluates as it reads."""

    def __init__(self, items: list[Token], lookup):
        self.items, self.pos, self.lookup = items, 0, lookup

    def peek(self) -> Token | None:
        return self.items[self.pos] if self.pos < len(self.items) else None

    def take(self) -> Token:
        token = self.peek()
        if token is None:
            raise FormulaError("formula ends too early")
        self.pos += 1
        return token

    def expression(self, min_precedence: int = 1):
        left = self.unary()
        while (op := self.peek()) and op.type == Token.OP_IN and PRECEDENCE.get(op.value, 0) >= min_precedence:
            self.take()
            right = self.expression(PRECEDENCE[op.value] + 1)
            left = _apply(op.value, left, right)
        return left

    def unary(self):
        token = self.peek()
        if token and token.type == Token.OP_PRE:
            self.take()
            value = _number(self.unary())
            return -value if token.value == "-" else value
        value = self.primary()
        while (post := self.peek()) and post.type == Token.OP_POST:  # 5% -> 0.05
            self.take()
            value = _number(value) / 100
        return value

    def primary(self):
        token = self.take()
        if token.type == Token.OPERAND:
            if token.subtype == Token.NUMBER:
                return float(token.value)
            if token.subtype == Token.TEXT:
                return token.value[1:-1].replace('""', '"')
            if token.subtype == Token.RANGE:
                return self.lookup(token.value)
            raise FormulaError(f"cannot use {token.value}")
        if token.type == Token.PAREN and token.subtype == Token.OPEN:
            value = self.expression()
            self._expect(Token.PAREN)
            return value
        if token.type == Token.FUNC and token.subtype == Token.OPEN:
            name, args = token.value[:-1].upper(), []
            if not self._at_close(Token.FUNC):
                args.append(self.expression())
                while (sep := self.peek()) and sep.type == Token.SEP:
                    self.take()
                    args.append(self.expression())
            self._expect(Token.FUNC)
            return _call(name, args)
        raise FormulaError(f"unexpected '{token.value}'")

    def _at_close(self, token_type) -> bool:
        token = self.peek()
        return token is not None and token.type == token_type and token.subtype == Token.CLOSE

    def _expect(self, token_type):
        if not self._at_close(token_type):
            raise FormulaError("missing closing bracket")
        self.take()


def _number(value) -> float:
    if value is None:
        return 0.0  # blank cell counts as 0, as in Excel
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        try:
            return float(value)
        except ValueError:
            pass
    raise FormulaError(f"#VALUE! ({value!r} is not a number)")


def _apply(op: str, left, right):
    if op == "&":
        return f"{'' if left is None else left}{'' if right is None else right}"
    if op == "=":
        return left == right
    if op == "<>":
        return left != right
    if op in ("<", ">", "<=", ">="):
        a, b = _number(left), _number(right)
        return a < b if op == "<" else a > b if op == ">" else a <= b if op == "<=" else a >= b
    a, b = _number(left), _number(right)
    if op == "+":
        return a + b
    if op == "-":
        return a - b
    if op == "*":
        return a * b
    if op == "/":
        if b == 0:
            raise FormulaError("#DIV/0!")
        return a / b
    return a ** b


def _call(name: str, args: list):
    if name == "SUM":
        total = 0.0
        for arg in args:
            if isinstance(arg, list):  # a range: SUM skips text and blanks
                total += sum(float(v) for v in arg if isinstance(v, (int, float)) and not isinstance(v, bool))
            else:
                total += _number(arg)
        return total
    if name == "ROUND" and len(args) == 2:
        places = int(_number(args[1]))
        # Excel rounds halves away from zero; Python's round() does not.
        return float(Decimal(repr(_number(args[0]))).quantize(Decimal(1).scaleb(-places), rounding=ROUND_HALF_UP))
    if name == "ABS" and len(args) == 1:
        return abs(_number(args[0]))
    raise FormulaError(f"function {name} is not supported yet")
