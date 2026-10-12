"""Reads the Stark workbook and reproduces its rolling order dates.

In the Orders sheet the order and shipping dates are formulas, not stored dates:
    G13 = TODAY()          (order 5020)
    G2  = G13 - 25         (every other order: TODAY() minus N days)
    H2  = G2 + 1           (shipping date = order date + 1 day)
These are parsed into DateRules (an offset from TODAY(), or from another cell). The rule is the
stored business data; the date it produces depends on the application date, exactly like the
sheet does when it is opened on a different day. A typed date is kept as a fixed date.

Delivery dates, contact-us dates and statuses are never read from cells; dates.py / engine.py calculate them.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta

import openpyxl

ORDERS = "Orders"
NEW_ORDERS = "New Orders"
ESTIMATES = "Shipping estimates"
PRICES = "Price List"

NEW_ORDER_HEADERS = ["New order number", "Original order number", "Type", "Date created", "Shipping date (same day)",
                     "Item", "Quantity", "Customer", "E-mail", "Address", "Shipping method", "Voucher"]
NEW_ORDER_TYPES = {
    "lost": "Lost order reship (Regular)",
    "damaged": "Wrong or damaged replacement (Express)",
    "exchange": "Exchange (Regular)",
}
# The only saved orders that may be cleared for retesting: a reshipment or a wrong/damaged replacement. Never an exchange.
CLEARABLE_TYPES = (NEW_ORDER_TYPES["lost"], NEW_ORDER_TYPES["damaged"])
# New Orders header row, right of the data columns: the highest new order number ever issued. It only goes up,
# so a number stays used after its order is cleared. Row 1 is never deleted or shifted.
LAST_NUMBER_LABEL_CELL, LAST_NUMBER_CELL = "N1", "O1"
LAST_NUMBER_LABEL = "Highest order number ever issued (do not edit)"


class FormulaError(ValueError):
    pass


@dataclass(frozen=True)
class DateRule:
    """A date as the workbook defines it: TODAY() + offset, or a fixed (typed) date + offset."""
    offset: int
    fixed: date | None = None          # None = relative to TODAY()
    formula: str = ""                  # the cell's own formula, for audit
    via: str = ""                      # cells followed to reach TODAY(), e.g. "G6 -> G13"

    def on(self, app_date: date) -> date:
        return (self.fixed if self.fixed else app_date) + timedelta(days=self.offset)

    @property
    def rolling(self) -> bool:
        return self.fixed is None

    def describe(self) -> str:
        if self.fixed:
            return f"fixed {self.fixed + timedelta(days=self.offset)}"
        return "TODAY()" if self.offset == 0 else f"TODAY() {'+' if self.offset > 0 else '-'} {abs(self.offset)} days"


@dataclass
class Item:
    qty: int
    product: str      # as written on the order, e.g. "hoverboard"
    variant: str      # e.g. "green"
    raw: str

    def label(self) -> str:
        return f"{self.product}, {self.variant}"


@dataclass
class Order:
    number: int
    email: str
    name: str
    address: str
    items: list[Item]
    method: str
    order_date: date | None            # generated for the book's application date
    ship_date: date | None
    source: str                        # "Orders" or "New Orders"
    original_number: int | None = None
    order_type: str | None = None
    order_rule: DateRule | None = None
    ship_rule: DateRule | None = None
    ship_after_order: int | None = None   # H = G + n relationship, when the shipping date is built from the order date
    ship_date_inferred: bool = False      # no shipping date in the sheet: customer-created rule (order date + 1) used

    @property
    def destination(self) -> str:
        return self.address.split(",")[-1].strip()

    @property
    def first_name(self) -> str:
        return self.name.strip().split(" ")[0] if self.name.strip() else "there"


@dataclass
class Book:
    path: str
    app_date: date                                  # what TODAY() means for this load
    orders: list[Order]
    new_orders: list[Order]
    estimates: dict[str, tuple[int, int]]
    prices: dict[tuple[str, str], int]
    warnings: list[str] = field(default_factory=list)
    last_number: int = 0                            # highest new order number ever issued (0 = none recorded)

    def all_orders(self) -> list[Order]:
        return self.orders + self.new_orders

    def x_for(self, destination: str, method: str) -> int:
        if destination not in self.estimates:
            raise KeyError(f"{destination!r} is not in Shipping estimates")
        reg, exp = self.estimates[destination]
        return reg if method == "Regular" else exp


def _as_date(v) -> date | None:
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    return None


_REF_OFFSET = re.compile(r"^=\s*\$?([A-Z]+)\$?(\d+)\s*(?:([+-])\s*(\d+))?\s*$")


def date_rule(ws, coord: str, _seen=None) -> DateRule | None:
    """Parse a date cell into a DateRule: typed date, =TODAY(), =REF, or =REF +/- n (as this workbook uses)."""
    _seen = _seen or []
    if coord in _seen:
        raise FormulaError(f"circular reference at {coord}")
    _seen = _seen + [coord]
    v = ws[coord].value
    if v is None or (isinstance(v, str) and not v.strip()):
        return None
    d = _as_date(v)
    if d:
        return DateRule(0, fixed=d, formula=str(d), via=coord)
    if not isinstance(v, str) or not v.startswith("="):
        raise FormulaError(f"{ws.title}!{coord} is not a date: {v!r}")
    if v.replace(" ", "").upper() == "=TODAY()":
        return DateRule(0, formula=v, via=coord)
    m = _REF_OFFSET.match(v)
    if not m:
        raise FormulaError(f"{ws.title}!{coord} uses an unsupported date formula: {v}")
    col, row, sign, n = m.groups()
    base = date_rule(ws, f"{col}{row}", _seen)
    if base is None:
        return None
    off = int(n or 0) * (-1 if sign == "-" else 1)
    return DateRule(base.offset + off, base.fixed, formula=v, via=f"{coord} -> {base.via}")


def eval_date_cell(ws, coord: str, today: date) -> date | None:
    rule = date_rule(ws, coord)
    return rule.on(today) if rule else None


def parse_items(text: str) -> list[Item]:
    items = []
    for part in str(text or "").split(";"):
        part = part.strip()
        if not part:
            continue
        m = re.match(r"^(\d+)\s+(.+?),\s*(.+)$", part)
        if not m:
            raise ValueError(f"Can't read item {part!r} (expected '<qty> <product>, <variant>')")
        items.append(Item(int(m.group(1)), m.group(2).strip(), m.group(3).strip(), part))
    return items


def load(path: str, app_date: date | None = None) -> Book:
    """app_date is what TODAY() means (default: the real current date)."""
    app_date = app_date or date.today()
    wb = openpyxl.load_workbook(path)
    warnings: list[str] = []

    est = {}
    ws = wb[ESTIMATES]
    for r in range(2, ws.max_row + 1):
        city, reg, exp = ws[f"A{r}"].value, ws[f"B{r}"].value, ws[f"C{r}"].value
        if city and isinstance(reg, (int, float)) and isinstance(exp, (int, float)):
            est[str(city).strip()] = (int(reg), int(exp))

    prices = {}
    ws = wb[PRICES]
    product = None
    for r in range(6, ws.max_row + 1):
        a, b, c = ws[f"A{r}"].value, ws[f"B{r}"].value, ws[f"C{r}"].value
        if a and str(a).strip():
            product = str(a).strip()
        if product and b and isinstance(c, (int, float)):
            prices[(product, str(b).strip())] = int(c)

    orders = []
    ws = wb[ORDERS]
    for r in range(2, ws.max_row + 1):
        num = ws[f"A{r}"].value
        if not isinstance(num, (int, float)):
            continue
        g_rule, h_rule = date_rule(ws, f"G{r}"), date_rule(ws, f"H{r}")
        inferred = False
        if h_rule is None and g_rule is not None:
            h_rule, inferred = DateRule(g_rule.offset + 1, g_rule.fixed, formula="(blank: order date + 1)",
                                        via="customer-created rule"), True
        link = None
        h_raw = ws[f"H{r}"].value
        m = _REF_OFFSET.match(h_raw) if isinstance(h_raw, str) else None
        if m and m.group(1) == "G" and int(m.group(2)) == r:
            link = int(m.group(4) or 0) * (-1 if m.group(3) == "-" else 1)
        if h_rule is None:
            warnings.append(f"Order {int(num)} has no order date or shipping date")
        orders.append(Order(int(num), str(ws[f"B{r}"].value or "").strip(), str(ws[f"C{r}"].value or "").strip(),
                            str(ws[f"D{r}"].value or "").strip(), parse_items(ws[f"E{r}"].value),
                            str(ws[f"F{r}"].value or "").strip(),
                            g_rule.on(app_date) if g_rule else None, h_rule.on(app_date) if h_rule else None, ORDERS,
                            order_rule=g_rule, ship_rule=h_rule, ship_after_order=1 if inferred else link,
                            ship_date_inferred=inferred))

    new_orders = []
    ws = wb[NEW_ORDERS]
    for r in range(2, ws.max_row + 1):
        num = ws[f"A{r}"].value
        if not isinstance(num, (int, float)):
            continue
        created = eval_date_cell(ws, f"D{r}", app_date)
        shipped = eval_date_cell(ws, f"E{r}", app_date) or created   # company-created: ships the day it's created
        prod, _, var = str(ws[f"F{r}"].value or "").partition(",")
        new_orders.append(Order(int(num), str(ws[f"I{r}"].value or ""), str(ws[f"H{r}"].value or ""),
                                str(ws[f"J{r}"].value or ""), [Item(int(ws[f"G{r}"].value), prod.strip(), var.strip(),
                                                                    str(ws[f"F{r}"].value))],
                                str(ws[f"K{r}"].value), created, shipped, NEW_ORDERS,
                                original_number=int(ws[f"B{r}"].value), order_type=str(ws[f"C{r}"].value)))

    nums = [o.number for o in orders + new_orders]
    dupes = sorted({n for n in nums if nums.count(n) > 1})
    if dupes:
        warnings.append(f"Duplicate order numbers: {dupes}")
    return Book(path, app_date, orders, new_orders, est, prices, warnings, last_number=_last_number(wb[NEW_ORDERS]))


def _last_number(ws) -> int:
    v = ws[LAST_NUMBER_CELL].value
    return int(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else 0


def _record_number(ws, number: int) -> None:
    """Remember that this order number has been issued. The recorded value never goes down."""
    ws[LAST_NUMBER_LABEL_CELL] = LAST_NUMBER_LABEL
    ws[LAST_NUMBER_CELL] = max(_last_number(ws), int(number))


def append_new_order(path: str, row: dict) -> None:
    """Log one company-created order in New Orders (raw data only; delivery dates are recalculated on read)."""
    wb = openpyxl.load_workbook(path)
    ws = wb[NEW_ORDERS]
    r = 2
    while ws[f"A{r}"].value not in (None, ""):
        r += 1
    values = [row["number"], row["original"], row["type"], row["created"], row["ship"], row["item"], row["qty"],
              row["name"], row["email"], row["address"], row["method"], row.get("voucher") or None]
    for c, v in enumerate(values, 1):
        ws.cell(r, c, v)
    ws[f"D{r}"].number_format = ws[f"E{r}"].number_format = "mmmm d, yyyy"
    ws[f"L{r}"].number_format = '"$"#,##0'
    _record_number(ws, row["number"])
    wb.save(path)


def remove_new_order(path: str, number: int, original_number: int) -> dict:
    """Delete one saved reshipment or replacement from New Orders so its original can be processed again
    (demonstrations and retesting only). The row must match on new order number, original order number and
    type; anything else raises ValueError and nothing is written. No other row or sheet is touched."""
    wb = openpyxl.load_workbook(path)
    ws = wb[NEW_ORDERS]
    rows = [r for r in range(2, ws.max_row + 1)
            if isinstance(ws[f"A{r}"].value, (int, float)) and int(ws[f"A{r}"].value) == int(number)]
    if len(rows) != 1:
        raise ValueError(f"Order {number} is not a saved new order (found {len(rows)} matching rows in New Orders)")
    r = rows[0]
    linked = ws[f"B{r}"].value
    if not isinstance(linked, (int, float)) or int(linked) != int(original_number):
        raise ValueError(f"Order {number} was not created from order {original_number}")
    if ws[f"C{r}"].value not in CLEARABLE_TYPES:
        raise ValueError(f"Order {number} is a {ws[f'C{r}'].value!r} order. Only a reshipment or a wrong or damaged "
                         f"replacement can be cleared")
    removed = {h: ws.cell(r, c).value for c, h in enumerate(NEW_ORDER_HEADERS, 1)}
    # Every number saved so far stays used, including in a workbook saved before numbers were recorded.
    _record_number(ws, max(int(ws[f"A{x}"].value) for x in range(2, ws.max_row + 1)
                           if isinstance(ws[f"A{x}"].value, (int, float))))
    ws.delete_rows(r)
    wb.save(path)
    return removed
