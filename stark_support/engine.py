"""Support decisions for Stark Industries orders.

Inputs: an order (raw data from the workbook), the customer's contact date, and the request.
Rules come from the workbook's Policies sheet. Every date other than the recorded order and
shipping dates is calculated here. Results carry internal facts (never shown to customers)
plus customer-facing paragraphs that reply.compose turns into one email.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date

from . import dates, workbook
from .workbook import Book, Item, Order

REFUND_DAYS = 10      # refund and wrong/damaged window, business days after shipping
EXCHANGE_DAYS = 14    # exchange window, business days after shipping

NOT_SHIPPED = "Not shipped yet"
NOT_YET_DUE = "Not yet due"
DUE = "Due today"
LATE = "Late, not yet eligible for reshipment"
RESHIP = "Eligible for lost-order reshipment"


def fmt(d: date) -> str:
    return f"{d:%A}, {d:%B} {d.day}"


@dataclass
class Request:
    kind: str                         # status | lost | refund | exchange | exchange_tracking | damaged | wrong | cancel
    order_number: int | None = None
    email: str | None = None
    on: date | None = None            # customer contact date: required (pass date.today() explicitly for today)
    item: int | str | None = None     # 1-based line number or text; blank = the only item
    qty: int | None = None            # blank = all of that item
    new_item: str | None = None       # exchanges: "Product (Variant)" or "product, variant"
    offered_on: date | None = None    # exchange_tracking: the day the exchange was offered (its request date)
    new_order_number: int | None = None
    tracking_number: str | None = None
    label_sent: bool = False          # only say a label is attached when it really is
    ask_remaining: bool = False       # customer asked what their order is worth after the refund
    force_new: bool = False           # staff override: create another order even though one already exists


@dataclass
class Result:
    request: Request
    decision: str
    order: Order | None = None
    facts: dict = field(default_factory=dict)       # internal only
    paragraphs: list[str] = field(default_factory=list)
    new_order: dict | None = None
    internal_note: str | None = None


class Desk:
    def __init__(self, book: Book, persist_path: str | None = None):
        self.book = book
        self.persist_path = persist_path   # set to log new orders in the workbook

    # ---------- identification ----------
    def orders_for_email(self, email: str) -> list[Order]:
        e = email.strip().lower()
        return [o for o in self.book.all_orders() if o.email.strip().lower() == e]

    def identify(self, req: Request) -> tuple[Order | None, Result | None]:
        if req.order_number is not None:
            match = [o for o in self.book.all_orders() if o.number == int(req.order_number)]
            if not match:
                return None, Result(req, "NOT_FOUND", paragraphs=[
                    f"I couldn't find an order with the number {req.order_number}. Could you double-check it and send it over again?"])
            o = match[0]
            if req.email and req.email.strip().lower() != o.email.lower():
                return None, Result(req, "EMAIL_MISMATCH", o, internal_note=f"Order {o.number} belongs to {o.email}, not {req.email}.",
                                    paragraphs=["I found that order number, but it isn't linked to the email address you wrote from. "
                                                "Could you confirm the email address used when the order was placed?"])
            return o, None
        if req.email:
            found = self.orders_for_email(req.email)
            if len(found) == 1:
                return found[0], None
            if not found:
                return None, Result(req, "NOT_FOUND", paragraphs=[
                    "I couldn't find an order under this email address. Could you send me your order number?"])
            return None, Result(req, "AMBIGUOUS", facts={"matches": [o.number for o in found]},
                                internal_note=f"{len(found)} orders share {req.email}: {[o.number for o in found]}",
                                paragraphs=["I can see more than one order under your email address. Could you let me know "
                                            "the order number this is about, so I'm looking at the right one?"])
        return None, Result(req, "NO_ID", paragraphs=["Could you send me your order number so I can look into this?"])

    # ---------- date analysis ----------
    def analyse(self, o: Order, on: date) -> dict:
        """Everything about an order's timeline as of the customer's contact date."""
        x = self.book.x_for(o.destination, o.method)
        est = dates.estimated_delivery(o.ship_date, o.method, x)
        contact = dates.contact_us_date(o.ship_date, o.method, x)
        if on < o.ship_date:
            state = NOT_SHIPPED
        elif on < est:
            state = NOT_YET_DUE
        elif on == est:
            state = DUE
        elif on < contact:
            state = LATE
        else:
            state = RESHIP
        late_cal = max(0, dates.calendar_days_between(est, on))
        return {
            "app_date": self.book.app_date,
            "order_date_rule": o.order_rule.describe() if o.order_rule else "recorded",
            "ship_date_rule": (f"order date + {o.ship_after_order} day" if o.ship_after_order is not None
                               else (o.ship_rule.describe() if o.ship_rule else "same day as created")),
            "order_date": o.order_date, "ship_date": o.ship_date, "ship_date_inferred": o.ship_date_inferred,
            "destination": o.destination, "method": o.method, "x": x,
            "day_type": "business days" if o.method == "Regular" else "calendar days",
            "expected_delivery": est,
            "two_x_end": dates.add_shipping_days(o.ship_date, o.method, 2 * x),
            "contact_us_date": contact,
            "contact_date": on,
            "calendar_days_since_shipping": dates.calendar_days_between(o.ship_date, on),
            "business_days_since_shipping": dates.business_days_between(o.ship_date, on),
            "shipping_days_since_shipping": dates.shipping_days_between(o.ship_date, on, o.method),
            "calendar_days_past_expected": late_cal,
            "business_days_past_expected": max(0, dates.business_days_between(est, on)),
            "days_until_reship_eligible": max(0, dates.calendar_days_between(on, contact)),
            "state": state,
            "eligible_action": {NOT_SHIPPED: "none yet; can't cancel or change", NOT_YET_DUE: "wait for delivery",
                                DUE: "wait for delivery", LATE: f"wait; reship possible from {contact}",
                                RESHIP: "lost-order reshipment (Regular)"}[state],
        }

    # ---------- helpers ----------
    def pick_item(self, o: Order, sel) -> int | None:
        if sel in (None, ""):
            return 0 if len(o.items) == 1 else None
        if isinstance(sel, int) or str(sel).isdigit():
            i = int(sel) - 1
            return i if 0 <= i < len(o.items) else None
        words = str(sel).lower().replace(",", " ").split()
        hits = [i for i, it in enumerate(o.items) if all(w in f"{it.product} {it.variant}".lower() for w in words)]
        return hits[0] if len(hits) == 1 else None

    def price_key(self, product: str, variant: str) -> tuple[str, str] | None:
        words = product.lower().split()
        hits = [k for k in self.book.prices if all(w in k[0].lower() for w in words) and k[1].lower() == variant.lower()]
        return hits[0] if len(hits) == 1 else None

    def parse_new_item(self, text: str) -> tuple[str, str] | None:
        t = text.strip()
        if "(" in t and t.endswith(")"):
            p, v = t[:-1].split("(", 1)
        else:
            p, _, v = t.partition(",")
        p = " ".join(w for w in p.split() if not w.isdigit())
        return self.price_key(p.strip(), v.strip())

    def order_value(self, o: Order) -> int | None:
        total = 0
        for it in o.items:
            k = self.price_key(it.product, it.variant)
            if not k:
                return None
            total += it.qty * self.book.prices[k]
        return total

    def key_or_fallback(self, it: Item) -> tuple[str, str]:
        return self.price_key(it.product, it.variant) or (it.product.title(), it.variant.title())

    @staticmethod
    def name_of(key: tuple[str, str]) -> str:
        return f"{key[0]} ({key[1]})"

    @staticmethod
    def is_vibranium(item: Item) -> bool:
        return "vibranium" in f"{item.product} {item.variant}".lower()

    @staticmethod
    def is_hoverboard(item: Item) -> bool:
        return "hoverboard" in item.product.lower()

    @staticmethod
    def qty_phrase(q: int, name: str) -> str:
        return f"the {name}" if q == 1 else f"{q} × {name}"

    def next_order_number(self) -> int:
        """Highest number in use or ever issued, plus one: a cleared order's number is never given out again."""
        return max(max(o.number for o in self.book.all_orders()), self.book.last_number) + 1

    def create_order(self, o: Order, kind: str, created: date, key, qty: int, new_number=None, voucher=None) -> dict:
        """Company-created order: separate number, ships the day it's created, method fixed by type."""
        method = "Express" if kind == "damaged" else "Regular"
        number = int(new_number) if new_number is not None else self.next_order_number()
        if any(x.number == number for x in self.book.all_orders()):
            raise ValueError(f"Order number {number} is already used")
        x = self.book.x_for(o.destination, method)
        item_text = f"{key[0].lower()}, {key[1].lower()}"
        row = {"number": number, "original": o.number, "type": workbook.NEW_ORDER_TYPES[kind], "created": created,
               "ship": created, "item": item_text, "qty": qty, "name": o.name, "email": o.email, "address": o.address,
               "method": method, "x": x, "voucher": voucher,
               "est": dates.estimated_delivery(created, method, x), "contact": dates.contact_us_date(created, method, x)}
        self.book.new_orders.append(Order(number, o.email, o.name, o.address, [Item(qty, key[0].lower(), key[1].lower(), item_text)],
                                          method, created, created, workbook.NEW_ORDERS, original_number=o.number,
                                          order_type=row["type"]))
        if self.persist_path:
            workbook.append_new_order(self.persist_path, row)
        return row

    def _new_order_lines(self, row: dict, on: date) -> list[str]:
        # The shipping method stays internal: the customer needs the dates and what to do next.
        what = "exchange order" if row["type"] == workbook.NEW_ORDER_TYPES["exchange"] else "replacement order"
        return [f"Your {what} number is {row['number']}. It's scheduled to ship on {fmt(row['ship'])} and should arrive by "
                f"{fmt(row['est'])}. If it hasn't arrived by {fmt(row['contact'])}, please contact us again so we can help "
                f"resolve the issue."]

    def _label_line(self, req, what=None):
        send = "send it back" if what is None else f"send back {what}"
        return (f"Please use the free return label attached to this email to {send}." if req.label_sent
                else f"We'll email you a free return label so you can {send}.")

    def existing_new_orders(self, o: Order, kind: str, key=None) -> list[Order]:
        """Company-created orders of this type already logged for original order o (optionally for one item)."""
        out = [n for n in self.book.new_orders if n.original_number == o.number and n.order_type == workbook.NEW_ORDER_TYPES[kind]]
        if key:
            out = [n for n in out if (n.items[0].product.lower(), n.items[0].variant.lower()) == (key[0].lower(), key[1].lower())]
        return out

    def _refer(self, req, o, prior: Order, decision: str, opener: str) -> Result:
        """Repeat enquiry: answer with the status of the existing new order (fixed dates); never create another one."""
        r = self._status(Request(**{**req.__dict__, "order_number": prior.number, "email": None}), prior, escalate=False)
        r.paragraphs.insert(0, opener)
        note = f"Original {o.number} already has {prior.order_type} order {prior.number}; no new order created."
        if r.decision == RESHIP:
            note += f" Order {prior.number} is itself past its contact-us date: review before sending anything else."
        r.internal_note = note
        r.decision = f"{decision} / {r.decision}"
        r.facts["existing_order"] = prior.number
        return r

    # ---------- dispatcher ----------
    def handle(self, req: Request) -> Result:
        if req.on is None:
            return Result(req, "NEED_CONTACT_DATE", internal_note="Enter the customer contact date (or choose today explicitly).")
        o, problem = self.identify(req)
        if problem:
            return problem
        if o.ship_date is None:
            return Result(req, "MISSING_DATES", o, internal_note=(
                f"Order {o.number} has no order or shipping date in the workbook, so nothing can be calculated."))
        if o.order_date and req.on < o.order_date:
            return Result(req, "CONTACT_BEFORE_ORDER", o, {"app_date": self.book.app_date, "order_date": o.order_date,
                                                           "contact_date": req.on}, internal_note=(
                f"CHECK: contact date {req.on} is before order {o.number}'s order date {o.order_date} "
                f"(generated for application date {self.book.app_date}). Check the contact date or the application date."))
        try:
            self.book.x_for(o.destination, o.method)
        except KeyError as e:
            return Result(req, "DATA_ERROR", o, internal_note=str(e))
        fn = {"status": self._status, "lost": self._lost, "refund": self._refund, "exchange": self._exchange,
              "exchange_tracking": self._exchange_tracking, "damaged": self._damaged, "wrong": self._damaged,
              "cancel": self._cancel}[req.kind]
        return fn(req, o)

    # ---------- shipping ----------
    def _status(self, req, o, escalate=True):
        a = self.analyse(o, req.on)
        if a["state"] == RESHIP and escalate:
            return self._lost(req, o)
        ship, est, cu = fmt(a["ship_date"]), fmt(a["expected_delivery"]), fmt(a["contact_us_date"])
        missing = req.kind == "lost"       # the customer said it's late or hasn't arrived, not just "where is it?"
        if a["state"] == NOT_SHIPPED:
            p = (f"Your order {o.number} hasn't shipped yet. It's scheduled to ship on {ship} and should arrive by {est}. "
                 f"If it hasn't arrived by {cu}, please get back in touch.")
        elif a["state"] == NOT_YET_DUE:
            lead = (f"Your order {o.number} is still within its delivery time, so it isn't late yet." if missing
                    else f"Your order {o.number} is on its way.")
            p = (f"{lead} It shipped on {ship} and should arrive by {est}. "
                 f"If it hasn't arrived by {cu}, please get back in touch and we'll take care of it.")
        elif a["state"] == DUE:
            p = (f"Your order {o.number} is due to arrive today. It shipped on {ship}. "
                 f"If it hasn't arrived by {cu}, please get back in touch and we'll take care of it.")
        elif a["state"] == RESHIP:
            p = (f"Your order {o.number} shipped on {ship} and was due by {est}, and it hasn't arrived yet. "
                 f"I'm sorry about that. We're looking into it and will follow up with you.")
        else:
            p = (f"I'm sorry, your order {o.number} is running late. It shipped on {ship} and was due by {est}. "
                 f"Deliveries sometimes take a few extra days, so it may still arrive. If it hasn't reached you by {cu}, "
                 f"please get back in touch and we'll send you a replacement.")
        return Result(req, a["state"], o, a, [p])

    def _lost(self, req, o):
        a = self.analyse(o, req.on)
        if a["state"] != RESHIP:
            r = self._status(req, o)
            r.internal_note = f"Not eligible for reshipment until {a['contact_us_date']}."
            return r
        prior = self.existing_new_orders(o, "lost")
        if prior and not req.force_new:
            return self._refer(req, o, prior[-1], "ALREADY_RESHIPPED",
                               f"We already sent a replacement for order {o.number} under order number {prior[-1].number}.")
        rows = [self.create_order(o, "lost", req.on, self.key_or_fallback(it), it.qty,
                                  req.new_order_number if len(o.items) == 1 else None) for it in o.items]
        paras = [f"I'm sorry your order {o.number} still hasn't arrived. There was an issue during shipping and we "
                 f"haven't been able to locate the package, so I've set up a replacement for you at no cost."]
        for row in rows:
            paras += self._new_order_lines(row, req.on)
        a["new_orders"] = rows
        return Result(req, RESHIP, o, a, paras, new_order=rows[0] if len(rows) == 1 else {"orders": rows})

    # ---------- item requests ----------
    def _item_checks(self, req, o):
        idx = self.pick_item(o, req.item)
        if idx is None:
            names = ", ".join(self.name_of(self.key_or_fallback(i)) for i in o.items)
            return None, None, Result(req, "PICK_ITEM", o, {"items": names}, [
                f"Your order {o.number} has a few different items ({names}). Which one is this about, and how many?"])
        item = o.items[idx]
        qty = item.qty if req.qty in (None, "") else int(req.qty)
        if qty < 1 or qty > item.qty:
            return None, None, Result(req, "BAD_QTY", o, {"ordered": item.qty}, [
                f"Your order has {item.qty} of that item, so I can process up to {item.qty}. How many would you like to send back?"])
        if req.on < o.ship_date:
            return None, None, self._cancel(req, o)
        return item, qty, None

    @staticmethod
    def _rest_note(o, item, qty):
        return " The other items on your order stay as they are." if (qty < item.qty or len(o.items) > 1) else ""

    def _refund(self, req, o):
        item, qty, early = self._item_checks(req, o)
        if early:
            return early
        key = self.price_key(item.product, item.variant)
        name = self.name_of(key) if key else item.label()
        dl = dates.business_deadline(o.ship_date, REFUND_DAYS)
        ex_dl = dates.business_deadline(o.ship_date, EXCHANGE_DAYS)
        facts = {"ship_date": o.ship_date, "refund_deadline": dl, "item": name, "qty": qty, "ordered_qty": item.qty,
                 "order_value": self.order_value(o)}
        if self.is_hoverboard(item):
            offer = ""
            if req.on <= ex_dl:
                facts["exchange_deadline"] = ex_dl
                offer = (f" If you'd like, I can help you exchange {'it' if qty == 1 else 'them'} for another item of the same price or less instead, "
                         f"as long as we hear from you by {fmt(ex_dl)}.")
            return Result(req, "REFUND_HOVER", o, facts, [
                f"Hoverboards can't be returned for a refund, so I'm not able to refund {self.qty_phrase(qty, name)}.{offer}"])
        if req.on > dl:
            offer = ""
            if req.on <= ex_dl and not self.is_vibranium(item):
                facts["exchange_deadline"] = ex_dl
                offer = (f" I can still help you exchange {'it' if qty == 1 else 'them'} for another item of the same "
                         f"price or less, as long as we hear from you by {fmt(ex_dl)}.")
            return Result(req, "REFUND_LATE", o, facts, [
                f"I'm sorry, I'm not able to offer a refund for {self.qty_phrase(qty, name)}. Refunds need to be requested "
                f"within 10 business days of shipping. Your order shipped on {fmt(o.ship_date)}, so the last day was "
                f"{fmt(dl)}.{offer}"])
        if not key:
            return Result(req, "PRICE_UNKNOWN", o, facts, internal_note=f"{item.raw} not matched to the Price List.")
        unit = self.book.prices[key]
        amount = qty * unit
        facts.update(unit_price=unit, refund_amount=amount)
        paras = [f"Yes, you can return {self.qty_phrase(qty, name)} for a refund.{self._rest_note(o, item, qty)}",
                 f"{self._label_line(req)} Once you've shipped it, reply with the tracking number and we'll send you a "
                 f"${amount} voucher for {'that item' if qty == 1 else 'those items'}."]
        if facts["order_value"] is not None:
            facts["remaining_value"] = facts["order_value"] - amount
            if req.ask_remaining:
                paras.append(f"Your order came to ${facts['order_value']} in total, so after the ${amount} refund, "
                             f"the items you're keeping are worth ${facts['remaining_value']}.")
        return Result(req, "REFUND_OK", o, facts, paras)

    def _exchange_eval(self, req, o, check_on: date):
        item, qty, early = self._item_checks(req, o)
        if early:
            return None, early
        key = self.price_key(item.product, item.variant)
        name = self.name_of(key) if key else item.label()
        dl = dates.business_deadline(o.ship_date, EXCHANGE_DAYS)
        facts = {"ship_date": o.ship_date, "exchange_deadline": dl, "item": name, "qty": qty,
                 "ordered_qty": item.qty, "request_date": check_on}
        if self.is_vibranium(item):
            rdl = dates.business_deadline(o.ship_date, REFUND_DAYS)
            offer = ""
            if req.on <= rdl and not self.is_hoverboard(item):
                facts["refund_deadline"] = rdl
                offer = f" If you'd like, you can return it for a refund instead, as long as we hear from you by {fmt(rdl)}."
            return None, Result(req, "EX_VIB", o, facts, [
                f"Because the {name} you bought is made of vibranium, it can't be exchanged for another product.{offer}"])
        if check_on > dl:
            return None, Result(req, "EX_LATE", o, facts, [
                f"I'm sorry, I'm not able to set up an exchange for {self.qty_phrase(qty, name)}. Exchanges need to be "
                f"requested within 14 business days of shipping. Your order shipped on {fmt(o.ship_date)}, so the last "
                f"day was {fmt(dl)}."])
        if not req.new_item:
            return None, Result(req, "EX_NEED_ITEM", o, facts, [
                f"Happy to help with an exchange for {self.qty_phrase(qty, name)}. Which item would you like instead?"])
        new_key = self.parse_new_item(req.new_item)
        if not key or not new_key:
            return None, Result(req, "PRICE_UNKNOWN", o, facts, internal_note="Item not matched to the Price List.",
                                paragraphs=[f"Could you tell me exactly which item and variant you'd like instead of the {name}?"])
        old_p, new_p = self.book.prices[key], self.book.prices[new_key]
        new_name = self.name_of(new_key)
        facts.update(new_item=new_name, unit_price=old_p, new_unit_price=new_p)
        if new_p > old_p:
            return None, Result(req, "EX_PRICE", o, facts, [
                f"The {new_name} costs ${new_p}, which is more than the ${old_p} {name} you bought, so I can't swap them. "
                f"An exchange has to be for an item of the same price or less, and paying the difference isn't an option.",
                "If you'd like to pick something of the same price or less, I'd be glad to set that up."])
        facts["voucher"] = qty * (old_p - new_p)
        return (item, key, new_key, qty, facts), None

    def _exchange(self, req, o):
        ok, res = self._exchange_eval(req, o, req.on)
        if res:
            return res
        item, key, new_key, qty, facts = ok
        v = facts["voucher"]
        vline = f" Since it costs less, you'll also get a ${v} voucher for the difference." if v else ""
        return Result(req, "EX_OK", o, facts, [
            f"Yes, we can exchange {self.qty_phrase(qty, self.name_of(key))} for {self.qty_phrase(qty, self.name_of(new_key))}."
            f"{self._rest_note(o, item, qty)}{vline}",
            f"{self._label_line(req)} Once you've shipped it, reply with the tracking number and I'll set up your new order."])

    def _exchange_tracking(self, req, o):
        new_key = self.parse_new_item(req.new_item) if req.new_item else None
        prior = self.existing_new_orders(o, "exchange", new_key)
        if prior and not req.force_new:
            return self._refer(req, o, prior[-1], "ALREADY_EXCHANGED",
                               f"Your exchange for order {o.number} is already set up under order number {prior[-1].number}.")
        ok, res = self._exchange_eval(req, o, req.offered_on or req.on)
        if res:
            return res
        if not req.tracking_number:
            return Result(req, "EX_NEED_TRACKING", o, paragraphs=[
                "Could you send me the tracking number for the return? I'll set up your new order as soon as I have it."])
        item, key, new_key, qty, facts = ok
        row = self.create_order(o, "exchange", req.on, new_key, qty, req.new_order_number, facts["voucher"] or None)
        paras = [f"Thanks for sending the tracking number. Your exchange order for "
                 f"{self.qty_phrase(qty, self.name_of(new_key))} is all set."]
        paras += self._new_order_lines(row, req.on)
        if facts["voucher"]:
            paras.append(f"You'll also get a ${facts['voucher']} voucher for the price difference.")
        facts.update(tracking=req.tracking_number, new_order=row)
        return Result(req, "EX_CREATED", o, facts, paras, new_order=row)

    def _damaged(self, req, o):
        item, qty, early = self._item_checks(req, o)
        if early:
            return early
        wrong = req.kind == "wrong"
        dl = dates.business_deadline(o.ship_date, REFUND_DAYS)
        key = self.key_or_fallback(item)
        prior = self.existing_new_orders(o, "damaged", key)
        if prior and not req.force_new:
            return self._refer(req, o, prior[-1], "ALREADY_REPLACED",
                               f"We already set up a replacement for the {self.name_of(key)} from order {o.number} "
                               f"under order number {prior[-1].number}.")
        facts = {"ship_date": o.ship_date, "deadline": dl, "item": self.name_of(key), "qty": qty,
                 "reason": "wrong item" if wrong else "damaged"}
        if req.on > dl:
            what = "the wrong item" if wrong else "a damaged item"
            return Result(req, "DM_LATE", o, facts, [
                f"I'm sorry you received {what}. Unfortunately, I'm not able to send a replacement. Replacements need to "
                f"be requested within 10 business days of shipping. Your order shipped on {fmt(o.ship_date)}, so the "
                f"last day was {fmt(dl)}."])
        row = self.create_order(o, "damaged", req.on, key, qty, req.new_order_number)
        opener = (f"I'm sorry you received the wrong item instead of {self.qty_phrase(qty, self.name_of(key))}."
                  if wrong else f"I'm sorry {self.qty_phrase(qty, self.name_of(key))} arrived damaged.")
        paras = [f"{opener} I've set up a replacement for you.",
                 *self._new_order_lines(row, req.on),
                 self._label_line(req, "the item you received" if wrong else "the damaged item")]
        facts["new_order"] = row
        return Result(req, "DM_CREATED", o, facts, paras, new_order=row)

    def _cancel(self, req, o):
        a = self.analyse(o, req.on)
        verb = "is scheduled to ship on" if req.on < o.ship_date else "shipped on"
        arrive = "should arrive by" if req.on <= a["expected_delivery"] else "was due by"
        # Only mention what is still open for this order's items on the contact date.
        options = [name for name, ok in (
            ("an exchange", any(not self.is_vibranium(i) for i in o.items)
             and req.on <= dates.business_deadline(o.ship_date, EXCHANGE_DAYS)),
            ("a return", any(not self.is_hoverboard(i) for i in o.items)
             and req.on <= dates.business_deadline(o.ship_date, REFUND_DAYS))) if ok]
        if req.kind != "cancel":     # an item request (refund, exchange, damaged, wrong) before the order has shipped
            what = {"refund": "a refund", "exchange": "an exchange", "exchange_tracking": "an exchange"}.get(
                req.kind, "a replacement")
            paras = [f"I can't set up {what} yet because your order {o.number} hasn't shipped. It's scheduled to ship on "
                     f"{fmt(o.ship_date)} and should arrive by {fmt(a['expected_delivery'])}."]
            if req.kind in ("damaged", "wrong"):
                paras.append("If there's a problem with it when it arrives, please get back in touch.")
            elif options:
                paras.append(f"Once it arrives, I'd be happy to help with {' or '.join(options)} if it isn't what you need.")
            return Result(req, "CANCEL_NO", o, a, paras)
        paras = [f"Unfortunately, we aren't able to cancel or change an order once it's been placed. Your order {o.number} "
                 f"{verb} {fmt(o.ship_date)} and {arrive} {fmt(a['expected_delivery'])}."]
        if options:
            when = "Once it arrives" if req.on <= a["expected_delivery"] else "If it isn't what you need"
            tail = " if it isn't what you need" if req.on <= a["expected_delivery"] else ""
            paras.append(f"{when}, I'd be happy to help with {' or '.join(options)}{tail}.")
        return Result(req, "CANCEL_NO", o, a, paras)
