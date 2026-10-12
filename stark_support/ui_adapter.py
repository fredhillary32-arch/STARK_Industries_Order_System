"""Plain functions between the browser page (app.py) and the engine. No Streamlit code here.

Every rule, date and amount comes from engine.Desk; this module only asks the engine and reshapes its
answers into plain dicts for display. Checks never write: only save_action / save_once pass a persist_path.
"""
from __future__ import annotations

import logging
import os
import uuid
from datetime import date, datetime

from . import Desk, Request, compose, dates, engine, load, workbook
from . import __main__ as cli
from . import reply
from .reply import lint

WORKBOOK_PATH = cli.DEFAULT_BOOK
log = logging.getLogger("stark_support.ui")

# What the customer is asking, in their words. The first value is the engine's request kind.
INQUIRIES = [
    ("status", "Where is my order?"),
    ("lost", "My order is late or hasn't arrived"),
    ("cancel", "Can I cancel or change my order?"),
    ("refund", "I want to return an item for a refund"),
    ("exchange", "I want to exchange an item"),
    ("exchange_tracking", "I've sent back my exchange item (tracking number)"),
    ("damaged", "My item arrived damaged"),
    ("wrong", "I received the wrong item"),
]
ACTIONS = INQUIRIES
ITEM_KINDS = ("refund", "exchange", "exchange_tracking", "damaged", "wrong")
LABEL_KINDS = ("refund", "exchange", "damaged", "wrong")
# Plain-language status for each engine state.
STATUS_PLAIN = {engine.NOT_SHIPPED: "Not shipped yet", engine.NOT_YET_DUE: "On schedule",
                engine.DUE: "On schedule, due today",
                engine.LATE: "Running late, waiting for reshipment eligibility",
                engine.RESHIP: "Eligible for reshipment"}
ACTION_FIELDS = {"item", "qty", "new_item", "offered_on", "tracking_number", "label_sent", "ask_remaining"}
STATE_COLOURS = {engine.NOT_SHIPPED: "gray", engine.NOT_YET_DUE: "blue", engine.DUE: "blue",
                 engine.LATE: "orange", engine.RESHIP: "red"}
# Decisions where the order's timeline can't be analysed.
NO_ANALYSIS = {"NEED_CONTACT_DATE", "MISSING_DATES", "CONTACT_BEFORE_ORDER", "DATA_ERROR"}


def local_today() -> date:
    return datetime.now().astimezone().date()


def show(d) -> str:
    return f"{d:%a} {d:%b} {d.day}, {d.year}" if isinstance(d, date) else ""


def show_long(d) -> str:
    return f"{d:%A}, {d:%B} {d.day}, {d.year}" if isinstance(d, date) else ""


def full_email(subject: str, body: str) -> str:
    """Exactly what Copy Email puts on the clipboard."""
    subject = (subject or "").strip()
    return (f"Subject: {subject}\n\n" if subject else "") + (body or "")


def check_email(view: dict, text: str) -> list[str]:
    """Problems in an (edited) email, judged against what the engine calculated for this result."""
    return reply.quality_issues(text, view.get("email_context") or {})


def catalogue(path=None) -> list[str]:
    """Items a customer can exchange for, straight from the Price List."""
    try:
        book = load(path or WORKBOOK_PATH, app_date=local_today())
        return [Desk.name_of(k) for k in book.prices]
    except Exception:
        log.exception("Could not read the Price List")
        return []


def prepare_reply(kind, identifier, contact_date, app_date=None, path=None, **fields) -> dict:
    """One inquiry, start to finish: find the order, apply the rules, draft the reply. Never writes."""
    if kind not in dict(INQUIRIES):
        return _error("Choose what the customer is asking.", contact_date, app_date)
    extra = set(fields) - ACTION_FIELDS
    if extra:
        return _error(f"Not available in this interface: {', '.join(sorted(extra))}.", contact_date, app_date)
    ident = str(identifier or "").strip()
    params = {"kind": kind, "on": contact_date, **fields}
    if ident.isdigit():
        params["order_number"] = int(ident)
    elif "@" in ident:
        params["email"] = ident
    elif ident:
        return _error("Enter an order number (digits only, for example 1650) or the customer's email address.",
                      contact_date, app_date)
    return _run(params, app_date, path, persist=False)


def lint_email(text: str) -> list[str]:
    """Forbidden internal terms found in an (edited) customer email, in readable form."""
    return [p.replace(r"\b", "").replace("\\", "") for p in lint(text or "")]


# ---------- public calls ----------
def check_order(identifier, contact_date, app_date=None, path=None) -> dict:
    """Look up an order by number or email and return its status. Never writes."""
    return prepare_reply("status", identifier, contact_date, app_date, path)


def run_action(kind, order_number, contact_date, app_date=None, path=None, **fields) -> dict:
    """Preview a customer request (refund, exchange, ...) on one order. Never writes."""
    if kind not in dict(ACTIONS):
        return _error(f"Unknown action: {kind}.", contact_date, app_date)
    extra = set(fields) - ACTION_FIELDS
    if extra:
        return _error(f"Not available in this interface: {', '.join(sorted(extra))}.", contact_date, app_date)
    params = {"kind": kind, "order_number": order_number, "on": contact_date, **fields}
    return _run(params, app_date, path, persist=False)


def save_action(params: dict, app_date=None, path=None) -> dict:
    """Run the same request again on a freshly loaded workbook and log any new order in New Orders."""
    params = {k: v for k, v in params.items() if k not in ("force_new", "new_order_number")}
    return _run(params, app_date, path, persist=True)


def remember_pending(state, result: dict) -> str | None:
    """Keep a previewed request under a one-time token so it can be saved exactly once."""
    if not result.get("ok") or not result.get("previews"):
        return None
    token = uuid.uuid4().hex
    state.setdefault("pending", {})[token] = result["request"]
    return token


def save_once(state, token, path=None) -> dict:
    """Save the request behind a token. A second call with the same token saves nothing and returns the first result."""
    done = state.setdefault("saved", {})
    if token in done:
        return {**done[token], "repeat": True}
    job = state.setdefault("pending", {}).get(token)
    if job is None:
        return _error("This preview has expired. Click Check Order and try again.", None, None)
    res = save_action(job["params"], job["app_date"], path)
    if res["ok"]:
        del state["pending"][token]
        done[token] = res
    return res


def clear_replacement(original_number, new_number, path=None) -> dict:
    """Demonstrations and retesting: remove one saved reshipment or replacement so its original order can be
    processed again. Refuses anything that is not a saved reshipment/replacement of exactly that original."""
    path = path or WORKBOOK_PATH
    name = os.path.basename(path)
    try:
        book = load(path, app_date=local_today())
        target = next((n for n in book.new_orders if n.number == int(new_number)), None)
        if any(o.number == int(new_number) for o in book.orders):
            return _error(f"Order {new_number} is an original order. Original orders are never cleared.", None, None)
        if target is None:
            return _error(f"Order {new_number} is not a saved order in the workbook. Nothing was changed.", None, None)
        if target.original_number != int(original_number):
            return _error(f"Order {new_number} was not created from order {original_number}. Nothing was changed.",
                          None, None)
        if target.order_type not in workbook.CLEARABLE_TYPES:
            return _error(f"Order {new_number} is an exchange order. Only a reshipment or a wrong or damaged item "
                          f"replacement can be cleared. Nothing was changed.", None, None)
        workbook.remove_new_order(path, target.number, target.original_number)
    except PermissionError:
        log.exception("Workbook is locked")
        return _error(f"Close {name} in Excel and click Confirm clear again.", None, None)
    except Exception as e:  # plain message; the traceback goes to the console
        log.exception("Clear failed")
        return _error(f"Something went wrong: {e}. Check the New Orders sheet before trying again.", None, None)
    return {"ok": True, "error": None, "cleared": target.number, "original": target.original_number,
            "message": f"Saved replacement order {target.number} was cleared from the workbook. Order "
                       f"{target.original_number} can be processed again. No new order was created."}


def timeline(order_number, contact_date, app_date=None, path=None) -> dict:
    """The order's dates in order, the contact date placed where it falls, plus linked New Orders."""
    res = check_order(order_number, contact_date, app_date, path)
    return {k: res.get(k) for k in ("ok", "error", "timeline", "linked_orders")}


# ---------- internals ----------
def _error(message, contact_date, app_date) -> dict:
    return {"ok": False, "error": message, "contact_date": contact_date, "app_date": app_date}


def _clean(params: dict) -> dict:
    out = {}
    for k, v in params.items():
        if isinstance(v, str):
            v = v.strip() or None
        out[k] = v
    if out.get("qty") is not None:
        out["qty"] = int(out["qty"])
    return out


def _run(params, app_date, path, persist) -> dict:
    path = path or WORKBOOK_PATH
    app_date = app_date or local_today()
    on = params.get("on")
    try:
        params = _clean(params)
        book = load(path, app_date=app_date)
        saved_before = list(book.new_orders)    # a preview adds its order to book.new_orders in memory only
        desk = Desk(book, persist_path=path if persist else None)
        req = Request(**params)
        r = desk.handle(req)
        linked_pool = book.new_orders if persist else saved_before
        return _view(r, req, desk, linked_pool, params, app_date, persist)
    except PermissionError:
        log.exception("Workbook is locked")
        name = os.path.basename(path)
        return _error(f"Close {name} in Excel and click Create order again." if persist else
                      f"The workbook could not be opened. Close {name} in Excel and try again.", on, app_date)
    except Exception as e:  # shown as a plain message; the traceback goes to the console
        log.exception("Request failed")
        tail = " Check the New Orders sheet before trying again." if persist else " Nothing was saved."
        return _error(f"Something went wrong: {e}.{tail}", on, app_date)


def _problem(r, req) -> str | None:
    d = r.decision
    if d == "NEED_CONTACT_DATE":
        return "Enter the date the customer contacted you. It is needed to work out the order's status."
    if d == "NO_ID":
        return "Enter an order number (for example 1650) or the customer's email address."
    if d == "NOT_FOUND":
        return (f"No order with the number {req.order_number} was found. Check the number and try again."
                if req.order_number is not None else f"No order was found for the email address {req.email}.")
    if d in ("EMAIL_MISMATCH", "MISSING_DATES", "DATA_ERROR"):
        return r.internal_note
    return None


def _rows(new_order) -> list[dict]:
    if not new_order:
        return []
    return list(new_order["orders"]) if "orders" in new_order else [new_order]


def _plain(v):
    if isinstance(v, dict):
        return {str(k): _plain(x) for k, x in v.items()}
    if isinstance(v, (list, tuple)):
        return [_plain(x) for x in v]
    return v if isinstance(v, (int, float, bool, type(None))) else str(v)


def _is_rolling(o) -> bool:
    return bool(o.order_rule and o.order_rule.rolling)


def _order(desk, o) -> dict:
    return {"number": o.number, "name": o.name, "email": o.email, "address": o.address,
            "destination": o.destination, "method": o.method, "source": o.source,
            "original_number": o.original_number, "order_type": o.order_type, "rolling": _is_rolling(o),
            "items": [{"line": i, "qty": it.qty, "name": desk.name_of(desk.key_or_fallback(it))}
                      for i, it in enumerate(o.items, 1)]}


def _key_dates(a: dict, rolling: bool) -> list[tuple[str, str]]:
    tag = "rolling" if rolling else "fixed"
    since = ("Not shipped yet" if a["state"] == engine.NOT_SHIPPED else
             f"{a['calendar_days_since_shipping']} calendar / {a['business_days_since_shipping']} business")
    late = ("Not late" if not a["calendar_days_past_expected"] else
            f"{a['calendar_days_past_expected']} calendar / {a['business_days_past_expected']} business")
    until = "Eligible now" if a["state"] == engine.RESHIP else str(a["days_until_reship_eligible"])
    return [(f"Order date ({tag})", show(a["order_date"])), (f"Shipping date ({tag})", show(a["ship_date"])),
            ("Expected delivery", show(a["expected_delivery"])), ("Contact-us date", show(a["contact_us_date"])),
            ("Customer contact date", show(a["contact_date"])), ("Days since shipping", since),
            ("Days late", late), ("Days until reship eligible", until)]


def _timeline(a: dict, rolling: bool) -> list[dict]:
    tag = "rolling" if rolling else "fixed"
    events = [(a["order_date"], 0, f"Ordered ({tag})"), (a["ship_date"], 1, f"Shipped ({tag})"),
              (a["expected_delivery"], 2, "Expected delivery"), (a["contact_us_date"], 3, "Contact-us date"),
              (a["contact_date"], 4, "Customer contacted us")]
    return [{"date": d, "label": label, "is_contact": i == 4} for d, i, label in sorted(e for e in events if e[0])]


def _linked(desk, o, on, pool) -> list[dict]:
    out = []
    for n in pool:
        if n.original_number != o.number:
            continue
        row = {"number": n.number, "type": n.order_type, "method": n.method, "created": n.order_date,
               "shipped": n.ship_date, "items": [f"{it.qty} × {desk.name_of(desk.key_or_fallback(it))}" for it in n.items]}
        try:
            a = desk.analyse(n, on)
            row.update(expected_delivery=a["expected_delivery"], contact_us_date=a["contact_us_date"], state=a["state"])
        except KeyError as e:
            row["note"] = str(e)
        out.append(row)
    return out


def _view(r, req, desk, linked_pool, params, app_date, saved) -> dict:
    book = desk.book
    out = {"ok": True, "error": None, "warning": None, "kind": req.kind, "decision": r.decision,
           "contact_date": req.on, "app_date": app_date, "paragraphs": list(r.paragraphs), "matches": [],
           "order": None, "analysis": None, "key_dates": [], "timeline": [], "linked_orders": [],
           "previews": [], "saved_orders": [], "clearable": [], "existing_order": r.facts.get("existing_order"),
           "email": "", "email_withheld": False, "catalogue": [desk.name_of(k) for k in book.prices],
           "request": {"params": params, "app_date": app_date},
           "internal": {"decision": r.decision, "facts": _plain(r.facts), "internal_note": r.internal_note,
                        "warnings": list(book.warnings)}}
    problem = _problem(r, req)
    if problem:
        return {**out, "ok": False, "error": problem}
    if r.decision == "AMBIGUOUS":
        found = {x.number: x for x in book.all_orders()}
        return {**out, "matches": list(r.facts["matches"]),
                "match_details": [{"number": n, "name": found[n].name, "destination": found[n].destination,
                                   "items": "; ".join(i["name"] if i["qty"] == 1 else f"{i['qty']} × {i['name']}"
                                                      for i in _order(desk, found[n])["items"])}
                                  for n in r.facts["matches"]]}

    o, _ = desk.identify(req)          # the order that was looked up (r.order can be an existing new order)
    out["order"] = _order(desk, o)
    if r.decision == "CONTACT_BEFORE_ORDER":
        out["warning"] = (f"The customer contact date ({show(req.on)}) is earlier than order {o.number}'s order date "
                          f"({show(o.order_date)}), which was generated for the application date {show(app_date)}. "
                          f"The application date is probably too late: set it to the day the customer wrote, "
                          f"or check the contact date.")
        return out
    if r.decision not in NO_ANALYSIS:
        a = desk.analyse(o, req.on)
        out["analysis"] = {"state": a["state"], "colour": STATE_COLOURS.get(a["state"], "gray"),
                           "eligible_action": a["eligible_action"]}
        out["key_dates"] = _key_dates(a, _is_rolling(o))
        out["timeline"] = _timeline(a, _is_rolling(o))
        out["linked_orders"] = _linked(desk, o, req.on, linked_pool)
        if o.source == workbook.ORDERS:      # saved reshipments/replacements of this original that may be cleared for retesting
            out["clearable"] = [{"number": n["number"], "original": o.number, "type": n["type"], "items": n["items"],
                                 "created": n["created"]}
                                for n in out["linked_orders"] if n["type"] in workbook.CLEARABLE_TYPES]

    rows = _rows(r.new_order)
    for row in rows:
        row["details"] = _new_order_details(row, o, saved)
    out["saved_orders" if saved else "previews"] = rows
    out["new_order_label"] = "Exchange order" if req.kind == "exchange_tracking" else "Replacement order"
    out["explanation"] = _explanation(r, o, rows, out["analysis"] and a, saved)
    out["subject"] = reply.subject(req.kind, o.number)
    out["email_context"] = reply.context(r, [o.number, out["existing_order"]], new_order_saved=saved)
    if rows and not saved:
        out["email_withheld"] = True     # like the CLI: no email may promise an order that doesn't exist yet
    else:
        out["email"] = compose([r])
        out["email_issues"] = reply.quality_issues(out["email"], out["email_context"]) if out["email"] else []
    if out["analysis"]:
        out["status"] = _status(a, o, linked_pool, rows, req.kind, saved)
    out["recommendation"] = _recommend(r, req, out, rows, saved)
    return out


def _status(a: dict, o, linked_pool, rows, kind, saved=False) -> dict:
    """The order's situation in plain words. Dates and state come from Desk.analyse."""
    state, cu = a["state"], show_long(a["contact_us_date"])
    label, colour = STATUS_PLAIN.get(state, state), STATE_COLOURS.get(state, "gray")
    reshipped = [n.number for n in linked_pool if n.original_number == o.number
                 and n.order_type == workbook.NEW_ORDER_TYPES["lost"]
                 and (saved or n.number not in {r["number"] for r in rows})]   # a preview isn't a real order
    late = f"{_days(a['calendar_days_past_expected'])} past the expected delivery date when the customer wrote."
    if state == engine.NOT_SHIPPED:
        note = "The order had not shipped yet when the customer wrote."
    elif state in (engine.NOT_YET_DUE, engine.DUE):
        note = "The order was within its delivery time when the customer wrote."
    elif state == engine.LATE:
        note = f"{late} {_days(a['days_until_reship_eligible']).capitalize()} until it can be reshipped."
    else:
        note = late
    if kind in ITEM_KINDS and state != engine.NOT_SHIPPED:
        # The customer is writing about an item they have, so lateness isn't the point: just the shipping facts.
        label, colour, note = "Shipped", "gray", ""
    if state == engine.RESHIP and reshipped:
        label = f"Replacement already sent (order {reshipped[-1]})"
        advice = f"The contact-us date has passed and replacement order {reshipped[-1]} was created for this order."
    elif state == engine.RESHIP:
        advice = (f"The contact-us date, {cu}, has passed. The package has not arrived, so the order qualifies "
                  f"for a replacement.")
    else:
        advice = (f"If the package has not arrived by {cu}, the customer should contact us again so we can help "
                  f"resolve the issue.")
    return {"label": label, "colour": colour, "note": note, "state": state, "advice": advice,
            "dates": [("Ships" if state == engine.NOT_SHIPPED else "Shipped", show_long(a["ship_date"])),
                      ("Expected delivery", show_long(a["expected_delivery"])), ("Contact-us date", cu)]}


def _new_order_details(row: dict, o, saved: bool) -> list[tuple[str, str]]:
    """What the agent reviews before confirming. Every value is the engine's."""
    number = str(row["number"]) if saved else f"{row['number']} (given when it is saved, so it can differ from this preview)"
    return [("New order number", number), ("Created from", f"order {o.number}, {o.name}"),
            ("Item", f"{row['qty']} × {row['item']}"), ("Ships", show_long(row["ship"])),
            ("Expected delivery", show_long(row["est"])), ("Contact-us date", show_long(row["contact"])),
            ("Shipping method (internal, not in the email)", row["method"])]


def _plural(n: int, unit: str) -> str:
    return f"{n} {unit[:-1] if n == 1 else unit}"


def _rule_lines(method, destination, x, ship, est, two_x_end, contact) -> list[str]:
    """The 2X + 1 calculation, step by step. The dates are the engine's; dates.add_shipping_days lists the days counted."""
    unit = "business days" if method == "Regular" else "calendar days"
    how = ("Business days skip Saturdays and Sundays." if method == "Regular"
           else "Calendar days include Saturdays and Sundays.")
    counted = ", ".join(f"day {i} {show(dates.add_shipping_days(ship, method, i))}" for i in range(1, 2 * x + 1))
    return [f"Shipping method and destination: {method} shipping to {destination}.",
            f"X for {method} shipping to {destination} is {_plural(x, unit)} (Shipping estimates sheet). {how}",
            f"The shipping date, {show_long(ship)}, is day zero. Counting begins the following day.",
            f"Expected delivery is X, {_plural(x, unit)} after the shipping date: {show_long(est)}.",
            f"2X is {_plural(2 * x, unit)} after the shipping date: {show_long(two_x_end)}.",
            f"Days counted: {counted}.",
            f"Plus 1 calendar day gives the contact-us date: {show_long(contact)}."]


def _explanation(r, o, rows, a, saved) -> list[dict]:
    """Internal only: why each contact-us date is what it is (the 2X + 1 rule)."""
    out = []
    if a:
        out.append({"title": f"Contact-us date for order {o.number}",
                    "lines": _rule_lines(a["method"], a["destination"], a["x"], a["ship_date"], a["expected_delivery"],
                                         a["two_x_end"], a["contact_us_date"])})
    f = r.facts
    if r.order is not None and r.order.number != o.number and "two_x_end" in f:      # an existing new order
        out.append({"title": f"Contact-us date for existing order {r.order.number}",
                    "lines": _rule_lines(f["method"], f["destination"], f["x"], f["ship_date"], f["expected_delivery"],
                                         f["two_x_end"], f["contact_us_date"])})
    for row in rows:
        why = {workbook.NEW_ORDER_TYPES["lost"]: "A reshipment always goes by Regular shipping, whatever the original used.",
               workbook.NEW_ORDER_TYPES["damaged"]: "A wrong or damaged item replacement always goes by Express shipping.",
               workbook.NEW_ORDER_TYPES["exchange"]: "An exchange always goes by Regular shipping."}.get(row["type"], "")
        out.append({"title": f"Contact-us date for new order {row['number']}" + ("" if saved else " (preview)"),
                    "lines": [f"{why} It ships the day it is created."] + _rule_lines(
                        row["method"], o.destination, row["x"], row["ship"], row["est"],
                        dates.add_shipping_days(row["ship"], row["method"], 2 * row["x"]), row["contact"])})
    return out


def _days(n) -> str:
    return f"{n} day" if n == 1 else f"{n} days"


def _recommend(r, req, out, rows, saved) -> dict:
    """What the agent should do now, keyed on the engine's decision. Values come from the engine's facts."""
    f, d = r.facts, r.decision
    label = ("Attach the free return label to this email." if req.label_sent else
             "Send the free return label in a separate email. This email says it will follow.")
    send = "Send the email below."
    if rows and saved:
        numbers = ", ".join(str(x["number"]) for x in rows)
        steps = [f"Order {numbers} is saved in the workbook.", send]
        if req.kind in ("damaged", "wrong"):
            steps.append(label)
        steps += [f"Issue the ${x['voucher']} voucher for the price difference." for x in rows if x.get("voucher")]
        return {"tone": "success", "title": f"Done: order {numbers} created", "steps": steps}
    if rows:
        what = "exchange" if req.kind == "exchange_tracking" else "replacement"
        return {"tone": "warning", "title": f"Create the {what} order ({rows[0]['method']} shipping)",
                "steps": [f"This order qualifies. Review the {what} order below, then click Create order and "
                          f"Confirm and save.",
                          "The customer email appears once the order is saved, so it never promises an order that "
                          "doesn't exist."]}
    if out["existing_order"]:
        n = out["existing_order"]
        steps = [f"Order {n} was already created for this. Do not create another.",
                 f"{send} It gives the customer the status of order {n}."]
        if d.endswith(engine.RESHIP):
            steps.append(f"Order {n} is itself past its contact-us date. Review it before sending anything else.")
        return {"tone": "warning" if d.endswith(engine.RESHIP) else "info",
                "title": f"No new order: order {n} already covers this", "steps": steps}
    if d in (engine.NOT_SHIPPED, engine.NOT_YET_DUE, engine.DUE):
        return {"tone": "info", "title": "No action needed, send the reply",
                "steps": [f"The order is on schedule. {send}",
                          f"A replacement is only possible from {show(f['contact_us_date'])}."]}
    if d == engine.LATE:
        return {"tone": "info", "title": "Do not reship yet, send the reply",
                "steps": [f"The order is late but not yet eligible for a replacement. {send}",
                          f"If the customer writes again on or after {show(f['contact_us_date'])}, check the order again "
                          f"to create the replacement."]}
    if d == "CANCEL_NO" and req.kind != "cancel":
        return {"tone": "warning", "title": "Not possible yet: the order hasn't shipped",
                "steps": [f"On the contact date this order had not shipped (it ships {show(f['ship_date'])}), so there is "
                          f"no item to act on. Check the contact date and the application date.",
                          "If the dates are right, send the email below."]}
    if d == "CANCEL_NO":
        return {"tone": "info", "title": "Change or cancellation unavailable",
                "steps": [f"Orders can't be cancelled or changed once placed. {send}"]}
    if d == "REFUND_OK":
        return {"tone": "success", "title": f"Refund allowed: ${f['refund_amount']} voucher after the return",
                "steps": [send, label, f"Issue the ${f['refund_amount']} voucher only after the customer replies with the "
                                       f"return tracking number."]}
    if d == "REFUND_HOVER":
        steps = [f"Hoverboards can't be refunded. {send}"]
        if f.get("exchange_deadline"):
            steps.append(f"The email offers an exchange instead, available until {show(f['exchange_deadline'])}.")
        return {"tone": "info", "title": "Refund not allowed (hoverboard)", "steps": steps}
    if d == "REFUND_LATE":
        steps = [f"The last day for a refund was {show(f['refund_deadline'])}. {send}"]
        if f.get("exchange_deadline"):
            steps.append(f"The email offers an exchange instead, available until {show(f['exchange_deadline'])}.")
        return {"tone": "info", "title": "Refund not allowed (too late)", "steps": steps}
    if d == "EX_OK":
        steps = [send, label, "Do not create the exchange order yet. When the customer sends the return tracking number, "
                              "choose \"I've sent back my exchange item\" to create it."]
        if f.get("voucher"):
            steps.append(f"A ${f['voucher']} voucher for the price difference goes with the exchange order.")
        return {"tone": "success", "title": "Exchange allowed", "steps": steps}
    if d == "EX_VIB":
        steps = [f"Vibranium items can't be exchanged. {send}"]
        if f.get("refund_deadline"):
            steps.append(f"The email offers a refund instead, available until {show(f['refund_deadline'])}.")
        return {"tone": "info", "title": "Exchange not allowed (vibranium item)", "steps": steps}
    if d == "EX_LATE":
        return {"tone": "info", "title": "Exchange not allowed (too late)",
                "steps": [f"The last day for an exchange was {show(f['exchange_deadline'])}. {send}"]}
    if d == "EX_PRICE":
        return {"tone": "info", "title": "Exchange not allowed (new item costs more)",
                "steps": [f"The {f['new_item']} is ${f['new_unit_price']}; the item bought was ${f['unit_price']}. {send}"]}
    if d == "EX_NEED_ITEM":
        return {"tone": "info", "title": "Ask which item they want instead",
                "steps": ["Choose the item they want above and check again, or send the email below to ask them."]}
    if d == "EX_NEED_TRACKING":
        return {"tone": "info", "title": "Ask for the return tracking number",
                "steps": ["Enter the tracking number above and check again, or send the email below to ask for it."]}
    if d == "DM_LATE":
        return {"tone": "info", "title": "Replacement not allowed (too late)",
                "steps": [f"The last day to report it was {show(f['deadline'])}. {send}"]}
    if d == "PICK_ITEM":
        return {"tone": "info", "title": "Say which item this is about",
                "steps": [f"This order has several items: {f['items']}.",
                          "Fill in \"Which item\" above and check again, or send the email below to ask the customer."]}
    if d == "BAD_QTY":
        return {"tone": "info", "title": "Quantity is more than was ordered",
                "steps": [f"The order has {f['ordered']} of that item. Correct the quantity above and check again."]}
    if d == "PRICE_UNKNOWN":
        return {"tone": "warning", "title": "Item not found in the Price List",
                "steps": [r.internal_note or "Check the item name against the Price List sheet."]}
    return {"tone": "info", "title": d, "steps": [send] if r.paragraphs else []}
