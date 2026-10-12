"""Turns engine results into one short, human customer email, and checks an email before it is sent."""
from __future__ import annotations

import re
from datetime import date

# Phrases and internals that must never reach a customer.
FORBIDDEN = [r"\b2X\b", r"\bX\b", r"\blost\b", r"WORKDAY", r"formula", r"as per", r"please be advised",
             r"we regret to inform", r"your request has been noted", r"we value your business",
             r"apologize for any inconvenience", r"\bpolicy\b"]

BLOCKING = {"AMBIGUOUS", "NOT_FOUND", "NO_ID", "EMAIL_MISMATCH"}

# First line: acknowledges what the customer wrote about. Left out when the answer already opens that way.
OPENERS = {
    "status": "Thanks for checking in on your order.",
    "lost": "Thanks for letting us know about your order.",
    "refund": "Thanks for getting in touch about a refund.",
    "exchange": "Thanks for getting in touch about an exchange.",
    "exchange_tracking": "Thanks for getting in touch about your exchange.",
    "damaged": "Thanks for getting back to us.",
    "wrong": "Thanks for getting back to us.",
    "cancel": "Thanks for getting in touch about your order.",
}
SUBJECTS = {
    "status": "Your order {n}: delivery update",
    "lost": "Your order {n}: delivery update",
    "refund": "Your refund request for order {n}",
    "exchange": "Your exchange request for order {n}",
    "exchange_tracking": "Your exchange for order {n}",
    "damaged": "About the damaged item from order {n}",
    "wrong": "About the wrong item from order {n}",
    "cancel": "Your request to cancel or change order {n}",
}
# The email must talk about what the customer asked.
TOPIC = {
    "status": r"arriv|ship", "lost": r"arriv|replacement", "refund": r"refund", "exchange": r"exchang",
    "exchange_tracking": r"exchang|tracking", "damaged": r"damaged|replacement", "wrong": r"wrong item|replacement",
    "cancel": r"cancel",
}
_MONTHS = "January|February|March|April|May|June|July|August|September|October|November|December"
_CLOSINGS = r"^(Best|Best regards|Kind regards|Regards|Thanks|Thank you|Sincerely|Warm regards),?$"


def compose(results, first_name: str | None = None, sign_off: str = "Stark Industries Support") -> str:
    """One greeting, every answer in order, one closing. If the order can't be identified,
    only the identification question is sent, since nothing else can be answered yet."""
    blocking = [r for r in results if r.decision in BLOCKING]
    use = blocking[:1] if blocking else results
    if first_name is None:
        first_name = next((r.order.first_name for r in use if r.order), "there")
    paras, seen = [], set()
    for r in use:
        for p in r.paragraphs:
            if p not in seen:
                seen.add(p)
                paras.append(p)
    body = "\n\n".join(paras)
    if not body:
        return ""  # internal-only result: nothing to send
    opener = "Thanks for reaching out." if blocking else OPENERS.get(use[0].request.kind, "Thanks for reaching out.")
    opener = "" if body.startswith(("Thanks", "I'm sorry")) else f"{opener}\n\n"
    return f"Hi {first_name},\n\n{opener}{body}\n\nBest,\n{sign_off}"


def subject(kind: str, order_number: int | None) -> str:
    if order_number is None:
        return "Your order enquiry"
    return SUBJECTS.get(kind, "Your order {n}").format(n=order_number)


def lint(text: str) -> list[str]:
    return [p for p in FORBIDDEN if re.search(p, text, flags=re.IGNORECASE)]


def _walk(v):
    if isinstance(v, dict):
        for x in v.values():
            yield from _walk(x)
    elif isinstance(v, (list, tuple)):
        for x in v:
            yield from _walk(x)
    else:
        yield v


MONEY_KEYS = {"unit_price", "new_unit_price", "refund_amount", "order_value", "remaining_value", "voucher"}


def _money(v):
    if isinstance(v, dict):
        for k, x in v.items():
            if k in MONEY_KEYS and isinstance(x, (int, float)) and not isinstance(x, bool):
                yield int(x)
            else:
                yield from _money(x)
    elif isinstance(v, (list, tuple)):
        for x in v:
            yield from _money(x)


def context(result, order_numbers=(), new_order_saved: bool = False) -> dict:
    """What an email about this result may state: the dates, amounts and order numbers the engine produced."""
    values = list(_walk(result.facts)) + list(_walk(result.new_order or {}))
    numbers = {int(n) for n in order_numbers if n is not None}
    if result.order:
        numbers.add(result.order.number)
        if result.order.original_number:
            numbers.add(result.order.original_number)
    for row in ([result.new_order] if result.new_order and "orders" not in result.new_order
                else (result.new_order or {}).get("orders", [])):
        numbers.add(row["number"])
    return {"kind": result.request.kind,
            "dates": sorted({f"{d:%B} {d.day}" for d in values if isinstance(d, date)}),
            "amounts": sorted(set(_money(result.facts)) | set(_money(result.new_order or {}))),
            "orders": sorted(numbers),
            "new_order_saved": bool(new_order_saved and result.new_order)}


def quality_issues(text: str, ctx: dict) -> list[str]:
    """Problems to fix before an email is sent. Empty list = it passed every check."""
    issues = []
    lines = [ln.strip() for ln in (text or "").strip().splitlines() if ln.strip()]
    if not lines:
        return ["The email is empty."]
    if not re.match(r"^(Hi|Hello|Dear)\b.+", lines[0]):
        issues.append("It doesn't start with a greeting (for example \"Hi Clint,\").")
    if not any(re.match(_CLOSINGS, ln, flags=re.IGNORECASE) for ln in lines[-3:]):
        issues.append("It doesn't end with a closing (for example \"Best,\" and your name).")
    if len(lines) < 4:
        issues.append("It has no message between the greeting and the closing.")
    found = [p.replace(r"\b", "").replace("\\", "") for p in lint(text)]
    if found:
        issues.append("It contains internal terms that must not reach a customer: " + ", ".join(found) + ".")
    topic = TOPIC.get(ctx.get("kind"))
    if topic and not re.search(topic, text, flags=re.IGNORECASE):
        issues.append("It no longer seems to answer the customer's question.")
    dates_ok = set(ctx.get("dates", []))
    bad = sorted({f"{m} {d}" for m, d in re.findall(rf"\b({_MONTHS}) (\d{{1,2}})\b", text)} - dates_ok)
    if bad:
        issues.append("It mentions a date the application did not calculate for this order: " + ", ".join(bad) + ".")
    bad = sorted({int(a) for a in re.findall(r"\$(\d+)", text)} - set(ctx.get("amounts", [])))
    if bad:
        issues.append("It mentions an amount the application did not calculate: " + ", ".join(f"${a}" for a in bad) + ".")
    bad = sorted({int(n) for n in re.findall(r"\border(?: number)?(?: is)? #?(\d{3,})", text, flags=re.IGNORECASE)}
                 - set(ctx.get("orders", [])))
    if bad:
        issues.append("It mentions an order number that isn't part of this result: " + ", ".join(map(str, bad)) + ".")
    if not ctx.get("new_order_saved") and re.search(
            r"(new|replacement|exchange) order number|(I've|I have|we've|we have) (already )?set up|"
            r"exchange order .{0,80}is all set",
            text, flags=re.IGNORECASE):
        issues.append("It says a new order exists, but no new order was saved for this result.")
    if re.search(r"(refund|voucher) (has been|was|is being) (processed|issued|sent)", text, flags=re.IGNORECASE):
        issues.append("It says a refund or voucher was already issued. The application never confirms that.")
    return issues
