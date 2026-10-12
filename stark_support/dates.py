"""Date rules, mirroring the workbook formulas exactly.

Orders!I  Regular: WORKDAY(ship, X)        Express: ship + X
Orders!J  Regular: WORKDAY(ship, 2X) + 1   Express: ship + 2X + 1
The ship date is day 0; counting starts the next day. The +1 is always a calendar day.
"""
from datetime import date, timedelta


def workday(start: date, days: int) -> date:
    """Excel WORKDAY(start, days) with no holidays: step forward over Mon-Fri only."""
    d = start
    while days > 0:
        d += timedelta(days=1)
        if d.weekday() < 5:
            days -= 1
    return d


def add_shipping_days(ship: date, method: str, days: int) -> date:
    if method == "Regular":
        return workday(ship, days)
    if method == "Express":
        return ship + timedelta(days=days)
    raise ValueError(f"Unknown shipping method: {method!r}")


def estimated_delivery(ship: date, method: str, x: int) -> date:
    return add_shipping_days(ship, method, x)


def contact_us_date(ship: date, method: str, x: int) -> date:
    """2X counted from the shipping date with the method's day type, then +1 calendar day."""
    return add_shipping_days(ship, method, 2 * x) + timedelta(days=1)


def business_deadline(ship: date, business_days: int) -> date:
    """Last allowed request day for refund / damaged (10) and exchange (14) windows."""
    return workday(ship, business_days)


def calendar_days_between(start: date, end: date) -> int:
    return (end - start).days


def business_days_between(start: date, end: date) -> int:
    """Mon-Fri days after `start` up to and including `end` (same counting as WORKDAY). Negative if end < start."""
    if end < start:
        return -business_days_between(end, start)
    n, d = 0, start
    while d < end:
        d += timedelta(days=1)
        if d.weekday() < 5:
            n += 1
    return n


def shipping_days_between(start: date, end: date, method: str) -> int:
    return business_days_between(start, end) if method == "Regular" else calendar_days_between(start, end)
