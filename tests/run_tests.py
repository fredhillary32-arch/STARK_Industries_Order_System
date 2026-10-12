"""The 24 required tests. Expected values were counted by hand on a calendar, not produced by the engine.

Runs on temporary copies of the sanitized demo workbook; the source workbook is never modified.
Usage: python3 tests/run_tests.py   (writes tests/test_report.md)
"""
import os
import re
import shutil
import subprocess
import sys
import tempfile
from datetime import date, timedelta

import openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
from stark_support import Desk, Request, compose, load  # noqa: E402
from stark_support import engine as E  # noqa: E402
from stark_support.reply import lint  # noqa: E402

BOOK = os.path.join(ROOT, "data", "stark_orders_demo.xlsx")


def find_soffice():
    """LibreOffice (used only by test R2): on PATH, or in its standard Windows install folders."""
    found = shutil.which("soffice") or shutil.which("soffice.exe")
    if found:
        return found
    for base in (os.environ.get("ProgramFiles", r"C:\Program Files"), os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")):
        exe = os.path.join(base, "LibreOffice", "program", "soffice.exe")
        if os.path.exists(exe):
            return exe
    return None
D = lambda m, d: date(2026, m, d)  # noqa: E731
TMP = tempfile.mkdtemp(prefix="stark-tests-")
results = []


APP_A = D(10, 8)    # application date (what TODAY() means) for the policy tests; 1650 = Sep 29 / Sep 30 on this day
APP_B = D(10, 30)   # second application date for the rolling tests
# Order-date offsets read by hand from the original formulas (G13 =TODAY(), G2 =G13 - 25, ...)
FORMULA_OFFSETS = {1024: -25, 1545: -20, 1614: -15, 1628: -10, 1650: -9, 1655: -6, 1814: -5, 1855: -5,
                   2314: -3, 3514: -3, 3520: -1, 5020: 0}


def fixture_copy():
    """Temp copy with two extra rolling test orders (dates shown for application date Oct 8):
    9001  multi-item, =G13-6 (Fri Oct 2), shipping =G14+3 (Mon Oct 5, not order date + 1), New York Regular
    9002  =G13-6 (Fri Oct 2), shipping date blank (customer-created rule: order date + 1), Toronto Regular"""
    p = os.path.join(TMP, f"fixture{len(os.listdir(TMP))}.xlsx")
    shutil.copy(BOOK, p)
    wb = openpyxl.load_workbook(p)
    demo_emails = {
        2: "multi-order-c@example.com", 3: "multi-order-a@example.com",
        5: "multi-order-a@example.com", 6: "single-order@example.com",
        7: "multi-order-b@example.com", 9: "multi-order-b@example.com",
        10: "multi-order-b@example.com", 11: "multi-order-c@example.com",
        12: "multi-order-a@example.com",
    }
    for row, email in demo_emails.items():
        wb["Orders"].cell(row, 2).value = email
    demo_names = {
        2: "Norman Osbourne", 3: "Clint Barton", 5: "Clint Barton",
        6: "M'Baku", 7: "Scott Lang", 8: "Nat", 9: "Scott Lang", 10: "Scott Lang",
        11: "Norman Osbourne", 12: "Bruce Banner", 13: "Miles Morales",
    }
    for row, name in demo_names.items():
        wb["Orders"].cell(row, 3).value = name
    for row in wb["New Orders"].iter_rows(min_row=2):   # tests start with no saved new orders, whatever the live workbook holds
        for cell in row:
            cell.value = None
    wb["New Orders"]["N1"] = wb["New Orders"]["O1"] = None   # ...and no recorded "highest order number ever issued"
    ws = wb["Orders"]
    for r, row in ((14, [9001, "pepper@example.com", "Pepper Potts", "Stark Tower, New York",
                         "2 hoverboard, grey; 3 quiver, aluminum; 1 helmet, regular", "Regular", "=G13-6", "=G14+3"]),
                   (15, [9002, "happy@example.com", "Happy Hogan", "1 Bay St., Toronto", "1 taser, low voltage", "Regular",
                         "=G13-6", None])):
        for c, v in enumerate(row, 1):
            ws.cell(r, c, v)
    wb.save(p)
    return p


def desk(path=None, persist=False):
    p = path or fixture_copy()
    return Desk(load(p, app_date=APP_A), persist_path=p if persist else None)


def record(n, title, inp, expected, actual, ok, why, reply=None):
    bad = lint(reply) if reply else []
    if bad:
        ok, why = False, f"{why} Customer text contains forbidden terms: {bad}"
    results.append((n, title, inp, expected, actual, "SKIPPED" if ok is None else "PASS" if ok else "FAIL", why, reply))


def F(r, *keys):
    return "; ".join(f"{k}={r.facts.get(k)}" for k in keys)


def run():
    # 1 Normal Regular
    d = desk(); r = d.handle(Request("status", 5020, on=D(10, 12)))
    ok = (r.decision == E.NOT_YET_DUE and r.facts["ship_date"] == D(10, 9) and r.facts["expected_delivery"] == D(10, 13)
          and r.facts["contact_us_date"] == D(10, 16) and r.facts["business_days_since_shipping"] == 1
          and r.facts["calendar_days_since_shipping"] == 3)
    record(1, "Normal Regular delivery", "5020 New York Regular, shipped Fri Oct 9; customer contacts Mon Oct 12",
           "Due Tue Oct 13 (2 business days, weekend skipped); contact-us Fri Oct 16; 3 calendar / 1 business day elapsed; not yet due",
           F(r, "state", "expected_delivery", "contact_us_date", "calendar_days_since_shipping", "business_days_since_shipping"),
           ok, "Business days skip Sat/Sun.", compose([r]))

    # 2 Normal Express
    d = desk(); r = d.handle(Request("status", 3514, on=D(10, 7)))
    ok = (r.decision == E.DUE and r.facts["expected_delivery"] == D(10, 7) and r.facts["contact_us_date"] == D(10, 9))
    record(2, "Normal Express delivery", "3514 New York Express, shipped Tue Oct 6; contacts Wed Oct 7",
           "Due Wed Oct 7 (1 calendar day) → due today; contact-us Fri Oct 9",
           F(r, "state", "expected_delivery", "contact_us_date"), ok, "Express counts calendar days.", compose([r]))

    # 3 Late but not eligible
    d = desk(); r = d.handle(Request("status", 1650, on=D(10, 8)))
    ok = (r.decision == E.LATE and r.facts["calendar_days_past_expected"] == 2 and r.facts["business_days_past_expected"] == 2
          and r.new_order is None and "Tuesday, October 13" in compose([r]))
    record(3, "Late but not yet eligible for reshipment", "1650, contacts Thu Oct 8",
           "Due Oct 6, so 2 days late; not eligible until Tue Oct 13; no new order",
           F(r, "state", "calendar_days_past_expected", "business_days_past_expected", "contact_us_date") + f"; new_order={r.new_order}",
           ok, "Late is not the same as reship-eligible.", compose([r]))

    # 4 Reship eligible (on the threshold day)
    d = desk(); r = d.handle(Request("status", 1650, on=D(10, 13)))
    n = r.new_order or {}
    ok = (r.decision == E.RESHIP and n.get("method") == "Regular" and n.get("ship") == D(10, 13)
          and n.get("est") == D(10, 19) and n.get("contact") == D(10, 24))
    record(4, "Reship eligible", "1650, contacts Tue Oct 13 (the contact-us date)",
           "Eligible; new Regular order ships Oct 13, due Mon Oct 19, contact-us Sat Oct 24",
           f"{r.decision}; new #{n.get('number')} {n.get('method')} ship {n.get('ship')} est {n.get('est')} contact {n.get('contact')}",
           ok, "On the contact-us date counts as eligible.", compose([r]))

    # 5 Far after the deadline
    d = desk(); r = d.handle(Request("status", 1024, on=D(10, 30)))
    n = r.new_order or {}
    ok = (r.decision == E.RESHIP and r.facts["calendar_days_since_shipping"] == 46 and r.facts["calendar_days_past_expected"] == 45
          and n.get("method") == "Regular" and n.get("est") == D(11, 3) and n.get("contact") == D(11, 6))
    record(5, "Customer contacts far after the deadline", "1024 New York Express, shipped Mon Sep 14; contacts Fri Oct 30",
           "46 calendar days since shipping, 45 past the Sep 15 due date; reship Regular from Oct 30: due Tue Nov 3, contact-us Fri Nov 6",
           F(r, "calendar_days_since_shipping", "calendar_days_past_expected") + f"; new {n.get('method')} est {n.get('est')} contact {n.get('contact')}",
           ok, "Still reshipped; the original Express doesn't carry over.", compose([r]))

    # 6 Multiple orders under one email
    d = desk(); r = d.handle(Request("status", email="multi-order-a@example.com", on=D(10, 8)))
    txt = compose([r])
    ok = r.decision == "AMBIGUOUS" and r.facts["matches"] == [1545, 1628, 3520] and not re.search(r"\b(1545|1628|3520)\b", txt)
    record(6, "Multiple orders under one email", "multi-order-a@example.com (Clint Barton ×2, Bruce Banner ×1), no order number",
           "No guess; asks for the order number; no other order details revealed", f"{r.decision}; matches={r.facts.get('matches')}",
           ok, "Email is not an order identifier.", txt)

    # 7 Partial refund
    d = desk(); r = d.handle(Request("refund", 9001, item="quiver", qty=1, on=D(10, 8)))
    o = r.order
    ok = r.decision == "REFUND_OK" and r.facts["qty"] == 1 and [i.qty for i in o.items] == [2, 3, 1]
    record(7, "Partial refund", "9001 (2 hoverboard grey; 3 quiver aluminum; 1 helmet regular), return 1 quiver, Oct 8",
           "Only 1 quiver processed; order lines unchanged (2/3/1)", f"{r.decision}; qty={r.facts.get('qty')}; lines={[i.qty for i in o.items]}",
           ok, "Requested quantity only.", compose([r]))

    # 8 Exact refund dollar amount
    d = desk(); r = d.handle(Request("refund", 9001, item="quiver", qty=2, on=D(10, 8)))
    txt = compose([r])
    ok = r.facts.get("refund_amount") == 40 and "$40 voucher" in txt
    record(8, "Exact refund dollar amount", "9001, return 2 of 3 quivers at $20", "Email says $40 voucher (2 × $20)",
           f"refund_amount={r.facts.get('refund_amount')}; '$40 voucher' in email={'$40 voucher' in txt}", ok,
           "Amount stated in dollars.", txt)

    # 9 Remaining order value
    d = desk(); r = d.handle(Request("refund", 9001, item="helmet", on=D(10, 8), ask_remaining=True))
    txt = compose([r])
    ok = (r.facts["order_value"] == 389 and r.facts["refund_amount"] == 29 and r.facts["remaining_value"] == 360
          and "$29" in txt and "$389" in txt and "$360" in txt)
    record(9, "Remaining order value", "9001, return the helmet ($29) and ask what's left",
           "Order $389 (2×150 + 3×20 + 29); refund $29; remaining $360, all stated",
           F(r, "order_value", "refund_amount", "remaining_value"), ok, "Refund, total and remaining kept distinct.", txt)

    # 10 Hoverboard restriction (partial, other items unaffected)
    d = desk(); r = d.handle(Request("refund", 1655, qty=2, on=D(10, 8)))
    r2 = d.handle(Request("refund", 9001, item="helmet", on=D(10, 8)))
    txt = compose([r])
    ok = r.decision == "REFUND_HOVER" and "Hoverboards can't be returned for a refund" in txt and "exchange" in txt and r2.decision == "REFUND_OK"
    record(10, "Hoverboard refund restriction", "1655 return 2 of 5 hoverboards (Oct 8); plus 9001 helmet on the same order as hoverboards",
           "Hoverboards refused with the reason, exchange offered (by Oct 22); helmet on 9001 still refundable",
           f"1655 {r.decision}; 9001 helmet {r2.decision}", ok, "Restriction applies per item.", txt)

    # 11 Vibranium original
    d = desk(); r = d.handle(Request("exchange", 1650, new_item="Quiver (Aluminum)", on=D(10, 8)))
    r2 = d.handle(Request("exchange", 3520, new_item="Quiver (Titanium)", on=D(10, 9)))
    txt = compose([r])
    ok = r.decision == "EX_VIB" and "Because the Vibranium Plate (Red) you bought is made of vibranium, it can't be exchanged" in txt and r2.decision == "EX_VIB"
    record(11, "Vibranium original exchange restriction", "1650 Vibranium Plate → Quiver; 3520 Recurve Bow (Vibranium) → Quiver",
           "Both refused, explaining the customer's own item is vibranium; 1650 offered a refund by Oct 14",
           f"1650 {r.decision}; 3520 {r2.decision}", ok, "Specific reason given.", txt)

    # 12 Non-vibranium → vibranium
    d = desk(); r = d.handle(Request("exchange", 1814, qty=1, new_item="Vibranium Plate (Red)", on=D(10, 8)))
    txt = compose([r])
    ok = r.decision == "EX_OK" and r.facts["voucher"] == 3 and "$3 voucher" in txt
    record(12, "Non-vibranium to vibranium exchange", "1814 Taser (High Voltage) $15 → Vibranium Plate (Red) $12, 1 of 3",
           "Allowed; $3 voucher", f"{r.decision}; voucher={r.facts.get('voucher')}", ok,
           "Vibranium as the new item is fine.", txt)

    # 13 Partial exchange (offer, then tracking → new order)
    d = desk(); r = d.handle(Request("exchange", 1628, qty=2, new_item="Waterproof Spray (Medium)", on=D(10, 8)))
    r2 = d.handle(Request("exchange_tracking", 1628, qty=2, new_item="Waterproof Spray (Medium)", offered_on=D(10, 8),
                          tracking_number="1ZTEST", on=D(10, 12)))
    n = r2.new_order or {}
    ok = (r.decision == "EX_OK" and r.facts["voucher"] == 10 and r2.decision == "EX_CREATED" and n.get("qty") == 2
          and n.get("method") == "Regular" and n.get("est") == D(10, 14) and n.get("contact") == D(10, 17)
          and r2.order.items[0].qty == 5)
    record(13, "Partial exchange", "1628 (5 quivers titanium $20, New Jersey): exchange 2 for Waterproof Spray (Medium) $15; "
           "offered Oct 8, tracking received Mon Oct 12",
           "2 of 5 exchanged, $10 voucher (2 × $5); new Regular order for 2 sprays ships Oct 12, due Wed Oct 14, contact-us Sat Oct 17; "
           "original still shows 5", f"{r.decision} voucher {r.facts.get('voucher')}; {r2.decision} qty {n.get('qty')} {n.get('method')} "
           f"est {n.get('est')} contact {n.get('contact')}; original qty {r2.order.items[0].qty}", ok,
           "Deadline checked against the offer date.", compose([r]) + "\n\n--- after tracking ---\n\n" + compose([r2]))

    # 14 Wrong product
    d = desk(); r = d.handle(Request("wrong", 3520, on=D(10, 12)))
    n = r.new_order or {}
    ok = (r.decision == "DM_CREATED" and n.get("method") == "Express" and n.get("ship") == D(10, 12)
          and n.get("est") == D(10, 13) and n.get("contact") == D(10, 15) and "wrong item" in compose([r]))
    record(14, "Wrong product replacement", "3520 Montreal Regular, shipped Oct 8; wrong item reported Mon Oct 12",
           "Express replacement ships Oct 12, due Tue Oct 13, contact-us Thu Oct 15; return label for the wrong item",
           f"{r.decision}; {n.get('method')} ship {n.get('ship')} est {n.get('est')} contact {n.get('contact')}", ok,
           "Wrong-item replacements ship Express.", compose([r]))

    # 15 Damaged product
    d = desk(); r = d.handle(Request("damaged", 2314, qty=1, on=D(10, 9)))
    n = r.new_order or {}
    ok = (r.decision == "DM_CREATED" and n.get("method") == "Express" and n.get("qty") == 1
          and n.get("est") == D(10, 11) and n.get("contact") == D(10, 14))
    record(15, "Damaged product replacement", "2314 Toronto Regular (3 hoverboards), 1 damaged, reported Fri Oct 9",
           "Express replacement for 1: ships Oct 9, due Sun Oct 11 (calendar days), contact-us Wed Oct 14",
           f"{r.decision}; {n.get('method')} qty {n.get('qty')} est {n.get('est')} contact {n.get('contact')}", ok,
           "Hoverboards can still be replaced when damaged.", compose([r]))

    # 16 Multiple new orders for one customer
    p = fixture_copy(); d = desk(p, persist=True)
    a = d.handle(Request("status", 1655, on=D(10, 9)))
    b = d.handle(Request("status", 1855, on=D(10, 9)))
    c = d.handle(Request("damaged", 2314, qty=1, on=D(10, 9)))
    again = d.handle(Request("status", 1655, on=D(10, 12)))
    reloaded = load(p, app_date=APP_A)
    nums = [x.new_order["number"] for x in (a, b, c)]
    logged = [(o.number, o.original_number, o.method) for o in reloaded.new_orders]
    originals = {o.number: (o.ship_date, [i.qty for i in o.items]) for o in reloaded.orders if o.number in (1655, 1855, 2314)}
    ok = (len(set(nums)) == 3 and logged == [(nums[0], 1655, "Regular"), (nums[1], 1855, "Regular"), (nums[2], 2314, "Express")]
          and originals == {1655: (D(10, 3), [5]), 1855: (D(10, 4), [5]), 2314: (D(10, 6), [3])}
          and again.decision.startswith("ALREADY_RESHIPPED") and len(reloaded.new_orders) == 3)
    record(16, "Multiple new orders for one customer", "Scott Lang (multi-order-b@example.com): 1655 and 1855 not arrived on Oct 9, "
           "2314 one hoverboard damaged on Oct 9; then asks about 1655 again on Oct 12",
           "3 separate new orders with unique numbers (2 Regular reships, 1 Express replacement), logged in New Orders; originals "
           "unchanged; the repeat question on 1655 points to its existing replacement instead of creating a fourth",
           f"numbers={nums}; logged={logged}; originals={originals}; repeat={again.decision}", ok,
           "No limit on new orders, no duplicates either.", compose([again]))

    # 17 Weekend date calculation
    d = desk()
    s1 = d.analyse(next(o for o in d.book.orders if o.number == 1545), D(9, 21))
    s2 = d.analyse(next(o for o in d.book.orders if o.number == 3520), D(10, 8))
    s3 = d.analyse(next(o for o in d.book.orders if o.number == 1614), D(9, 26))
    ok = (s1["expected_delivery"] == D(9, 22) and s1["contact_us_date"] == D(9, 25) and s2["two_x_end"] == D(10, 16)
          and s2["contact_us_date"] == D(10, 17) and s3["expected_delivery"] == D(9, 27) and s3["contact_us_date"] == D(10, 1))
    record(17, "Weekend date calculation",
           "1545 NJ Regular shipped Sat Sep 19; 3520 Montreal Regular shipped Thu Oct 8; 1614 Albuquerque Express shipped Thu Sep 24",
           "1545: due Tue Sep 22, contact-us Fri Sep 25. 3520: 2X ends Fri Oct 16, +1 → Sat Oct 17. "
           "1614: due Sun Sep 27, contact-us Thu Oct 1",
           f"1545 {s1['expected_delivery']}/{s1['contact_us_date']}; 3520 {s2['two_x_end']}/{s2['contact_us_date']}; "
           f"1614 {s3['expected_delivery']}/{s3['contact_us_date']}", ok,
           "Saturday ship day is day 0; +1 can land on a weekend; Express counts weekends.",
           compose([d.handle(Request("status", 1614, on=D(9, 26)))]))

    # 18 Business-day calculation
    d = desk(); a = d.analyse(next(o for o in d.book.orders if o.number == 1650), D(10, 30))
    rf = d.handle(Request("refund", 1650, on=D(10, 14))); rf2 = d.handle(Request("refund", 1650, on=D(10, 15)))
    ok = (a["business_days_since_shipping"] == 22 and a["calendar_days_since_shipping"] == 30 and a["two_x_end"] == D(10, 12)
          and rf.decision == "REFUND_OK" and rf2.decision == "REFUND_LATE" and rf.facts["refund_deadline"] == D(10, 14))
    record(18, "Business-day date calculation", "1650 shipped Wed Sep 30: business days to Oct 30; refund on Oct 14 vs Oct 15",
           "22 business days vs 30 calendar days; 2X (8 business days) ends Mon Oct 12; refund window ends Wed Oct 14 "
           "(OK on the 14th, late on the 15th)",
           f"business={a['business_days_since_shipping']} calendar={a['calendar_days_since_shipping']} 2X end={a['two_x_end']}; "
           f"Oct 14 {rf.decision}, Oct 15 {rf2.decision}", ok, "Business and calendar counts kept separate.", compose([rf2]))

    # 19 Customer-created shipping date
    d = desk()
    o2 = next(o for o in d.book.orders if o.number == 9002)
    o1 = next(o for o in d.book.orders if o.number == 9001)
    ok = (o2.ship_date == D(10, 3) and o2.ship_date_inferred and o1.ship_date == D(10, 5) and not o1.ship_date_inferred)
    r = d.handle(Request("status", 9002, on=D(10, 5)))
    record(19, "Customer-created shipping date",
           "9002 ordered Fri Oct 2 with no shipping date recorded; 9001 ordered Fri Oct 2 with Mon Oct 5 recorded",
           "9002 ships Sat Oct 3 (day after ordering); 9001 keeps its recorded Oct 5 (not recalculated)",
           f"9002 ship {o2.ship_date} (inferred={o2.ship_date_inferred}); 9001 ship {o1.ship_date} (inferred={o1.ship_date_inferred})",
           ok, "Rule only fills a missing date.", compose([r]))

    # 20 Company-created shipping date
    p = fixture_copy(); d = desk(p, persist=True)
    r = d.handle(Request("status", 1628, on=D(10, 8)))
    num = r.new_order["number"]
    reloaded = load(p, app_date=APP_A)
    new = next(o for o in reloaded.new_orders if o.number == num)
    later = Desk(reloaded).handle(Request("status", num, on=D(10, 13)))
    ok = (new.order_date == D(10, 8) and new.ship_date == D(10, 8) and later.facts["expected_delivery"] == D(10, 12)
          and later.facts["contact_us_date"] == D(10, 15) and later.decision == E.LATE)
    record(20, "Company-created shipping date", "1628 reshipped on Thu Oct 8 (saved to New Orders), then asked about on Tue Oct 13",
           "New order created and shipped Oct 8 (same day); due Mon Oct 12, contact-us Thu Oct 15; on Oct 13 it's late, not yet reshippable",
           f"created {new.order_date}, shipped {new.ship_date}; est {later.facts['expected_delivery']}; contact {later.facts['contact_us_date']}; {later.decision}",
           ok, "Company orders ship the day they're made.", compose([later]))

    # 21 Dynamic contact date
    d = desk()
    seq = {D(10, 5): E.NOT_YET_DUE, D(10, 6): E.DUE, D(10, 7): E.LATE, D(10, 8): E.LATE, D(10, 13): E.RESHIP, D(10, 30): E.RESHIP}
    got = {k: d.analyse(next(o for o in d.book.orders if o.number == 1650), k)["state"] for k in seq}
    missing = d.handle(Request("status", 1650))
    today = d.analyse(next(o for o in d.book.orders if o.number == 1650), date.today())
    ok = got == seq and missing.decision == "NEED_CONTACT_DATE" and today["contact_date"] == date.today()
    record(21, "Dynamic customer contact date", "1650 evaluated on Oct 5, 6, 7, 8, 13, 30; no date given; today chosen explicitly",
           "Not yet due / due / late / late / reship / reship; no date → asks for one; today only when chosen",
           f"{ {str(k): v for k, v in got.items()} }; no date → {missing.decision}; today → {today['contact_date']}", ok,
           "Contact date drives the result.")

    # 22 Order 1650, contact date entered manually
    d = desk(); r = d.handle(Request("status", 1650, on=D(10, 7)))
    ok = r.decision == E.LATE and r.facts["contact_date"] == D(10, 7) and r.facts["calendar_days_past_expected"] == 1
    record(22, "Order 1650 with a manually entered contact date", "1650, contact date Wed Oct 7",
           "1 day past the Oct 6 due date; late, reship possible from Oct 13",
           F(r, "contact_date", "state", "calendar_days_past_expected", "contact_us_date"), ok, "Entered date respected.", compose([r]))

    # 23 Order 1650 contacted Oct 30
    d = desk(); r = d.handle(Request("status", 1650, on=D(10, 30)))
    n = r.new_order or {}
    f = r.facts
    ok = (f["order_date"] == D(9, 29) and f["ship_date"] == D(9, 30) and f["x"] == 4 and f["expected_delivery"] == D(10, 6)
          and f["two_x_end"] == D(10, 12) and f["contact_us_date"] == D(10, 13) and f["calendar_days_since_shipping"] == 30
          and f["business_days_since_shipping"] == 22 and f["calendar_days_past_expected"] == 24
          and f["business_days_past_expected"] == 18 and r.decision == E.RESHIP and n.get("method") == "Regular"
          and n.get("ship") == D(10, 30) and n.get("est") == D(11, 5) and n.get("contact") == D(11, 12))
    record(23, "Order 1650 contacted October 30, 2026",
           "1650: ordered Sep 29, shipped Sep 30, Johannesburg Regular (4 business days); contact date Oct 30",
           "Due Tue Oct 6; 2X ends Mon Oct 12 → contact-us Tue Oct 13; 30 calendar / 22 business days since shipping; "
           "24 calendar / 18 business days late; eligible; reship Regular ships Fri Oct 30, due Thu Nov 5, contact-us Thu Nov 12",
           F(r, "order_date", "ship_date", "x", "expected_delivery", "two_x_end", "contact_us_date", "calendar_days_since_shipping",
             "business_days_since_shipping", "calendar_days_past_expected", "business_days_past_expected", "state")
           + f"; reship {n.get('method')} ship {n.get('ship')} est {n.get('est')} contact {n.get('contact')}",
           ok, "Calculated from Sep 30 + Johannesburg + Regular + policy + Oct 30.", compose([r]))

    # 24 Stale calculated date in the workbook
    up = openpyxl.load_workbook(open(BOOK, "rb"))
    stale_est, stale_cu = up["Orders 2"]["I6"].value.date(), up["Orders 2"]["J6"].value.date()
    clean = openpyxl.load_workbook(BOOK)["Orders"]
    clean_headers = [c.value for c in clean[1] if c.value]
    clean_formulas = sorted({c.column_letter for row in clean.iter_rows() for c in row if isinstance(c.value, str) and c.value.startswith("=")})
    d = desk(); r = d.handle(Request("status", 1650, on=D(10, 8)))
    ok = (stale_est == D(10, 1) and stale_cu == D(10, 8) and r.facts["expected_delivery"] == D(10, 6)
          and r.facts["contact_us_date"] == D(10, 13) and r.decision == E.LATE and r.new_order is None
          and clean_headers[-1] == "Shipping date" and clean_formulas == ["G", "H"])
    record(24, "Pre-filled calculated dates in the workbook are not used",
           f"Workbook 'Orders 2' row 1650 has typed delivery {stale_est} and 2X+1 {stale_cu}; application date Oct 8, customer contacts Oct 8",
           "System ignores those cells: due Oct 6, contact-us Oct 13, so on Oct 8 it's late but NOT reshipped "
           "(the pre-filled Oct 8 would have wrongly triggered a reship). Clean Orders keeps only raw data plus the rolling "
           "order/shipping formulas (G, H); no derived columns",
           f"est {r.facts['expected_delivery']}; contact {r.facts['contact_us_date']}; {r.decision}; new order {r.new_order}; "
           f"Orders columns: {clean_headers}; formula columns: {clean_formulas}", ok,
           "Derived dates always come from the engine.", compose([r]))


    rolling_tests()
    new_order_tests()


def rolling_tests():
    """R1-R5: the order dates themselves roll with TODAY(), exactly as the workbook formulas say."""
    a, b, now = load(BOOK, app_date=APP_A), load(BOOK, app_date=APP_B), load(BOOK)
    ga = {o.number: o for o in a.orders}
    gb = {o.number: o for o in b.orders}
    gn = {o.number: o for o in now.orders}

    # R1 every order follows its own formula offset on both application dates
    rows, ok = [], True
    for n, off in FORMULA_OFFSETS.items():
        oa, ob = ga[n], gb[n]
        good = (oa.order_rule.offset == off and oa.order_date == APP_A + timedelta(off) and ob.order_date == APP_B + timedelta(off)
                and (ob.order_date - oa.order_date).days == (APP_B - APP_A).days == 22
                and oa.ship_after_order == 1 and oa.ship_date == oa.order_date + timedelta(1) and ob.ship_date == ob.order_date + timedelta(1))
        ok &= good
        rows.append(f"{n}: rule TODAY(){off:+d} → {oa.order_date:%b %d}/{oa.ship_date:%b %d} on Oct 8, "
                    f"{ob.order_date:%b %d}/{ob.ship_date:%b %d} on Oct 30{'' if good else ' MISMATCH'}")
    record("R1", "Order and shipping dates roll with TODAY() for all 12 orders",
           "Application date Oct 8 vs Oct 30; offsets read by hand from the original formulas",
           "Order date = application date + the order's formula offset; shipping date = order date + 1; "
           "every date moves exactly 22 days between the two application dates",
           "; ".join(rows), ok, "The stored data is the offset, not the date.")

    # R2 cross-check against the spreadsheet itself (LibreOffice recalculation with TODAY() pinned)
    if not find_soffice():
        record("R2", "Application matches the spreadsheet's own recalculation",
               "Original upload recalculated in LibreOffice with TODAY() pinned to Oct 8, then Oct 30",
               "All 12 order/shipping dates identical to the application's generated dates on both days",
               "Not run: LibreOffice is not installed on this computer", None,
               "Install LibreOffice (libreoffice.org) to run this cross-check; every other test is independent of it.")
        sheet = None
    else:
        sheet = {APP_A: libreoffice_dates(APP_A), APP_B: libreoffice_dates(APP_B)}
        ok = all(sheet[app][n] == (bk[n].order_date, bk[n].ship_date)
                 for app, bk in ((APP_A, ga), (APP_B, gb)) for n in FORMULA_OFFSETS)
        record("R2", "Application matches the spreadsheet's own recalculation",
               "Original upload recalculated in LibreOffice with TODAY() pinned to Oct 8, then Oct 30",
               "All 12 order/shipping dates identical to the application's generated dates on both days",
               f"1650: sheet {' / '.join(map(str, sheet[APP_A][1650]))} vs app {ga[1650].order_date} / {ga[1650].ship_date} on Oct 8; "
               f"sheet {' / '.join(map(str, sheet[APP_B][1650]))} vs app {gb[1650].order_date} / {gb[1650].ship_date} on Oct 30; "
               f"all 12 orders × 2 dates × 2 days equal={ok}",
               ok, "Same formula, same answer.")

    # R3 1650 with application date Oct 30 and contact date Oct 30
    r = Desk(b).handle(Request("status", 1650, on=D(10, 30)))
    f = r.facts
    ok = (f["order_date"] == D(10, 21) and f["ship_date"] == D(10, 22) and f["expected_delivery"] == D(10, 28)
          and f["contact_us_date"] == D(11, 4) and f["calendar_days_since_shipping"] == 8 and f["business_days_since_shipping"] == 6
          and f["calendar_days_past_expected"] == 2 and r.decision == E.LATE and r.new_order is None)
    record("R3", "Order 1650 when the application date is Oct 30",
           "Application date Oct 30 (1650 = TODAY() - 9), customer contact date Oct 30",
           "Ordered Wed Oct 21, shipped Thu Oct 22; due Wed Oct 28; contact-us Wed Nov 4; 8 calendar / 6 business days since "
           "shipping; 2 days late; not yet reshippable (compare test 23, same contact date on application date Oct 8)",
           F(r, "order_date", "ship_date", "expected_delivery", "contact_us_date", "calendar_days_since_shipping",
             "business_days_since_shipping", "calendar_days_past_expected", "state"), ok,
           "Application date moves the order; the contact date is judged against the moved dates.", compose([r]))

    # R4 contact date is independent of the application date
    r = Desk(b).handle(Request("status", 1650, on=D(10, 8)))
    ok = r.decision == "CONTACT_BEFORE_ORDER" and not r.paragraphs
    record("R4", "Contact date earlier than the generated order date",
           "Application date Oct 30 (1650 ordered Oct 21), contact date Oct 8",
           "Flagged internally (contact before order date); no customer email",
           f"{r.decision}; note: {r.internal_note}", ok, "Application date and contact date stay separate inputs.")

    # R5 =TODAY() order follows the application date, including the real current date by default
    ok = (ga[5020].order_date == APP_A and gb[5020].order_date == APP_B and gn[5020].order_date == date.today()
          and now.app_date == date.today())
    record("R5", "Order 5020 (=TODAY()) and the default application date",
           "Load with Oct 8, Oct 30, and with no application date given",
           "5020's order date equals the application date each time; with none given, the real current date is used",
           f"Oct 8 → {ga[5020].order_date}; Oct 30 → {gb[5020].order_date}; default → {gn[5020].order_date} (today {date.today()})",
           ok, "TODAY() is never hard-coded.")


APP_C = D(12, 15)   # a much later application date for the fixed new-order tests
# Hand-counted values for the three company-created orders (ship date = creation date = day 0):
#   1814 damaged, contact Mon Oct 12: Express, New York X=1 → due Tue Oct 13; contact-us Oct 12 + 2 + 1 = Thu Oct 15
#   1628 exchange, contact Wed Oct 14: Regular, New Jersey X=2 → due Fri Oct 16; 4 business days = Tue Oct 20, +1 = Wed Oct 21
#   1650 lost, contact Fri Oct 30: Regular, Johannesburg X=4 → due Thu Nov 5; 8 business days = Wed Nov 11, +1 = Thu Nov 12
NEW_EXPECTED = {  # kind: (original, created/shipped, method, due, contact-us, item, qty)
    "damaged": (1814, D(10, 12), "Express", D(10, 13), D(10, 15), "taser, high voltage", 3),
    "exchange": (1628, D(10, 14), "Regular", D(10, 16), D(10, 21), "quiver, aluminum", 5),
    "lost": (1650, D(10, 30), "Regular", D(11, 5), D(11, 12), "vibranium plate, red", 1),
}


def new_order_tests():
    """N1-N9: company-created orders keep their real creation/shipping dates while the originals keep rolling."""
    from stark_support import workbook as W
    p = fixture_copy()
    d = desk(p, persist=True)                       # application date Oct 8; new orders are written to the workbook
    made = {
        "damaged": d.handle(Request("damaged", 1814, on=D(10, 12))),
        "exchange": d.handle(Request("exchange_tracking", 1628, on=D(10, 14), offered_on=D(10, 14),
                                     new_item="quiver, aluminum", tracking_number="1Z999TEST")),
        "lost": d.handle(Request("lost", 1650, on=D(10, 30))),
    }
    nums = {k: r.new_order["number"] for k, r in made.items()}
    books = {app: load(p, app_date=app) for app in (APP_A, APP_B, APP_C, date.today())}
    by_num = {app: {o.number: o for o in bk.all_orders()} for app, bk in books.items()}
    originals = {o.number for o in books[APP_A].orders}

    # N1 unique numbers, linked to their originals
    bk = books[APP_A]
    all_nums = [o.number for o in bk.all_orders()]
    ok = (all(r.decision in ("DM_CREATED", "EX_CREATED", E.RESHIP) for r in made.values())
          and len(set(nums.values())) == 3 and not set(nums.values()) & originals and len(all_nums) == len(set(all_nums))
          and not any("Duplicate" in w for w in bk.warnings) and sorted(nums.values()) == [9003, 9004, 9005])
    links = []
    for k, n in nums.items():
        o, orig = by_num[APP_A][n], by_num[APP_A][NEW_EXPECTED[k][0]]
        good = (o.original_number == orig.number and o.order_type == W.NEW_ORDER_TYPES[k] and o.email == orig.email
                and o.name == orig.name and o.address == orig.address)
        ok &= good
        links.append(f"{n} → original {o.original_number} ({o.order_type}){'' if good else ' MISMATCH'}")
    record("N1", "New orders get unique numbers and stay linked to their originals",
           "Damaged 1814 (Oct 12), exchange 1628 (Oct 14), reship 1650 (Oct 30); saved to New Orders, then reloaded",
           "Three new numbers (next free: 9003, 9004, 9005), none reused, each row naming its original order and copying its customer",
           "; ".join(links) + f"; {len(all_nums)} order numbers, all unique", ok,
           "Originals are never overwritten; the link is the Original order number column.")

    # N2 creation and shipping dates never move with the application date
    ws = openpyxl.load_workbook(p)["New Orders"]
    cells_ok = all(hasattr(ws[f"{c}{r}"].value, "year") and not str(ws[f"{c}{r}"].value).startswith("=")
                   for r in range(2, 5) for c in "DE")
    rows, ok = [], cells_ok
    for k, n in nums.items():
        exp = NEW_EXPECTED[k][1]
        seen = {(by_num[app][n].order_date, by_num[app][n].ship_date) for app in books}
        good = seen == {(exp, exp)}
        ok &= good
        rows.append(f"{n}: {exp:%b %d}/{exp:%b %d} on every application date{'' if good else f' MISMATCH {seen}'}")
    record("N2", "New order dates stay fixed when the application date advances",
           f"Workbook reloaded with application dates Oct 8, Oct 30, Dec 15 and today ({date.today()})",
           "Created = shipped = the day it was set up, identical on all four loads; stored as typed dates, not formulas",
           "; ".join(rows) + f"; typed date cells={cells_ok}", ok, "Company-created orders record what actually happened.")

    # N3 delivery and contact-us dates come from the fixed shipping date
    rows, ok = [], True
    for k, n in nums.items():
        _, ship, method, due, cu, _, _ = NEW_EXPECTED[k]
        seen = {(f["expected_delivery"], f["contact_us_date"])
                for f in (Desk(bk).analyse(by_num[app][n], D(12, 31)) for app, bk in books.items())}
        good = seen == {(due, cu)} and made[k].new_order["est"] == due and made[k].new_order["contact"] == cu
        ok &= good
        rows.append(f"{n} ({method}): due {due:%a %b %d}, contact-us {cu:%a %b %d}{'' if good else f' MISMATCH {seen}'}")
    record("N3", "Delivery and contact-us dates use the new order's own shipping date",
           "Same three orders analysed under all four application dates",
           "1814 replacement: Tue Oct 13 / Thu Oct 15; 1628 exchange: Fri Oct 16 / Wed Oct 21; 1650 reship: Thu Nov 5 / Thu Nov 12",
           "; ".join(rows), ok, "Same estimate and contact-us rules as every other order, counted from the fixed ship day.")

    # N4-N6 shipping method by type
    for t, k, title, why in (("N4", "lost", "Lost-order reshipment ships Regular", "Reships always go Regular, whatever the original used."),
                             ("N5", "damaged", "Wrong or damaged replacement ships Express", "Replacements always go Express, even though 1814 was Regular."),
                             ("N6", "exchange", "Exchange order ships Regular", "Exchanges always go Regular.")):
        orig, ship, method, *_ = NEW_EXPECTED[k]
        n = nums[k]
        cell = next(ws[f"K{r}"].value for r in range(2, 5) if ws[f"A{r}"].value == n)
        got = {by_num[app][n].method for app in books}
        ok = got == {method} and cell == method and made[k].new_order["method"] == method
        record(t, title, f"Order {n}, created from {orig} ({by_num[APP_A][orig].method}) on {ship:%b %d}",
               f"Shipping method {method}", f"Workbook: {cell}; loaded: {', '.join(sorted(got))}", ok, why,
               compose([made[k]], by_num[APP_A][orig].first_name))

    # N7 repeat enquiries refer to the existing order (application date has moved on)
    rows_before = sum(1 for r in range(2, 50) if ws[f"A{r}"].value not in (None, ""))
    repeats = [
        ("lost", Desk(load(p, app_date=APP_B), persist_path=p).handle(Request("lost", 1650, on=D(11, 10))), "ALREADY_RESHIPPED / " + E.LATE),
        ("damaged", Desk(load(p, app_date=D(10, 9)), persist_path=p).handle(Request("damaged", 1814, on=D(10, 13))), "ALREADY_REPLACED / " + E.DUE),
        ("wrong", Desk(load(p, app_date=D(10, 9)), persist_path=p).handle(Request("wrong", 1814, on=D(10, 13))), "ALREADY_REPLACED / " + E.DUE),
        ("exchange", Desk(load(p, app_date=D(10, 9)), persist_path=p).handle(
            Request("exchange_tracking", 1628, on=D(10, 15), offered_on=D(10, 14), new_item="quiver, aluminum",
                    tracking_number="1Z999TEST")), "ALREADY_EXCHANGED / " + E.NOT_YET_DUE),
    ]
    wb2 = openpyxl.load_workbook(p)["New Orders"]
    rows_after = sum(1 for r in range(2, 50) if wb2[f"A{r}"].value not in (None, ""))
    target = {"lost": nums["lost"], "damaged": nums["damaged"], "wrong": nums["damaged"], "exchange": nums["exchange"]}
    ok = rows_before == rows_after == 3 and all(
        r.new_order is None and r.decision == want and r.facts.get("existing_order") == target[k] and str(target[k]) in compose([r])
        for k, r, want in repeats)
    record("N7", "Repeat enquiries point to the existing new order instead of creating another",
           "Reloaded later: 1650 lost again (app Oct 30, contact Nov 10); 1814 damaged and wrong-item again (app Oct 9, contact Oct 13); "
           "1628 tracking sent again (app Oct 9, contact Oct 15)",
           f"No new rows (New Orders stays at 3); replies name {nums['lost']}, {nums['damaged']}, {nums['damaged']}, {nums['exchange']} "
           "with that order's status (late / due today / due today / on its way)",
           "; ".join(f"{k}: {r.decision} → {r.facts.get('existing_order')}" for k, r, _ in repeats) + f"; rows {rows_before} → {rows_after}",
           ok, "A second order is only created on a deliberate staff override (force_new).",
           compose([repeats[0][1]], "Bruce"))

    # N8 originals keep rolling in the same workbook
    rows, ok = [], True
    for n, off in FORMULA_OFFSETS.items():
        dts = [by_num[app][n].order_date for app in (APP_A, APP_B, APP_C)]
        good = dts == [APP_A + timedelta(off), APP_B + timedelta(off), APP_C + timedelta(off)]
        ok &= good
        if n in (1628, 1650, 1814) or not good:
            rows.append(f"{n}: {dts[0]:%b %d} → {dts[1]:%b %d} → {dts[2]:%b %d}{'' if good else ' MISMATCH'}")
    record("N8", "Original orders still roll after new orders are logged",
           "Same saved workbook (with the three new orders) loaded on Oct 8, Oct 30 and Dec 15",
           "All 12 originals = application date + their formula offset (1650: Sep 29 → Oct 21 → Dec 06) while 9003-9005 stay put",
           "; ".join(rows) + f"; all 12 ok={ok}", ok, "Rolling formulas and fixed new-order dates live side by side.")

    # N9 contact date is a separate input from the application date
    n = nums["lost"]
    seen = {app: Desk(bk).handle(Request("status", n, on=D(11, 3))) for app, bk in books.items()}
    late = Desk(books[APP_A]).handle(Request("status", n, on=D(11, 9)))
    ok = (len({(r.decision, r.facts["expected_delivery"], r.facts["contact_us_date"], r.facts["app_date"] == app)
               for app, r in seen.items()}) == 1
          and seen[APP_A].decision == E.NOT_YET_DUE and late.decision == E.LATE
          and len({r.facts["app_date"] for r in seen.values()}) == 4)
    record("N9", "Contact date stays separate from the application date",
           f"Status of {n}: contact Nov 3 under four application dates; then contact Nov 9 on Oct 8",
           "Same answer (on its way, due Thu Nov 5) whatever the application date; changing only the contact date to Nov 9 makes it late",
           f"Nov 3 → {', '.join(f'{a:%b %d}: {r.decision}' for a, r in seen.items())}; Nov 9 → {late.decision}", ok,
           "The application date only moves rolling originals; the contact date decides the customer's answer.",
           compose([seen[APP_A]], "Bruce"))


def libreoffice_dates(app):
    """Recalculate the sanitized demo workbook with TODAY() replaced by DATE(app) and read Orders G/H back."""
    src = os.path.join(TMP, f"orig_{app:%m%d}.xlsx")
    wb = openpyxl.load_workbook(open(BOOK, "rb"))
    ws = wb["Orders"]
    for r in range(2, 14):
        for col in "GH":
            v = ws[f"{col}{r}"].value
            if isinstance(v, str) and "TODAY()" in v.upper():
                ws[f"{col}{r}"] = re.sub(r"TODAY\(\)", f"DATE({app.year},{app.month},{app.day})", v, flags=re.I)
    wb.save(src)
    out = os.path.join(TMP, f"lo_{app:%m%d}")
    subprocess.run([find_soffice(), "--headless", "--calc", "--convert-to", "xlsx", "--outdir", out, src],
                   capture_output=True, timeout=180)
    ws = openpyxl.load_workbook(os.path.join(out, os.path.basename(src)), data_only=True)["Orders"]
    dd = lambda v: v.date() if hasattr(v, "date") else v  # noqa: E731
    return {ws[f"A{r}"].value: (dd(ws[f"G{r}"].value), dd(ws[f"H{r}"].value)) for r in range(2, 14)}


def write_report():
    passed = sum(1 for r in results if r[5] == "PASS")
    skipped = sum(1 for r in results if r[5] == "SKIPPED")
    out = ["# Stark support system: test report", "",
           f"**{passed} of {len(results)} passed**{f', {skipped} skipped' if skipped else ''} (24 required tests, rolling tests R1-R5, fixed new-order tests N1-N9). Every test uses an explicit customer contact date. "
           "The policy tests use application date Oct 8, 2026 (TODAY() for the rolling order dates). R1-R5 prove the original dates roll; N1-N9 prove company-created orders keep their real dates. Expected values were "
           "counted by hand. Tests run on temporary copies; the clean workbook is not modified. Every customer email is also "
           "checked for internal terms (X, 2X, \"lost\", formulas, policy jargon).", ""]
    for n, title, inp, exp, act, res, why, reply in results:
        out += [f"## Test {n}: {title}", "", f"- **Input:** {inp}", f"- **Expected result:** {exp}",
                f"- **Actual result:** {act}", f"- **Pass/Fail:** {res}", f"- **Reason:** {why}", ""]
        out += (["**Customer response:**", "", "```text", reply, "```", ""] if reply else
                ["**Customer response:** none (internal date check only)", ""])
    with open(os.path.join(ROOT, "tests", "test_report.md"), "w", encoding="utf-8") as fh:
        fh.write("\n".join(out))
    print(f"{passed}/{len(results)} passed" + (f", {skipped} skipped" if skipped else ""))
    print("Full report: " + os.path.join(ROOT, "tests", "test_report.md"))
    for r in results:
        if r[5] == "SKIPPED":
            print("SKIPPED", r[0], r[1], "|", r[4])
        elif r[5] != "PASS":
            print("FAIL", r[0], r[1], "|", r[4], "|", r[6])


if __name__ == "__main__":
    try:
        run()
    finally:
        write_report()
        shutil.rmtree(TMP, ignore_errors=True)
