"""Command line: python3 -m stark_support <kind> [order] --contact-date YYYY-MM-DD [options]

Two separate dates:
  --app-date      what TODAY() means for the rolling order/shipping dates (default: the real current date)
  --contact-date  the day the customer contacted us (required; or --today)

Examples
  python3 -m stark_support status 1650 --app-date 2026-10-08 --contact-date 2026-10-30
  python3 -m stark_support status --email multi-order-b@example.com --today
  python3 -m stark_support refund 1855 --qty 2 --contact-date 2026-10-08 --ask-remaining
  python3 -m stark_support exchange 1814 --qty 1 --new-item "Vibranium Plate (Red)" --contact-date 2026-10-08
  python3 -m stark_support exchange_tracking 1814 --qty 1 --new-item "Vibranium Plate (Red)" --tracking 1Z99 \\
      --offered-on 2026-10-08 --contact-date 2026-10-12 --save
  python3 -m stark_support damaged 3514 --contact-date 2026-10-08 --save     (or: wrong)

Without --save, new orders are previewed only and no email is printed, so nothing claims an order
that doesn't exist yet.
"""
import argparse
import os
from datetime import date, datetime

from . import Desk, Request, compose, load

DATA_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data")
LOCAL_BOOK = os.path.join(DATA_DIR, "stark_orders_clean.xlsx")
DEMO_BOOK = os.path.join(DATA_DIR, "stark_orders_demo.xlsx")
DEFAULT_BOOK = LOCAL_BOOK if os.path.exists(LOCAL_BOOK) else DEMO_BOOK


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("kind", choices=["status", "lost", "refund", "exchange", "exchange_tracking", "damaged", "wrong", "cancel"])
    ap.add_argument("order", nargs="?", type=int)
    ap.add_argument("--email")
    when = ap.add_mutually_exclusive_group(required=True)
    when.add_argument("--contact-date", type=date.fromisoformat, help="date the customer contacted us")
    when.add_argument("--today", action="store_true", help="use today's date as the contact date")
    ap.add_argument("--item")
    ap.add_argument("--qty", type=int)
    ap.add_argument("--new-item")
    ap.add_argument("--offered-on", type=date.fromisoformat)
    ap.add_argument("--tracking")
    ap.add_argument("--new-order-number", type=int)
    ap.add_argument("--label-sent", action="store_true")
    ap.add_argument("--force-new", action="store_true", help="create another order even if one already exists")
    ap.add_argument("--ask-remaining", action="store_true")
    ap.add_argument("--app-date", type=date.fromisoformat,
                    help="what TODAY() means for the rolling order dates (default: the real current date)")
    ap.add_argument("--book", default=DEFAULT_BOOK)
    ap.add_argument("--save", action="store_true", help="log any new order in New Orders")
    a = ap.parse_args()

    on = datetime.now().astimezone().date() if a.today else a.contact_date
    book = load(a.book, app_date=a.app_date)
    desk = Desk(book, persist_path=a.book if a.save else None)
    r = desk.handle(Request(a.kind, a.order, a.email, on, a.item, a.qty, a.new_item, a.offered_on,
                            a.new_order_number, a.tracking, a.label_sent, a.ask_remaining, a.force_new))
    print(f"[internal] decision: {r.decision}")
    for k, v in r.facts.items():
        print(f"[internal] {k}: {v}")
    if r.internal_note:
        print(f"[internal] note: {r.internal_note}")
    for w in book.warnings:
        print(f"[internal] warning: {w}")
    if r.new_order and not a.save:
        print(f"[internal] new order PREVIEW ONLY (add --save to create it): {r.new_order}")
        print("\n(Email not shown until the new order is saved.)")
        return
    if r.paragraphs:
        print("\n" + compose([r]))


if __name__ == "__main__":
    main()
