"""Simple question-and-answer menu for the order system (started by Start_STARK_Order_System.bat).

It only collects answers and passes them to the normal command line (python -m stark_support ...),
so every calculation and rule is exactly the same as running the command yourself.
"""
import contextlib
import io
import sys
from datetime import date

from . import __main__ as cli

KINDS = [
    ("status", "Check an order's status / delivery date"),
    ("lost", "Order hasn't arrived (lost order reshipment)"),
    ("refund", "Refund request"),
    ("exchange", "Exchange request (first message)"),
    ("exchange_tracking", "Exchange: customer sent the return tracking number"),
    ("damaged", "Item arrived damaged"),
    ("wrong", "Wrong item received"),
    ("cancel", "Cancel or change an order"),
]


def ask(prompt, default=""):
    v = input(f"{prompt}{f' [{default}]' if default else ''}: ").strip()
    return v or default


def ask_date(prompt, default_today=True):
    while True:
        v = ask(prompt + " (YYYY-MM-DD" + (", Enter = today" if default_today else ", Enter = skip") + ")")
        if not v:
            return date.today().isoformat() if default_today else ""
        try:
            return date.fromisoformat(v).isoformat()
        except ValueError:
            print("  Please type the date like 2026-10-30.")


def yes(prompt):
    return ask(prompt + " (y/n)", "n").lower().startswith("y")


def run_cli(args):
    """Run the normal command line with these arguments; return what it printed."""
    out = io.StringIO()
    old = sys.argv
    sys.argv = ["stark_support"] + args
    try:
        with contextlib.redirect_stdout(out):
            cli.main()
    except SystemExit:
        pass
    finally:
        sys.argv = old
    return out.getvalue()


def one_request():
    print()
    for i, (_, label) in enumerate(KINDS, 1):
        print(f"  {i}. {label}")
    print("  0. Exit")
    choice = ask("Choose")
    if choice in ("0", "q", ""):
        return False
    if not choice.isdigit() or not 1 <= int(choice) <= len(KINDS):
        print("  Please choose a number from the list.")
        return True
    kind = KINDS[int(choice) - 1][0]

    args = [kind]
    order = ask("Order number (Enter to look up by email instead)")
    if order:
        args.append(order)
    email = ask("Customer email (Enter to skip)" if order else "Customer email")
    if email:
        args += ["--email", email]
    args += ["--contact-date", ask_date("Date the customer contacted us")]
    app = ask_date("Application date, what TODAY() means for the order dates", default_today=True)
    if app != date.today().isoformat():
        args += ["--app-date", app]

    if kind in ("refund", "exchange", "exchange_tracking", "damaged", "wrong"):
        item = ask("Which item (line number or name, Enter = the only item)")
        if item:
            args += ["--item", item]
        qty = ask("How many (Enter = all of that item)")
        if qty:
            args += ["--qty", qty]
    if kind in ("exchange", "exchange_tracking"):
        new_item = ask('Item they want instead, e.g. "Waterproof Spray (Medium)"')
        if new_item:
            args += ["--new-item", new_item]
    if kind == "exchange_tracking":
        tracking = ask("Return tracking number")
        if tracking:
            args += ["--tracking", tracking]
        offered = ask_date("Date the exchange was first requested", default_today=False)
        if offered:
            args += ["--offered-on", offered]
    if kind in ("refund", "exchange", "damaged", "wrong") and yes("Is a return label attached to your email?"):
        args.append("--label-sent")
    if kind == "refund" and yes("Did the customer ask what the rest of the order is worth?"):
        args.append("--ask-remaining")

    text = run_cli(args)
    print("\n" + text)
    if "PREVIEW ONLY" in text:
        if yes("Create this new order and save it in the workbook?"):
            print("\n" + run_cli(args + ["--save"]))
        else:
            print("Nothing was saved.")
    return True


def main():
    print("STARK INDUSTRIES order system")
    print(f"Workbook: {cli.DEFAULT_BOOK}")
    print("Lines starting with [internal] are for you; the email to send is printed after them.")
    while True:
        try:
            if not one_request():
                break
        except (KeyboardInterrupt, EOFError):
            break
        except Exception as e:  # keep the menu open on a bad entry or a locked workbook
            print(f"\nSomething went wrong: {e}")
            if isinstance(e, PermissionError):
                print("If the workbook is open in Excel, close it and try again.")


if __name__ == "__main__":
    main()
