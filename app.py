"""Browser page for the order system (started by Start_Order_Support_App.bat).

Only lays out widgets. Every answer comes from stark_support.ui_adapter, which calls the engine.
Flow: enter the inquiry, check the order, read the recommended action, review the email, copy it.
"""
import logging

import streamlit as st

from stark_support import ui_adapter as ui

log = logging.getLogger("stark_support.ui")
S = st.session_state
TITLE = "STARK INDUSTRIES Order Support"
CHECK = "Check Order and Prepare Reply"
INQUIRIES = dict(ui.INQUIRIES)
EXTRA_KEYS = ("item", "qty", "new_item", "tracking", "offered", "label", "rest")

# Copy Email reads the subject and email exactly as they are on screen, edits included.
COPY_SCRIPT = """<script>
(() => {
  if (window.__starkCopyReady) return;
  window.__starkCopyReady = true;
  const say = (text) => {
    let box = document.getElementById('stark-copy-msg');
    if (!box) {
      box = document.createElement('div');
      box.id = 'stark-copy-msg';
      box.style.cssText = 'position:fixed;right:24px;bottom:24px;z-index:99999;padding:10px 16px;border-radius:8px;' +
        'background:#1f2937;color:#fff;font:14px sans-serif;box-shadow:0 2px 8px rgba(0,0,0,.3)';
      document.body.appendChild(box);
    }
    box.textContent = text;
    box.style.display = 'block';
    clearTimeout(window.__starkCopyTimer);
    window.__starkCopyTimer = setTimeout(() => { box.style.display = 'none'; }, 3500);
  };
  document.addEventListener('click', async (event) => {
    if (!event.target.closest || !event.target.closest('.st-key-copy_email button')) return;
    const subject = document.querySelector('.st-key-email_subject input');
    const body = document.querySelector('.st-key-email_body textarea');
    if (!body) { say('There is no email to copy yet.'); return; }
    const head = subject && subject.value.trim() ? 'Subject: ' + subject.value.trim() + '\\n\\n' : '';
    const text = head + body.value;
    let ok = false;
    try { await navigator.clipboard.writeText(text); ok = true; } catch (e) {
      const area = document.createElement('textarea');
      area.value = text; document.body.appendChild(area); area.select();
      try { ok = document.execCommand('copy'); } catch (e2) {}
      area.remove();
    }
    say(ok ? 'Email copied (subject and message).' : 'Could not copy. Select the email text and press Ctrl+C.');
  }, true);
})();
</script>"""


def esc(text) -> str:
    return str(text).replace("$", "\\$")   # a bare $ would start a formula in Streamlit markdown


def md_table(headers, rows):
    lines = ["| " + " | ".join(headers) + " |", "|" + "---|" * len(headers)]
    lines += ["| " + " | ".join(esc(c) for c in row) + " |" for row in rows]
    st.markdown("\n".join(lines))


# ---------- what was typed ----------
def current_inputs() -> dict:
    kind = S.get("inquiry")
    fields = {}
    if kind in ui.ITEM_KINDS:
        fields["item"] = S.get("item") or None
        fields["qty"] = S.get("qty")
    if kind in ("exchange", "exchange_tracking"):
        fields["new_item"] = S.get("new_item")
    if kind == "exchange_tracking":
        fields["tracking_number"] = S.get("tracking") or None
        fields["offered_on"] = S.get("offered")
    if kind in ui.LABEL_KINDS:
        fields["label_sent"] = bool(S.get("label"))
    if kind == "refund":
        fields["ask_remaining"] = bool(S.get("rest"))
    return {"kind": kind, "ident": (S.get("ident") or "").strip(), "contact": S.get("contact"), "app": S.get("app"),
            "fields": fields}


# ---------- callbacks (every engine call happens here, once per click) ----------
def show_result(res):
    S.result = res
    S.token = ui.remember_pending(S, res)
    S.confirm = False
    S.cancelled = False
    S.save_error = None
    S.notice = S.clear_target = S.clear_error = S.clear_note = None
    S.run_id = S.get("run_id", 0) + 1
    if res.get("ok") and res.get("email"):
        S.email_subject, S.email_body = res["subject"], res["email"]


def do_check():
    S.checked = current_inputs()
    i = S.checked
    show_result(ui.prepare_reply(i["kind"], i["ident"], i["contact"], i["app"], **i["fields"]))


def pick_match(number):
    S.ident = str(number)
    do_check()


def ask_confirm(value):
    S.confirm = value
    S.cancelled = not value


def do_save():
    res = ui.save_once(S, S.get("token"))
    if not res["ok"]:
        S.save_error = res["error"]
        return
    show_result(res)


def ask_clear(number):
    S.clear_target, S.clear_error, S.clear_note = number, None, None


def cancel_clear():
    S.clear_target, S.clear_error, S.clear_note = None, None, "Cancelled. Nothing was changed."


def do_clear():
    """Remove the one saved replacement the user confirmed, then check the original order again (a preview only)."""
    if not S.get("clear_target") or not S.get("result") or not S.result.get("order"):
        return
    done = ui.clear_replacement(S.result["order"]["number"], S.clear_target)
    if not done["ok"]:
        S.clear_error = done["error"]
        return
    i = S.checked
    show_result(ui.prepare_reply(i["kind"], i["ident"], i["contact"], i["app"], **i["fields"]))
    S.notice = done["message"]


def new_order():
    for k in ("result", "checked", "token", "confirm", "cancelled", "save_error", "notice", "clear_target", "clear_error",
              "clear_note", "email_subject", "email_body"):
        S.pop(k, None)
    # Widgets keep their value in the browser unless it is set here explicitly.
    S.ident, S.inquiry, S.contact, S.app = "", ui.INQUIRIES[0][0], None, ui.local_today()
    S.item, S.qty, S.new_item = "", None, None
    S.tracking, S.offered, S.label, S.rest = "", None, False, False


# ---------- sections ----------
def inquiry_section():
    with st.container(border=True):
        st.subheader("Customer inquiry")
        a, b = st.columns([2, 3])
        a.text_input("Order number", key="ident", placeholder="For example 1650",
                     help="No order number? Type the customer's email address instead.")
        kind = b.selectbox("What is the customer asking?", list(INQUIRIES), format_func=INQUIRIES.get, key="inquiry",
                           help="Pick the closest match. It decides which rules are checked and what the email answers.")
        c, d = st.columns(2)
        c.date_input("Date the customer contacted us", value=None, key="contact",
                     help="Required. The day the customer wrote, which is not always today.")
        d.date_input("Application date (normally today)", value=ui.local_today(), key="app",
                     help="The original orders' order and shipping dates are counted back from this date. "
                          "Leave it on today unless you are reviewing an older case.")
        if kind in ui.ITEM_KINDS:
            e, f = st.columns(2)
            e.text_input("Which item (only if the order has several)", key="item", placeholder="For example quiver")
            f.number_input("How many (leave empty for all of them)", min_value=1, step=1, value=None, key="qty")
        if kind in ("exchange", "exchange_tracking"):
            st.selectbox("Item they want instead", S.catalogue, index=None, placeholder="Choose an item", key="new_item")
        if kind == "exchange_tracking":
            e, f = st.columns(2)
            e.text_input("Return tracking number", key="tracking")
            f.date_input("Date the exchange was first requested (optional)", value=None, key="offered")
        if kind in ui.LABEL_KINDS:
            st.checkbox("I'm attaching the free return label to this email", key="label")
        if kind == "refund":
            st.checkbox("The customer asked what the rest of the order is worth", key="rest")
        go, again, _ = st.columns([3, 2, 4])
        go.button(CHECK, type="primary", on_click=do_check, key="check")
        again.button("Check Another Order", on_click=new_order, key="another")
        st.caption("Checking is a preview. Nothing is saved until you click Create order and confirm.")


def fact_lines(pairs):
    st.markdown("  \n".join(f"**{label}:** {esc(value)}" for label, value in pairs))


def status_section(res):
    st.subheader("1. Order status")
    status = res.get("status")
    if not status:
        st.markdown(esc(res["decision"]))
        return
    st.markdown(f"#### :{status['colour']}[{status['label']}]")
    if status["note"]:
        st.markdown(esc(status["note"]))
    fact_lines(status["dates"])
    st.markdown(f"**{esc(status['advice'])}**")


def action_section(res):
    st.subheader("2. Recommended action")
    rec = res["recommendation"]
    box = {"success": st.success, "warning": st.warning}.get(rec["tone"], st.info)
    box(f"**{esc(rec['title'])}**")
    for step in rec["steps"]:
        st.markdown(f"- {esc(step)}")


def new_order_section(res):
    """The new order, full width: what will be (or was) written, and the only buttons that save."""
    rows = res["saved_orders"] or res["previews"]
    if not rows:
        return
    what = res["new_order_label"]
    with st.container(border=True):
        if res["saved_orders"]:
            st.success(f"**{what} saved.** It is now in the New Orders sheet of the workbook.")
        else:
            st.warning(f"**{what}: preview only, not saved.** Nothing is written to the workbook until you click "
                       f"Confirm and save.")
        for row in rows:
            fact_lines(row["details"])
        if res["saved_orders"]:
            return
        if S.get("save_error"):
            st.error(S.save_error)
        if not S.get("confirm"):
            if S.get("cancelled"):
                st.info("Cancelled. Nothing was saved.")
            st.button("Create order", key="create", type="primary", on_click=ask_confirm, args=(True,))
            return
        st.markdown(f"**Save this {what.lower()} to the workbook?**")
        with st.container(horizontal=True):
            st.button("Confirm and save", key="save", type="primary", on_click=do_save)
            st.button("Cancel", key="cancel", on_click=ask_confirm, args=(False,))


def demo_section(res):
    """Demonstrations and testing only: remove a saved reshipment or replacement so the scenario can be repeated."""
    items = res.get("clearable") or []
    if not items:
        return
    original = res["order"]["number"]
    with st.expander("Demonstration and testing: clear a saved replacement", expanded=bool(S.get("clear_target") or S.get("clear_note"))):
        st.caption("Not part of handling a customer inquiry. Use this only to remove a saved replacement so the same "
                   "order can be demonstrated or tested again. It changes the saved data in the workbook.")
        if S.get("clear_note"):
            st.info(S.clear_note)
        for c in items:
            st.markdown(f"**Order {c['number']}** · {esc(c['type'])} · {esc('; '.join(c['items']))} · "
                        f"created {ui.show(c['created'])} from order {original}")
            if S.get("clear_target") != c["number"]:
                st.button("Clear saved replacement for retesting", key=f"clear_{c['number']}", on_click=ask_clear,
                          args=(c["number"],))
                continue
            with st.container(border=True):
                st.warning(f"**Clear saved replacement order {c['number']}?**  \nOriginal order: {original}  \n"
                           f"Saved replacement order: {c['number']}  \n"
                           f"Clearing removes this one saved order from the New Orders sheet, so that order {original} "
                           f"can be processed again. No other order is changed and no new order is created.")
                if S.get("clear_error"):
                    st.error(S.clear_error)
                with st.container(horizontal=True):
                    st.button("Confirm clear", key="confirm_clear", type="primary", on_click=do_clear)
                    st.button("Cancel", key="cancel_clear", on_click=cancel_clear)


def email_section(res):
    st.subheader("3. Customer email")
    if res["email_withheld"]:
        st.info("The email appears after you create the order.")
        return
    if not res["email"]:
        st.info("There is no email to send for this result.")
        return
    st.text_input("Subject", key="email_subject")
    st.text_area("Email (you can edit it here before copying)", key="email_body", height=300)
    issues = ui.check_email(res, S.get("email_body", ""))
    st.button("Copy Email", key="copy_email")
    if issues:
        st.warning("**Check before sending:**\n" + "\n".join(f"- {esc(i)}" for i in issues))
    else:
        st.caption("Checked against this order: greeting and closing, dates, amounts, order numbers and wording.")


def details_section(res):
    o = res["order"]
    with st.expander("4. Order details"):
        kind = "original order" if o["source"] == "Orders" else f"{o['order_type']}, created from order {o['original_number']}"
        left, right = st.columns(2)
        with left:
            st.markdown("**Customer and order**")
            md_table(["", ""], [["Order", f"{o['number']} ({kind})"], ["Name", o["name"]], ["Email", o["email"]],
                                ["Address", o["address"]], ["Shipping", o["method"]],
                                ["Items", "; ".join(f"{i['qty']} × {i['name']}" for i in o["items"])]])
        with right:
            if res["key_dates"]:
                st.markdown("**Key dates**")
                md_table(["", ""], res["key_dates"])
        if res["timeline"]:
            st.markdown("**Timeline**")
            for e in res["timeline"]:
                text = f"{ui.show(e['date'])} · {e['label']}"
                st.markdown(f"- **{text}**" if e["is_contact"] else f"- {text}")
        if res["linked_orders"]:
            st.markdown("**New orders already created for this order (fixed dates)**")
            for n in res["linked_orders"]:
                lines = [f"Created and shipped {ui.show(n['created'])}"]
                if "state" in n:
                    lines += [f"Expected delivery {ui.show(n['expected_delivery'])}",
                              f"Contact-us date {ui.show(n['contact_us_date'])}", f"On the contact date: {n['state']}"]
                st.markdown(f"- Order {n['number']} · {n['type']} · {'; '.join(n['items'])}: " + ", ".join(lines))


def internal_section(res):
    info = res["internal"]
    with st.expander("5. Internal explanation (not for the customer)"):
        for part in res.get("explanation", []):
            st.markdown(f"**{esc(part['title'])}**")
            st.markdown("\n".join(f"- {esc(line)}" for line in part["lines"]))
        st.markdown(f"**Decision:** {esc(info['decision'])}")
        if info["internal_note"]:
            st.markdown(f"**Note:** {esc(info['internal_note'])}")
        for w in info["warnings"]:
            st.markdown(f"**Workbook warning:** {esc(w)}")
        st.markdown("**Values the decision was based on**")
        st.json(info["facts"], expanded=False)


def result_area():
    res = S.get("result")
    if not res:
        st.info(f"Fill in the customer inquiry above, then click **{CHECK}**.")
        return
    if res["ok"] and S.get("checked") != current_inputs():
        st.warning(f"You changed the inquiry. Click **{CHECK}** to update the result below.")
    if not res["ok"]:
        st.error(res["error"])
        return
    if res["matches"]:
        st.warning("More than one order uses this email address. Which order is the customer writing about?")
        for m in res["match_details"]:
            st.button(f"Order {m['number']}: {m['name']}, {m['items']}, {m['destination']}", key=f"match_{m['number']}",
                      on_click=pick_match, args=(m["number"],))
        return
    o = res["order"]
    saved = "Saved" if res["saved_orders"] else "Preview"
    st.markdown(f"**{saved}** · Order {o['number']}, {esc(o['name'])} · {esc(INQUIRIES.get(res['kind'], ''))} · "
                f"contacted {ui.show(res['contact_date'])} · application date {ui.show(res['app_date'])}")
    if S.get("notice"):
        st.success(S.notice)
    if res["warning"]:
        st.warning(res["warning"])
        internal_section(res)
        return
    left, right = st.columns(2)
    with left:
        status_section(res)
    with right:
        action_section(res)
    new_order_section(res)
    demo_section(res)
    email_section(res)
    details_section(res)
    internal_section(res)


st.set_page_config(page_title=TITLE, layout="wide")
st.title(TITLE)
if "catalogue" not in S:
    S.catalogue = ui.catalogue()
try:
    inquiry_section()
    result_area()
except Exception as e:  # plain message on the page, traceback in the console
    log.exception("Page failed")
    st.error(f"Something went wrong while showing this result: {e}. Nothing was saved. Click {CHECK} again.")
st.html(COPY_SCRIPT, unsafe_allow_javascript=True)
st.caption("To stop the app, close its black console window or press Ctrl+C in it. Closing this browser tab does not stop it.")
