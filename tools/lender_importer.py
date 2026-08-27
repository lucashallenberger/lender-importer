"""Lender Importer — Streamlit page.

Reuses core.Api (the Salesforce + Excel + matching logic) unchanged; this module
only provides the web UI. Each user signs into Salesforce in the app; the session
is held only in st.session_state (lives while the browser tab is open) and nothing
is written to the server. That Salesforce login is itself the access gate.
"""

import base64
import json
import os

import streamlit as st

from core import Api

STORE_KEY = "lender_login"


def _api() -> Api:
    if "lender_api" not in st.session_state:
        st.session_state["lender_api"] = Api()
    return st.session_state["lender_api"]


def _storage():
    """Browser localStorage handle for 'remember me' (optional; None if unavailable)."""
    if os.getenv("QCP_NO_LOCALSTORAGE"):    # headless tests have no browser to answer
        return None
    try:
        from streamlit_local_storage import LocalStorage
        return LocalStorage()
    except Exception:
        return None


def _ls(ls, method, *args):
    """Call a localStorage method, swallowing anything it throws.

    'Remember me' is a convenience and the component is a third-party iframe;
    a Streamlit upgrade underneath it must not be able to take the whole
    sign-in page down.
    """
    if ls is None:
        return None
    try:
        return getattr(ls, method)(*args)
    except Exception:  # noqa: BLE001
        return None


def _encode(d):
    return base64.b64encode(json.dumps(d).encode()).decode()


def _decode(s):
    try:
        return json.loads(base64.b64decode(s).decode())
    except Exception:
        return None


def _try_remembered(api, ls):
    """If a saved login exists in the browser, sign in with it automatically."""
    if ls is None or st.session_state.get("lender_tried_remember"):
        return
    st.session_state["lender_tried_remember"] = True
    blob = _ls(ls, "getItem", STORE_KEY)
    creds = _decode(blob) if blob else None
    if creds:
        r = api.connect_with(creds.get("username"), creds.get("password"),
                             creds.get("security_token"), creds.get("domain"))
        if r["ok"]:
            st.session_state["lender_user"] = creds.get("username")
            st.rerun()
        else:
            _ls(ls, "deleteItem", STORE_KEY)   # stale (password/token changed)
            # Keep it: the sign-in form explains what Salesforce objected to,
            # rather than silently showing a blank login after a failed retry.
            st.session_state["lender_remember_error"] = r


def _reset():
    """Reset the import flow but keep the Salesforce session (stay signed in)."""
    for k in ("lender_stage", "lender_analyze", "lender_deal", "lender_prop"):
        st.session_state.pop(k, None)


# Salesforce answers most bad logins with one vague code, so spell out what
# each one actually means here — including the causes specific to running on a
# server rather than on someone's office machine.
_LOGIN_HINTS = {
    "INVALID_LOGIN":
        "Salesforce returns this for several different problems and won't say which:\n\n"
        "- **The security token is stale.** Salesforce issues a new token every time the "
        "password changes, including a routine expiry — the old one stops working that "
        "day. Reset it (Salesforce → Settings → Reset My Security Token) and paste the "
        "new one.\n"
        "- **The login is coming from an untrusted IP.** This app runs on Streamlit "
        "Cloud, not your office, and its outbound address changes without warning. If "
        "your profile has Login IP Ranges set, the address below has to fall inside them.\n"
        "- The username, password or org domain is simply wrong.",
    "LOGIN_MUST_USE_SECURITY_TOKEN":
        "The security token is missing or no longer valid — paste a freshly reset one.",
    "API_DISABLED_FOR_ORG":
        "API access is off for the org, or this user's profile is missing the "
        "**API Enabled** permission. An admin has to grant it.",
    "INVALID_OPERATION_WITH_EXPIRED_PASSWORD":
        "The Salesforce password has expired. Change it in Salesforce first, then reset "
        "the security token — changing the password invalidates the old one.",
}


def _egress_ip():
    """The address Salesforce sees this app logging in from, for Login IP Ranges.
    Best effort, asked once per session, and never fatal."""
    if "lender_egress_ip" not in st.session_state:
        try:
            import requests
            st.session_state["lender_egress_ip"] = requests.get(
                "https://api.ipify.org", timeout=4).text.strip()
        except Exception:  # noqa: BLE001
            st.session_state["lender_egress_ip"] = None
    return st.session_state["lender_egress_ip"]


def _show_login_error(r):
    code = r.get("code") or "error"
    st.error(f"**Sign-in failed — {code}**\n\n{r.get('error') or 'No detail returned.'}")
    if r.get("stage") == "object":
        obj = os.environ.get("SF_OBJECT", "ascendix__DealSource__c")
        st.info(f"The Salesforce login worked. Reading the `{obj}` object is what failed — "
                "this user most likely has no access to it.")
        return
    hint = _LOGIN_HINTS.get(code)
    if hint:
        st.info(hint)
    ip = _egress_ip()
    if ip:
        st.caption(f"Signing in from {ip} — the address Salesforce sees.")


def _login(api, ls):
    st.subheader("Sign in to Salesforce")
    st.caption("Need a token? Salesforce → Settings → Reset My Security Token.")
    u = st.text_input("Username (email)")
    p = st.text_input("Password", type="password")
    t = st.text_input("Security token")
    d = st.text_input("Org / domain", value="ascendixre-1500.my.salesforce.com")
    remember = st.checkbox("Remember me on this computer", value=False,
                           help="Saves your login in this browser so a refresh won't sign you out. "
                                "Only use on your own machine.")
    if st.button("Sign in", type="primary"):
        with st.spinner("Connecting to Salesforce…"):
            r = api.connect_with(u, p, t, d)
        if r["ok"]:
            st.session_state["lender_user"] = u
            if remember and ls is not None:
                # Saved on the *next* run, not here. A component call only
                # reaches the browser if the script finishes, and the st.rerun()
                # below aborts this one — the write would never leave.
                st.session_state["lender_remember_creds"] = {
                    "username": u, "password": p, "security_token": t, "domain": d}
            st.rerun()
        else:
            _show_login_error(r)


def _sign_out(ls):
    _ls(ls, "deleteItem", STORE_KEY)
    for k in list(st.session_state.keys()):
        if k.startswith("lender_"):
            st.session_state.pop(k, None)
    st.rerun()


# ── pages ─────────────────────────────────────────────────────────────────
def render():
    st.header("Lender Importer")
    api = _api()
    ls = _storage()
    if api.sf is None:                       # not signed in this session yet
        _try_remembered(api, ls)             # auto sign-in from saved login, if any
    if api.sf is None:
        stale = st.session_state.pop("lender_remember_error", None)
        if stale:                            # the saved login stopped working
            st.warning("The login saved in this browser no longer works, so it has been "
                       "cleared. Salesforce said:")
            _show_login_error(stale)
        _login(api, ls)
        return
    saving = st.session_state.pop("lender_remember_creds", None)
    if saving is not None:                   # deferred from the sign-in run
        _ls(ls, "setItem", STORE_KEY, _encode(saving))
    c1, c2 = st.columns([4, 1])
    c1.caption(f"Signed in as {st.session_state.get('lender_user', '?')} · {api.api_name}")
    if c2.button("Sign out"):
        _sign_out(ls)

    stage = st.session_state.setdefault("lender_stage", "setup")
    {"setup": _setup, "questions": _questions, "review": _review, "done": _done}[stage](api)


def _setup(api):
    up = st.file_uploader("Lender list (.xlsx)", type=["xlsx", "xlsm"])
    sheet = st.text_input("Sheet (name or number)", value="2")

    deal = st.session_state.get("lender_deal")
    q = st.text_input("Search for the deal", placeholder="Type part of the deal name…")
    if q and len(q) >= 2 and not deal:
        deals = api.search_deals(q)
        if not deals:
            st.caption("No deals match.")
        else:
            pick = st.selectbox("Matching deals", [d["name"] for d in deals])
            if st.button("Use this deal"):
                st.session_state["lender_deal"] = next(d for d in deals if d["name"] == pick)
                st.rerun()
    if deal:
        c1, c2 = st.columns([4, 1])
        c1.success(f"Deal: {deal['name']}")
        if c2.button("Change"):
            st.session_state.pop("lender_deal", None); st.rerun()

    prop = st.text_input("Property (used in each record name)", placeholder="e.g. 1234 Main")
    if prop:
        st.caption(f'→ e.g. "REIT - {prop}"')

    ready = up is not None and deal and prop
    if st.button("Analyze", type="primary", disabled=not ready):
        with st.spinner("Reading file & matching accounts/contacts…"):
            res = api.load_excel_bytes(up.getvalue(), up.name, sheet=sheet.strip() or None)
            if not res["ok"]:
                st.error(res["error"]); return
            out = api.analyze(deal["id"], deal["name"], prop.strip())
        st.session_state["lender_prop"] = prop.strip()
        st.session_state["lender_analyze"] = out
        st.session_state["lender_stage"] = "questions" if out["questions"] else "review"
        st.rerun()


SKIP_LABEL = "— skip (leave blank) —"


def _opt_label(c):
    return c["name"] + (f"  ·  {c['score']}%" if c["score"] is not None else "")


def _questions(api):
    out = st.session_state["lender_analyze"]
    qs = out["questions"]
    st.write(f"A few things to confirm ({len(qs)}):")
    with st.form("qform"):
        picks = {}
        for q in qs:
            title = (f'Which account is "{q["typed"]}"?' if q["kind"] == "account"
                     else f'Which person is "{q["typed"]}"?' + (f' (row {q["row"]})' if q["row"] else ""))
            # labels parallel to id_by_label so we can map the selection back to an id
            id_by_label = {SKIP_LABEL: None}
            for c in q["candidates"]:
                id_by_label[_opt_label(c)] = c["id"]
            sel = st.radio(title, list(id_by_label.keys()), key="q_" + q["id"])
            picks[q["id"]] = id_by_label[sel]
        remember = st.checkbox("Remember these choices", value=True)
        submit = st.form_submit_button("Confirm & continue", type="primary")
    if submit:
        for q in qs:
            api.answer(q["id"], picks[q["id"]], remember)
        deal = st.session_state["lender_deal"]
        out2 = api.analyze(deal["id"], deal["name"], st.session_state["lender_prop"])
        st.session_state["lender_analyze"] = out2
        st.session_state["lender_stage"] = "questions" if out2["questions"] else "review"
        st.rerun()


def _review(api):
    out = st.session_state["lender_analyze"]
    m = out["summary"]
    st.caption(f"Deal: {out['deal']}")
    c = st.columns(4)
    c[0].metric("Ready", m["ready"]); c[1].metric("Auto-fixed", m["auto_fixed"])
    c[2].metric("You chose", m["you_chose"]); c[3].metric("Blank contact", m["blank_contacts"])
    if out["preview"]:
        st.dataframe([{
            "Name": p["name"], "Interest": p["interest"],
            "Account": "set" if p["account"] else "blank",
            "Contact": "set" if p["contact"] else "blank"} for p in out["preview"]],
            use_container_width=True, hide_index=True)
    c1, c2 = st.columns([1, 1])
    if c1.button("Start over"):
        _reset(); st.rerun()
    if c2.button(f"Upload {m['ready']} records to Salesforce", type="primary"):
        with st.spinner("Uploading… this can take a minute."):
            st.session_state["lender_upload"] = api.upload()
        st.session_state["lender_stage"] = "done"; st.rerun()


def _done(api):
    r = st.session_state.get("lender_upload", {})
    if r.get("ok"):
        st.success(f"{r['created']} records created on {r['deal']}" +
                   (f" · {r['failed']} failed" if r.get("failed") else ""))
        if r.get("failed"):
            st.write("\n".join(f"row {e['row']}: {e['error']}" for e in r["errors"]))
    else:
        st.error(f"Upload failed: {r.get('error')}")
    if st.button("Import another", type="primary"):
        _reset(); st.rerun()
