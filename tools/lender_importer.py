"""Lender Importer — Streamlit page.

Reuses core.Api (the Salesforce + Excel + matching logic) unchanged; this module
only provides the web UI. Gated behind an app password (APP_PASSWORD secret).
Salesforce credentials come from Replit Secrets via src/sf_client (env-aware).
"""

import os
import streamlit as st

from core import Api


# ── login gate ────────────────────────────────────────────────────────────
def _gate() -> bool:
    if st.session_state.get("lender_authed"):
        return True
    st.subheader("Sign in")
    expected = os.environ.get("APP_PASSWORD", "")
    if not expected:
        st.error("APP_PASSWORD is not set. Add it in Replit Secrets to enable this tool.")
        return False
    pw = st.text_input("App password", type="password")
    if st.button("Unlock", type="primary"):
        if pw == expected:
            st.session_state["lender_authed"] = True
            st.rerun()
        else:
            st.error("Wrong password.")
    return False


def _api() -> Api:
    if "lender_api" not in st.session_state:
        st.session_state["lender_api"] = Api()
    return st.session_state["lender_api"]


def _reset():
    for k in ("lender_api", "lender_stage", "lender_analyze", "lender_deal", "lender_prop"):
        st.session_state.pop(k, None)


# ── pages ─────────────────────────────────────────────────────────────────
def render():
    st.header("Lender Importer")
    if not _gate():
        return
    api = _api()
    try:
        api.connect()
    except SystemExit as e:
        st.error(str(e)); return
    except Exception as e:  # noqa: BLE001
        st.error(f"Salesforce connection failed: {e}"); return
    st.caption(f"Connected · {api.api_name}")

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
