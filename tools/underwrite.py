"""Underwriter — Streamlit page.

Drop ALL the deal's PDFs in one place. Claude classifies each one (rent roll /
tax bill / operating statement) and routes it to the matching engine. Review
each document's extraction step by step — or flip review mode off if you're
feeling lazy — then build ONE workbook:

  Property Info      — cover sheet (APN web lookup, editable)
  Rent Roll + Source — worksheet linked to the faithful transcription
  Historicals        — combined + per-statement tabs
  Tax <year> / RE Taxes

No Pro Forma tab — that stays manual for now.
"""

import io
import re
from pathlib import Path

import pandas as pd
import streamlit as st

from tools import hist_llm
from tools import prop_info as PI
from tools import rent_roll as RR
from tools import tax_parser as TX
from tools import statements as ST

KINDS = ["rent_roll", "tax_bill", "summary", "detail", "other"]
KIND_LABELS = {"rent_roll": "🏘️ Rent roll", "tax_bill": "🏠 Tax bill",
               "summary": "📊 Statement (summary)", "detail": "📊 Statement (detail)",
               "other": "❓ Other / skip"}


# ── classification ──────────────────────────────────────────────────────────
def classify_pdf(fname: str, pdf_bytes: bytes) -> str:
    """Heuristics first (free); Claude fallback when ambiguous."""
    try:
        text = ST._text(pdf_bytes)[:5000]
    except Exception:
        text = ""
    t, n = text.upper(), fname.upper()
    undoubled = text[::2].upper()          # char-doubled headers ("CCaasshh FFllooww")

    if "RENT ROLL" in t or "RENT ROLL" in undoubled or "RENTROLL" in n.replace(" ", "") or "RENT ROLL" in n:
        return "rent_roll"
    if ("SECURED PROPERTY TAX" in t or "ANNUAL PROPERTY TAX" in t
            or ("TAX" in n and "GENERAL TAX LEVY" in t) or "TAXABLE VALUE" in t):
        return "tax_bill"
    if any(k in t or k in undoubled for k in ("CASH FLOW", "INCOME STATEMENT", "OPERATING STATEMENT", "PROFIT AND LOSS")):
        return ST.detect_kind(pdf_bytes)
    if hist_llm.available():
        try:
            schema = {"type": "object",
                      "properties": {"kind": {"type": "string",
                                              "enum": ["rent_roll", "tax_bill", "operating_statement", "other"]}},
                      "required": ["kind"], "additionalProperties": False}
            resp = hist_llm._client().messages.create(
                model=hist_llm.MODEL, max_tokens=2000, thinking={"type": "adaptive"},
                messages=[{"role": "user", "content":
                           "Classify this real-estate document from its filename and first-page "
                           f"text.\n\nFilename: {fname}\n\nText:\n{text[:3000]}"}],
                output_config={"format": {"type": "json_schema", "schema": schema}},
            )
            kind = hist_llm._json_response(resp)["kind"]
            if kind == "operating_statement":
                return ST.detect_kind(pdf_bytes)
            return kind
        except Exception:
            pass
    return "other"


# ── per-kind parsing ────────────────────────────────────────────────────────
def _parse_doc(kind: str, fname: str, pdf_bytes: bytes, ai_on: bool):
    if kind == "rent_roll":
        if not ai_on:
            raise RuntimeError("rent roll extraction needs ANTHROPIC_API_KEY")
        return RR.extract_rent_roll(pdf_bytes)
    if kind == "tax_bill":
        pdf_path = TX.OUTPUT_DIR / fname
        pdf_path.write_bytes(pdf_bytes)
        d = TX.parse_pdf(pdf_path)
        if ai_on and TX._weak_tax(d):
            try:
                d = TX._merge_tax(d, TX.extract_tax_bill(pdf_bytes))
            except Exception:
                pass
        shot = TX.pdf_to_screenshot(pdf_path)
        return {"data": d, "shot": str(shot) if shot else None}
    if kind == "summary":
        rows = ST.parse_summary(pdf_bytes)
        n_items = sum(1 for r in rows if r.get("amount") is not None
                      and not r.get("total") and not r.get("net"))
        if n_items < 3 and ai_on:
            rows = hist_llm.extract_statement(pdf_bytes)
        return {"rows": rows}
    if kind == "detail":
        d = ST.parse_detail(pdf_bytes)
        if not d["cats"] and ai_on:
            return {"rows": hist_llm.extract_statement(pdf_bytes), "as_summary": True}
        return d
    return None


def _stmt_label(fname):
    m = re.search(r"(20\d{2})", fname)
    lbl = m.group(1) if m else Path(fname).stem[:12]
    if re.search(r"ytd", fname, re.I):
        lbl += " YTD"
    return lbl


# ── workbook assembly (no Pro Forma — that stays manual) ────────────────────
def build_underwriting(prop_info, rr_data, tax_bills, summaries, detail, use_llm=True):
    from openpyxl import Workbook
    wb = Workbook(); wb.remove(wb.active)
    if prop_info and any(v is not None for v in prop_info.values()):
        PI.build_sheet(wb.create_sheet("Property Info"), prop_info)
    if rr_data:
        RR.build_into(wb, rr_data, name="Rent Roll", src_name="Rent Roll Source")
    if summaries or detail:
        ST.build_into(wb, summaries, detail, use_llm=use_llm, combined_title="Historicals")
    if tax_bills:
        TX.build_tax_into(wb, tax_bills, prefix="Tax ", combined_name="RE Taxes")
    if not wb.sheetnames:
        raise RuntimeError("nothing to build — no parsable documents")
    buf = io.BytesIO(); wb.save(buf); return buf.getvalue()


# ── review widgets per kind ─────────────────────────────────────────────────
def _review_rent_roll(key, parsed):
    c1, c2, c3 = st.columns(3)
    pname = c1.text_input("Property", value=parsed.get("property_name") or "", key=f"uwrr_pn_{key}")
    csz = c2.text_input("City/State/ZIP", value=parsed.get("city_state_zip") or "", key=f"uwrr_csz_{key}")
    apn = c3.text_input("APN", value=parsed.get("apn") or "", key=f"uwrr_apn_{key}")
    snote = st.text_input("Source note", value=parsed.get("source_note") or "", key=f"uwrr_sn_{key}")
    df = st.data_editor(RR._df_from_units(parsed["units"]), num_rows="dynamic",
                        use_container_width=True, key=f"uwrr_units_{key}")
    return {"property_name": pname, "city_state_zip": csz, "apn": apn or None,
            "source_note": snote, "units": RR._units_from_df(df)}


def _review_tax(key, parsed):
    d = parsed["data"]
    c1, c2, c3 = st.columns(3)
    apn = c1.text_input("APN", value=d.get("apn") or "", key=f"uwtx_apn_{key}")
    year = c2.number_input("Tax year", value=int(d.get("tax_year") or 0), step=1, key=f"uwtx_yr_{key}")
    hc = c3.number_input("Annual tax total", value=float(d.get("property_tax_hardcoded") or 0.0),
                         step=0.01, key=f"uwtx_hc_{key}")
    cl, cr = st.columns(2)
    with cl:
        st.markdown("**Mill Rates**")
        mill_df = st.data_editor(pd.DataFrame(d["mill_rates"], columns=["Agency", "Rate"]),
                                 num_rows="dynamic", use_container_width=True, key=f"uwtx_mill_{key}")
    with cr:
        st.markdown("**Direct Assessments**")
        da_df = st.data_editor(pd.DataFrame(d["direct_assessments"], columns=["Assessment", "Amount"]),
                               num_rows="dynamic", use_container_width=True, key=f"uwtx_da_{key}")
    t1, t2, t3 = st.columns(3)
    land = t1.number_input("Land", value=int(d["taxable_value"].get("land") or 0), step=1, key=f"uwtx_l_{key}")
    impr = t2.number_input("Improvements", value=int(d["taxable_value"].get("improvements") or 0), step=1, key=f"uwtx_i_{key}")
    pers = t3.number_input("Pers Property", value=int(d["taxable_value"].get("pers_property") or 0), step=1, key=f"uwtx_p_{key}")
    edited = {"apn": apn or None, "tax_year": int(year) or None,
              "mill_rates": [(str(a), float(r)) for a, r in mill_df.itertuples(index=False)
                             if str(a).strip() and pd.notna(r)],
              "direct_assessments": [(str(n_), float(v)) for n_, v in da_df.itertuples(index=False)
                                     if str(n_).strip() and pd.notna(v)],
              "taxable_value": {"land": land or None, "improvements": impr or None,
                                "pers_property": pers or None},
              "property_tax_hardcoded": hc or None}
    return {"data": edited, "shot": parsed.get("shot")}


def _review_stmt(key, fname, parsed):
    label = st.text_input("Label (column header)", value=_stmt_label(fname), key=f"uwst_lbl_{key}")
    if "rows" in parsed:
        rows = parsed["rows"]
        items = [(r["label"], r.get("amount")) for r in rows if r.get("amount") is not None]
        st.dataframe(pd.DataFrame(items, columns=["Line", "Amount"]).head(30),
                     use_container_width=True, hide_index=True)
    else:
        st.caption(f"{len(parsed['cats'])} categories · {len(parsed['months'])} months "
                   f"({parsed['months'][0][1]} – {parsed['months'][-1][1]})" if parsed["months"] else "")
    return label


# ── UI ───────────────────────────────────────────────────────────────────────
def render():
    st.header("📐 Underwriter")
    st.caption("Drop every PDF for the deal below. Claude sorts them (rent roll / tax "
               "bill / statement), each is parsed by its tool, you glance over the "
               "results, and out comes ONE workbook — sources and worksheets all linked.")

    ai_on = hist_llm.available()
    if not ai_on:
        st.warning("Set ANTHROPIC_API_KEY in the app secrets — classification fallback, "
                   "rent roll reading, and weak-parse rescue all use Claude.")

    review = st.toggle("🔍 Review before building (turn off if you're feeling lazy)",
                       value=True, key="uw_review")

    files = st.file_uploader("Deal PDFs", type=["pdf"], accept_multiple_files=True,
                             key="uw_files")
    docs = st.session_state.setdefault("uw_docs", {})

    if files:
        live = set()
        for f in files:
            key = f"{f.name}:{f.size}"
            live.add(key)
            if key not in docs:
                data = f.getvalue()
                with st.spinner(f"Classifying & parsing {f.name}…"):
                    kind = classify_pdf(f.name, data)
                    try:
                        parsed = _parse_doc(kind, f.name, data, ai_on)
                    except Exception as e:  # noqa: BLE001
                        st.error(f"{f.name}: {e}")
                        parsed = None
                docs[key] = {"fname": f.name, "bytes": data, "kind": kind, "parsed": parsed}
        for k in list(docs):
            if k not in live:
                del docs[k]

    if not docs:
        st.info("Drop the deal's PDFs to begin.")
        return

    # ── classification summary (overridable) ─────────────────────────────
    st.subheader("1 · Documents")
    for key, doc in docs.items():
        c1, c2 = st.columns([3, 2])
        c1.markdown(f"`{doc['fname']}`")
        sel = c2.selectbox("Type", KINDS, index=KINDS.index(doc["kind"]),
                           format_func=lambda k: KIND_LABELS[k],
                           key=f"uw_kind_{key}", label_visibility="collapsed")
        if sel != doc["kind"]:                       # user override -> reparse
            with st.spinner(f"Re-reading {doc['fname']} as {KIND_LABELS[sel]}…"):
                try:
                    doc["parsed"] = _parse_doc(sel, doc["fname"], doc["bytes"], ai_on)
                    doc["kind"] = sel
                except Exception as e:  # noqa: BLE001
                    st.error(f"{doc['fname']}: {e}")

    # ── property info (APN web lookup) ───────────────────────────────────
    st.subheader("2 · Property info")
    # prefill APN + ZIP from whatever was parsed
    apn_guess, zip_guess = "", ""
    for doc in docs.values():
        p = doc.get("parsed")
        if doc["kind"] == "rent_roll" and p:
            if p.get("apn") and not apn_guess:
                apn_guess = p["apn"]
            m = re.search(r"\b(\d{5})\b", p.get("city_state_zip") or "")
            if m and not zip_guess:
                zip_guess = m.group(1)
        if doc["kind"] == "tax_bill" and p and p["data"].get("apn") and not apn_guess:
            apn_guess = p["data"]["apn"]
    c1, c2, c3 = st.columns([2, 2, 1])
    apn_in = c1.text_input("APN (best)", value=st.session_state.get("uw_apn", apn_guess), key="uw_apn")
    addr_in = c2.text_input("Street address (if no APN)", key="uw_addr",
                            placeholder="e.g. 14 Brooks Ave")
    zip_in = c3.text_input("ZIP code", value=st.session_state.get("uw_zip", zip_guess), key="uw_zip",
                           help="Disambiguates the address — '14 Brooks Ave' exists in a dozen "
                                "cities; the ZIP pins down which one is yours.")
    if st.button("🔎 Look up property info on the web",
                 disabled=not (ai_on and (apn_in.strip() or addr_in.strip()))):
        if apn_in.strip():
            st.session_state.pop("uw_cands", None)
            with st.spinner(f"Researching APN {apn_in}…"):
                try:
                    st.session_state["uw_pinfo"] = PI.fetch(apn_in, hint=addr_in, zip_code=zip_in)
                except Exception as e:  # noqa: BLE001
                    st.error(f"Lookup failed: {e}")
        else:
            # address only -> find candidates first, user picks the right city
            with st.spinner(f"Finding properties matching “{addr_in}”…"):
                try:
                    st.session_state["uw_cands"] = PI.candidates(addr_in, zip_code=zip_in)
                except Exception as e:  # noqa: BLE001
                    st.error(f"Search failed: {e}")
    cands = st.session_state.get("uw_cands")
    if cands:
        labels = [f"{c.get('address')}, {c.get('city_state_zip') or '?'}"
                  + (f" — APN {c['apn']}" if c.get("apn") else "")
                  + (f"  ({c['note']})" if c.get("note") else "")
                  for c in cands]
        sel = st.selectbox("Multiple properties match — which one is yours?",
                           range(len(labels)), format_func=lambda i: labels[i], key="uw_cand_sel")
        if st.button("✓ Use this property"):
            c = cands[sel]
            zc = re.search(r"\b(\d{5})\b", c.get("city_state_zip") or "")
            with st.spinner("Researching the selected property…"):
                try:
                    st.session_state["uw_pinfo"] = PI.fetch(
                        c.get("apn") or "",
                        hint=f"{c.get('address')}, {c.get('city_state_zip') or ''}",
                        zip_code=zc.group(1) if zc else zip_in)
                    st.session_state.pop("uw_cands", None)
                except Exception as e:  # noqa: BLE001
                    st.error(f"Lookup failed: {e}")
    pinfo = st.session_state.get("uw_pinfo")
    if pinfo and pinfo.get("verified_address"):
        st.info(f"📍 Verified as: **{pinfo['verified_address']}** — if that's the wrong "
                "property, fix the ZIP and re-run the lookup.")
    if pinfo and review:
        cols3 = st.columns(3)
        fields = PI.PROP_FIELDS[:2] + [("Address line 2", "address_line2")] + PI.PROP_FIELDS[2:] + PI.BLDG_FIELDS
        edit = {}
        for j, (label, k) in enumerate(fields):
            v = pinfo.get(k)
            edit[k] = cols3[j % 3].text_input(label, value="" if v is None else str(v),
                                              key=f"uw_pif_{k}")
        for k, v in edit.items():
            v = (v or "").strip()
            if v == "":
                edit[k] = None
            else:
                try:
                    edit[k] = float(v) if "." in v else int(v)
                except ValueError:
                    edit[k] = v
        st.session_state["uw_pinfo_edit"] = edit

    # ── step-by-step review ──────────────────────────────────────────────
    edited_docs = {}
    if review:
        st.subheader("3 · Review each document")
        for key, doc in docs.items():
            if doc["kind"] == "other" or not doc.get("parsed"):
                continue
            with st.expander(f"{KIND_LABELS[doc['kind']]} — {doc['fname']}", expanded=False):
                if doc["kind"] == "rent_roll":
                    edited_docs[key] = _review_rent_roll(key, doc["parsed"])
                elif doc["kind"] == "tax_bill":
                    edited_docs[key] = _review_tax(key, doc["parsed"])
                else:
                    edited_docs[key] = _review_stmt(key, doc["fname"], doc["parsed"])

    # ── build ────────────────────────────────────────────────────────────
    st.subheader("4 · Build" if review else "2 · Build")
    if st.button("🏗️ Build workbook", type="primary"):
        rr_data, tax_bills, summaries, detail = None, [], [], None
        for key, doc in docs.items():
            p = doc.get("parsed")
            if not p or doc["kind"] == "other":
                continue
            if doc["kind"] == "rent_roll":
                data = edited_docs.get(key, p) if review else p
                if rr_data:
                    st.warning(f"Multiple rent rolls — using the first, skipping {doc['fname']}.")
                else:
                    rr_data = data
            elif doc["kind"] == "tax_bill":
                e = edited_docs.get(key, p) if review else p
                d = e["data"]
                tax_bills.append((d, d.get("apn") or "unknown", e.get("shot"),
                                  TX._bill_year(d, doc["fname"])))
            else:
                label = edited_docs.get(key) if review else None
                label = label or _stmt_label(doc["fname"])
                if doc["kind"] == "summary" or p.get("as_summary"):
                    summaries.append({"label": label, "rows": p["rows"]})
                else:
                    if detail:
                        st.warning(f"Multiple detail statements — using the first, skipping {doc['fname']}.")
                    else:
                        detail = {"label": label, **{k: p[k] for k in ("cats", "totals", "months")}}
        summaries.sort(key=lambda s: s["label"])
        try:
            with st.spinner("Assembling workbook…"):
                xb = build_underwriting(st.session_state.get("uw_pinfo_edit") or pinfo,
                                        rr_data, tax_bills, summaries, detail, use_llm=ai_on)
        except Exception as e:  # noqa: BLE001
            st.error(f"Build failed: {e}")
            return
        parts = [lbl for lbl, ok in [("property info", pinfo), ("rent roll", rr_data),
                                     (f"{len(tax_bills)} tax bill(s)", tax_bills),
                                     (f"{len(summaries)} statement(s)" + (" + detail" if detail else ""),
                                      summaries or detail)] if ok]
        st.success("Built from: " + ", ".join(parts))
        prop = (rr_data or {}).get("property_name") or (pinfo or {}).get("property_name") or "deal"
        stem = re.sub(r"[^0-9A-Za-z]+", "_", str(prop)).strip("_") or "deal"
        st.session_state["uw_xb"] = (xb, f"{stem}_Underwriting.xlsx")

    if "uw_xb" in st.session_state:
        xb, fname = st.session_state["uw_xb"]
        st.download_button("⬇ Download workbook (.xlsx)", data=xb, file_name=fname,
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True)
