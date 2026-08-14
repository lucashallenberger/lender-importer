"""Underwriter — Streamlit page.

Drop ALL the deal's documents in one place — PDFs, and the Excel/CSV rent rolls
and P&Ls that show up just as often. Claude classifies each one (rent roll /
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
import zlib
from pathlib import Path

import pandas as pd
import streamlit as st

from tools import hist_llm
from tools import prop_info as PI
from tools import rent_roll as RR
from tools import tabular
from tools import tax_parser as TX
from tools import statements as ST

KINDS = ["rent_roll", "tax_bill", "summary", "detail", "other"]
KIND_LABELS = {"rent_roll": "🏘️ Rent roll", "tax_bill": "🏠 Tax bill",
               "summary": "📊 Statement (summary)", "detail": "📊 Statement (detail)",
               "other": "❓ Other / skip"}


# ── classification ──────────────────────────────────────────────────────────
def classify_doc(fname: str, data: bytes) -> str:
    """Heuristics first (free); Claude fallback when ambiguous. Works off the
    document's text — the PDF text layer, or the workbook rendering."""
    sheet = tabular.is_spreadsheet(data)
    try:
        text = ST._text(data)[:5000]
    except Exception:
        text = ""
    t, n = text.upper(), fname.upper()
    undoubled = text[::2].upper()          # char-doubled headers ("CCaasshh FFllooww")

    if ("RENT ROLL" in t or "RENT ROLL" in undoubled or "RENT ROLL" in n
            or "RENTROLL" in n.replace(" ", "").replace("_", "")
            or re.search(r"\bRR\b", n)):
        return "rent_roll"
    if not sheet and ("SECURED PROPERTY TAX" in t or "ANNUAL PROPERTY TAX" in t
                      or ("TAX" in n and "GENERAL TAX LEVY" in t) or "TAXABLE VALUE" in t):
        return "tax_bill"
    STMT = ("CASH FLOW", "INCOME STATEMENT", "OPERATING STATEMENT", "PROFIT AND LOSS",
            "PROFIT & LOSS", "P&L", "P & L")
    if any(k in t or k in undoubled for k in STMT) or any(k in n for k in STMT) \
            or re.search(r"\bT-?12\b", n) or "OPERATING EXPENSES" in t or "TOTAL INCOME" in t:
        return ST.detect_kind(data)
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
                return ST.detect_kind(data)
            if kind == "tax_bill" and sheet:
                return "other"                 # tax bills are always the printed PDF
            return kind
        except Exception:
            pass
    return "other"


classify_pdf = classify_doc          # old name, kept for callers/tests


# ── per-kind parsing ────────────────────────────────────────────────────────
def _parse_doc(kind: str, fname: str, data: bytes, ai_on: bool, auto: bool = False):
    """auto=True lets a spreadsheet statement pick its own shape: if it turns out to
    be a monthly grid (a T-12), it comes back as a month-column detail statement
    instead of a single flattened column. The type dropdown stays an override."""
    sheet = tabular.is_spreadsheet(data)
    if kind == "rent_roll":
        if not ai_on:
            raise RuntimeError("rent roll extraction needs ANTHROPIC_API_KEY")
        return RR.extract_rent_roll(data)
    if kind == "tax_bill":
        if sheet:
            raise RuntimeError(
                "tax bills have to be the PDF of the printed bill — the parser reads "
                "the mill rates, direct assessments and taxable values off it, and "
                "screenshots it for the source tab")
        pdf_path = TX.OUTPUT_DIR / fname
        pdf_path.write_bytes(data)
        d = TX.parse_pdf(pdf_path)
        if ai_on and TX._weak_tax(d):
            try:
                d = TX._merge_tax(d, TX.extract_tax_bill(data))
            except Exception:
                pass
        try:                                   # the bill names its county — free field
            d["county"] = TX.county_from_text(ST._text(data))
        except Exception:  # noqa: BLE001
            pass
        shot = TX.pdf_to_screenshot(pdf_path)
        return {"data": d, "shot": str(shot) if shot else None}
    if kind in ("summary", "detail") and sheet:
        # The regex parsers were built for PDF text layouts; a workbook goes
        # straight to Claude, which reads every period column in one pass.
        if not ai_on:
            raise RuntimeError("reading a spreadsheet statement needs ANTHROPIC_API_KEY")
        got = hist_llm.extract_statement_periods(data)
        periods, rows = got.get("periods") or [], got.get("rows") or []
        if kind == "detail" or auto:
            det = ST.detail_from_periods(periods, rows)
            if det:
                return {**det, "as_detail": True}
            if kind == "detail":
                raise RuntimeError(
                    "no monthly columns found in this workbook — it reads as a "
                    "single-period statement; leave it set to Statement (summary)")
        return {"rows": ST.summary_from_periods(periods, rows), "as_summary": True}
    if kind == "summary":
        rows = ST.parse_summary(data)
        n_items = sum(1 for r in rows if r.get("amount") is not None
                      and not r.get("total") and not r.get("net"))
        if n_items < 3:
            if not ai_on:
                raise RuntimeError(
                    "this statement has no readable text layer (it looks scanned) and "
                    "AI is off — set ANTHROPIC_API_KEY so Claude can read it")
            rows = hist_llm.extract_statement(data)
        return {"rows": rows}
    if kind == "detail":
        d = ST.parse_detail(data)
        if not d["cats"]:
            if not ai_on:
                raise RuntimeError(
                    "no transactions could be read from this statement and AI is off — "
                    "set ANTHROPIC_API_KEY so Claude can read it")
            return {"rows": hist_llm.extract_statement(data), "as_summary": True}
        return d
    return None


def _stmt_label(fname):
    m = re.search(r"(20\d{2})", fname)
    lbl = m.group(1) if m else Path(fname).stem[:12]
    if re.search(r"ytd", fname, re.I):
        lbl += " YTD"
    return lbl


# ── workbook assembly (no Pro Forma — that stays manual) ────────────────────
def _seed_pinfo(docs) -> dict:
    """Property facts the parsed documents already carry. The cover sheet is
    built from this even when the web lookup was never run or came back empty —
    a tab with the APN and unit count beats no tab at all."""
    layers = []
    for doc in docs.values():
        p = doc.get("parsed")
        if not p:
            continue
        if doc["kind"] == "rent_roll":
            layers.append(PI.from_rent_roll(p))
        elif doc["kind"] == "tax_bill":
            d = p.get("data") or {}
            layers.append({"apn": d.get("apn"), "county": d.get("county")})
    return PI.merge(*layers)


def build_underwriting(prop_info, rr_data, tax_bills, summaries, details, use_llm=True):
    from openpyxl import Workbook
    from tools import xl as XL
    wb = Workbook(); wb.remove(wb.active)
    if PI.has_any(prop_info):
        PI.build_sheet(XL.new_sheet(wb, "Property Info")[0], prop_info)
    if rr_data and rr_data.get("units"):
        RR.build_into(wb, rr_data, name="W - RR", src_name="S - RR")
    if summaries or details:
        ST.build_into(wb, summaries, details, use_llm=use_llm, combined_title="W - Historicals")
    if tax_bills:
        TX.build_tax_into(wb, tax_bills, prefix="S - Tax ", combined_name="W - RE Taxes",
                          always_combined=True)
    if not wb.sheetnames:
        raise RuntimeError("nothing to build — no parsable documents")
    _order_sheets(wb)
    buf = io.BytesIO(); wb.save(buf); return buf.getvalue()


def _order_sheets(wb):
    """Cover sheet, then every worksheet, then every source — the W-/S- naming
    says which is which, so the tab strip should read that way too. Sheets are
    referenced by name in formulas, so reordering is safe."""
    def rank(name):
        return 0 if name == "Property Info" else 1 if name.startswith("W - ") else 2
    wb._sheets.sort(key=lambda ws: rank(ws.title))   # stable: keeps build order within a rank
    wb.active = 0                                    # open on the cover, not wherever the index landed


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
    rec = TX.tax_recon(d)
    if rec and rec[2] <= 1.0:
        st.caption(f"✅ Self-check: rate × value + assessments = ${rec[0]:,.2f} "
                   f"vs printed ${rec[1]:,.2f}.")
    elif rec:
        st.warning(f"⚠️ Doesn't reconcile — computed ${rec[0]:,.2f} vs printed "
                   f"${rec[1]:,.2f} ({rec[2]:.1f}% off). Check the values below.")
    else:
        st.warning("⚠️ No printed annual total to check against — verify the values below.")
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
    st.caption("Drop every document for the deal below — PDF, Excel or CSV. Claude sorts "
               "them (rent roll / tax bill / statement), each is parsed by its tool, you "
               "glance over the results, and out comes ONE workbook — sources and "
               "worksheets all linked. (Tax bills have to be the PDF of the printed bill.)")

    ai_on = hist_llm.available()
    if not ai_on:
        st.warning("Set ANTHROPIC_API_KEY in the app secrets — classification fallback, "
                   "rent roll reading, and weak-parse rescue all use Claude.")

    review = st.toggle("🔍  Review before building (turn off if you're feeling lazy)",
                       value=True, key="uw_review")

    files = st.file_uploader("Deal documents (PDF / Excel / CSV)",
                             type=tabular.UPLOAD_TYPES, accept_multiple_files=True,
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
                    parsed, err = None, None
                    try:
                        kind = classify_doc(f.name, data)
                    except Exception as e:  # noqa: BLE001
                        kind, err = "other", str(e)
                    if not err:                    # keep the detected type either
                        try:                       # way, so a retry is one click
                            parsed = _parse_doc(kind, f.name, data, ai_on, auto=True)
                            if (parsed or {}).get("as_detail"):
                                kind = "detail"    # a monthly grid, detected on read
                        except Exception as e:  # noqa: BLE001
                            err = str(e)
                # the error is kept on the doc, not just flashed: a Streamlit
                # rerun would otherwise wipe it and the file would sit there mute
                docs[key] = {"fname": f.name, "bytes": data, "kind": kind,
                             "parsed": parsed, "error": err}
        for k in list(docs):
            if k not in live:
                del docs[k]

    if not docs:
        st.info("Drop the deal's documents to begin.")
        return

    # ── classification summary (overridable) ─────────────────────────────
    st.subheader("1 · Documents")
    for key, doc in docs.items():
        c1, c2 = st.columns([3, 2])
        c1.markdown(f"`{doc['fname']}`")
        sel = c2.selectbox("Type", KINDS, index=KINDS.index(doc["kind"]),
                           format_func=lambda k: KIND_LABELS[k],
                           key=f"uw_kind_{key}", label_visibility="collapsed")
        # User override -> reparse. The chosen type is adopted either way, so a
        # failure doesn't re-fire (and re-bill a Claude call) on every rerun.
        if sel != doc["kind"]:
            with st.spinner(f"Re-reading {doc['fname']} as {KIND_LABELS[sel]}…"):
                try:
                    doc["parsed"] = _parse_doc(sel, doc["fname"], doc["bytes"], ai_on)
                    doc["error"] = None
                except Exception as e:  # noqa: BLE001
                    doc["parsed"], doc["error"] = None, str(e)
                doc["kind"] = sel
        if doc.get("error"):
            c1.caption(f"⚠️ {doc['error']} — left out of the workbook unless you "
                       "pick a type it can be read as")

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
    seed = _seed_pinfo(docs)
    # what the documents already say about the address, so the researcher has a
    # second handle on the parcel even when the APN box is all that's filled in
    doc_hint = ", ".join(x for x in [seed.get("address_line1") or seed.get("property_name"),
                                     seed.get("address_line2")] if x)
    if st.button("🔎  Look up property info on the web",
                 disabled=not (ai_on and (apn_in.strip() or addr_in.strip()))):
        if apn_in.strip():
            st.session_state.pop("uw_cands", None)
            with st.spinner(f"Researching APN {apn_in}…"):
                try:
                    st.session_state["uw_pinfo"] = PI.fetch(
                        apn_in, hint=addr_in.strip() or doc_hint, zip_code=zip_in)
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
        if st.button("✓  Use this property"):
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
    elif pinfo is not None and not PI.has_any(pinfo):
        st.warning("The web lookup came back empty. The Property Info tab is still built "
                   "from the documents — fill in the blanks below.")

    # Documents first, lookup on top of them: the tab gets built either way.
    base = PI.merge(seed, pinfo)
    if review:
        st.caption("Property Info tab — from the documents, plus the web lookup where it "
                   "found something. Edit freely; blanks stay blank (gold = fill-in).")
        cols3 = st.columns(3)
        fields = PI.PROP_FIELDS[:2] + [("Address line 2", "address_line2")] + PI.PROP_FIELDS[2:] + PI.BLDG_FIELDS
        # New source data must reach the widgets: Streamlit ignores `value=` once
        # a key exists, so the key carries a fingerprint of what seeded it.
        sig = zlib.crc32("|".join(str(base.get(k)) for k in PI.ALL_KEYS).encode())
        edit = {}
        for j, (label, k) in enumerate(fields):
            v = base.get(k)
            edit[k] = cols3[j % 3].text_input(label, value="" if v is None else str(v),
                                              key=f"uw_pif_{k}_{sig}")
        for k, v in edit.items():
            v = (v or "").strip()
            if v == "":
                edit[k] = None
            else:
                try:
                    edit[k] = float(v) if "." in v else int(v)
                except ValueError:
                    edit[k] = v
        # edit wins verbatim (so a cleared field stays cleared); base carries the rest
        st.session_state["uw_pinfo_edit"] = {**base, **edit}

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
    if st.button("🏗️  Build workbook", type="primary"):
        rolls, tax_bills, summaries, details = [], [], [], []
        for key, doc in docs.items():
            p = doc.get("parsed")
            if not p or doc["kind"] == "other":
                continue
            if doc["kind"] == "rent_roll":
                rolls.append(edited_docs.get(key, p) if review else p)
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
                else:                       # every monthly statement keeps its months
                    details.append({"label": label,
                                    **{k: p[k] for k in ("cats", "totals", "months")}})
        rr_data = RR.combine(rolls)
        if len(rolls) > 1:
            st.info(f"Combined {len(rolls)} rent rolls into one worksheet — "
                    f"{len(rr_data['units'])} units. Unit numbers that appeared in more "
                    "than one building carry a building prefix.")
        summaries.sort(key=lambda s: s["label"])
        dupes = {l for l in (s["label"] for s in summaries)
                 if [s["label"] for s in summaries].count(l) > 1}
        if dupes:
            st.warning("Two or more statements share the label " +
                       ", ".join(f"'{d}'" for d in sorted(dupes)) +
                       " — their columns will be indistinguishable on the combined tab. "
                       "Give them distinct labels in the review step above.")
        # rebuild the seed here: rent-roll edits from the review step (unit count,
        # corrected APN) belong on the cover sheet too
        final_pinfo = PI.merge(_seed_pinfo(docs), PI.from_rent_roll(rr_data), pinfo)
        if review and st.session_state.get("uw_pinfo_edit"):
            # what the user saw in the editor is what gets written
            final_pinfo = {**final_pinfo, **st.session_state["uw_pinfo_edit"]}
        try:
            with st.spinner("Assembling workbook…"):
                xb = build_underwriting(final_pinfo, rr_data, tax_bills, summaries,
                                        details, use_llm=ai_on)
        except Exception as e:  # noqa: BLE001
            st.error(f"Build failed: {e}")
            return
        parts = [lbl for lbl, ok in [("property info", PI.has_any(final_pinfo)),
                                     ("rent roll", rr_data),
                                     (f"{len(tax_bills)} tax bill(s)", tax_bills),
                                     (f"{len(summaries)} statement(s)"
                                      + (f" + {len(details)} month-by-month" if details else ""),
                                      summaries or details)] if ok]
        st.success("Built from: " + ", ".join(parts))
        if ST.LAST_LLM_ERROR:
            st.warning("The AI pass over the statements failed, so the Historicals tab fell back "
                       "to the built-in rules — categories aren't aligned across years and line "
                       "items aren't classified. The numbers themselves are unaffected.\n\n"
                       f"`{ST.LAST_LLM_ERROR}`")
        prop = (rr_data or {}).get("property_name") or final_pinfo.get("property_name") or "deal"
        stem = re.sub(r"[^0-9A-Za-z]+", "_", str(prop)).strip("_") or "deal"
        st.session_state["uw_xb"] = (xb, f"{stem}_Underwriting.xlsx")

    if "uw_xb" in st.session_state:
        xb, fname = st.session_state["uw_xb"]
        st.download_button("⬇  Download workbook (.xlsx)", data=xb, file_name=fname,
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True)
