"""Historicals — Streamlit page. Upload property income statements (summary or
transaction-detail PDFs, or Excel/CSV P&Ls); download a workbook with
per-statement tabs + a combined tab (every number linked to its source,
categories aligned, missing lines in red)."""

import re
import streamlit as st

from tools import statements as S
from tools import hist_llm
from tools import tabular


def _tab_label(file_label, tab, n_tabs):
    """One statement per file keeps the label the user typed. A file holding
    several takes each tab's own name where that names the period ('2024',
    '25 YTD') — that name is what the column header should read."""
    if n_tabs < 2 or not tab:
        return file_label
    tab = tab.strip()
    return tab if re.search(r'\d', tab) else f"{file_label} {tab}".strip()


def _guess_label(name):
    m = re.search(r'(20\d{2})', name)
    yr = m.group(1) if m else name.rsplit('.', 1)[0][:12]
    if re.search(r'ytd', name, re.I):
        yr = f'{yr} YTD'
    return yr


def render():
    st.header("📊 Historicals")
    st.caption("Upload income statements — summary or transaction-detail PDFs, or Excel/CSV "
               "P&Ls. You get one tab per "
               "statement plus a combined tab: categories aligned across years, every number a live "
               "formula back to its source, and any year missing a line shown in red.")

    ai_on = hist_llm.available()
    if ai_on:
        st.caption("🤖  AI assist: **on** — Claude aligns variant category names, classifies line "
                   "items, and reads statements the built-in parser can't.")
    else:
        st.caption("AI assist: off (set ANTHROPIC_API_KEY in the app secrets to enable smarter "
                   "matching, classification, and extraction of unfamiliar statement formats).")

    files = st.file_uploader("Statements (PDF / Excel / CSV)", type=tabular.UPLOAD_TYPES,
                             accept_multiple_files=True)
    if not files:
        st.info("Upload two or more statements to begin.")
        return

    st.write("**Confirm each statement** (label = the column header; type auto-detected):")
    items = []
    for i, f in enumerate(files):
        data = f.getvalue()
        c = st.columns([3, 2, 2])
        c[0].markdown(f"`{f.name}`")
        # keys carry the index: two uploads can share a name, and duplicate
        # widget keys crash the page
        label = c[1].text_input("Label", value=_guess_label(f.name), key=f"lbl_{i}_{f.name}")
        sheet = tabular.is_spreadsheet(data)
        auto = S.detect_kind(data)
        kind = c[2].selectbox("Type", ["summary", "detail"],
                              index=0 if auto == "summary" else 1, key=f"knd_{i}_{f.name}",
                              disabled=sheet,
                              help="Spreadsheets are read by Claude line by line — "
                                   "the summary/detail split only applies to PDFs."
                                   if sheet else None)
        items.append((label, kind, data))

    if not st.button("Build workbook", type="primary"):
        return

    summaries, details = [], []
    ai_extracted, from_sheets = [], []
    with st.spinner("Parsing statements & assembling…" + (" (AI assist on)" if ai_on else "")):
        for label, kind, data in items:
            try:
                if tabular.is_spreadsheet(data):
                    # the regex parsers read PDF text layouts; workbooks go to Claude,
                    # which reads every period column — a monthly grid (T-12) keeps
                    # its months, however many of them arrive
                    if not ai_on:
                        st.error(f"'{label}' is a spreadsheet — reading it needs "
                                 "ANTHROPIC_API_KEY set in the app secrets.")
                        return
                    # A workbook often holds one statement per tab — a year each.
                    # Each is read on its own: asked for the whole file at once,
                    # the reader returns the tabs merged into a single column.
                    # Tabs without figures (cover, alert, notes) are skipped
                    # before they cost a read.
                    cand = hist_llm.statement_tabs(data) or [None]
                    for tab in cand:
                        got = hist_llm.extract_statement_periods(data, sheet=tab)
                        periods, rows = got.get("periods") or [], got.get("rows") or []
                        if not any(a is not None for r in rows
                                   for a in (r.get("amounts") or [])):
                            continue              # nothing to put in a column
                        lbl = _tab_label(label, tab, len(cand))
                        det = S.detail_from_periods(periods, rows)
                        from_sheets.append(lbl + (" (monthly)" if det else ""))
                        if det:
                            details.append({"label": lbl, **det})
                        else:
                            summaries.append({"label": lbl,
                                              "rows": S.summary_from_periods(periods, rows)})
                elif kind == "summary":
                    rows = S.parse_summary(data)
                    # weak parse (unfamiliar format)? -> let Claude read the PDF
                    n_items = sum(1 for r in rows if r.get("amount") is not None
                                  and not r.get("total") and not r.get("net"))
                    if n_items < 3 and ai_on:
                        rows = hist_llm.extract_statement(data)
                        ai_extracted.append(label)
                    summaries.append({"label": label, "rows": rows})
                else:
                    d = S.parse_detail(data)
                    if not d["cats"] and ai_on:   # detail parser found nothing
                        rows = hist_llm.extract_statement(data)
                        ai_extracted.append(label)
                        summaries.append({"label": label, "rows": rows})
                    else:
                        details.append({"label": label, **d})
            except Exception as e:  # noqa: BLE001
                st.error(f"Could not parse '{label}': {e}")
                return
        summaries.sort(key=lambda s: s["label"])
        labels = [s["label"] for s in summaries]
        dupes = sorted({l for l in labels if labels.count(l) > 1})
        if dupes:
            st.warning("Two or more statements share the label " +
                       ", ".join(f"'{d}'" for d in dupes) + " — their columns will be "
                       "indistinguishable on the combined tab. Give them distinct labels above.")
        xlsx = S.build_workbook(summaries, details)
    if S.LAST_LLM_ERROR:
        st.warning("The AI pass failed, so the combined tab fell back to the built-in rules — "
                   "category names aren't aligned across years, line items aren't classified, "
                   "and total rows aren't bolded. The numbers themselves are unaffected.\n\n"
                   f"`{S.LAST_LLM_ERROR}`")
    if ai_extracted:
        st.info("Read by AI (built-in parser couldn't handle the format): " + ", ".join(ai_extracted)
                + ". Double-check these against the source PDFs.")
    if from_sheets:
        st.info("Read from the spreadsheet by AI: " + ", ".join(from_sheets)
                + ". Double-check these against the original workbook.")

    st.success(f"Built — {len(summaries)} summary statement(s)"
               + "".join(f" + {d['label']} month-by-month ({len(d['months'])} months)"
                         for d in sorted(details, key=lambda d: d['label']))
               + ".")
    for d in sorted(details, key=lambda d: d["label"]):
        cats, tots = d["cats"], d["totals"] or {}
        recon = [c for c in cats
                 if abs(round(sum(cats[c].values()), 2) - tots.get(c, sum(cats[c].values()))) >= 0.01]
        if recon:
            st.warning(f"{d['label']} — categories where the monthly sum ≠ the printed total "
                       "(review these — likely manual/cash-basis adjustments): " + ", ".join(recon))
    st.download_button("⬇  Download workbook (.xlsx)", data=xlsx, file_name="Historicals.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
