"""Historicals — Streamlit page. Upload property income statements (summary or
transaction-detail PDFs); download a workbook with per-statement tabs + a combined
tab (every number linked to its source, categories aligned, missing lines in red)."""

import re
import streamlit as st

from tools import statements as S
from tools import hist_llm


def _guess_label(name):
    m = re.search(r'(20\d{2})', name)
    yr = m.group(1) if m else name.rsplit('.', 1)[0][:12]
    if re.search(r'ytd', name, re.I):
        yr = f'{yr} YTD'
    return yr


def render():
    st.header("Historicals — statement combiner")
    st.caption("Upload income statements (summary or transaction-detail PDFs). You get one tab per "
               "statement plus a combined tab: categories aligned across years, every number a live "
               "formula back to its source, and any year missing a line shown in red.")

    ai_on = hist_llm.available()
    if ai_on:
        st.caption("🤖 AI assist: **on** — Claude aligns variant category names, classifies line "
                   "items, and reads statements the built-in parser can't.")
    else:
        st.caption("AI assist: off (set ANTHROPIC_API_KEY in the app secrets to enable smarter "
                   "matching, classification, and extraction of unfamiliar statement formats).")

    files = st.file_uploader("Statements (PDF)", type=["pdf"], accept_multiple_files=True)
    if not files:
        st.info("Upload two or more statements to begin.")
        return

    st.write("**Confirm each statement** (label = the column header; type auto-detected):")
    items = []
    for f in files:
        data = f.getvalue()
        c = st.columns([3, 2, 2])
        c[0].markdown(f"`{f.name}`")
        label = c[1].text_input("Label", value=_guess_label(f.name), key="lbl_" + f.name)
        auto = S.detect_kind(data)
        kind = c[2].selectbox("Type", ["summary", "detail"],
                              index=0 if auto == "summary" else 1, key="knd_" + f.name)
        items.append((label, kind, data))

    if not st.button("Build workbook", type="primary"):
        return

    summaries, detail = [], None
    ai_extracted = []
    with st.spinner("Parsing statements & assembling…" + (" (AI assist on)" if ai_on else "")):
        for label, kind, data in items:
            try:
                if kind == "summary":
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
                        detail = {"label": label, **d}
            except Exception as e:  # noqa: BLE001
                st.error(f"Could not parse '{label}': {e}")
                return
        summaries.sort(key=lambda s: s["label"])
        xlsx = S.build_workbook(summaries, detail)
    if ai_extracted:
        st.info("Read by AI (built-in parser couldn't handle the format): " + ", ".join(ai_extracted)
                + ". Double-check these against the source PDFs.")

    st.success(f"Built — {len(summaries)} summary statement(s)"
               + (f" + 1 detail ({len(detail['months'])} months)" if detail else "") + ".")
    if detail:
        recon = [c for c in detail["cats"]
                 if abs(round(sum(detail["cats"][c].values()), 2) - detail["totals"].get(c, sum(detail["cats"][c].values()))) >= 0.01]
        if recon:
            st.warning("Categories where the monthly sum ≠ the printed total (review these — likely "
                       "manual/cash-basis adjustments): " + ", ".join(recon))
    st.download_button("⬇ Download workbook (.xlsx)", data=xlsx, file_name="Historicals.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
