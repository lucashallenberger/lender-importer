"""Historicals — Streamlit page. Upload property income statements (summary or
transaction-detail PDFs); download a workbook with per-statement tabs + a combined
tab (every number linked to its source, categories aligned, missing lines in red)."""

import re
import streamlit as st

from tools import statements as S


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
    with st.spinner("Parsing statements & assembling…"):
        for label, kind, data in items:
            try:
                if kind == "summary":
                    summaries.append({"label": label, "rows": S.parse_summary(data)})
                else:
                    detail = {"label": label, **S.parse_detail(data)}
            except Exception as e:  # noqa: BLE001
                st.error(f"Could not parse '{label}': {e}")
                return
        summaries.sort(key=lambda s: s["label"])
        xlsx = S.build_workbook(summaries, detail)

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
