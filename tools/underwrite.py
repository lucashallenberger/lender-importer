"""Underwriting (Beta) — Streamlit page.

Upload everything for ONE property — rent roll PDF, tax bill PDF(s), operating
statement PDF(s) — and get ONE workbook:

  Pro Forma          — the synthesis tab: GPR -> vacancy -> EGI -> OpEx -> NOI ->
                       valuation (cap rate) -> debt/DSCR. Every number is either
                       a live formula into the tabs below or a clearly-marked
                       yellow ASSUMPTION cell. Nothing is invented.
  Rent Roll (+Source)— from the Rent Roll Parser engine
  Historicals        — combined + per-statement tabs from the Historicals engine
  Tax <year> / RE Taxes — from the Tax Bill Parser engine

Any missing document type degrades to editable assumption cells.
"""

import io
import re
from pathlib import Path

import streamlit as st

from tools import hist_llm
from tools import rent_roll as RR
from tools import tax_parser as TX
from tools import statements as ST

YELLOW = "FFF2CC"        # assumption/input cells
NAVY = "203864"
GREEN = "E2EFDA"
MONEY0 = '"$"#,##0'
PCT = '0.00%'

OPEX_BUCKETS = [c for c in hist_llm.CLASSES if c.startswith("OpEx")]


def _build_proforma(ws, name, rr_refs, tax_meta, hist_meta):
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter

    bold = Font(bold=True)
    hdr = Font(bold=True, color="FFFFFF")
    navy = PatternFill("solid", fgColor=NAVY)
    green = PatternFill("solid", fgColor=GREEN)
    yellow = PatternFill("solid", fgColor=YELLOW)
    gray = Font(size=9, color="808080")

    def section(r, title):
        c = ws.cell(r, 1, title); c.font = hdr; c.fill = navy
        for col in (2, 3):
            ws.cell(r, col).fill = navy

    def put(r, label, value, fmt=MONEY0, src="", inp=False, is_bold=False):
        ws.cell(r, 1, label).font = bold if is_bold else Font()
        c = ws.cell(r, 2, value)
        c.number_format = fmt
        if inp:
            c.fill = yellow
        if is_bold:
            c.font = bold
        if src:
            ws.cell(r, 3, src).font = gray
        return r + 1

    ws.cell(1, 1, f"UNDERWRITING PRO FORMA — {name}").font = Font(bold=True, size=14)
    ws.cell(2, 1, "Yellow cells are assumptions — edit them. Everything else is a live "
                  "formula into the source tabs.").font = gray

    r = 4
    # ── INCOME ──────────────────────────────────────────────────────────
    section(r, "INCOME (annual)"); r += 1
    gpr_r = r
    if rr_refs:
        r = put(r, "Gross Potential Rent", f"={rr_refs['gpr_annual_cell']}", src="Rent Roll")
    else:
        r = put(r, "Gross Potential Rent", 0, src="no rent roll uploaded — enter", inp=True)
    vac_r = r
    r = put(r, "Vacancy %", 0.05, fmt=PCT, src="assumption", inp=True)
    vacloss_r = r
    r = put(r, "Vacancy Loss", f"=-B{gpr_r}*B{vac_r}")
    # historicals annualization factor (YTD detail may span ≠ 12 months)
    oth_r = r
    if hist_meta and hist_meta.get("val_letter"):
        H, V = hist_meta["title"], hist_meta["val_letter"]
        f0, f1 = hist_meta["first_row"], hist_meta["last_row"]
        annf = round(12 / hist_meta["val_months"], 4) if hist_meta.get("val_months") else 1.0
        annf_r = r
        r = put(r, "Historicals annualization ×", annf, fmt="0.0000",
                src=f"12 / {hist_meta.get('val_months', 12)} months of '{hist_meta.get('val_label','')}'", inp=True)
        oth_r = r
        r = put(r, "Other Income",
                f"=SUMIF('{H}'!${hist_meta['cls_letter']}${f0}:${hist_meta['cls_letter']}${f1},"
                f"\"Income - Other Income\",'{H}'!${V}${f0}:${V}${f1})*B{annf_r}",
                src=f"Historicals ({hist_meta.get('val_label','')})")
    else:
        annf_r = None
        r = put(r, "Other Income", 0, src="no statements uploaded — enter", inp=True)
    egi_r = r
    r = put(r, "Effective Gross Income", f"=B{gpr_r}+B{vacloss_r}+B{oth_r}", is_bold=True)
    r += 1

    # ── EXPENSES ────────────────────────────────────────────────────────
    section(r, "OPERATING EXPENSES (annual)"); r += 1
    exp_rows = []
    for bucket in OPEX_BUCKETS:
        short = bucket.split(" - ", 1)[1]
        if bucket == "OpEx - RE Taxes" and tax_meta:
            hc, pf = tax_meta["hardcoded_cell"], tax_meta["performula_cell"]
            exp_rows.append(r)
            r = put(r, short, f'=IF({hc}<>"",{hc},{pf})', src=f"Tax bill ({tax_meta['sheet']})")
        elif hist_meta and hist_meta.get("val_letter"):
            H, V = hist_meta["title"], hist_meta["val_letter"]
            f0, f1 = hist_meta["first_row"], hist_meta["last_row"]
            exp_rows.append(r)
            r = put(r, short,
                    f"=SUMIF('{H}'!${hist_meta['cls_letter']}${f0}:${hist_meta['cls_letter']}${f1},"
                    f"\"{bucket}\",'{H}'!${V}${f0}:${V}${f1})*B{annf_r}",
                    src=f"Historicals ({hist_meta.get('val_label','')})")
        else:
            exp_rows.append(r)
            r = put(r, short, 0, src="enter", inp=True)
    opex_r = r
    r = put(r, "Total Operating Expenses", f"=SUM(B{exp_rows[0]}:B{exp_rows[-1]})", is_bold=True)
    r += 1

    # ── NOI + VALUATION ─────────────────────────────────────────────────
    section(r, "NET OPERATING INCOME"); r += 1
    noi_r = r
    r = put(r, "NOI", f"=B{egi_r}-B{opex_r}", is_bold=True)
    r += 1
    section(r, "VALUATION"); r += 1
    cap_r = r
    r = put(r, "Cap Rate", 0.055, fmt=PCT, src="assumption", inp=True)
    val_r = r
    r = put(r, "Implied Value", f"=IF(B{cap_r}=0,0,B{noi_r}/B{cap_r})", is_bold=True)
    units_r = r
    if rr_refs:
        r = put(r, "Units", f"={rr_refs['total_cell']}", fmt="0", src="Rent Roll")
    else:
        r = put(r, "Units", 0, fmt="0", src="enter", inp=True)
    r = put(r, "Value / Unit", f"=IF(B{units_r}=0,0,B{val_r}/B{units_r})")
    r += 1

    # ── DEBT ────────────────────────────────────────────────────────────
    section(r, "DEBT"); r += 1
    loan_r = r
    r = put(r, "Loan Amount", 0, src="assumption", inp=True)
    rate_r = r
    r = put(r, "Interest Rate", 0.065, fmt=PCT, src="assumption", inp=True)
    am_r = r
    r = put(r, "Amortization (years)", 30, fmt="0", src="assumption", inp=True)
    ads_r = r
    r = put(r, "Annual Debt Service",
            f"=IF(B{loan_r}=0,0,-PMT(B{rate_r}/12,B{am_r}*12,B{loan_r})*12)")
    dscr_r = r
    r = put(r, "DSCR", f'=IF(B{ads_r}=0,"",B{noi_r}/B{ads_r})', fmt='0.00"x"')
    r = put(r, "Cash Flow After Debt Service", f"=B{noi_r}-B{ads_r}", is_bold=True)

    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 40
    ws.freeze_panes = "A4"


def build_underwriting(name, rr_data, tax_bills, summaries, detail, use_llm=True):
    """Assemble the single deal workbook. Returns .xlsx bytes."""
    from openpyxl import Workbook
    wb = Workbook(); wb.remove(wb.active)
    pf = wb.create_sheet("Pro Forma")     # leftmost

    rr_refs = RR.build_into(wb, rr_data, name="Rent Roll", src_name="Rent Roll Source") if rr_data else None
    hist_meta = (ST.build_into(wb, summaries, detail, use_llm=use_llm, combined_title="Historicals")
                 if (summaries or detail) else None)
    tax_meta = TX.build_tax_into(wb, tax_bills, prefix="Tax ", combined_name="RE Taxes") if tax_bills else None

    _build_proforma(pf, name, rr_refs, tax_meta, hist_meta)
    buf = io.BytesIO(); wb.save(buf); return buf.getvalue()


# ── UI ───────────────────────────────────────────────────────────────────────
def render():
    st.header("📐 Underwriting (Beta)")
    st.caption("Upload everything for one property — rent roll, tax bills, operating "
               "statements. You get ONE workbook: a Pro Forma tab where every number is "
               "a live formula into the parsed source tabs, and assumptions are yellow "
               "input cells. Missing documents become editable inputs instead.")

    ai_on = hist_llm.available()
    if not ai_on:
        st.warning("Set ANTHROPIC_API_KEY in the app secrets — the rent roll reader "
                   "requires it, and tax/statement extraction is much stronger with it.")

    name = st.text_input("Property name", placeholder="e.g. 2821 Sierra St")

    c1, c2, c3 = st.columns(3)
    rr_file = c1.file_uploader("Rent roll (PDF)", type=["pdf"], key="uw_rr")
    tax_files = c2.file_uploader("Tax bills (PDF)", type=["pdf"], accept_multiple_files=True, key="uw_tax")
    stmt_files = c3.file_uploader("Operating statements (PDF)", type=["pdf"],
                                  accept_multiple_files=True, key="uw_stmt")

    stmt_items = []
    for f in stmt_files or []:
        data = f.getvalue()
        cols = st.columns([3, 2, 2])
        cols[0].markdown(f"`{f.name}`")
        m = re.search(r"(20\d{2})", f.name)
        label = cols[1].text_input("Label", value=(m.group(1) if m else Path(f.name).stem[:12]),
                                   key="uw_lbl_" + f.name)
        auto = ST.detect_kind(data)
        kind = cols[2].selectbox("Type", ["summary", "detail"],
                                 index=0 if auto == "summary" else 1, key="uw_knd_" + f.name)
        stmt_items.append((label, kind, data))

    if not st.button("🏗️ Build underwriting workbook", type="primary",
                     disabled=not (rr_file or tax_files or stmt_items)):
        return

    prop = name.strip() or "Property"
    rr_data, tax_bills, summaries, detail = None, [], [], None
    with st.spinner("Parsing documents & assembling…"):
        try:
            if rr_file:
                if not ai_on:
                    st.error("Rent roll extraction needs ANTHROPIC_API_KEY."); return
                rr_data = RR.extract_rent_roll(rr_file.getvalue())
            for f in tax_files or []:
                pdf_path = TX.OUTPUT_DIR / f.name
                pdf_bytes = bytes(f.getbuffer())
                pdf_path.write_bytes(pdf_bytes)
                d = TX.parse_pdf(pdf_path)
                if ai_on and TX._weak_tax(d):
                    try:
                        d = TX._merge_tax(d, TX.extract_tax_bill(pdf_bytes))
                    except Exception:
                        pass
                shot = TX.pdf_to_screenshot(pdf_path)
                tax_bills.append((d, d.get("apn") or "unknown",
                                  str(shot) if shot else None,
                                  TX._bill_year(d, f.name)))
            for label, kind, data in stmt_items:
                if kind == "summary":
                    rows = ST.parse_summary(data)
                    n_items = sum(1 for r in rows if r.get("amount") is not None
                                  and not r.get("total") and not r.get("net"))
                    if n_items < 3 and ai_on:
                        rows = hist_llm.extract_statement(data)
                    summaries.append({"label": label, "rows": rows})
                else:
                    d = ST.parse_detail(data)
                    if not d["cats"] and ai_on:
                        summaries.append({"label": label, "rows": hist_llm.extract_statement(data)})
                    else:
                        detail = {"label": label, **d}
            summaries.sort(key=lambda s: s["label"])
            xb = build_underwriting(prop, rr_data, tax_bills, summaries, detail, use_llm=ai_on)
        except Exception as e:  # noqa: BLE001
            st.error(f"Build failed: {e}")
            return

    parts = [p for p, ok in [("rent roll", rr_data), ("tax bills", tax_bills),
                             ("statements", summaries or detail)] if ok]
    st.success(f"Built {prop} underwriting workbook from: {', '.join(parts)}. "
               "Open the Pro Forma tab and fill the yellow assumption cells.")
    stem = re.sub(r"[^0-9A-Za-z]+", "_", prop).strip("_") or "deal"
    st.download_button("⬇ Download underwriting workbook (.xlsx)", data=xb,
                       file_name=f"{stem}_Underwriting.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       use_container_width=True)
