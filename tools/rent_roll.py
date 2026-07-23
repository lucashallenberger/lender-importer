"""Rent Roll Parser — Streamlit page.

Upload a rent-roll PDF (any property-management format). Get one workbook with:
  • Source     — a faithful PDF→Excel transcription of the rent roll
  • Worksheet  — the underwriting rent-roll format, with every shared value a live
                 formula back to the Source sheet (in-place rent, tenant, unit,
                 annual rent, totals, occupancy).

Extraction uses Claude (rent rolls vary too much for a regex parser) via the same
plumbing as the Historicals tool; degrades to a clear message if no API key.
"""

import io
import re
from datetime import datetime

import pandas as pd
import streamlit as st

from tools import hist_llm

# ── styling constants (kept close to the reference screenshots) ─────────────
NAVY   = "203864"
GREEN  = "E2EFDA"
PEACH  = "FCE4D6"
BLUE   = "D9E1F2"

MONEY      = '"$"#,##0.00'
MONEY0     = '"$"#,##0'
MONEY_DASH = '"$"#,##0.00;-"$"#,##0.00;"-"'   # zero shows as "-"
PCT        = '0.0%'

UNIT_COLS = ["Unit", "Unit Type", "Monthly Rent", "Status", "Lease Name",
             "Lease Status", "Move-In Date", "Lease End", "Lease Type", "Notes"]


# ── extraction ──────────────────────────────────────────────────────────────
_EXTRACT_PROMPT = """This is a property RENT ROLL. Transcribe it faithfully — do
not compute, merge, or invent anything.

Return the property header and one entry per UNIT (include vacant units).

Header:
- property_name: building name/street (e.g. "2821 Sierra St"), or "" if absent
- city_state_zip: e.g. "Los Angeles, CA 90031", or ""
- apn: assessor parcel number if present, else null
- source_note: one line describing the report source exactly as printed
  (software, who generated it, dates), or ""

For each unit:
- unit: unit number/label exactly as printed (e.g. "01")
- unit_type: floor-plan/type exactly as printed (e.g. "1/1/d", "Single")
- monthly_rent: the unit's monthly rent as a number. For a VACANT unit use the
  target/asking/market rent shown for it (or 0 if none is shown).
- status: "Occupied" or "Vacant"
- lease_name: tenant/lease name, or "" if vacant/blank
- lease_status: e.g. "Active", or ""
- move_in_date: move-in / lease start date as printed (e.g. "10/01/2013"), or ""
- lease_end: lease end as printed — a date or "MTM", or ""
- lease_type: e.g. "MTM", "Fixed", or ""

Use the primary rent column if several are shown. Keep the units in printed order."""

_SCHEMA = {
    "type": "object",
    "properties": {
        "property_name": {"type": "string"},
        "city_state_zip": {"type": "string"},
        "apn": {"anyOf": [{"type": "string"}, {"type": "null"}]},
        "source_note": {"type": "string"},
        "units": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "unit": {"type": "string"},
                    "unit_type": {"type": "string"},
                    "monthly_rent": {"anyOf": [{"type": "number"}, {"type": "null"}]},
                    "status": {"type": "string"},
                    "lease_name": {"type": "string"},
                    "lease_status": {"type": "string"},
                    "move_in_date": {"type": "string"},
                    "lease_end": {"type": "string"},
                    "lease_type": {"type": "string"},
                },
                "required": ["unit", "unit_type", "monthly_rent", "status",
                             "lease_name", "lease_status", "move_in_date",
                             "lease_end", "lease_type"],
                "additionalProperties": False,
            },
        },
    },
    "required": ["property_name", "city_state_zip", "apn", "source_note", "units"],
    "additionalProperties": False,
}


def extract_rent_roll(pdf_bytes: bytes) -> dict:
    """Claude reads the rent-roll PDF and returns the structured dict above."""
    import base64
    resp = hist_llm._client().messages.create(
        model=hist_llm.MODEL,
        max_tokens=16000,
        thinking={"type": "adaptive"},
        messages=[{
            "role": "user",
            "content": [
                {"type": "document",
                 "source": {"type": "base64", "media_type": "application/pdf",
                            "data": base64.standard_b64encode(pdf_bytes).decode()}},
                {"type": "text", "text": _EXTRACT_PROMPT},
            ],
        }],
        output_config={"format": {"type": "json_schema", "schema": _SCHEMA}},
    )
    return hist_llm._json_response(resp)


# ── helpers ─────────────────────────────────────────────────────────────────
def _is_vacant(u: dict) -> bool:
    return "vac" in str(u.get("status", "")).lower()


def _norm_type(t: str) -> str:
    """Present a bed/bath type in '1 + 1' form; 'Studio'/'Single' -> 'Studio'."""
    if not t:
        return ""
    s = t.lower()
    if "studio" in s or "single" in s:
        return "Studio"
    m = re.match(r'\s*(\d+)\s*[/x+]\s*(\d+)', s)
    if m:
        return f"{m.group(1)} + {m.group(2)}"
    return t


def _num(x):
    try:
        return float(x)
    except (TypeError, ValueError):
        return 0.0


def _beds(u):
    """Bedroom count from the unit type; Studio/Single -> 0, else leading int."""
    s = str(u.get("unit_type", "")).lower()
    if "studio" in s or "single" in s:
        return 0
    m = re.match(r'\s*(\d+)', s)
    return int(m.group(1)) if m else None


def _movein(u):
    s = str(u.get("move_in_date", "")).strip()
    for fmt in ("%m/%d/%Y", "%m/%d/%y", "%Y-%m-%d", "%m-%d-%Y"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            pass
    return None


def _auto_notes(units):
    """Analytical flags per unit (biggest/smallest, highest/lowest rent, tenure,
    vacancy) — written into the Notes column instead of extra columns.
    Returns {index: note-string}. Only labels a superlative when it's unique."""
    n = len(units)
    tags = {i: [] for i in range(n)}

    for i, u in enumerate(units):
        if _is_vacant(u):
            tags[i].append("Vacant — target/asking rent")

    # biggest / smallest unit by bedroom count (only when there's a mix)
    beds = {i: _beds(u) for i, u in enumerate(units)}
    known = [v for v in beds.values() if v is not None]
    if known and len(set(known)) > 1:
        big = [i for i, v in beds.items() if v == max(known)]
        small = [i for i, v in beds.items() if v == min(known)]
        if len(big) == 1:
            tags[big[0]].append("Biggest unit")
        if len(small) == 1:
            tags[small[0]].append("Smallest unit")

    # highest / lowest in-place rent among occupied units with a real rent
    rents = {i: _num(u.get("monthly_rent")) for i, u in enumerate(units)
             if not _is_vacant(u) and _num(u.get("monthly_rent")) > 0}
    if len(rents) > 1:
        hi = max(rents, key=rents.get)
        lo = min(rents, key=rents.get)
        tags[hi].append("Highest rent")
        if lo != hi:
            tags[lo].append("Lowest rent")

    # longest-tenured / newest lease among occupied units with a parseable date
    dates = {i: _movein(u) for i, u in enumerate(units)
             if not _is_vacant(u) and _movein(u)}
    if len(dates) > 1:
        oldest = min(dates, key=dates.get)
        newest = max(dates, key=dates.get)
        tags[oldest].append("Longest-tenured")
        if newest != oldest:
            tags[newest].append("Newest lease")

    return {i: " · ".join(t) for i, t in tags.items()}


# ── workbook: Source sheet (faithful transcription) ─────────────────────────
def _build_source(ws, data):
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter

    hdr_font = Font(bold=True, color="FFFFFF")
    navy = PatternFill("solid", fgColor=NAVY)
    green = PatternFill("solid", fgColor=GREEN)
    peach = PatternFill("solid", fgColor=PEACH)
    blue = PatternFill("solid", fgColor=BLUE)
    bold = Font(bold=True)
    title = Font(bold=True, size=14)
    sub = Font(size=11, color="595959")
    note = Font(italic=True, size=9, color="808080")

    units = data["units"]
    n = len(units)

    # title block
    ws.cell(1, 1, f"{data.get('property_name') or 'Property'} — Rent Roll").font = title
    subline = " · ".join(x for x in [data.get("city_state_zip"),
                                     f"APN {data['apn']}" if data.get("apn") else None,
                                     f"{n} Units"] if x)
    ws.cell(2, 1, subline).font = sub
    if data.get("source_note"):
        ws.cell(3, 1, f"Source: {data['source_note']}").font = note

    headers = ["Unit", "Type", "Monthly Rent", "Status", "Lease Name",
               "Lease Status", "Move-In Date"]
    hrow = 5
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(hrow, c, h)
        cell.font = hdr_font; cell.fill = navy
        cell.alignment = Alignment(horizontal="center")

    ds = hrow + 1
    for i, u in enumerate(units):
        r = ds + i
        vac = _is_vacant(u)
        ws.cell(r, 1, u.get("unit", "")).alignment = Alignment(horizontal="center")
        ws.cell(r, 2, u.get("unit_type", ""))
        rc = ws.cell(r, 3, _num(u.get("monthly_rent"))); rc.number_format = MONEY
        ws.cell(r, 4, "Vacant" if vac else "Occupied")
        ws.cell(r, 5, "" if vac else u.get("lease_name", ""))
        ws.cell(r, 6, "" if vac else u.get("lease_status", ""))
        ws.cell(r, 7, u.get("move_in_date", ""))
        if vac:
            for c in range(1, 8):
                ws.cell(r, c).fill = peach
                ws.cell(r, c).font = Font(italic=True, color="C00000")
    de = ds + n - 1

    # TOTAL row
    tr = de + 1
    ws.cell(tr, 1, "TOTAL").font = bold
    ws.cell(tr, 2, f"{n} units").font = bold
    tc = ws.cell(tr, 3, f"=SUM(C{ds}:C{de})"); tc.font = bold; tc.number_format = MONEY
    for c in range(1, 8):
        ws.cell(tr, c).fill = blue

    # Rent Summary block (live formulas)
    sr = tr + 2
    ws.cell(sr, 1, "Rent Summary").font = bold
    rows = [
        ("Units — Total", n, None),
        ("Units — Occupied", f'=COUNTIF(D{ds}:D{de},"Occupied")', None),
        ("Units — Vacant", f'=COUNTIF(D{ds}:D{de},"Vacant")', None),
        ("Occupancy %", f'=C{sr+2}/C{sr+1}', PCT),
        ("In-Place Monthly Rent (occupied only)", f'=SUMIF(D{ds}:D{de},"Occupied",C{ds}:C{de})', MONEY),
        ("Target Rent on Vacant Unit(s)", f'=SUMIF(D{ds}:D{de},"Vacant",C{ds}:C{de})', MONEY),
        ("Gross Potential Monthly Rent (in-place + target)", f'=SUM(C{ds}:C{de})', MONEY),
        ("Annualized Gross Potential Rent", f'=C{sr+7}*12', MONEY),
    ]
    for j, (label, val, fmt) in enumerate(rows, start=1):
        rr = sr + j
        ws.cell(rr, 1, label)
        cell = ws.cell(rr, 3, val)
        if fmt:
            cell.number_format = fmt
        for c in range(1, 8):
            ws.cell(rr, c).fill = green
    ws.cell(sr, 1).fill = green

    widths = [8, 12, 16, 12, 22, 14, 14]
    for c, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(c)].width = w

    return {
        "data_start": ds, "data_end": de,
        "occupancy_cell": f"'Source'!C{sr+4}",
        "total_cell": f"'Source'!C{sr+1}",
        "occupied_cell": f"'Source'!C{sr+2}",
        "vacant_cell": f"'Source'!C{sr+3}",
    }


# ── workbook: Worksheet sheet (underwriting format, sourced) ────────────────
def _build_worksheet(ws, data, refs):
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter

    hdr_font = Font(bold=True, color="FFFFFF")
    navy = PatternFill("solid", fgColor=NAVY)
    green = PatternFill("solid", fgColor=GREEN)
    peach = PatternFill("solid", fgColor=PEACH)
    bold = Font(bold=True)
    note = Font(italic=True, size=9, color="808080")

    headers = ["UNIT", "TENANT", "UNIT TYPE", "UNIT TYPE", "IN-PLACE RENT ($)",
               "MARKET RENT ($)", "LEASE TYPE", "LEASE START", "LEASE END",
               "IN-PLACE VACANCY", "LOSS TO LEASE", "ANNUAL IN-PLACE RENT", "NOTES"]
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(1, c, h)
        cell.font = hdr_font; cell.fill = navy
        cell.alignment = Alignment(horizontal="center", wrap_text=True, vertical="center")

    units = data["units"]
    n = len(units)
    ds = refs["data_start"]
    auto_notes = _auto_notes(units)

    for i, u in enumerate(units):
        wr = 2 + i           # worksheet row
        sr = ds + i          # matching Source row
        vac = _is_vacant(u)
        ws.cell(wr, 1, f"='Source'!A{sr}").alignment = Alignment(horizontal="center")
        ws.cell(wr, 2, "VACANT" if vac else f"='Source'!E{sr}")
        ws.cell(wr, 3, f"='Source'!B{sr}")
        ws.cell(wr, 4, _norm_type(u.get("unit_type", "")))
        e = ws.cell(wr, 5, 0 if vac else f"='Source'!C{sr}"); e.number_format = MONEY
        f = ws.cell(wr, 6, None); f.number_format = MONEY          # Market Rent — input
        ws.cell(wr, 7, u.get("lease_type", ""))
        ws.cell(wr, 8, f"='Source'!G{sr}")
        ws.cell(wr, 9, u.get("lease_end", ""))
        jv = ws.cell(wr, 10, f"='Source'!C{sr}" if vac else 0); jv.number_format = MONEY_DASH
        k = ws.cell(wr, 11, f'=IF(F{wr}="",0,F{wr}-E{wr})'); k.number_format = MONEY_DASH
        li = ws.cell(wr, 12, f"=E{wr}*12"); li.number_format = MONEY0
        ws.cell(wr, 13, u["notes"] if "notes" in u else auto_notes.get(i, ""))
        if vac:
            for c in range(1, 14):
                ws.cell(wr, c).fill = peach
                if ws.cell(wr, c).value in (None, 0) or c in (1, 2, 4):
                    ws.cell(wr, c).font = Font(italic=True, color="C00000")

    last = 1 + n
    # Monthly Total
    mt = last + 1
    ws.cell(mt, 1, "Monthly Total").font = bold
    for col in (5, 6, 10, 11, 12):
        L = get_column_letter(col)
        cell = ws.cell(mt, col, f"=SUM({L}2:{L}{last})")
        cell.font = bold
        cell.number_format = MONEY0 if col == 12 else MONEY_DASH if col in (10, 11) else MONEY0
    # Annual Total
    at = mt + 1
    ws.cell(at, 1, "Annual Total").font = bold
    for col, formula in ((5, f"=E{mt}*12"), (6, f"=F{mt}*12"),
                         (10, f"=J{mt}*12"), (12, f"=L{mt}")):
        cell = ws.cell(at, col, formula); cell.font = bold; cell.number_format = MONEY0
    for r in (mt, at):
        for c in range(1, 14):
            ws.cell(r, c).fill = green

    # Unit summary (references the Source sheet)
    us = at + 2
    summ = [("Total Units", refs["total_cell"], None),
            ("Occupied Units", refs["occupied_cell"], None),
            ("Vacant Units", refs["vacant_cell"], None),
            ("Occupancy", refs["occupancy_cell"], PCT)]
    for j, (label, ref, fmt) in enumerate(summ):
        rr = us + j
        ws.cell(rr, 2, label).font = bold
        cell = ws.cell(rr, 5, f"={ref}")
        if fmt:
            cell.number_format = fmt
        for c in range(2, 6):
            ws.cell(rr, c).fill = green

    # source note
    if data.get("source_note"):
        ws.cell(us + 5, 2, f"Source: {data['source_note']}  ·  "
                           "Market Rent left blank for input; vacant unit shows target/asking rent.").font = note

    widths = [7, 20, 10, 10, 15, 15, 11, 13, 12, 15, 14, 17, 34]
    for c, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.row_dimensions[1].height = 30


def build_workbook(data: dict) -> bytes:
    """Return .xlsx bytes: a Worksheet tab (underwriting format, sourced) followed
    by a Source tab (faithful transcription the Worksheet links back to)."""
    from openpyxl import Workbook
    wb = Workbook(); wb.remove(wb.active)
    ws = wb.create_sheet("Worksheet")
    src = wb.create_sheet("Source")
    refs = _build_source(src, data)
    _build_worksheet(ws, data, refs)
    buf = io.BytesIO(); wb.save(buf); return buf.getvalue()


# ── UI ───────────────────────────────────────────────────────────────────────
def _df_from_units(units):
    notes = _auto_notes(units)
    return pd.DataFrame([{
        "Unit": u.get("unit", ""), "Unit Type": u.get("unit_type", ""),
        "Monthly Rent": _num(u.get("monthly_rent")),
        "Status": "Vacant" if _is_vacant(u) else "Occupied",
        "Lease Name": u.get("lease_name", ""), "Lease Status": u.get("lease_status", ""),
        "Move-In Date": u.get("move_in_date", ""), "Lease End": u.get("lease_end", ""),
        "Lease Type": u.get("lease_type", ""), "Notes": notes.get(i, ""),
    } for i, u in enumerate(units)], columns=UNIT_COLS)


def _units_from_df(df):
    out = []
    for row in df.itertuples(index=False):
        d = dict(zip(UNIT_COLS, row))
        if not str(d["Unit"]).strip():
            continue
        out.append({
            "unit": str(d["Unit"]).strip(), "unit_type": str(d["Unit Type"]).strip(),
            "monthly_rent": _num(d["Monthly Rent"]), "status": str(d["Status"]).strip(),
            "lease_name": str(d["Lease Name"]).strip(), "lease_status": str(d["Lease Status"]).strip(),
            "move_in_date": str(d["Move-In Date"]).strip(), "lease_end": str(d["Lease End"]).strip(),
            "lease_type": str(d["Lease Type"]).strip(), "notes": str(d["Notes"]).strip(),
        })
    return out


def render():
    st.header("🏘️ Rent Roll Parser")
    st.caption("Upload a rent-roll PDF. You get one workbook: a **Source** tab (faithful "
               "PDF→Excel transcription) and a **Worksheet** tab (underwriting format) whose "
               "shared values link by formula back to the Source tab.")

    if not hist_llm.available():
        st.warning("This tool reads rent rolls with Claude — set **ANTHROPIC_API_KEY** in the "
                   "app secrets to enable it. (Rent-roll layouts vary too much for a fixed parser.)")
        return

    files = st.file_uploader("Rent roll PDF(s)", type=["pdf"], accept_multiple_files=True)
    if not files:
        st.info("Upload a rent-roll PDF to begin.")
        return

    parsed = st.session_state.setdefault("rr_parsed", {})
    for f in files:
        key = f"{f.name}:{f.size}"
        if key not in parsed:
            with st.spinner(f"Reading {f.name} with Claude…"):
                try:
                    parsed[key] = extract_rent_roll(f.getvalue())
                except Exception as e:  # noqa: BLE001
                    st.error(f"Could not read {f.name}: {e}")
                    continue
    # drop removed files
    live = {f"{f.name}:{f.size}" for f in files}
    for k in list(parsed):
        if k not in live:
            del parsed[k]

    for f in files:
        key = f"{f.name}:{f.size}"
        data = parsed.get(key)
        if not data:
            continue
        with st.expander(f"📄 {f.name}  —  {data.get('property_name') or '?'}",
                         expanded=len(files) == 1):
            c1, c2, c3 = st.columns(3)
            pname = c1.text_input("Property", value=data.get("property_name") or "", key=f"pn_{key}")
            csz = c2.text_input("City / State / ZIP", value=data.get("city_state_zip") or "", key=f"csz_{key}")
            apn = c3.text_input("APN", value=data.get("apn") or "", key=f"apn_{key}")
            snote = st.text_input("Source note", value=data.get("source_note") or "", key=f"sn_{key}")

            st.markdown("**Units** (edit any cell to correct the extraction)")
            df = st.data_editor(_df_from_units(data["units"]), num_rows="dynamic",
                                use_container_width=True, key=f"units_{key}")

            edited = {"property_name": pname, "city_state_zip": csz, "apn": apn or None,
                      "source_note": snote, "units": _units_from_df(df)}

            if edited["units"]:
                xlsx = build_workbook(edited)
                stem = re.sub(r'[^0-9A-Za-z]+', "_", (pname or f.name)).strip("_") or "rent_roll"
                st.download_button("⬇ Download workbook (.xlsx)", data=xlsx,
                                   file_name=f"{stem}_RentRoll.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                   key=f"dl_{key}", use_container_width=True)
            else:
                st.info("No units yet — add at least one row to build the workbook.")
