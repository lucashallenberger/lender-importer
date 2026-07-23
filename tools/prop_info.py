"""Property Info — a cover sheet with PROPERTY INFORMATION / BUILDING INFORMATION
blocks (user's exact format: blue headers, gold value cells), plus a web lookup:
give an APN (and optional name/address hint) and Claude researches the county
assessor / ZIMAS / listing sites via web search and fills what it can find.
Unknown fields stay blank (gold = fill-in)."""

import json
import re

from tools import hist_llm

# (row label, dict key) in exact sheet order
PROP_FIELDS = [
    ("Property Name", "property_name"),
    ("Property Address", "address_line1"),      # line 2 goes on the next row
    ("County", "county"),
    ("APN", "apn"),
    ("Zoning", "zoning"),
    ("Land Area (acres)", "land_acres"),
    ("Land Area (SF)", "land_sf"),
    ("Parking", "parking"),
    ("Number of Spaces", "parking_spaces"),
]
BLDG_FIELDS = [
    ("Property Type", "property_type"),
    ("Year Built", "year_built"),
    ("Number of Buildings", "num_buildings"),
    ("Number of Stories", "num_stories"),
    ("Number of Units", "num_units"),
    ("Gross SF", "gross_sf"),
    ("Leasable SF", "leasable_sf"),
]
ALL_KEYS = [k for _, k in PROP_FIELDS] + ["address_line2"] + [k for _, k in BLDG_FIELDS]

_PROMPT = """Research this property using web search and report what you find.

APN (Assessor's Parcel Number): {apn}
{hint}

Check the county assessor's parcel records (for Los Angeles County: the LA County
Assessor portal, and ZIMAS zimas.lacity.org for zoning), plus listing/data sites
(Redfin, Zillow, LoopNet, PropertyShark) as needed. Prefer official assessor data
when sources disagree.

When done, output ONLY a JSON object (no other text after it) with these keys —
use null for anything you could not verify; do NOT guess:
- property_name: short name, usually the street address (e.g. "2821 Sierra")
- address_line1: street address (e.g. "2821 N SIERRA ST")
- address_line2: city, state zip (e.g. "Lincoln Heights, CA 90031")
- county: e.g. "Los Angeles"
- apn: echo the APN, digits/dashes as commonly written
- zoning: zoning code (e.g. "[Q]R1-1D-HCR")
- land_acres: number
- land_sf: number
- parking: e.g. "Surface", "Garage", or null
- parking_spaces: number or null
- property_type: e.g. "Multifamily", "Office", "Retail"
- year_built: e.g. 1923 or "1923/2026" if renovated
- num_buildings: number
- num_stories: number
- num_units: number
- gross_sf: number
- leasable_sf: number or null"""


def fetch(apn: str, hint: str = "") -> dict:
    """Web-research the APN and return the property-info dict (missing keys None)."""
    resp = hist_llm._client().messages.create(
        model=hist_llm.MODEL,
        max_tokens=8000,
        thinking={"type": "adaptive"},
        tools=[{"type": "web_search_20250305", "name": "web_search", "max_uses": 8}],
        messages=[{"role": "user", "content": _PROMPT.format(
            apn=apn, hint=f"Known name/address hint: {hint}" if hint.strip() else "")}],
    )
    text = "".join(b.text for b in resp.content if b.type == "text")
    # the answer ends with a flat JSON object (often ```json fenced) — take the last
    info = {}
    for raw in re.findall(r"\{[^{}]*\}", text, re.S)[::-1]:
        try:
            info = json.loads(raw)
            break
        except Exception:
            continue
    return {k: info.get(k) for k in ALL_KEYS}


def build_sheet(ws, info: dict):
    """Write the Property Info sheet in the user's exact layout."""
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

    blue = PatternFill("solid", fgColor="305496")
    gold = PatternFill("solid", fgColor="FFC000")
    hdr = Font(bold=True, color="FFFFFF")
    thin = Side(style="thin", color="404040")
    box = Border(top=thin, bottom=thin, left=thin, right=thin)
    right = Alignment(horizontal="right", vertical="center")

    def header(row, c0, c1, title):
        ws.merge_cells(start_row=row, start_column=c0, end_row=row, end_column=c1)
        c = ws.cell(row, c0, title); c.font = hdr; c.fill = blue
        for cc in range(c0, c1 + 1):
            ws.cell(row, cc).fill = blue; ws.cell(row, cc).border = box

    def field(row, lc, vc, label, val, fmt=None):
        ws.cell(row, lc, label).border = box
        for cc in range(lc + 1, vc + 1):
            cell = ws.cell(row, cc)
            cell.fill = gold; cell.border = box
        c = ws.cell(row, vc)
        if val is not None and val != "":
            c.value = val
            if fmt and isinstance(val, (int, float)):
                c.number_format = fmt
        c.alignment = right

    NUMFMT = {"land_acres": "0.00", "land_sf": "#,##0.00", "parking_spaces": "0",
              "year_built": "0", "num_buildings": "0", "num_stories": "0",
              "num_units": "0", "gross_sf": "#,##0", "leasable_sf": "#,##0"}

    # left block — PROPERTY INFORMATION (labels A, values C)
    header(1, 1, 3, "PROPERTY INFORMATION")
    r = 2
    for label, key in PROP_FIELDS:
        field(r, 1, 3, label, info.get(key), NUMFMT.get(key))
        r += 1
        if key == "address_line1":               # second address line on its own row
            field(r, 1, 3, "", info.get("address_line2"))
            r += 1

    # right block — BUILDING INFORMATION (labels E, values G)
    header(1, 5, 7, "BUILDING INFORMATION")
    r = 2
    for label, key in BLDG_FIELDS:
        field(r, 5, 7, label, info.get(key), NUMFMT.get(key))
        r += 1

    for col, w in (("A", 20), ("B", 6), ("C", 24), ("D", 3), ("E", 22), ("F", 8), ("G", 18)):
        ws.column_dimensions[col].width = w
