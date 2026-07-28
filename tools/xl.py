"""Shared workbook helpers.

Sheet titles are a sharp edge in openpyxl:
  • "[]:*?/\\" raise ValueError — and our titles come from filenames and
    statement labels, so they really do contain slashes;
  • >31 chars is accepted but warns and can make Excel refuse the file;
  • a duplicate title is SILENTLY renamed ("S - 2024" -> "S - 20241"), which
    would leave every ='S - 2024'!A1 formula pointing at a sheet that no longer
    exists (#REF!).

new_sheet() handles all three and returns the title actually used, so callers
build their cross-sheet formulas against the real name.
"""

import re

LIMIT = 31
_BAD = re.compile(r"[\[\]:*?/\\']")


def safe_title(base, limit=LIMIT):
    """An Excel-legal sheet title (illegal chars -> '-', trimmed to 31)."""
    t = _BAD.sub("-", str(base if base not in (None, "") else "Sheet"))
    t = re.sub(r"\s+", " ", t).strip(" '")
    return (t[:limit].strip() or "Sheet")


def new_sheet(wb, base, limit=LIMIT):
    """Create a sheet with a legal, unique title. Returns (worksheet, title)."""
    title = safe_title(base, limit)
    if title in wb.sheetnames:
        for i in range(2, 500):
            suffix = f" ({i})"
            cand = title[: limit - len(suffix)].strip() + suffix
            if cand not in wb.sheetnames:
                title = cand
                break
    ws = wb.create_sheet(title)
    return ws, ws.title
