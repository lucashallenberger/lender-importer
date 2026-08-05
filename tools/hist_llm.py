"""Optional Claude-powered layer for the Historicals tool.

Activates only when ANTHROPIC_API_KEY is present (environment or Streamlit
secrets). Everything degrades gracefully to the deterministic path without it.

Capabilities:
  1. extract_statement(pdf_bytes)          — Claude reads a statement PDF the
     deterministic parser can't handle (other software, scans) and returns the
     same row structure parse_summary produces.
  2. match_labels(sources, targets)        — align variant line-item names
     across statements ("City Reg." <-> "City Registration Fee").
  3. classify_labels(labels)               — underwriting bucket per line item.

Answers are cached in tools/data/hist_learned.json so repeat builds are free.
"""

import base64
import json
import os
from pathlib import Path

MODEL = "claude-opus-5"

CLASSES = [
    "Income - Rental Income", "Income - Other Income",
    "OpEx - R&M", "OpEx - Utilities", "OpEx - Management Fee",
    "OpEx - RE Taxes", "OpEx - G&A", "OpEx - Insurance",
    "CapEx - Capital Improvements", "CapEx - Leasing Commissions",
    "Non-OpEx - Mortgage Interest", "Non-OpEx - Legal & Professional",
    "Non-OpEx - Other",
]

_CACHE_PATH = Path(__file__).parent / "data" / "hist_learned.json"


# ---------------------------------------------------------------- plumbing
def api_key():
    key = os.environ.get("ANTHROPIC_API_KEY")
    if key:
        return key
    try:  # Streamlit Cloud exposes secrets via st.secrets
        import streamlit as st
        return st.secrets.get("ANTHROPIC_API_KEY")
    except Exception:
        return None


def available():
    return bool(api_key())


def _client():
    import anthropic
    return anthropic.Anthropic(api_key=api_key())


def _load_cache():
    try:
        with open(_CACHE_PATH) as fh:
            d = json.load(fh)
            d.setdefault("aliases", {}); d.setdefault("classes", {})
            return d
    except Exception:
        return {"aliases": {}, "classes": {}}


def _save_cache(d):
    try:
        _CACHE_PATH.parent.mkdir(parents=True, exist_ok=True)
        with open(_CACHE_PATH, "w") as fh:
            json.dump(d, fh, indent=1, sort_keys=True)
    except Exception:
        pass


def check_refusal(resp):
    """Opus 5 ships cybersecurity/bio classifiers that can decline a request: the
    call returns 200 with stop_reason 'refusal' and no usable content. Surface it
    as a readable error instead of an IndexError three frames down."""
    if getattr(resp, "stop_reason", None) == "refusal":
        why = getattr(getattr(resp, "stop_details", None), "explanation", None)
        raise RuntimeError("Claude declined to process this document"
                           + (f" — {why}" if why else "")
                           + ". If it's an ordinary property document, re-upload it; "
                             "otherwise enter the numbers by hand.")
    return resp


def _text_of(resp):
    return "".join(b.text for b in check_refusal(resp).content if b.type == "text")


def _json_response(resp):
    return json.loads(next(b.text for b in check_refusal(resp).content if b.type == "text"))


_SHEET_INTRO = """The document below is a SPREADSHEET rendered as text: one line per
row, cells separated by ' | ', each sheet introduced by a '=== SHEET: name ===' banner.
Empty cells render as blanks. Formula cells show their computed value where the file
carries one; a cell still showing a formula (e.g. '=E2*12') is a derived column the
file never stored a result for — ignore it and use the underlying values instead.
Read the rest exactly as you would read the printed report."""


def content_blocks(data, prompt, transform=None):
    """Message content for a source document. PDFs go through as a document
    block; workbooks and CSVs are rendered to text first (the API can't read an
    .xlsx), so every extractor accepts either without caring which it got.
    `transform` gets a crack at the rendering — a caller that knows some of the
    sheet is noise can drop it before it's paid for."""
    from tools import tabular
    if tabular.kind(data) == "pdf":
        return [{"type": "document",
                 "source": {"type": "base64", "media_type": "application/pdf",
                            "data": base64.standard_b64encode(bytes(data)).decode()}},
                {"type": "text", "text": prompt}]
    rendered = tabular.to_text(data)
    if transform:
        rendered = transform(rendered)
    return [{"type": "text", "text": f"{_SHEET_INTRO}\n\n{rendered}"},
            {"type": "text", "text": prompt}]


# ---------------------------------------------------------------- 1) extraction
_EXTRACT_PROMPT = """This is a property operating/income statement (cash flow report).
Extract every financial line in top-to-bottom order.

For each line give:
- label: the line-item name exactly as printed (trimmed)
- amount: the period amount as a number (negative if negative), or null for
  section headers with no amount
- kind: "item" (a normal line item), "total" (a subtotal/total line),
  "net" (a net income/NOI/cash-flow line), or "section" (a header like
  Income/Expenses with no amount)

Rules: use the main period column if there are multiple columns. Do not invent,
merge, or skip lines. Do not compute anything — transcribe only."""


def extract_statement(data):
    """Claude-side extraction -> rows shaped like parse_summary's output.
    Accepts a statement PDF or a spreadsheet/CSV export of one."""
    schema = {
        "type": "object",
        "properties": {
            "rows": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {
                        "label": {"type": "string"},
                        "amount": {"anyOf": [{"type": "number"}, {"type": "null"}]},
                        "kind": {"type": "string", "enum": ["item", "total", "net", "section"]},
                    },
                    "required": ["label", "amount", "kind"],
                    "additionalProperties": False,
                },
            }
        },
        "required": ["rows"],
        "additionalProperties": False,
    }
    resp = _client().messages.create(
        model=MODEL,
        max_tokens=16000,
        thinking={"type": "adaptive"},
        messages=[{"role": "user", "content": content_blocks(data, _EXTRACT_PROMPT)}],
        output_config={"format": {"type": "json_schema", "schema": schema}},
    )
    rows = []
    for r in _json_response(resp)["rows"]:
        row = {"label": r["label"].strip(), "amount": r["amount"],
               "total": r["kind"] == "total", "net": r["kind"] == "net"}
        if r["kind"] == "section":
            row["section"] = True; row["amount"] = None
        rows.append(row)
    return rows


# ------------------------------------------------- 1b) period-aware extraction
_PERIODS_PROMPT = """This is a property operating/income statement (P&L, cash flow
report, or T-12). Extract every financial line in top-to-bottom order.

First identify the statement's VALUE COLUMNS ("periods"):
- A monthly statement (T-12) has one column per month — list them in printed order,
  e.g. ["Jan-25","Feb-25",...]. Include a trailing Total / YTD column only if one is
  actually printed, labelled exactly as printed.
- A single-period statement has one value column — return one label for it (or "" if
  the column is unlabelled).

For each line give:
- label: the line-item name exactly as printed (trimmed)
- amounts: one number per period column, in the SAME ORDER as `periods` (negative if
  negative); use null where the cell is blank
- kind: "item" (a normal line item), "total" (a subtotal/total line), "net" (a net
  income/NOI/cash-flow line), or "section" (a header like Income/Expenses with no
  amounts)

Rules: do not invent, merge, or skip lines. Do not compute anything — transcribe only."""


def extract_statement_periods(data):
    """Read a statement keeping every period column separate. Returns
    {'periods': [label...], 'rows': [{'label','amounts','kind'}...]} — the caller
    decides whether that's a monthly grid or a single-period statement."""
    schema = {
        "type": "object",
        "properties": {
            "periods": {"type": "array", "items": {"type": "string"}},
            "rows": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {
                        "label": {"type": "string"},
                        "amounts": {"type": "array",
                                    "items": {"anyOf": [{"type": "number"}, {"type": "null"}]}},
                        "kind": {"type": "string", "enum": ["item", "total", "net", "section"]},
                    },
                    "required": ["label", "amounts", "kind"],
                    "additionalProperties": False,
                },
            },
        },
        "required": ["periods", "rows"],
        "additionalProperties": False,
    }
    # A T-12 is 12 numbers per line over ~40 lines, so this one runs long — stream
    # it (the SDK refuses a non-streaming request this size, and would time out).
    with _client().messages.stream(
        model=MODEL,
        max_tokens=32000,
        thinking={"type": "adaptive"},
        messages=[{"role": "user", "content": content_blocks(data, _PERIODS_PROMPT)}],
        output_config={"format": {"type": "json_schema", "schema": schema}},
    ) as stream:
        return _json_response(stream.get_final_message())


# ---------------------------------------------------------------- 2) matching
def match_labels(sources, targets):
    """Map each source label to the target label that is the SAME line item, or
    None. Returns {source: target-or-None}. Cached per (source, target-set)."""
    if not sources or not targets:
        return {}
    cache = _load_cache()
    tkey = "|".join(sorted(targets))
    out, todo = {}, []
    for s in sources:
        hit = cache["aliases"].get(f"{s}::{tkey}")
        if hit is not None:
            out[s] = hit if hit else None
        else:
            todo.append(s)
    if todo:
        schema = {
            "type": "object",
            "properties": {"matches": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {
                        "source": {"type": "string"},
                        "target": {"anyOf": [{"type": "string"}, {"type": "null"}]},
                    },
                    "required": ["source", "target"],
                    "additionalProperties": False,
                },
            }},
            "required": ["matches"], "additionalProperties": False,
        }
        prompt = (
            "These are line-item names from property operating statements produced by "
            "different reports of the same property.\n\n"
            f"SOURCE names:\n{json.dumps(todo, indent=1)}\n\n"
            f"TARGET names:\n{json.dumps(sorted(targets), indent=1)}\n\n"
            "For each SOURCE name, return the TARGET name that refers to the SAME "
            "line item/category (naming variants, abbreviations, or nested forms like "
            "'DWP:DWP Electric' vs 'DWP Electric'), or null if no target is the same "
            "item. Never match two genuinely different categories."
        )
        resp = _client().messages.create(
            model=MODEL, max_tokens=8000, thinking={"type": "adaptive"},
            messages=[{"role": "user", "content": prompt}],
            output_config={"format": {"type": "json_schema", "schema": schema}},
        )
        tset, sset = set(targets), set(todo)
        for m in _json_response(resp)["matches"]:
            s = m["source"]
            if s not in sset:            # model must echo inputs exactly — drop junk
                continue
            t = m["target"] if m["target"] in tset else None
            out[s] = t
            cache["aliases"][f"{s}::{tkey}"] = t or ""
        _save_cache(cache)
    return out


# ---------------------------------------------------------------- 3) classification
def classify_labels(labels):
    """Assign each label one of CLASSES. Returns {label: class}. Cached."""
    if not labels:
        return {}
    cache = _load_cache()
    out, todo = {}, []
    for l in labels:
        if l in cache["classes"]:
            out[l] = cache["classes"][l]
        else:
            todo.append(l)
    if todo:
        schema = {
            "type": "object",
            "properties": {"items": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {
                        "label": {"type": "string"},
                        "classification": {"type": "string", "enum": CLASSES},
                    },
                    "required": ["label", "classification"],
                    "additionalProperties": False,
                },
            }},
            "required": ["items"], "additionalProperties": False,
        }

        def ask(labels):
            prompt = (
                "Classify each property operating-statement line item into exactly one "
                "underwriting bucket. Echo every label EXACTLY as given, one entry per "
                "label, no additions.\n\nBuckets:\n" + "\n".join(f"- {c}" for c in CLASSES) +
                f"\n\nLine items:\n{json.dumps(labels, indent=1)}\n\n"
                "Notes: utilities (power/water/sewer/gas/trash/internet) -> 'OpEx - Utilities'. "
                "Routine repairs, maintenance, cleaning, pest, gardening, supplies -> 'OpEx - R&M'. "
                "Mortgage/debt service -> 'Non-OpEx - Mortgage Interest'. Deposits, escrow, "
                "prepayments, one-time non-operating flows -> 'Non-OpEx - Other'."
            )
            resp = _client().messages.create(
                model=MODEL, max_tokens=8000, thinking={"type": "adaptive"},
                messages=[{"role": "user", "content": prompt}],
                output_config={"format": {"type": "json_schema", "schema": schema}},
            )
            return _json_response(resp)["items"]

        pending = list(todo)
        for _attempt in range(2):            # one retry for anything dropped/mangled
            pset = set(pending)
            for m in ask(pending):
                if m["label"] in pset:       # accept only exact echoes of inputs
                    out[m["label"]] = m["classification"]
                    cache["classes"][m["label"]] = m["classification"]
            pending = [l for l in pending if l not in out]
            if not pending:
                break
        _save_cache(cache)
    return out


# ---------------------------------------------------------------- 4) row roles
def label_roles(labels):
    """Classify each label's ROLE on an income statement: 'total' (any total,
    subtotal, or net line — 'Total X', 'Net Operating Income', 'NOI', 'Gross
    Income', 'Net Income'), 'section' (a bare header with no amount like 'Income'
    or 'Expenses'), or 'item' (a normal line item). Returns {label: role}.
    Cached in hist_learned.json under 'roles'."""
    if not labels:
        return {}
    cache = _load_cache(); cache.setdefault("roles", {})
    out, todo = {}, []
    for l in labels:
        if l in cache["roles"]:
            out[l] = cache["roles"][l]
        else:
            todo.append(l)
    if todo:
        schema = {
            "type": "object",
            "properties": {"items": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {
                        "label": {"type": "string"},
                        "role": {"type": "string", "enum": ["item", "total", "section"]},
                    },
                    "required": ["label", "role"],
                    "additionalProperties": False,
                },
            }},
            "required": ["items"], "additionalProperties": False,
        }

        def ask(labels):
            prompt = (
                "For each property income-statement line label, give its ROLE:\n"
                "- 'total' for any total, subtotal, or net line (e.g. 'Total Income', "
                "'Total Expenses', 'Total Operating Expenses', 'Net Operating Income', "
                "'NOI', 'Gross Income', 'Net Income')\n"
                "- 'section' for a bare section header with no amount (e.g. 'Income', "
                "'Expenses', 'Operating Expenses')\n"
                "- 'item' for a normal line item\n"
                "Echo every label EXACTLY as given, one entry per label, no additions.\n\n"
                f"Labels:\n{json.dumps(labels, indent=1)}"
            )
            resp = _client().messages.create(
                model=MODEL, max_tokens=8000, thinking={"type": "adaptive"},
                messages=[{"role": "user", "content": prompt}],
                output_config={"format": {"type": "json_schema", "schema": schema}},
            )
            return _json_response(resp)["items"]

        pending = list(todo)
        for _attempt in range(2):
            pset = set(pending)
            for m in ask(pending):
                if m["label"] in pset:
                    out[m["label"]] = m["role"]
                    cache["roles"][m["label"]] = m["role"]
            pending = [l for l in pending if l not in out]
            if not pending:
                break
        _save_cache(cache)
    return out
