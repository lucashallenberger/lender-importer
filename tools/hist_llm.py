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

MODEL = "claude-opus-4-8"

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


def _json_response(resp):
    return json.loads(next(b.text for b in resp.content if b.type == "text"))


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


def extract_statement(pdf_bytes):
    """Claude-side extraction -> rows shaped like parse_summary's output."""
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
        messages=[{
            "role": "user",
            "content": [
                {"type": "document",
                 "source": {"type": "base64", "media_type": "application/pdf",
                            "data": base64.standard_b64encode(pdf_bytes).decode()}},
                {"type": "text", "text": _EXTRACT_PROMPT},
            ],
        }],
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
