"""QCP Tools — combined Streamlit app.

The sidebar switches between the tools:
  • Underwriter       — drop a deal's documents, get one linked workbook
  • Public UW         — strip a finished workbook down to the tabs that go out
  • Rent Roll Parser  — a rent roll (PDF/Excel) to a sourced worksheet
  • Tax Bill Parser   — LA County secured property tax bills to Excel
  • Historicals       — combine operating statements across years
  • Lender Importer   — bulk-create Salesforce Deal Source records

The tool list is drag-to-reorder and the arrangement is remembered per browser.

Run locally:  streamlit run app.py
"""

import json
import os

import streamlit as st

# Must be the first Streamlit call.
st.set_page_config(page_title="QCP Tools", page_icon="🏢", layout="wide",
                   initial_sidebar_state="expanded")

from tools import (tax_parser, lender_importer, historicals, rent_roll,  # noqa: E402
                   underwrite, public_uw)

# id -> (sidebar label, render, caption). The id is what gets saved, so labels
# can be reworded without stranding somebody's saved order.
# NB: the gap after each emoji is a non-breaking space — Streamlit collapses
# ordinary runs of whitespace in a widget label and the icon ends up glued on.
TOOLS = {
    "underwriter": ("📐   Underwriter", underwrite.render,
                    "Deal documents → one workbook"),
    "public_uw": ("📤   Public UW", public_uw.render,
                  "Finished workbook → public tabs"),
    "rent_roll": ("🏘️   Rent Roll Parser", rent_roll.render,
                  "Rent roll → sourced worksheet"),
    "tax_bill": ("🏠   Tax Bill Parser", tax_parser.render,
                 "Tax bills → rates & assessments"),
    "historicals": ("📊   Historicals", historicals.render,
                    "Statements → combined tab"),
    "lender": ("🏦   Lender Importer", lender_importer.render,
               "Lender list → Salesforce"),
}
DEFAULT_ORDER = list(TOOLS)
ORDER_KEY = "qcp_tool_order"

_CSS = """
<style>
  /* tighten the default top gutter so page headers sit higher */
  .block-container { padding-top: 2.6rem; padding-bottom: 4rem; max-width: 1500px; }

  /* sidebar: quieter chrome, roomier tool list */
  [data-testid="stSidebar"] { border-right: 1px solid rgba(49,51,63,.12); }
  [data-testid="stSidebar"] .block-container { padding-top: 1.6rem; }
  [data-testid="stSidebar"] [role="radiogroup"] > label { padding: .18rem 0; }
  [data-testid="stSidebar"] hr { margin: .9rem 0; }
  [data-testid="stSidebar"] [data-testid="stExpander"] summary { font-size: .82rem; }

  /* headers */
  h1, h2, h3 { letter-spacing: -.01em; }
  h2 { padding-top: .4rem; }

  /* numbers should read as numbers */
  [data-testid="stMetricValue"] { font-size: 1.7rem; font-variant-numeric: tabular-nums; }
  [data-testid="stMetricLabel"] { opacity: .75; }
  [data-testid="stDataFrame"] { font-variant-numeric: tabular-nums; }

  /* controls */
  .stButton > button, .stDownloadButton > button { border-radius: .5rem; font-weight: 600; }
  [data-testid="stFileUploaderDropzone"] { border-radius: .6rem; }

  /* The drag-to-reorder list is a component iframe, and a component first
     rendered inside a collapsed expander measures itself at zero height and
     never re-measures once shown. Give this one — and only this one, the
     localStorage handle is an iframe too — the height its contents need. */
  iframe[title="streamlit_sortables.sortable_items"],
  div:has(> iframe[title="streamlit_sortables.sortable_items"]) {
    height: SORT_Hpx !important; min-height: SORT_Hpx !important;
  }
</style>
""".replace("SORT_H", str(len(TOOLS) * 45 + 14))

# drag handles for the arrange panel
_SORT_CSS = """
.sortable-component { background: transparent; font-size: .86rem; }
.sortable-item {
  background: #fff; color: #1f2333; border: 1px solid rgba(49,51,63,.15);
  border-radius: .45rem; padding: .34rem .6rem; margin: .18rem 0; cursor: grab;
}
.sortable-item:hover { border-color: #4f46e5; }
"""


# ── saved order (per browser, via localStorage) ─────────────────────────────
def _store():
    """localStorage handle, or None. Its own component key so it can't collide
    with the Lender Importer's 'remember me' storage."""
    if os.getenv("QCP_NO_LOCALSTORAGE"):        # headless tests have no browser
        return None
    try:
        from streamlit_local_storage import LocalStorage
        return LocalStorage(key="qcp_order_store")
    except Exception:  # noqa: BLE001
        return None


def _clean(order) -> list:
    """A saved order, made safe: known ids only, and any tool added since it was
    saved appended in its default position rather than disappearing."""
    seen = [t for t in (order or []) if t in TOOLS]
    return list(dict.fromkeys(seen)) + [t for t in DEFAULT_ORDER if t not in seen]


def _load_order() -> list:
    """Read once per session — after that the session copy is the truth."""
    if "tool_order" not in st.session_state:
        saved = None
        try:
            ls = _store()
            raw = ls.getItem(ORDER_KEY) if ls else None
            saved = json.loads(raw) if raw else None
        except Exception:  # noqa: BLE001
            saved = None
        st.session_state["tool_order"] = _clean(saved)
    return st.session_state["tool_order"]


def _save_order(order):
    st.session_state["tool_order"] = order
    try:
        ls = _store()
        if ls:
            ls.setItem(ORDER_KEY, json.dumps(order), key="qcp_order_set")
    except Exception:  # noqa: BLE001
        pass                                     # order still holds for this session


st.markdown(_CSS, unsafe_allow_html=True)
order = _load_order()

with st.sidebar:
    st.markdown("### 🏢 QCP Tools")
    st.caption("Underwriting workbooks, end to end.")
    st.divider()

    # The tool list is drawn into this slot *after* the arrange panel has run,
    # so a drag takes effect in the same pass. Re-running instead would abort
    # the localStorage write mid-flight and the order would never be saved.
    nav = st.container()

    with st.expander("⇅   Arrange"):
        labels = [TOOLS[t][0] for t in order]
        try:
            from streamlit_sortables import sort_items
            st.caption("Drag to reorder — your arrangement is remembered on this browser.")
            dragged = sort_items(labels, direction="vertical",
                                 custom_style=_SORT_CSS, key="arrange")
            new_order = [order[labels.index(l)] for l in dragged if l in labels]
        except Exception:  # noqa: BLE001 — component missing: pick the order instead
            st.caption("Click the tools in the order you want them.")
            chosen = st.multiselect("Order", labels, default=labels,
                                    label_visibility="collapsed", key="arrange_pick")
            new_order = [order[labels.index(l)] for l in chosen if l in labels]
        if order != DEFAULT_ORDER and st.button("Reset to default",
                                                use_container_width=True):
            new_order = list(DEFAULT_ORDER)
        new_order = _clean(new_order)
        if new_order != order:
            _save_order(new_order)
            order = new_order

    with nav:
        labels = [TOOLS[t][0] for t in order]
        current = st.session_state.get("tool_current", order[0])
        index = order.index(current) if current in order else 0
        picked_label = st.radio("Tool", labels, index=index, label_visibility="collapsed",
                                captions=[TOOLS[t][2] for t in order])
        picked = order[labels.index(picked_label)]
        st.session_state["tool_current"] = picked

    st.divider()
    try:
        from tools import hist_llm
        on = hist_llm.available()
        st.caption(("🟢   AI on · " + hist_llm.MODEL) if on else
                   "⚪   AI off · set ANTHROPIC_API_KEY in the app secrets")
    except Exception:  # noqa: BLE001
        pass

TOOLS[picked][1]()
