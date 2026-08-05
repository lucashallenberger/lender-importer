"""QCP Tools — combined Streamlit app.

The sidebar switches between the tools:
  • Underwriter       — drop a deal's documents, get one linked workbook
  • Public UW         — strip a finished workbook down to the tabs that go out
  • Rent Roll Parser  — a rent roll (PDF/Excel) to a sourced worksheet
  • Tax Bill Parser   — LA County secured property tax bills to Excel
  • Historicals       — combine operating statements across years
  • Lender Importer   — bulk-create Salesforce Deal Source records

Run locally:  streamlit run app.py
"""

import streamlit as st

# Must be the first Streamlit call.
st.set_page_config(page_title="QCP Tools", page_icon="🏢", layout="wide",
                   initial_sidebar_state="expanded")

from tools import (tax_parser, lender_importer, historicals, rent_roll,  # noqa: E402
                   underwrite, public_uw)

# label -> (render, sidebar caption). Ordered by how a deal moves through them:
# build the workbook, publish it, then the single-purpose parsers.
# NB: the gap after each emoji is a non-breaking space — Streamlit collapses
# ordinary runs of whitespace in a widget label and the icon ends up glued on.
TOOLS = {
    "📐  Underwriter": (underwrite.render, "Deal documents → one workbook"),
    "📤  Public UW": (public_uw.render, "Finished workbook → public tabs"),
    "🏘️  Rent Roll Parser": (rent_roll.render, "Rent roll → sourced worksheet"),
    "🏠  Tax Bill Parser": (tax_parser.render, "Tax bills → rates & assessments"),
    "📊  Historicals": (historicals.render, "Statements → combined tab"),
    "🏦  Lender Importer": (lender_importer.render, "Lender list → Salesforce"),
}

_CSS = """
<style>
  /* tighten the default top gutter so page headers sit higher */
  .block-container { padding-top: 2.6rem; padding-bottom: 4rem; max-width: 1500px; }

  /* sidebar: quieter chrome, roomier tool list */
  [data-testid="stSidebar"] { border-right: 1px solid rgba(49,51,63,.12); }
  [data-testid="stSidebar"] .block-container { padding-top: 1.6rem; }
  [data-testid="stSidebar"] [role="radiogroup"] > label { padding: .18rem 0; }
  [data-testid="stSidebar"] hr { margin: .9rem 0; }

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
</style>
"""
st.markdown(_CSS, unsafe_allow_html=True)

with st.sidebar:
    st.markdown("### 🏢 QCP Tools")
    st.caption("Underwriting workbooks, end to end.")
    st.divider()
    choice = st.radio("Tool", list(TOOLS), label_visibility="collapsed",
                      captions=[c for _, c in TOOLS.values()])
    st.divider()
    try:
        from tools import hist_llm
        on = hist_llm.available()
        st.caption(("🟢  AI on · " + hist_llm.MODEL) if on else
                   "⚪  AI off · set ANTHROPIC_API_KEY in the app secrets")
    except Exception:  # noqa: BLE001
        pass

TOOLS[choice][0]()
