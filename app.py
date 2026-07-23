"""QCP Tools — combined Streamlit app.

A sidebar switches between two tools:
  • Tax Bill Parser  — parse LA County secured property tax bill PDFs to Excel
  • Lender Importer   — bulk-create Salesforce Deal Source records from a lender list

Run locally:  streamlit run app.py
On Replit:    configured via .replit (streamlit run app.py ...)
"""

import streamlit as st

# Must be the first Streamlit call.
st.set_page_config(page_title="QCP Tools", page_icon="🏢", layout="wide")

from tools import tax_parser, lender_importer, historicals, rent_roll, underwrite  # noqa: E402

PAGES = {
    "Underwriting (Beta)": underwrite.render,
    "Tax Bill Parser": tax_parser.render,
    "Rent Roll Parser": rent_roll.render,
    "Lender Importer": lender_importer.render,
    "Historicals": historicals.render,
}

with st.sidebar:
    st.markdown("## QCP Tools")
    choice = st.radio("Choose a tool", list(PAGES.keys()), label_visibility="collapsed")
    st.divider()

PAGES[choice]()
