"""
Translator App — PDF + Word [Streamlit]
Entry point: set page config, inject CSS, kiểm tra password, render 2 tab.

Mọi logic backend & UI chi tiết nằm trong các module riêng:
- config.py        : hằng số (API key, model, ngôn ngữ, ...)
- styles.py        : CSS dark theme
- auth.py          : password gate + logout
- gemini.py        : Gemini client + JSON parser + thread helper
- ui_common.py     : timer box, log adder, stat box
- pdf_backend.py   : extract / translate / render PDF
- pdf_tab.py       : UI tab PDF
- word_backend.py  : extract / translate / render DOCX
- word_tab.py      : UI tab Word
"""
import streamlit as st

import styles
import auth
import pdf_tab
import word_tab


st.set_page_config(
    page_title="Translator — Vi Nguyen",
    page_icon="⬡",
    layout="centered",
)
styles.inject()

if not auth.check_password():
    st.stop()

# ── Header ────────────────────────────────────────────────────────────────────
st.markdown("## ⬡ Translator")
st.markdown(
    "<span style='color:#64748b;font-size:0.9rem'>"
    "PDF & Word — Powered by Gemini — Vi Nguyen</span>",
    unsafe_allow_html=True,
)
_, col_lo = st.columns([6, 1])
with col_lo:
    auth.logout_button()
st.divider()

# ── Tabs ──────────────────────────────────────────────────────────────────────
tab_pdf, tab_word = st.tabs(["📄 Dịch PDF", "📝 Dịch Word"])
with tab_pdf:
    pdf_tab.render()
with tab_word:
    word_tab.render()
