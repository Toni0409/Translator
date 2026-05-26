"""
Translator App — Word [Streamlit]
Entry point: set page config, inject CSS, kiểm tra password, render tab Word.

Module:
- config.py        : hằng số (API key, model, ngôn ngữ, ...)
- styles.py        : CSS dark theme
- auth.py          : password gate + logout
- gemini.py        : Gemini client + JSON parser + thread helper
- ui_common.py     : timer box, log adder, stat box
- word_backend.py  : extract / translate / render DOCX
- word_tab.py      : UI tab Word

Tính năng cũ (PDF, So sánh / Đánh giá) đã chuyển vào `archive/` — không build.
"""
import streamlit as st

import styles
import auth
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
    "Word — Powered by Gemini — Vi Nguyen</span>",
    unsafe_allow_html=True,
)
_, col_lo = st.columns([6, 1])
with col_lo:
    auth.logout_button()
st.divider()

# ── Word tab ──────────────────────────────────────────────────────────────────
word_tab.render()
