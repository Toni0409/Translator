"""
Translator App — Word [Streamlit]
Entry point: set page config, inject CSS, kiểm tra password, render tab Word.

Mọi logic backend & UI chi tiết nằm trong các module riêng:
- config.py        : hằng số (API key, model, ngôn ngữ, ...)
- styles.py        : CSS dark theme
- auth.py          : password gate + logout
- gemini.py        : Gemini client + JSON parser + thread helper
- ui_common.py     : timer box, log adder, stat box
- word_backend.py  : extract / translate / render DOCX
- word_tab.py      : UI tab Word

Các module đang ngủ (bật lại qua feature flag trong config.py):
- pdf_backend.py / pdf_tab.py   : PDF_TAB_ENABLED
- review_*.py                   : REVIEW_TAB_ENABLED
"""
import streamlit as st

import config
import styles
import auth
import word_tab

if config.PDF_TAB_ENABLED:
    import pdf_tab  # noqa: F401  (đang ngủ — bật qua config.PDF_TAB_ENABLED)

if config.REVIEW_TAB_ENABLED:
    import review_tab  # noqa: F401  (đang ngủ — bật qua config.REVIEW_TAB_ENABLED)


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

# ── Tabs ──────────────────────────────────────────────────────────────────────
tab_labels = []
tab_renderers = []

if config.PDF_TAB_ENABLED:
    tab_labels.append("📄 Dịch PDF")
    tab_renderers.append(pdf_tab.render)

tab_labels.append("📝 Dịch Word")
tab_renderers.append(word_tab.render)

if config.REVIEW_TAB_ENABLED:
    tab_labels.append("🔍 So sánh / Đánh giá")
    tab_renderers.append(review_tab.render)

if len(tab_labels) == 1:
    # Chỉ còn 1 tab → render trực tiếp, không cần st.tabs
    tab_renderers[0]()
else:
    for tab, render in zip(st.tabs(tab_labels), tab_renderers):
        with tab:
            render()
