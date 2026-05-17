"""CSS injection cho Streamlit UI — dark theme."""
import streamlit as st

_CSS = """
<style>
    .stApp { background-color: #0f1117; color: #e2e8f0; }
    .block-container { max-width: 820px; padding-top: 2rem; }
    h1, h2, h3 { color: #e2e8f0 !important; }
    p, li, span { color: #e2e8f0; }

    div[data-testid="stFileUploader"] {
        border: 1.5px dashed #4a5080 !important; border-radius: 10px !important;
        padding: 8px !important; background: #1a1d27 !important;
    }
    div[data-testid="stFileUploader"] label { color: #e2e8f0 !important; }
    [data-testid="stFileUploaderDropzone"] { background-color: #242838 !important; border-color: #4a5080 !important; }
    [data-testid="stFileUploaderDropzone"] > div { background-color: #242838 !important; }
    [data-testid="stFileUploaderDropzone"] span { color: #e2e8f0 !important; }
    [data-testid="stFileUploaderDropzone"] button { background-color: #2d3149 !important; color: #e2e8f0 !important; border-color: #4a5080 !important; }
    [data-testid="stFileUploaderDropzone"] p { color: #94a3b8 !important; }

    label, .stSelectbox label, .stTextInput label,
    div[data-testid="stWidgetLabel"] p { color: #e2e8f0 !important; font-weight: 500; }
    .stSelectbox div[data-baseweb="select"] > div { background: #1a1d27 !important; color: #e2e8f0 !important; border-color: #4a5080 !important; }
    .stTextInput input { background: #1a1d27 !important; color: #e2e8f0 !important; border-color: #4a5080 !important; }
    input::placeholder { color: #64748b !important; }

    .stat-box { background: #1a1d27; border-radius: 10px; padding: 14px 18px; text-align: center; border: 1px solid #4a5080; }
    .stat-val { font-size: 1.4rem; font-weight: bold; color: #e2e8f0; }
    .stat-lbl { font-size: 0.78rem; color: #94a3b8; margin-top: 4px; }

    .timer-box { background: #1a1d27; border-radius: 10px; padding: 12px 18px; text-align: center; border: 1px solid #6c63ff; margin-bottom: 10px; }
    .timer-val { font-size: 2rem; font-weight: bold; color: #a78bfa; font-family: 'Courier New', monospace; }
    .timer-lbl { font-size: 0.78rem; color: #94a3b8; margin-top: 2px; }
    .timer-status { font-size: 0.85rem; color: #818cf8; margin-top: 6px; }

    .log-box { background: #1a1d27; border-radius: 8px; padding: 14px 18px; font-family: 'Courier New', monospace; font-size: 0.83rem; max-height: 340px; overflow-y: auto; border: 1px solid #4a5080; white-space: pre-wrap; color: #c4cde0; line-height: 1.6; }

    .stButton > button { background-color: #6c63ff !important; color: white !important; border: none !important; border-radius: 8px !important; padding: 10px 28px !important; font-weight: bold !important; font-size: 1rem !important; width: 100% !important; }
    .stButton > button:hover { background-color: #a78bfa !important; }
    .stButton > button:disabled { background-color: #2d3149 !important; color: #64748b !important; }

    .stDownloadButton > button { background-color: #059669 !important; color: white !important; border-radius: 8px !important; font-weight: bold !important; font-size: 1rem !important; width: 100% !important; }
    .stDownloadButton > button:hover { background-color: #10b981 !important; }

    div[data-testid="stProgressBar"] > div { background-color: #1a1d27 !important; }
    div[data-testid="stProgressBar"] > div > div { background-color: #6c63ff !important; }

    hr { border-color: #2d3149 !important; }
    div[data-testid="stAlert"] { background: #1a1d27 !important; border-color: #4a5080 !important; color: #e2e8f0 !important; }

    .stTabs [data-baseweb="tab-list"] { background: #1a1d27; border-radius: 10px; padding: 4px; gap: 4px; }
    .stTabs [data-baseweb="tab"] { color: #94a3b8 !important; background: transparent; border-radius: 8px; font-weight: 600; font-size: 1rem; padding: 8px 24px; }
    .stTabs [aria-selected="true"] { color: #e2e8f0 !important; background: #6c63ff !important; }

    .login-box { background: #1a1d27; border: 1px solid #4a5080; border-radius: 14px; padding: 2.5rem 2rem; max-width: 380px; margin: 4rem auto 0 auto; text-align: center; }
    .login-title { font-size: 1.6rem; font-weight: bold; color: #e2e8f0; margin-bottom: 0.3rem; }
    .login-sub   { font-size: 0.85rem; color: #64748b; margin-bottom: 1.8rem; }
</style>
"""


def inject():
    """Gắn CSS dark theme vào trang."""
    st.markdown(_CSS, unsafe_allow_html=True)
