"""Hằng số dùng chung cho toàn app."""
import streamlit as st

# ── Secrets ───────────────────────────────────────────────────────────────────
API_KEY      = st.secrets.get("GEMINI_API_KEY", "")
APP_PASSWORD = st.secrets.get("APP_PASSWORD", "")

# ── PDF ───────────────────────────────────────────────────────────────────────
PDF_MODEL    = "gemini-2.5-flash"
PRICE_INPUT  = 0.10       # USD / 1M input tokens (sẽ ×10 vì model tier)
PRICE_OUTPUT = 0.40       # USD / 1M output tokens
USD_TO_VND   = 25_400
PDF_DELAY    = 0.3        # delay giữa các page để né rate limit

MAX_RETRIES  = 5
RETRY_CODES  = ("429", "resource_exhausted", "quota", "rate")

UNICODE_FONTS = [
    "Carlito-Regular.ttf",
    "/usr/share/fonts/truetype/crosextra/Carlito-Regular.ttf",
    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    "C:/Windows/Fonts/Arial.ttf",
    "C:/Windows/Fonts/calibri.ttf",
]

BOLD_FONT_PAIRS = [
    ("Carlito-Regular.ttf", "Carlito-Bold.ttf"),
    ("Arial.ttf",           "Arialbd.ttf"),
    ("calibri.ttf",         "calibrib.ttf"),
    ("times.ttf",           "timesbd.ttf"),
    ("DejaVuSans.ttf",      "DejaVuSans-Bold.ttf"),
]

# ── Word ──────────────────────────────────────────────────────────────────────
WORD_MODELS = [
    "gemini-2.5-flash-lite",
    "gemini-2.5-flash",
    "gemini-2.0-flash",
    "gemini-2.0-flash-lite",
]
CHUNK_SIZE         = 25
MAX_WORD_TOKENS    = 65_536
NO_TRANSLATE_ROLES = {"header", "footer"}

# ── Ngôn ngữ ──────────────────────────────────────────────────────────────────
LANGUAGES = ["Tiếng Việt", "Tiếng Anh", "Tiếng Nhật", "Tiếng Trung", "Tiếng Pháp", "Tiếng Đức"]

LANG_EN = {
    "Tiếng Việt":  "Vietnamese",
    "Tiếng Anh":   "English",
    "Tiếng Nhật":  "Japanese",
    "Tiếng Trung": "Chinese",
    "Tiếng Pháp":  "French",
    "Tiếng Đức":   "German",
}
