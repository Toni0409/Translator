"""Hằng số dùng chung cho toàn app."""
import streamlit as st

# ── Secrets ───────────────────────────────────────────────────────────────────
API_KEY      = st.secrets.get("GEMINI_API_KEY", "")
APP_PASSWORD = st.secrets.get("APP_PASSWORD", "")

# ── Feature flags ─────────────────────────────────────────────────────────────
# Tắt tab "So sánh / Đánh giá" — đang ngủ, chưa dùng. Bật lại bằng True.
REVIEW_TAB_ENABLED = False
# Tắt tab "Dịch PDF" — chỉ giữ Dịch Word. Bật lại bằng True.
PDF_TAB_ENABLED = False

# ── PDF ───────────────────────────────────────────────────────────────────────
PDF_MODEL    = "gemini-3.5-flash"
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
    "gemini-3.5-flash",
    "gemini-2.5-flash-lite",
    "gemini-2.5-flash",
    "gemini-2.0-flash",
    "gemini-2.0-flash-lite",
]
MAX_WORD_TOKENS    = 8_192     # output limit thực tế của Gemini Flash family
                               # (65k cũ gây JSON cut khi chunk lớn)
MAX_WORD_WORKERS   = 4         # số chunk dịch song song
CHUNK_RETRIES      = 3         # số lần retry mỗi chunk khi lỗi
TARGET_CHUNK_CHARS = 8_000     # mục tiêu ký tự / chunk (adaptive)
MIN_CHUNK_BLOCKS   = 8         # tối thiểu paragraph / chunk
MAX_CHUNK_BLOCKS   = 40        # tối đa paragraph / chunk

# H/F detection: text lặp lại ≥ HF_REPEAT_THRESHOLD lần (và dài ≥ HF_REPEAT_MIN_CHARS)
# → đánh dấu là body_repeated (mặc định không dịch — bấm nút mới dịch)
HF_REPEAT_THRESHOLD = 3
HF_REPEAT_MIN_CHARS = 10

# Roles KHÔNG dịch mặc định ở lần đầu — user phải bấm nút riêng để dịch
NO_TRANSLATE_ROLES = {"header", "footer", "body_repeated"}

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
