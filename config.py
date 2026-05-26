"""Hằng số dùng chung cho toàn app (Word)."""
import streamlit as st

# ── Secrets ───────────────────────────────────────────────────────────────────
API_KEY      = st.secrets.get("GEMINI_API_KEY", "")
APP_PASSWORD = st.secrets.get("APP_PASSWORD", "")

# ── Pricing (Gemini 3.5 Flash tier) ───────────────────────────────────────────
PRICE_INPUT  = 0.10       # USD / 1M input tokens (sẽ ×10 vì model tier)
PRICE_OUTPUT = 0.40       # USD / 1M output tokens
USD_TO_VND   = 25_400

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
# App tập trung 2 hướng dịch: Anh ↔ Việt cho tài liệu kỹ thuật thang máy/thang cuốn.
LANGUAGES = ["Tiếng Anh", "Tiếng Việt"]

LANG_EN = {
    "Tiếng Anh":  "English",
    "Tiếng Việt": "Vietnamese",
}

# (label, source_lang, target_lang) — dùng cho radio "Hướng dịch"
TRANSLATION_DIRECTIONS = [
    ("Anh → Việt", "English",    "Vietnamese"),
    ("Việt → Anh", "Vietnamese", "English"),
]
