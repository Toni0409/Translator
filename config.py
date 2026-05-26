"""Hằng số dùng chung cho toàn app (Word)."""
import streamlit as st


# ── Secrets (safe getter — P3.4) ──────────────────────────────────────────────
def _safe_secret(key: str, default: str = "") -> str:
    """Đọc Streamlit secret an toàn khi `.streamlit/secrets.toml` không tồn tại
    (CLI smoke test, unit test). App vẫn báo lỗi rõ khi password/API thiếu lúc
    runtime — chỉ tránh crash ở import time.
    """
    try:
        return st.secrets.get(key, default)
    except Exception:
        return default


API_KEY      = _safe_secret("GEMINI_API_KEY", "")
APP_PASSWORD = _safe_secret("APP_PASSWORD", "")

# ── Pricing (Gemini 3.5 Flash tier) ───────────────────────────────────────────
# Nguồn: Google AI pricing page, kiểm tra 2026-05-26.
# `gemini-3.5-flash` Standard Paid: $1.50 input / $9.00 output per 1M tokens.
# Đây là cost source-of-truth duy nhất — `calc_cost()` trong ui_common dùng giá này.
# Nếu Google đổi giá → cập nhật cả comment ngày + 2 giá trị bên dưới.
PRICE_INPUT  = 1.50       # USD / 1M input tokens (gemini-3.5-flash)
PRICE_OUTPUT = 9.00       # USD / 1M output tokens (gemini-3.5-flash)
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

# Roles KHÔNG dịch mặc định ở lần đầu — user phải bấm nút riêng để dịch.
# `media_only` = paragraph chỉ chứa ảnh/drawing/object, không có text →
# giữ nguyên XML, không gửi AI, không thay runs (P2).
NO_TRANSLATE_ROLES = {"header", "footer", "body_repeated", "media_only"}

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
