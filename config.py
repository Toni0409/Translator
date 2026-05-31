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
MAX_WORD_WORKERS   = 16        # số chunk dịch song song (ThreadPoolExecutor).
                               # Đây là ĐÒN BẨY TỐC ĐỘ chính: thời gian dịch ≈
                               # ⌈số_chunk / workers⌉ × latency_mỗi_chunk.
                               # Tăng workers KHÔNG đổi chất lượng — mỗi chunk dịch
                               # độc lập, cùng prompt + glossary + doc context.
                               # 16 an toàn cho Gemini paid Tier 1+ (~1000 RPM);
                               # retry/backoff đã xử lý 429 nếu lỡ chạm rate limit.
                               # → Gặp 429 liên tục thì hạ về 8; tier cao + doc
                               #   rất lớn có thể nâng 24–32.
CHUNK_RETRIES      = 3         # số lần retry mỗi chunk khi lỗi
# Dịch bù (coverage backfill): sau khi dịch 1 chunk, block nào model TRẢ VỀ THIẾU
# (JSON rỗng/hỏng/bỏ sót item) hoặc bản dịch CÒN tiếng nguồn sẽ được dịch lại theo
# sub-chunk nhỏ. Đây là lưới an toàn chống "dịch sót" — nguyên nhân chính khiến cả
# 1 chunk đôi khi giữ nguyên tiếng Anh khi model trả response lỗi/rỗng.
COVERAGE_RETRY_ROUNDS = 2      # số vòng dịch bù tối đa (0 = tắt)
RETRY_SUBCHUNK_BLOCKS = 6      # kích thước sub-chunk khi dịch bù (nhỏ → ít lỗi lặp lại)
TARGET_CHUNK_CHARS = 8_000     # mục tiêu ký tự / chunk (adaptive)
MIN_CHUNK_BLOCKS   = 8         # tối thiểu paragraph / chunk
MAX_CHUNK_BLOCKS   = 40        # tối đa paragraph / chunk
# H/F detection: text lặp lại ≥ HF_REPEAT_THRESHOLD lần (và dài ≥ HF_REPEAT_MIN_CHARS)
# → đánh dấu là body_repeated. Vẫn dịch (xem `NO_TRANSLATE_ROLES`).
HF_REPEAT_THRESHOLD = 3
HF_REPEAT_MIN_CHARS = 10

# Roles KHÔNG dịch mặc định.
# `media_only` = paragraph chỉ chứa ảnh/drawing/object, không có text →
# giữ nguyên XML, không gửi AI, không thay runs.
# (Header / footer / body_repeated giờ dịch mặc định — không tách nút riêng nữa.)
NO_TRANSLATE_ROLES = {"media_only"}

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
