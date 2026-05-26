"""Domain glossary cho tài liệu kỹ thuật thang máy / thang cuốn.

Hai phần:
- `STANDARDS_KEEP_AS_IS`: tên tiêu chuẩn / part-number giữ nguyên không dịch.
- Seed JSON (`data/glossary_<name>.json`): từ điển 2 chiều en↔vi.

Flow tích hợp (xem P4.4–P4.7 trong DEV_PLAN):
1. `detect_subdomain(blocks)` → tập subdomain {"elevator", "escalator"}.
2. `seed_for_direction(subdomains, source_lang, target_lang)` → dict {term: tr}.
3. Frontend gọi `build_glossary(... seed=seed)` để merge seed vào trước AI-extract.
"""
from __future__ import annotations

import json
import os
import re
from functools import lru_cache


STANDARDS_KEEP_AS_IS = {
    "EN 81-20", "EN 81-50", "EN 81-1", "EN 81-2",
    "ISO 22201", "ISO 14798", "ISO 25745",
    "ASME A17.1", "ASME A17.3",
    "GB 7588", "GB 7588-2003", "GB 16899",
    "TCVN 6395", "TCVN 6396", "TCVN 6397", "TCVN 7628",
    "JIS A 4302",
}


_DATA_DIR = os.path.join(os.path.dirname(__file__), "data")


@lru_cache(maxsize=8)
def load_seed(name: str) -> dict:
    """Đọc `data/glossary_<name>.json` → dict {"en_vi": {...}, "vi_en": {...}}.

    Trả `{"en_vi": {}, "vi_en": {}}` nếu file thiếu hoặc parse fail.
    """
    path = os.path.join(_DATA_DIR, f"glossary_{name}.json")
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        if not isinstance(data, dict):
            return {"en_vi": {}, "vi_en": {}}
        return {
            "en_vi": dict(data.get("en_vi") or {}),
            "vi_en": dict(data.get("vi_en") or {}),
        }
    except Exception:
        return {"en_vi": {}, "vi_en": {}}


# Keyword detect — case-insensitive. Vietnamese variants có dấu để khớp chính xác.
_ELEVATOR_KEYWORDS = (
    "elevator", "lift", "hoistway", "cabin", "counterweight",
    "machine room", "machine-room-less", "traction machine", "sheave",
    "thang máy", "thang may", "giếng thang", "gieng thang",
    "đối trọng", "doi trong", "phòng máy", "phong may",
    "máy kéo", "may keo",
)
_ESCALATOR_KEYWORDS = (
    "escalator", "moving walkway", "handrail", "comb plate", "step chain",
    "balustrade", "newel", "skirt panel",
    "thang cuốn", "thang cuon", "tay vịn", "tay vin",
    "tấm lược", "tam luoc", "bậc thang", "bac thang",
    "xích bậc", "xich bac",
)

# Build regex per subdomain — word boundaries để tránh false positive
# (e.g. "elevation" không khớp "elevator").
def _compile_kw(words: tuple[str, ...]) -> re.Pattern:
    parts = sorted({re.escape(w) for w in words}, key=len, reverse=True)
    return re.compile(r"(?<![\w])(?:" + "|".join(parts) + r")(?![\w])",
                      re.IGNORECASE)


_ELEV_RE = _compile_kw(_ELEVATOR_KEYWORDS)
_ESC_RE  = _compile_kw(_ESCALATOR_KEYWORDS)


def detect_subdomain(blocks: list[dict]) -> set[str]:
    """Trả set subdomain trong {"elevator", "escalator"}.

    Heuristic: scan tối đa 200 blocks đầu, đếm số block hit từng subdomain.
    Yêu cầu ≥ 2 hit để counted (giảm false positive với tài liệu chỉ nhắc thoáng).
    """
    elev_hits = esc_hits = 0
    for b in blocks[:200]:
        text = b.get("text", "")
        if not text:
            continue
        if _ELEV_RE.search(text):
            elev_hits += 1
        if _ESC_RE.search(text):
            esc_hits += 1
    found = set()
    if elev_hits >= 2:
        found.add("elevator")
    if esc_hits >= 2:
        found.add("escalator")
    return found


def seed_for_direction(subdomains: set[str],
                       source_lang: str,
                       target_lang: str) -> dict:
    """Trả dict {term: translation} cho hướng dịch.

    - English → Vietnamese: dùng `en_vi`.
    - Vietnamese → English: dùng `vi_en`.
    - Hướng khác: trả {} (P0/P1 chỉ support 2 hướng này nhưng vẫn defensive).
    """
    if not subdomains:
        return {}
    direction_key = None
    if source_lang == "English" and target_lang == "Vietnamese":
        direction_key = "en_vi"
    elif source_lang == "Vietnamese" and target_lang == "English":
        direction_key = "vi_en"
    if direction_key is None:
        return {}
    merged: dict = {}
    # Sort subdomains để ổn định (escalator sau elevator nếu có cả 2)
    for name in sorted(subdomains):
        seed = load_seed(name)
        for k, v in (seed.get(direction_key) or {}).items():
            if k and v and k not in merged:
                merged[k] = v
    return merged
