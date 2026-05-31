"""Shared Gemini client + low-level helpers (generate, token usage, JSON parsing)."""
import json

import streamlit as st
from google import genai
from google.genai import types as gtypes

from config import API_KEY


@st.cache_resource
def get_client():
    """Single Gemini client (cached) cho Word tab."""
    return genai.Client(api_key=API_KEY)


def generate(client, model: str, prompt: str,
             max_output_tokens: int = 65_536, temperature: float = 0.1):
    """Wrapper gọn cho `client.models.generate_content`."""
    return client.models.generate_content(
        model=model,
        contents=prompt,
        config=gtypes.GenerateContentConfig(
            max_output_tokens=max_output_tokens,
            temperature=temperature,
        ),
    )


def usage_tokens(resp) -> tuple[int, int]:
    """Trích (in_tokens, out_tokens) từ response."""
    meta  = getattr(resp, "usage_metadata", None)
    in_t  = getattr(meta, "prompt_token_count",     0) or 0
    out_t = getattr(meta, "candidates_token_count", 0) or 0
    return in_t, out_t


def parse_json_loose(raw: str):
    """
    Parse JSON tolerant với markdown fence + JSON bị cắt giữa chừng.
    Hỗ trợ cả dict và list ở root.
    """
    raw = (raw or "").strip()
    if raw.startswith("```"):
        raw = raw.split("```")[1]
        if raw.startswith("json"):
            raw = raw[4:]
    raw = raw.strip()

    try:
        return json.loads(raw)
    except Exception:
        pass

    # Thử cắt phần đuôi bị truncate cho dict
    for end_token in ("},\n", "}, \n", "},"):
        last = raw.rfind(end_token)
        if last > 0:
            try:
                return json.loads(raw[:last + 1] + "]")
            except Exception:
                pass

    # Trích array trong nhiễu
    s, e = raw.find("["), raw.rfind("]")
    if s != -1 and e > s:
        candidate = raw[s:e + 1]
        try:
            return json.loads(candidate)
        except Exception:
            last = candidate.rfind("},")
            if last == -1:
                last = candidate.rfind("}")
            if last > 0:
                try:
                    return json.loads(candidate[:last + 1].rstrip(",") + "]")
                except Exception:
                    pass
    return None
