"""P5 (extra): Smoke check direction resolver + ZIP name sanitizer.

Chạy: python tests/smoke_direction.py
"""
from __future__ import annotations

import os
import sys

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from tests._helpers import setup_env, Report


cleanup = setup_env()
try:
    from word_tab import _resolve_langs, _safe_zip_name
    from word_backend import _build_chunk_prompt

    report = Report("smoke_direction")

    # ── _resolve_langs ──────────────────────────────────────────────────
    lbl, src, tgt = _resolve_langs("Anh → Việt")
    report.check("Anh→Việt resolves", src == "English" and tgt == "Vietnamese",
                 f"{lbl} | {src} → {tgt}")
    lbl, src, tgt = _resolve_langs("Việt → Anh")
    report.check("Việt→Anh resolves", src == "Vietnamese" and tgt == "English",
                 f"{lbl} | {src} → {tgt}")
    lbl, src, tgt = _resolve_langs("Unknown label")
    report.check("Unknown label fallback to first direction",
                 src == "English" and tgt == "Vietnamese")

    # ── _build_chunk_prompt header ─────────────────────────────────────
    chunk = [{"id": "p1", "text": "Hello", "role": "paragraph"}]
    p_with = _build_chunk_prompt(chunk, "Vietnamese", "Doc ctx",
                                 source_lang="English")
    p_without = _build_chunk_prompt(chunk, "Vietnamese", "Doc ctx")
    report.check("prompt uses 'from English into Vietnamese' when source_lang",
                 "from English into Vietnamese" in p_with)
    report.check("prompt falls back to 'into Vietnamese' without source_lang",
                 "Translate these Word document blocks into Vietnamese" in p_without)

    # ── _safe_zip_name ─────────────────────────────────────────────────
    used: set[str] = set()
    safe1 = _safe_zip_name("hello.docx", "Tiếng Việt", used)
    used.add(safe1)
    safe_traverse = _safe_zip_name("../../etc/passwd", "Tiếng Anh", used)
    safe_path = _safe_zip_name("a/b/c.docx", "Tiếng Việt", used)
    safe_dup  = _safe_zip_name("hello.docx", "Tiếng Việt", used)

    report.check("plain name sanitized + suffix",
                 safe1.endswith(".docx") and "/" not in safe1 and ".." not in safe1,
                 f"got {safe1!r}")
    report.check("path traversal removed",
                 ".." not in safe_traverse and "/" not in safe_traverse,
                 f"got {safe_traverse!r}")
    report.check("subdir separator removed",
                 "/" not in safe_path,
                 f"got {safe_path!r}")
    report.check("duplicate gets unique suffix",
                 safe_dup != safe1,
                 f"first={safe1!r} dup={safe_dup!r}")

finally:
    cleanup()


sys.exit(report.summary())
