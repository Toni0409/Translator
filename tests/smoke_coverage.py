"""Coverage backfill smoke checks — guard against silent "dịch sót".

Mô phỏng đúng lỗi quan sát được: 1 chunk lớn bị model trả response RỖNG/HỎNG
(parse ra []), nhưng sub-chunk nhỏ thì dịch OK. translate_chunk_with_retry phải
tự dịch bù để KHÔNG đoạn nào bị giữ nguyên tiếng nguồn.

Chạy: python tests/smoke_coverage.py
"""
from __future__ import annotations

import json
import os
import re
import sys

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from tests._helpers import setup_env, Report

cleanup = setup_env()
try:
    import word_backend as wb
    from word_backend import translate_chunk_with_retry, _english_residue

    # setup_env() stubs gemini.parse_json_loose → None; restore a real parser so
    # translate_chunk can actually read the fake model JSON.
    wb.parse_json_loose = lambda raw: json.loads(raw)

    report = Report("smoke_coverage")

    def _ids_from_prompt(prompt: str) -> list[str]:
        payload = prompt[prompt.rfind("Blocks:") + len("Blocks:"):].strip()
        try:
            return [x["id"] for x in json.loads(payload)]
        except Exception:
            return re.findall(r'"id":\s*"([^"]+)"', payload)

    def make_fake(big_threshold: int, calls: list, always_empty: bool = False):
        """Fake call_gemini: chunk >= big_threshold block → trả '[]' (mô phỏng
        response rỗng/hỏng); sub-chunk nhỏ → dịch sang tiếng Việt thật."""
        def fake(client, prompt):
            ids = _ids_from_prompt(prompt)
            calls.append(len(ids))
            if always_empty or len(ids) >= big_threshold:
                return "[]", 5, 5
            out = [{"id": i, "text": f"Đây là bản dịch tiếng Việt cho đoạn {i}."}
                   for i in ids]
            return json.dumps(out, ensure_ascii=False), 5, 5
        return fake

    EN = ("Release the stop switch and move the elevator down during inspection. "
          "Check the dynamic function before stepping onto the car top.")
    chunk40 = [{"id": f"B{i}", "text": EN, "role": "table_cell",
                "table_cell": (19, 1, 2)} for i in range(40)]

    # ── T1: big chunk trả rỗng → coverage backfill cứu toàn bộ ──────────
    calls = []
    wb.call_gemini = make_fake(big_threshold=20, calls=calls)
    result, ti, to, err = translate_chunk_with_retry(
        None, chunk40, "Vietnamese", "ctx", source_lang="English")
    translated = [k for k, v in result.items() if v.startswith("Đây là bản dịch")]
    residue = [k for k, v in result.items() if _english_residue(v)]
    report.check("all 40 blocks present in result", len(result) == 40,
                 f"got {len(result)}")
    report.check("all 40 blocks translated via sub-chunk backfill",
                 len(translated) == 40, f"translated={len(translated)}")
    report.check("0 blocks left with English residue (was ~23 before fix)",
                 len(residue) == 0, f"residue={len(residue)}")
    report.check("no fatal error reported (fully recovered)", err is None,
                 f"err={err}")
    report.check("did fall back to small sub-chunks",
                 any(n <= 6 for n in calls), f"call sizes={calls[:12]}")

    # ── T2: happy path — chunk dịch hết ngay vòng 1, KHÔNG dịch bù thừa ──
    calls2 = []
    wb.call_gemini = make_fake(big_threshold=999, calls=calls2)  # never "fails"
    small = [{"id": f"S{i}", "text": EN, "role": "paragraph"} for i in range(5)]
    result2, _, _, err2 = translate_chunk_with_retry(
        None, small, "Vietnamese", "ctx", source_lang="English")
    report.check("happy path: all translated", len(result2) == 5)
    report.check("happy path: exactly ONE API call (no wasted backfill)",
                 len(calls2) == 1, f"calls={calls2}")
    report.check("happy path: no residue", err2 is None)

    # ── T3: last resort — mọi call đều rỗng → giữ gốc + báo lỗi ─────────
    calls3 = []
    wb.call_gemini = make_fake(big_threshold=20, calls=calls3, always_empty=True)
    result3, _, _, err3 = translate_chunk_with_retry(
        None, chunk40, "Vietnamese", "ctx", source_lang="English")
    report.check("last resort: all blocks still present (original kept)",
                 len(result3) == 40)
    report.check("last resort: originals preserved (no data loss)",
                 all(result3[b["id"]] == b["text"] for b in chunk40))
    report.check("last resort: error surfaced for logging", err3 is not None,
                 f"err={err3}")

finally:
    cleanup()

sys.exit(report.summary())
