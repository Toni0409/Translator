"""P5.5: Smoke check domain detection + seed glossary merge.

Chạy: python tests/smoke_domain.py
"""
from __future__ import annotations

import os
import sys

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from tests._helpers import setup_env, Report


cleanup = setup_env()
try:
    from domain_glossary import (
        STANDARDS_KEEP_AS_IS, load_seed, detect_subdomain, seed_for_direction,
    )
    from word_backend import build_glossary, build_doc_context

    report = Report("smoke_domain")

    # ── load_seed ────────────────────────────────────────────────────────
    elev = load_seed("elevator")
    esc  = load_seed("escalator")
    report.check("elevator seed has en_vi+vi_en",
                 bool(elev["en_vi"]) and bool(elev["vi_en"]),
                 f"{len(elev['en_vi'])} en_vi, {len(elev['vi_en'])} vi_en")
    report.check("escalator seed has en_vi+vi_en",
                 bool(esc["en_vi"]) and bool(esc["vi_en"]),
                 f"{len(esc['en_vi'])} en_vi, {len(esc['vi_en'])} vi_en")
    report.check("missing seed → empty",
                 load_seed("nonexistent") == {"en_vi": {}, "vi_en": {}})

    # ── detect_subdomain ────────────────────────────────────────────────
    elev_blocks = [
        {"text": "The elevator car is at the landing.", "role": "paragraph"},
        {"text": "Hoistway dimensions comply with EN 81-20.", "role": "paragraph"},
        {"text": "Counterweight rail inspection.", "role": "paragraph"},
    ]
    esc_blocks = [
        {"text": "The escalator handrail moves with the steps.", "role": "paragraph"},
        {"text": "Comb plate replacement procedure.", "role": "paragraph"},
        {"text": "Step chain tension must be within spec.", "role": "paragraph"},
    ]
    nodom_blocks = [
        {"text": "This is a software documentation file.", "role": "paragraph"},
        {"text": "Functions, classes, and modules.", "role": "paragraph"},
    ]
    vn_blocks = [
        {"text": "Thang máy chở khách 1000kg, tốc độ 1.6 m/s.", "role": "paragraph"},
        {"text": "Giếng thang phải đáp ứng TCVN 6395.", "role": "paragraph"},
        {"text": "Đối trọng có khối lượng cân bằng đầy đủ.", "role": "paragraph"},
    ]

    report.check("detect elevator (English)",
                 detect_subdomain(elev_blocks) == {"elevator"})
    report.check("detect escalator (English)",
                 detect_subdomain(esc_blocks) == {"escalator"})
    report.check("non-domain → empty",
                 detect_subdomain(nodom_blocks) == set())
    report.check("detect elevator (Vietnamese)",
                 detect_subdomain(vn_blocks) == {"elevator"},
                 f"got {detect_subdomain(vn_blocks)}")

    # ── seed_for_direction ──────────────────────────────────────────────
    sd_en_vi = seed_for_direction({"elevator"}, "English", "Vietnamese")
    sd_vi_en = seed_for_direction({"elevator"}, "Vietnamese", "English")
    sd_none  = seed_for_direction(set(), "English", "Vietnamese")
    sd_both  = seed_for_direction({"elevator", "escalator"}, "English", "Vietnamese")

    report.check("EN→VI elevator seed populated", len(sd_en_vi) > 0,
                 f"{len(sd_en_vi)} entries, 'hoistway'→{sd_en_vi.get('hoistway')!r}")
    report.check("VI→EN elevator seed populated", len(sd_vi_en) > 0,
                 f"{len(sd_vi_en)} entries")
    report.check("empty subdomain → empty seed", sd_none == {})
    report.check("multi-subdomain merges", len(sd_both) > len(sd_en_vi),
                 f"both={len(sd_both)} vs elev_only={len(sd_en_vi)}")
    report.check("hoistway has Vietnamese translation in seed",
                 sd_en_vi.get("hoistway") == "giếng thang")

    # ── build_glossary precedence ──────────────────────────────────────
    g = build_glossary(client=None, blocks=elev_blocks,
                       target_lang="Vietnamese", source_lang="English",
                       seed=sd_en_vi)
    report.check("seed entries survive AI-stub failure",
                 "hoistway" in g and g["hoistway"] == "giếng thang",
                 f"glossary size={len(g)}")

    # ── build_doc_context with subdomain ────────────────────────────────
    ctx = build_doc_context(elev_blocks, source_lang="English",
                            subdomains={"elevator"})
    report.check("domain style block injected for elevator",
                 "elevator/escalator engineering" in ctx)
    report.check("standards listed in context",
                 "EN 81-20" in ctx)
    report.check("units listed in context",
                 "m/s" in ctx)

    ctx_no = build_doc_context(nodom_blocks, source_lang="English",
                               subdomains=set())
    report.check("non-domain doc does NOT get full domain style",
                 "Preserve units verbatim" not in ctx_no)

    # ── STANDARDS_KEEP_AS_IS ────────────────────────────────────────────
    report.check("EN 81-20 in standards", "EN 81-20" in STANDARDS_KEEP_AS_IS)
    report.check("TCVN 6395 in standards", "TCVN 6395" in STANDARDS_KEEP_AS_IS)
    report.check("ASME A17.1 in standards", "ASME A17.1" in STANDARDS_KEEP_AS_IS)

finally:
    cleanup()


sys.exit(report.summary())
