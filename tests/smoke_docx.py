"""P5.3 + P5.4: Smoke check DOCX extract/apply/validate + media count preservation.

Chạy: python tests/smoke_docx.py
"""
from __future__ import annotations

import os
import sys

# Bootstrap sys.path so this script can be run via `python tests/smoke_docx.py`
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from tests._helpers import (
    setup_env, make_docx_with_image, make_docx_with_inline_image, make_docx_plain,
    Report,
)


cleanup = setup_env()
try:
    from word_backend import (
        extract_docx_blocks, apply_translations,
        validate_docx_output, _count_docx_media,
    )

    report = Report("smoke_docx")

    # ── Case 1: image-only paragraph ─────────────────────────────────────
    src = make_docx_with_image()
    blocks = extract_docx_blocks(src)
    media_only = [b for b in blocks if b["role"] == "media_only"]
    text_blocks = [b for b in blocks if b["role"] != "media_only"
                   and b["role"] != "image_alt"]
    report.check("image-only paragraph detected as media_only",
                 len(media_only) >= 1,
                 f"found {len(media_only)} media_only blocks")
    report.check("text paragraphs still extracted",
                 len(text_blocks) >= 2,
                 f"found {len(text_blocks)} text blocks")

    translations = {b["id"]: "[dịch] " + b["text"] for b in text_blocks}
    out = apply_translations(src, blocks, translations)
    src_media = _count_docx_media(src)
    out_media = _count_docx_media(out)
    report.check("media files preserved",
                 src_media["media_files"] == out_media["media_files"],
                 f"src={src_media['media_files']} out={out_media['media_files']}")
    report.check("w:drawing count preserved",
                 src_media["drawing"] == out_media["drawing"],
                 f"src={src_media['drawing']} out={out_media['drawing']}")

    val = validate_docx_output(out, original_bytes=src)
    report.check("validate passes (image-only case)",
                 val["valid"],
                 f"errors={val['errors']}")

    # ── Case 2: inline image trong paragraph có bold ─────────────────────
    src2 = make_docx_with_inline_image()
    blocks2 = extract_docx_blocks(src2)
    img_blocks = [b for b in blocks2 if b["role"] != "image_alt"]
    translations2 = {b["id"]: "Đậm | sau ảnh" for b in img_blocks}
    out2 = apply_translations(src2, blocks2, translations2)
    src2_media = _count_docx_media(src2)
    out2_media = _count_docx_media(out2)
    report.check("inline image preserved through translate-apply",
                 src2_media["media_files"] == out2_media["media_files"],
                 f"src={src2_media['media_files']} out={out2_media['media_files']}")
    report.check("inline w:drawing preserved",
                 src2_media["drawing"] == out2_media["drawing"])

    val2 = validate_docx_output(out2, original_bytes=src2)
    report.check("validate passes (inline image case)",
                 val2["valid"],
                 f"errors={val2['errors']}")

    # ── Case 3: validate detects media loss ──────────────────────────────
    # Apply to text version → so sánh với image version → media loss
    src3 = make_docx_plain()
    val_bad = validate_docx_output(src3, original_bytes=src)   # src has image, src3 doesn't
    report.check("validate detects media file loss",
                 not val_bad["valid"],
                 f"valid={val_bad['valid']}, errors={val_bad['errors']}")

    # ── Case 4: validate without original (backward-compat) ──────────────
    val_no_orig = validate_docx_output(out)
    report.check("validate works without original_bytes",
                 val_no_orig["valid"],
                 f"valid={val_no_orig['valid']}")

finally:
    cleanup()


sys.exit(report.summary())
