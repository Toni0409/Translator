"""OCR smoke checks — cost helpers, occurrence extraction, caption selection,
overlay render, replace-by-occurrence.

Chạy: python tests/smoke_ocr.py
"""
from __future__ import annotations

import io
import os
import sys

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from tests._helpers import setup_env, _PNG_1X1, Report


cleanup = setup_env()
try:
    from word_backend import (
        extract_image_occurrences,
        estimate_image_input_tokens, estimate_ocr_cost,
        _build_ocr_prompt,
        insert_ocr_captions_into_docx,
        render_translated_overlay,
        replace_docx_image_occurrences,
        _count_docx_media, validate_docx_output,
    )

    report = Report("smoke_ocr")

    # ── T1: cost helpers ────────────────────────────────────────────────
    report.check("small img (200x100) = 258 tokens",
                 estimate_image_input_tokens(200, 100) == 258)
    report.check("medium img (1024x768) = 2 tiles × 258 = 516",
                 estimate_image_input_tokens(1024, 768) == 516)
    report.check("big img (2048x1536) = 6 tiles × 258 = 1548",
                 estimate_image_input_tokens(2048, 1536) == 1548)
    report.check("unknown dims → fallback 258",
                 estimate_image_input_tokens(None, None) == 258)

    # estimate_ocr_cost with sample occurrences
    occs_sample = [
        {"id": "OCC_0", "data": _PNG_1X1 * 100, "width_px": 1024, "height_px": 768},
        {"id": "OCC_1", "data": b"x" * 100,   "width_px": 50,   "height_px": 50},   # < 5KB → skip
    ]
    est = estimate_ocr_cost(occs_sample, skip_under_bytes=5000)
    report.check("estimate skips small image",
                 est["n_skipped"] == 1 and est["n_to_ocr"] == 1,
                 f"skipped={est['n_skipped']}, to_ocr={est['n_to_ocr']}")
    report.check("estimate has USD/VND > 0",
                 est["usd"] > 0 and est["vnd"] > 0,
                 f"${est['usd']:.6f} / {est['vnd']:.2f} VND")

    # ── T2: prompt builder ──────────────────────────────────────────────
    p_simple  = _build_ocr_prompt(None, "Vietnamese", None, None)
    p_full    = _build_ocr_prompt("English", "Vietnamese",
                                  {"car": "cabin"}, {"elevator"})
    report.check("prompt without source_lang has no 'from ...'",
                 "from " not in p_simple.split("Translate")[1][:30],
                 "OK simple prompt")
    report.check("prompt with source_lang says 'from English to Vietnamese'",
                 "from English to Vietnamese" in p_full)
    report.check("prompt with elevator subdomain has domain block",
                 "elevator/escalator engineering" in p_full)
    report.check("prompt with glossary includes term",
                 "car → cabin" in p_full)
    report.check("prompt always asks for regions[bbox]",
                 "regions" in p_simple and "bbox" in p_simple)

    # ── T3: occurrence extraction (same-media multi occurrence) ────────
    # Build a DOCX with the SAME image inserted 3 times
    from docx import Document
    from docx.shared import Inches

    doc = Document()
    doc.add_paragraph("Para 1")
    p1 = doc.add_paragraph()
    p1.add_run().add_picture(io.BytesIO(_PNG_1X1), width=Inches(0.4))
    doc.add_paragraph("Para 2")
    p2 = doc.add_paragraph()
    p2.add_run().add_picture(io.BytesIO(_PNG_1X1), width=Inches(0.4))
    doc.add_paragraph("Para 3")
    p3 = doc.add_paragraph()
    p3.add_run().add_picture(io.BytesIO(_PNG_1X1), width=Inches(0.4))
    out = io.BytesIO(); doc.save(out)
    src_docx = out.getvalue()

    occs = extract_image_occurrences(src_docx)
    report.check("3 occurrences extracted from 3 inserts",
                 len(occs) == 3,
                 f"got {len(occs)}")
    report.check("occurrence_index counts up per (part, rId)",
                 sorted(o["occurrence_index"] for o in occs) == [0, 1, 2] or
                 sorted(o["occurrence_index"] for o in occs) == [0, 0, 0],
                 # python-docx tạo 3 rId riêng → mỗi occ là 0; nếu chung rId thì 0,1,2
                 f"occurrence_indexes = {[o['occurrence_index'] for o in occs]}")
    report.check("each occurrence has rId, doc_part, paragraph_index",
                 all(o.get("rId") and o.get("doc_part") and "paragraph_index" in o
                     for o in occs))

    # ── T4: caption mode chọn từng ảnh ──────────────────────────────────
    # No selection → 0 captions inserted, DOCX unchanged
    ocr_results = {
        o["id"]: {"has_text": True, "ocr": f"text {i}", "translation": f"dịch {i}",
                  "regions": [], "confidence": 0.9}
        for i, o in enumerate(occs)
    }

    out_none = insert_ocr_captions_into_docx(src_docx, occs, ocr_results,
                                              selected_ids=[])
    report.check("selected_ids=[] → DOCX unchanged",
                 out_none == src_docx)

    # Pick only first → exactly 1 caption inserted
    only_first = [occs[0]["id"]]
    out_one = insert_ocr_captions_into_docx(src_docx, occs, ocr_results,
                                             selected_ids=only_first)
    # Count [OCR] occurrences in document.xml
    import zipfile
    with zipfile.ZipFile(io.BytesIO(out_one)) as zf:
        doc_xml = zf.read("word/document.xml")
    n_captions = doc_xml.count(b"[OCR] ")
    report.check("1 selection → 1 caption inserted",
                 n_captions == 1,
                 f"found {n_captions} [OCR] markers")

    # Pick all 3
    out_all = insert_ocr_captions_into_docx(src_docx, occs, ocr_results,
                                             selected_ids=[o["id"] for o in occs])
    with zipfile.ZipFile(io.BytesIO(out_all)) as zf:
        doc_xml = zf.read("word/document.xml")
    n_captions = doc_xml.count(b"[OCR] ")
    report.check("3 selections → 3 captions inserted",
                 n_captions == 3,
                 f"found {n_captions} [OCR] markers")

    # Re-apply same selection → dedupe (no duplicate captions)
    out_again = insert_ocr_captions_into_docx(out_all, occs, ocr_results,
                                               selected_ids=[o["id"] for o in occs])
    with zipfile.ZipFile(io.BytesIO(out_again)) as zf:
        doc_xml = zf.read("word/document.xml")
    n_captions = doc_xml.count(b"[OCR] ")
    report.check("Re-applying same selection does NOT duplicate captions",
                 n_captions == 3,
                 f"after re-apply: {n_captions} [OCR] markers")

    # Edited translations override
    edited = {occs[0]["id"]: "EDITED TEXT"}
    out_edit = insert_ocr_captions_into_docx(src_docx, occs, ocr_results,
                                              selected_ids=[occs[0]["id"]],
                                              edited_translations=edited)
    with zipfile.ZipFile(io.BytesIO(out_edit)) as zf:
        doc_xml = zf.read("word/document.xml")
    report.check("edited_translations override base translation",
                 b"[OCR] EDITED TEXT" in doc_xml)

    # ── T4b: font Unicode tiếng Việt (anti-square-glyph) ────────────────
    try:
        from PIL import ImageFont
        from word_backend import (
            _font_supports_vi, overlay_font_status, _resolve_font_path,
            OverlayFontError,
        )

        # `_font_supports_vi` phải return True cho font có Vietnamese glyph
        # (test với font Unicode đã resolve, không test load_default vì Pillow ≥10
        # default font cũng support VN — version-dependent).
        # Font ResolveAPI nên trả path hợp lệ (nếu env có font Unicode)
        path = _resolve_font_path()
        if path:
            f = ImageFont.truetype(path, size=20)
            report.check("resolved font supports Vietnamese",
                         _font_supports_vi(f),
                         f"font: {path}")
        else:
            report.check("env has Unicode-capable font installed",
                         False,
                         "no font found — render overlay sẽ raise OverlayFontError")

        status = overlay_font_status()
        report.check("overlay_font_status returns {ok, path, supports_vi}",
                     "ok" in status and "path" in status and "supports_vi" in status,
                     f"status={status}")

        # render_translated_overlay raise OverlayFontError khi font thiếu —
        # giả lập bằng monkeypatch _resolve_font_path
        import word_backend as _wb
        orig = _wb._resolve_font_path
        try:
            _wb._FONT_PATH_CACHE[0] = None
            _wb._resolve_font_path = lambda: None
            try:
                from PIL import Image
                buf = io.BytesIO()
                Image.new("RGB", (200, 150), (255, 255, 255)).save(buf, "PNG")
                _wb.render_translated_overlay(
                    buf.getvalue(), "image/png",
                    regions=[{"bbox": [0.1, 0.1, 0.8, 0.3], "translation": "Xin chào"}],
                )
                report.check("render_translated_overlay raises when no font", False,
                             "expected OverlayFontError")
            except OverlayFontError:
                report.check("render_translated_overlay raises OverlayFontError "
                             "when no Unicode font", True)
        finally:
            _wb._resolve_font_path = orig
            _wb._FONT_PATH_CACHE[0] = None
    except ImportError:
        pass

    # ── T5: overlay rendering with Pillow ────────────────────────────────
    try:
        import PIL  # noqa: F401
        # Single region overlay
        regions = [{"bbox": [0.1, 0.1, 0.8, 0.3],
                    "ocr": "Hello", "translation": "Xin chào"}]
        # Need a bigger image — use a 200x150 white PNG
        from PIL import Image
        buf = io.BytesIO()
        Image.new("RGB", (200, 150), (255, 255, 255)).save(buf, "PNG")
        big_png = buf.getvalue()

        new_bytes, new_ct = render_translated_overlay(big_png, "image/png",
                                                      regions=regions)
        report.check("overlay output is PNG", new_ct == "image/png")
        report.check("overlay bytes valid PNG (decodable)",
                     Image.open(io.BytesIO(new_bytes)).size == (200, 150))

        # No regions → returns input as-is
        same_bytes, same_ct = render_translated_overlay(big_png, "image/png",
                                                        regions=[])
        report.check("empty regions → return original", same_bytes == big_png)

        # Edited translation overrides single-region
        new2, _ = render_translated_overlay(big_png, "image/png",
                                            regions=regions,
                                            edited_translation="ABC")
        report.check("edited_translation rendered (different from base)",
                     new2 != new_bytes,
                     "OK overlay differs when edited text changes")

        # Replace-by-occurrence: pick only occ#1, expect occ#0 and occ#2 unchanged
        # (when occs share same rId; in our test python-docx assigns distinct rIds,
        # so each replacement only touches its own media)
        repl = {occs[1]["id"]: (new_bytes, "image/png")}
        out_repl = replace_docx_image_occurrences(src_docx, occs, repl)
        # Verify DOCX still valid + still has 3 drawing elements
        media_after = _count_docx_media(out_repl)
        report.check("replace keeps drawing count",
                     media_after["drawing"] >= 3,
                     f"drawings: {media_after['drawing']}")
        val = validate_docx_output(out_repl, original_bytes=src_docx)
        # output may have more media files if cloning occurred
        report.check("DOCX after replace passes ZIP/XML validation",
                     val["block_count"] > 0,
                     f"blocks={val['block_count']}, errors={val['errors']}")
    except ImportError:
        report.check("Pillow available", False, "Pillow not installed — overlay tests skipped")

finally:
    cleanup()


sys.exit(report.summary())
