"""
Word backend: extract DOCX paragraphs, gọi Gemini (fallback chain + retry +
parallel chunks), ghi lại DOCX giữ nguyên format runs.

3 nhóm logic:
1. Extract  : `extract_docx_blocks`, `iter_all_paragraphs`, `_detect_role`,
              `chunk_blocks` (adaptive theo char count)
2. Translate: `translate_chunk` (raw, raises), `translate_chunk_with_retry`,
              `translate_parallel` (background worker entry)
3. Render   : `apply_translations`, `replace_paragraph_text_keep_format`
"""
import hashlib
import io
import json
import re as _re
import re
import time
import threading
from collections import Counter
from concurrent.futures import ThreadPoolExecutor, as_completed

from docx import Document
from docx.text.paragraph import Paragraph
from docx.oxml.ns import qn

from config import (
    WORD_MODELS, MAX_WORD_TOKENS, NO_TRANSLATE_ROLES,
    CHUNK_RETRIES, TARGET_CHUNK_CHARS, MIN_CHUNK_BLOCKS, MAX_CHUNK_BLOCKS,
    HF_REPEAT_THRESHOLD, HF_REPEAT_MIN_CHARS,
)
from gemini import generate, usage_tokens, parse_json_loose


# Module-level cache cho model đang work (persist giữa các Streamlit rerun + thread).
# Cần lock vì 4 worker thread cùng đọc/ghi.
_working_model: list = [None]
_model_lock = threading.Lock()


# ══════════════════════════════════════════════════════════════════════════════
# PARAGRAPH ITERATION
# ══════════════════════════════════════════════════════════════════════════════
def _iter_paragraphs_in_table(table):
    for row in table.rows:
        for cell in row.cells:
            for para in cell.paragraphs:
                yield para
            for nested in cell.tables:
                yield from _iter_paragraphs_in_table(nested)


def _iter_footnote_endnote_paragraphs(doc):
    """
    Yield (Paragraph, hint) cho footnotes và endnotes.
    hint = "footnote" | "endnote".
    Bỏ qua separator/continuation (id = -1, 0).
    """
    SKIP_IDS = {"-1", "0"}
    for rel in doc.part.rels.values():
        rt = rel.reltype
        if "footnotes" not in rt and "endnotes" not in rt:
            continue
        hint = "footnote" if "footnotes" in rt else "endnote"
        try:
            root = rel.target_part._element
            for fn in root:
                if fn.get(qn("w:id"), "") in SKIP_IDS:
                    continue
                for p_elem in fn.iter(qn("w:p")):
                    yield Paragraph(p_elem, parent=doc), hint
        except Exception:
            pass


def _iter_comment_paragraphs(doc):
    """Yield (Paragraph, "comment") for all comment body paragraphs."""
    for rel in doc.part.rels.values():
        if "comments" not in rel.reltype:
            continue
        try:
            root = rel.target_part._element
            for comment in root:
                for p_elem in comment.iter(qn("w:p")):
                    yield Paragraph(p_elem, parent=doc), "comment"
        except Exception:
            pass


def _iter_blocks_with_meta(doc):
    """
    Single-walk generator → yield (paragraph, meta_dict) theo thứ tự ổn định.

    Meta = {
      "role_hint":  "body" | "header" | "footer" | "textbox"
                    | "footnote" | "endnote",
      "in_table":   bool,
      "table_cell": (T#, R#, C#) tuple 1-based, hoặc None,
    }

    Lý do single-walk: lxml element proxy có Python `id()` KHÔNG ổn định
    (cùng XML element được wrap nhiều proxy khác nhau ở các thời điểm khác nhau,
    và id có thể collide). → KHÔNG dùng `id(elem)` làm dict key cross-walk.
    → Dùng position index (idx) khi cần map extract ↔ apply.
    """
    t_counter = [0]

    def _walk_table(table, t_idx, role_hint):
        for r_idx, row in enumerate(table.rows, 1):
            for c_idx, cell in enumerate(row.cells, 1):
                for para in cell.paragraphs:
                    yield para, {
                        "role_hint": role_hint,
                        "in_table":  True,
                        "table_cell": (t_idx, r_idx, c_idx),
                    }
                for nested in cell.tables:
                    yield from _walk_table(nested, t_idx, role_hint)

    # Body paragraphs
    for para in doc.paragraphs:
        hint = "toc" if _is_toc_paragraph(para) else "body"
        yield para, {"role_hint": hint, "in_table": False, "table_cell": None}
    # Body tables
    for tbl in doc.tables:
        t_counter[0] += 1
        yield from _walk_table(tbl, t_counter[0], "body")
    # Headers + footers
    for section in doc.sections:
        try:
            for para in section.header.paragraphs:
                yield para, {"role_hint": "header", "in_table": False, "table_cell": None}
            for tbl in section.header.tables:
                t_counter[0] += 1
                yield from _walk_table(tbl, t_counter[0], "header")
        except Exception:
            pass
        try:
            for para in section.footer.paragraphs:
                yield para, {"role_hint": "footer", "in_table": False, "table_cell": None}
            for tbl in section.footer.tables:
                t_counter[0] += 1
                yield from _walk_table(tbl, t_counter[0], "footer")
        except Exception:
            pass
    # Text-boxes / shapes trong body document part
    # (text-boxes trong header/footer parts chưa cover ở iteration này)
    # parent=doc để `para.style` resolve được qua doc.part.styles
    for txbx in doc.element.iter(qn("w:txbxContent")):
        for p_elem in txbx.iter(qn("w:p")):
            yield Paragraph(p_elem, parent=doc), {
                "role_hint": "textbox", "in_table": False, "table_cell": None,
            }
    # Footnotes + endnotes
    for para, hint in _iter_footnote_endnote_paragraphs(doc):
        yield para, {"role_hint": hint, "in_table": False, "table_cell": None}
    # VML text-boxes (Word 2003/older format) — w:txbxContent usually already caught above;
    # this handles the rare case where VML v:textbox has direct w:p without w:txbxContent
    _VML_TB = "{urn:schemas-microsoft-com:vml}textbox"
    _seen_p = {id(p_elem) for txbx in doc.element.iter(qn("w:txbxContent"))
               for p_elem in txbx.iter(qn("w:p"))}
    for vml_tb in doc.element.iter(_VML_TB):
        for p_elem in vml_tb.iter(qn("w:p")):
            if id(p_elem) not in _seen_p:
                yield Paragraph(p_elem, parent=doc), {
                    "role_hint": "textbox", "in_table": False, "table_cell": None,
                }
    # Comments (word/comments.xml)
    for para, hint in _iter_comment_paragraphs(doc):
        yield para, {"role_hint": hint, "in_table": False, "table_cell": None}


def iter_all_paragraphs(doc):
    """
    Lặp qua toàn bộ paragraph: body → tables → headers → footers → text-boxes.
    Thin wrapper around `_iter_blocks_with_meta` — guarantee cùng thứ tự với extract.
    """
    for para, _meta in _iter_blocks_with_meta(doc):
        yield para


# ══════════════════════════════════════════════════════════════════════════════
# INLINE FORMAT — encode/decode bold/italic/underline qua tag để AI giữ được
# ══════════════════════════════════════════════════════════════════════════════
_TAG_RE      = re.compile(r"</?[biu]>")
_TAG_STRIP   = re.compile(r"</?[biu]>")
_ESC_MAP     = (("&", "&amp;"), ("<", "&lt;"), (">", "&gt;"))
_UNESC_MAP   = (("&lt;", "<"), ("&gt;", ">"), ("&amp;", "&"))


def _esc(s: str) -> str:
    for a, b in _ESC_MAP:
        s = s.replace(a, b)
    return s


def _unesc(s: str) -> str:
    for a, b in _UNESC_MAP:
        s = s.replace(a, b)
    return s


def runs_to_tagged_text(paragraph) -> str:
    """
    Convert paragraph runs → tagged text. VD: '<b>Important</b>: read <i>carefully</i>'.
    Escape & < > trong text gốc để không xung đột với tag.
    Bao gồm text bên trong w:hyperlink (URL được giữ nguyên khi apply).
    Math equations (m:oMath) → <MATH/> placeholder.
    Field runs (w:fldChar, w:instrText) are skipped — kept as-is.
    """
    from docx.text.run import Run
    _MATH_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
    parts = []
    for child in paragraph._p.iterchildren():
        # Math equations — preserve verbatim with placeholder
        if child.tag == f"{{{_MATH_NS}}}oMath":
            parts.append("<MATH/>")
            continue
        tag = child.tag.split("}", 1)[-1]
        if tag == "r":
            # Skip field instruction/delimiter runs
            if child.find(qn("w:instrText")) is not None:
                continue
            fld = child.find(qn("w:fldChar"))
            if fld is not None:
                continue
            run = Run(child, paragraph)
            text = run.text
            if not text:
                continue
            text = _esc(text)
            if run.bold:      text = f"<b>{text}</b>"
            if run.italic:    text = f"<i>{text}</i>"
            if run.underline: text = f"<u>{text}</u>"
            parts.append(text)
        elif tag == "hyperlink":
            # Include hyperlink inner text — URL stays in XML, only text is translated
            for r_elem in child.iterchildren(qn("w:r")):
                if r_elem.find(qn("w:instrText")) is not None:
                    continue
                if r_elem.find(qn("w:fldChar")) is not None:
                    continue
                run = Run(r_elem, paragraph)
                text = run.text
                if not text:
                    continue
                text = _esc(text)
                if run.bold:      text = f"<b>{text}</b>"
                if run.italic:    text = f"<i>{text}</i>"
                if run.underline: text = f"<u>{text}</u>"
                parts.append(text)
        elif tag == "ins":
            # Track changes: inserted text — include for translation
            for r_elem in child.iterchildren(qn("w:r")):
                if r_elem.find(qn("w:instrText")) is not None:
                    continue
                if r_elem.find(qn("w:fldChar")) is not None:
                    continue
                run = Run(r_elem, paragraph)
                text = run.text
                if not text:
                    continue
                text = _esc(text)
                if run.bold:      text = f"<b>{text}</b>"
                if run.italic:    text = f"<i>{text}</i>"
                if run.underline: text = f"<u>{text}</u>"
                parts.append(text)
        # w:del: intentionally skipped — deleted text not translated
    return "".join(parts).strip()


def has_inline_format(tagged: str) -> bool:
    return _TAG_RE.search(tagged) is not None


def strip_tags(tagged: str) -> str:
    """Bỏ tag, unescape — dùng khi fallback về plain text."""
    return _unesc(_TAG_STRIP.sub("", tagged))


def _parse_tagged(tagged: str) -> list[tuple[str, bool, bool, bool]]:
    """
    Parse tagged text → list of (text, bold, italic, underline).
    Tolerant với tag mismatch (ignore).
    """
    segments: list = []
    stack:    list = []
    last           = 0

    def flush(text):
        if text:
            segments.append((_unesc(text), "b" in stack, "i" in stack, "u" in stack))

    for m in _TAG_RE.finditer(tagged):
        if m.start() > last:
            flush(tagged[last:m.start()])
        tag = m.group()
        if tag.startswith("</"):
            if stack and stack[-1] == tag[2]:
                stack.pop()
        else:
            stack.append(tag[1])
        last = m.end()
    if last < len(tagged):
        flush(tagged[last:])
    return segments


def _paragraph_has_non_run_children(paragraph) -> bool:
    """True nếu paragraph có hyperlink/field/... — không rebuild runs an toàn được."""
    SAFE = {"pPr", "r", "hyperlink", "ins", "del",
            "bookmarkStart", "bookmarkEnd", "proofErr", "fldSimple",
            "rPr", "pPrChange", "oMath"}
    for child in paragraph._p.iterchildren():
        tag = child.tag.split("}", 1)[-1]
        if tag not in SAFE:
            return True
    return False


def replace_paragraph_text_keep_format(paragraph, new_text: str):
    """
    Replace paragraph text keeping format of first run.
    Handles hyperlinks: collects ALL w:r elements (regular + inside w:hyperlink)
    so translated text is written correctly even in hyperlink-only paragraphs.
    Hyperlink URLs (relationships) are untouched; only inner text changes.
    """
    # Collect all w:r elements in document order, including inside w:hyperlink
    # Field runs (w:fldChar, w:instrText) are excluded — keep original XML
    all_r = []
    for child in paragraph._p.iterchildren():
        ctag = child.tag.split("}", 1)[-1]
        if ctag == "r":
            # Field runs — skip from translation
            has_fld = (child.find(qn("w:fldChar")) is not None or
                       child.find(qn("w:instrText")) is not None)
            if has_fld:
                continue
            all_r.append(child)
        elif ctag == "hyperlink":
            for r_elem in child.iterchildren(qn("w:r")):
                has_fld = (r_elem.find(qn("w:fldChar")) is not None or
                           r_elem.find(qn("w:instrText")) is not None)
                if not has_fld:
                    all_r.append(r_elem)
        elif ctag == "ins":
            # Track changes insertion: include runs for translation
            for r_elem in child.iterchildren(qn("w:r")):
                has_fld = (r_elem.find(qn("w:fldChar")) is not None or
                           r_elem.find(qn("w:instrText")) is not None)
                if not has_fld:
                    all_r.append(r_elem)
        # del: skip — deleted text untouched

    if not all_r:
        paragraph.add_run(new_text)
        return

    # Find first run that has non-empty text (use as formatting template)
    first_r = next(
        (r for r in all_r
         if any((t.text or "").strip() for t in r.findall(qn("w:t")))),
        all_r[0],
    )

    # Write new_text into first_r
    t_elems = first_r.findall(qn("w:t"))
    if t_elems:
        t_elems[0].text = new_text
        if new_text and (new_text[0] == " " or new_text[-1] == " "):
            t_elems[0].set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        for t in t_elems[1:]:
            t.text = ""
    else:
        from lxml import etree
        t = etree.SubElement(first_r, qn("w:t"))
        t.text = new_text
        if new_text and (new_text[0] == " " or new_text[-1] == " "):
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")

    # Clear all other runs (including hyperlink runs)
    for r_elem in all_r:
        if r_elem is first_r:
            continue
        for t in r_elem.findall(qn("w:t")):
            t.text = ""


def replace_paragraph_with_tagged(paragraph, tagged_text: str):
    """
    Path 2 (new): parse <b><i><u> tags trong tagged_text và rebuild runs để
    giữ inline format. Inherit font/size/color từ run đầu tiên có text.
    Fallback về `replace_paragraph_text_keep_format` nếu paragraph có
    hyperlink/field, hoặc không parse được tag.
    """
    if _paragraph_has_non_run_children(paragraph):
        replace_paragraph_text_keep_format(paragraph, strip_tags(tagged_text))
        return

    segments = _parse_tagged(tagged_text)
    if not segments:
        replace_paragraph_text_keep_format(paragraph, strip_tags(tagged_text))
        return

    # Lấy template format từ run đầu có text
    runs     = paragraph.runs
    template = next((r for r in runs if r.text), runs[0] if runs else None)
    tpl_name = tpl_size = tpl_color = None
    if template is not None:
        try: tpl_name  = template.font.name
        except Exception: pass
        try: tpl_size  = template.font.size
        except Exception: pass
        try: tpl_color = template.font.color.rgb
        except Exception: pass

    # Xóa hết runs cũ
    p_elem = paragraph._p
    for run in list(runs):
        try: p_elem.remove(run._r)
        except Exception: pass

    # Thêm runs mới theo segments
    for text, bold, italic, underline in segments:
        if not text:
            continue
        run = paragraph.add_run(text)
        if bold:      run.bold      = True
        if italic:    run.italic    = True
        if underline: run.underline = True
        if tpl_name:  run.font.name = tpl_name
        if tpl_size:  run.font.size = tpl_size
        if tpl_color:
            try: run.font.color.rgb = tpl_color
            except Exception: pass


# ══════════════════════════════════════════════════════════════════════════════
# ROLE DETECTION (rule-based, không gọi AI)
# ══════════════════════════════════════════════════════════════════════════════
def _detect_role(para, in_table: bool, in_header: bool, in_footer: bool) -> str:
    if in_header: return "header"
    if in_footer: return "footer"
    style = (para.style.name or "").lower()
    if "toc" in style:     return "toc"
    if "heading" in style: return "section_heading"
    if in_table:           return "table_cell"
    try:
        pPr = para._p.pPr
        if pPr is not None and pPr.numPr is not None:
            return "bullet"
    except Exception:
        pass
    if style in ("title", "subtitle"):
        return "title"
    if any(k in style for k in ("caption", "note", "warning", "caution")):
        return "note"
    return "paragraph"


# ══════════════════════════════════════════════════════════════════════════════
# TOC DETECTION
# ══════════════════════════════════════════════════════════════════════════════
def _is_toc_paragraph(paragraph) -> bool:
    """Detect if paragraph is part of a Table of Contents."""
    try:
        style = (paragraph.style.name or "").lower()
        if style.startswith("toc"):
            return True
    except Exception:
        pass
    # Check for hyperlink to _Toc bookmark
    try:
        for h in paragraph._p.iter(qn("w:hyperlink")):
            anchor = h.get(qn("w:anchor"), "")
            if anchor.startswith("_Toc"):
                return True
    except Exception:
        pass
    return False


def _link_toc_to_headings(blocks: list[dict]) -> int:
    """Set _toc_mirror on TOC blocks pointing to matching heading block id."""
    headings = {}
    for b in blocks:
        if b["role"] == "section_heading":
            key = b["text"].strip().lower()
            if key and key not in headings:
                headings[key] = b["id"]
    linked = 0
    for b in blocks:
        if b["role"] != "toc":
            continue
        raw = b["text"]
        # Strip trailing tab+digits or dots+digits (page number suffix)
        m = _re.match(r"^(.*?)(?:[\t\s\.]+\d+)?$", raw)
        head_text = (m.group(1) if m else raw).strip().lower()
        if head_text in headings:
            b["_toc_mirror"] = headings[head_text]
            linked += 1
    return linked


# ══════════════════════════════════════════════════════════════════════════════
# EXTRACT
# ══════════════════════════════════════════════════════════════════════════════
def _mark_repeating_as_hf(blocks: list[dict],
                          threshold: int = HF_REPEAT_THRESHOLD,
                          min_chars: int = HF_REPEAT_MIN_CHARS) -> int:
    """
    Phát hiện block trong body có text lặp lại ≥ threshold lần
    (text dài ≥ min_chars) → đánh dấu role = 'body_repeated'.

    Đây là heuristic phổ biến để bắt header/footer ẩn trong body
    (vd: tiêu đề chapter chạy trên mỗi trang, watermark "Confidential",
    section banner...).

    Trả về số block bị đánh dấu.
    """
    counter: Counter = Counter()
    for b in blocks:
        if b["role"] in ("header", "footer"):
            continue   # đã là H/F thật
        if len(b["text"]) < min_chars:
            continue   # text quá ngắn, dễ false positive
        counter[b["text"]] += 1

    repeating_texts = {t for t, c in counter.items() if c >= threshold}
    marked = 0
    for b in blocks:
        if b["role"] in ("header", "footer"):
            continue
        if b["text"] in repeating_texts:
            b["role"] = "body_repeated"
            marked += 1
    return marked


def _iter_image_alt_texts(doc):
    """
    Yield (element, attr_name, text) for every image alt-text/title in the doc.
    Caller can later set element.set(attr_name, translated_text).

    Covers:
    - drawingML wp:docPr (descr, title)
    - drawingML pic:cNvPr (descr, title)
    - VML v:shape (alt)
    """
    WP_NS  = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture"
    VML_NS = "urn:schemas-microsoft-com:vml"

    for tag in (f"{{{WP_NS}}}docPr", f"{{{PIC_NS}}}cNvPr"):
        for el in doc.element.iter(tag):
            for attr in ("descr", "title"):
                val = (el.get(attr) or "").strip()
                if val:
                    yield el, attr, val

    for el in doc.element.iter(f"{{{VML_NS}}}shape"):
        val = (el.get("alt") or "").strip()
        if val:
            yield el, "alt", val


def extract_docx_blocks(docx_bytes: bytes) -> list[dict]:
    """
    Đọc DOCX → list block {id, text, role, para_idx, table_cell, ...}.

    Single-walk extraction (xem `_iter_blocks_with_meta` để biết tại sao).
    Role:
    - header / footer / textbox : từ meta hint
    - body_repeated             : text lặp lại ≥ 3 lần trong body (heuristic H/F)
    - section_heading / title / bullet / table_cell / note / toc / paragraph
    """
    doc = Document(io.BytesIO(docx_bytes))
    blocks = []
    for idx, (para, meta) in enumerate(_iter_blocks_with_meta(doc)):
        text = para.text.strip()
        if not text:
            continue
        role = _detect_role(
            para,
            meta["in_table"],
            meta["role_hint"] == "header",
            meta["role_hint"] == "footer",
        )
        # Override special roles (chỉ khi không phải H/F thật)
        if meta["role_hint"] == "textbox" and role not in ("header", "footer"):
            role = "textbox"
        if meta["role_hint"] in ("footnote", "endnote") and role not in ("header", "footer"):
            role = meta["role_hint"]
        if meta["role_hint"] == "comment" and role not in ("header", "footer"):
            role = "comment"
        if meta["role_hint"] == "toc" and role not in ("header", "footer"):
            role = "toc"
        tagged = runs_to_tagged_text(para)
        blocks.append({
            "id":          f"p{idx}",
            "text":        text,
            "text_tagged": tagged,
            "has_format":  has_inline_format(tagged),
            "role":        role,
            "para_idx":    idx,
            "table_cell":  meta["table_cell"],   # (T,R,C) tuple hoặc None
        })
    _mark_repeating_as_hf(blocks)
    _link_toc_to_headings(blocks)

    # Image alt-texts (drawingML + VML). Appended AFTER body blocks so
    # `apply_translations` can re-iterate `_iter_image_alt_texts` and match by order.
    for i, (_, _, alt_text) in enumerate(_iter_image_alt_texts(doc)):
        blocks.append({
            "id":          f"IMG_ALT_{i}",
            "text":        alt_text,
            "text_tagged": alt_text,
            "has_format":  False,
            "role":        "image_alt",
            "para_idx":    -1,
            "table_cell":  None,
        })

    return blocks


def count_by_role(blocks: list[dict]) -> dict:
    """Stat helper: đếm số block theo từng role + tổng theo nhóm."""
    counter = Counter(b["role"] for b in blocks)
    return {
        "total":         len(blocks),
        "body":          sum(c for r, c in counter.items() if r not in NO_TRANSLATE_ROLES),
        "header":        counter.get("header", 0),
        "footer":        counter.get("footer", 0),
        "body_repeated": counter.get("body_repeated", 0),
        "hf_total":      sum(counter.get(r, 0) for r in NO_TRANSLATE_ROLES),
        "footnote":      counter.get("footnote", 0),
        "endnote":       counter.get("endnote", 0),
        "comment":       counter.get("comment", 0),
        "by_role":       dict(counter),
    }


_NOUN_PHRASE_RE = re.compile(r"\b[A-Z][a-zA-Z\-]+(?:\s+[A-Z]?[a-zA-Z\-]+){0,3}\b")
_TECH_TERM_RE   = re.compile(r"\b[A-Z][a-zA-Z]{4,}\b")


def build_glossary(client, blocks: list[dict], target_lang: str,
                   source_lang: str | None = None,
                   top_n: int = 30, min_repeat: int = 3) -> dict:
    """
    One-shot AI call để dịch top-N thuật ngữ lặp lại trong doc.
    Inject vào mọi chunk prompt để dịch consistent cross-chunk.

    Heuristic candidates:
    - Noun phrases (capitalized, 1-4 words) — "Hoistway Door", "Control Panel"
    - Single technical terms (CamelCase, ≥5 chars) — "Inverter", "Calibration"

    Trả về {} nếu API fail hoặc không tìm thấy term.
    """
    text_all = " ".join(b["text"] for b in blocks
                        if b["role"] not in NO_TRANSLATE_ROLES)
    if not text_all:
        return {}

    candidates = _NOUN_PHRASE_RE.findall(text_all) + _TECH_TERM_RE.findall(text_all)
    counter    = Counter(candidates)
    top_terms  = [w for w, c in counter.most_common(top_n * 2) if c >= min_repeat][:top_n]
    if not top_terms:
        return {}

    src_clause = f"from {source_lang} " if source_lang else ""
    prompt = f"""Translate these technical terms {src_clause}to {target_lang}.
Return ONLY a JSON object: {{"term": "translation", ...}}

Rules:
- Use formal, consistent terminology suitable for technical documents.
- Keep proper nouns / company names / product codes / brand names UNCHANGED.
- Keep numbers, units, model codes UNCHANGED.
- Each term gets exactly ONE canonical translation.

Terms ({len(top_terms)}):
{json.dumps(top_terms, ensure_ascii=False)}"""

    try:
        raw, _, _ = call_gemini(client, prompt)
    except Exception:
        return {}
    parsed = parse_json_loose(raw)
    if not isinstance(parsed, dict):
        return {}
    # Filter — chỉ giữ entry hợp lệ
    return {
        str(k).strip(): str(v).strip()
        for k, v in parsed.items()
        if k and v and isinstance(v, str) and v.strip() != str(k).strip()
    }


def build_doc_context(blocks: list[dict], source_lang: str | None = None) -> str:
    """Tóm tắt cấu trúc tài liệu (title + headings + TOC + domain hint).

    `source_lang` được nhận để inject vào context nếu cần (Phase 1: không thay đổi
    output hiện tại, nhưng đảm bảo signature consistent).
    """
    titles   = [b["text"] for b in blocks if b["role"] == "title"][:2]
    headings = [b["text"] for b in blocks if b["role"] == "section_heading"][:10]
    tocs     = [b["text"] for b in blocks if b["role"] == "toc"][:8]

    lines = []
    if titles:   lines.append("Document title: " + " | ".join(titles))
    if headings: lines.append("Main sections: "  + " | ".join(headings[:6]))
    if tocs:     lines.append("TOC: "            + " | ".join(tocs[:5]))
    ctx = "\n".join(lines) if lines else "Technical document."

    all_text = " ".join(b["text"] for b in blocks[:30]).lower()
    if any(k in all_text for k in ("elevator", "lift", "hoistway", "schindler", "inventio")):
        ctx += "\nDomain: elevator/lift engineering."
    elif any(k in all_text for k in ("safety", "standard", "regulation", "iso", "en 81")):
        ctx += "\nDomain: safety standards."
    elif any(k in all_text for k in ("software", "api", "code", "function")):
        ctx += "\nDomain: software/IT."
    return ctx


# ══════════════════════════════════════════════════════════════════════════════
# ADAPTIVE CHUNKING — chia theo char count, có min/max paragraph count
# ══════════════════════════════════════════════════════════════════════════════
def chunk_blocks(blocks: list[dict],
                 target_chars: int = TARGET_CHUNK_CHARS,
                 min_count: int   = MIN_CHUNK_BLOCKS,
                 max_count: int   = MAX_CHUNK_BLOCKS) -> list[list[dict]]:
    """
    Chia list block thành chunks sao cho:
    - Mỗi chunk có tổng char ~ target_chars
    - Tối thiểu min_count, tối đa max_count paragraph / chunk
    """
    chunks: list[list[dict]] = []
    current: list[dict]      = []
    cur_chars = 0
    for b in blocks:
        size = len(b["text"]) + 30   # +overhead JSON wrapping
        full = (cur_chars + size > target_chars and len(current) >= min_count) \
               or len(current) >= max_count
        if current and full:
            chunks.append(current)
            current, cur_chars = [], 0
        current.append(b)
        cur_chars += size
    if current:
        chunks.append(current)
    return chunks


# ══════════════════════════════════════════════════════════════════════════════
# GEMINI CALL với fallback chain + token tracking
# ══════════════════════════════════════════════════════════════════════════════
def call_gemini(client, prompt: str) -> tuple[str, int, int]:
    """
    Gọi Gemini với fallback chain. Trả về (text, in_tokens, out_tokens).
    Raise RuntimeError nếu tất cả model đều fail. Thread-safe.
    """
    with _model_lock:
        cur = _working_model[0]
    models = ([cur] if cur else []) + [m for m in WORD_MODELS if m != cur]
    last_err = "no models"
    for model in models:
        try:
            resp = generate(client, model, prompt,
                            max_output_tokens=MAX_WORD_TOKENS, temperature=0.1)
            text = (resp.text or "").strip()
            if not text:
                last_err = "empty response"
                continue
            in_t, out_t = usage_tokens(resp)
            with _model_lock:
                _working_model[0] = model
            return text, in_t, out_t
        except Exception as e:
            last_err = str(e)
    raise RuntimeError(f"All Word models failed: {last_err}")


def get_working_model() -> str | None:
    with _model_lock:
        return _working_model[0]


# ══════════════════════════════════════════════════════════════════════════════
# TRANSLATE chunk (raw + retry)
# ══════════════════════════════════════════════════════════════════════════════
_TABLE_PREFIX_RE = re.compile(r"^\s*\(T\d+\s+R\d+\s+C\d+\)\s*")


def _format_block_text(b: dict) -> str:
    """Prefix (T#R#C#) cho table cell, dùng tagged text nếu có inline format."""
    text = b.get("text_tagged") or b["text"]
    tc   = b.get("table_cell")
    if tc:
        return f"(T{tc[0]} R{tc[1]} C{tc[2]}) {text}"
    return text


def _build_chunk_prompt(chunk: list[dict], target_lang: str, doc_context: str,
                        glossary: dict | None = None,
                        custom_rules: dict | None = None,
                        source_lang: str | None = None) -> str:
    """Build prompt. Dùng `text_tagged` nếu có (preserve inline format)."""
    payload = json.dumps(
        [{
            "id":   b["id"],
            "text": _format_block_text(b),
            "role": b.get("role", "paragraph"),
        } for b in chunk],
        ensure_ascii=False,
    )

    has_format = any(b.get("has_format") for b in chunk)
    has_tables = any(b.get("table_cell") for b in chunk)

    format_rule = (
        "- Some texts contain HTML-like tags <b>, <i>, <u> marking bold/italic/underline. "
        "PRESERVE these tags EXACTLY in the translation, wrapping the equivalent translated words. "
        "Do NOT add new tags where source has none. Do NOT use markdown (** or _).\n"
        if has_format else ""
    )
    table_rule = (
        "- Table cells are prefixed with (T# R# C#) — table/row/column context. "
        "DO NOT include this prefix in the translation output. "
        "Cells in the same column (same C#) MUST use the same terminology. "
        "Row 1 (R1) is the header — keep concise. "
        "Preserve numbers, units, model codes in cells verbatim.\n"
        if has_tables else ""
    )

    glossary_section = ""
    if glossary:
        # Limit glossary size để không bloat prompt
        items = list(glossary.items())[:50]
        glossary_str = "\n".join(f"  {en} → {vi}" for en, vi in items)
        glossary_section = (
            f"\nGlossary (USE THESE EXACT translations for consistency across the document):\n"
            f"{glossary_str}\n"
        )

    custom_rules_section = ""
    if custom_rules:
        extra = (custom_rules.get("_extra") or "").strip()
        if extra:
            custom_rules_section += f"\nAdditional instruction: {extra}\n"
        lines = [f"- For {role}: {rule.strip()}"
                 for role, rule in custom_rules.items()
                 if role != "_extra" and isinstance(rule, str) and rule.strip()]
        if lines:
            custom_rules_section += (
                "\nCustom rules (override defaults for these roles):\n"
                + "\n".join(lines) + "\n"
            )

    translate_header = (
        f"Translate these Word document blocks from {source_lang} into {target_lang}."
        if source_lang else
        f"Translate these Word document blocks into {target_lang}."
    )

    return f"""{translate_header}

Document context:
{doc_context}{glossary_section}{custom_rules_section}

Rules:
- Return ONLY valid JSON array. No markdown, no explanation.
- Keep "id" exactly as given.
- Translate ALL words into {target_lang}. Leave nothing in the source language.
- Preserve: numbers, units, product codes, model names, document IDs, revision codes.
- Do NOT translate company names (e.g. INVENTIO AG, Schindler, Otis).
- Preserve ALL punctuation, whitespace, tabs (\\t), line breaks (\\n).
- Match source sentence count — do NOT merge or split sentences.
{format_rule}{table_rule}- Preserve <MATH/> placeholders exactly as-is (they represent math equations).
- Preserve <FIELD>...</FIELD> content unchanged.
- For toc: translate heading text only; preserve numbering, dot leaders (...), page numbers.
- For bullet: concise imperative, keep bullet symbol.
- For heading/section_heading: concise technical, match source brevity.
- For table_cell: concise, consistent across row.
- For note: keep WARNING/NOTE/CAUTION label verbatim.
- For header/footer/body_repeated: concise like a page header/footer; preserve page numbers ("Page 1 of 10"), dates, document IDs verbatim.
- For footnote/endnote: translate fully as body text; preserve footnote reference numbers (¹²³ or superscript digits) verbatim.
- For comment: translate the comment text naturally as body text.
- For image_alt: concise description (max ~100 chars), suitable as screen-reader alt-text.

Format:
[{{"id":"...","text":"translated text"}}]

Blocks:
{payload}"""


def translate_chunk(client, chunk: list[dict], target_lang: str,
                    doc_context: str,
                    glossary: dict | None = None,
                    custom_rules: dict | None = None,
                    source_lang: str | None = None) -> tuple[dict, int, int]:
    """
    Dịch 1 chunk → (translations_dict, in_tokens, out_tokens).
    RAISE exception nếu API fail. Block bị model bỏ sót → fallback giữ text gốc.
    """
    prompt = _build_chunk_prompt(
        chunk, target_lang, doc_context,
        glossary=glossary, custom_rules=custom_rules,
        source_lang=source_lang,
    )
    raw, in_t, out_t = call_gemini(client, prompt)
    parsed = parse_json_loose(raw) or []
    if not isinstance(parsed, list):
        parsed = []
    result = {}
    for item in parsed:
        if isinstance(item, dict) and item.get("id") and item.get("text") is not None:
            text = str(item["text"])
            text = _TABLE_PREFIX_RE.sub("", text)   # AI lỡ giữ prefix → strip
            result[item["id"]] = text
    for b in chunk:
        # Fallback: nếu có tagged thì dùng tagged để preserve format khi không dịch được
        result.setdefault(b["id"], b.get("text_tagged") or b["text"])
    return result, in_t, out_t


def translate_chunk_with_retry(client, chunk: list[dict], target_lang: str,
                               doc_context: str,
                               retries: int = CHUNK_RETRIES,
                               glossary: dict | None = None,
                               custom_rules: dict | None = None,
                               fallback_individually: bool = True,
                               source_lang: str | None = None,
                               ) -> tuple[dict, int, int, str | None]:
    """
    Wrapper với exponential backoff. Trả về (translations, in_t, out_t, error).
    Nếu hết retries:
    - Nếu chunk có >1 block: thử dịch lại từng block riêng lẻ (smart fallback).
    - Nếu block đơn hoặc smart fallback cũng fail → giữ text gốc + error message.
    """
    last_err = None
    tok_in_total = tok_out_total = 0
    for attempt in range(retries):
        try:
            t, in_t, out_t = translate_chunk(
                client, chunk, target_lang, doc_context,
                glossary=glossary, custom_rules=custom_rules,
                source_lang=source_lang,
            )
            return t, in_t, out_t, None
        except Exception as e:
            last_err = str(e)
            if attempt < retries - 1:
                time.sleep(2 ** attempt)   # 1s, 2s, 4s

    # Final failure — smart fallback: retry each block individually
    if fallback_individually and len(chunk) > 1:
        result = {}
        for block in chunk:
            try:
                t_single, in_t, out_t = translate_chunk(
                    client, [block], target_lang, doc_context,
                    glossary=glossary, custom_rules=custom_rules,
                    source_lang=source_lang,
                )
                result.update(t_single)
                tok_in_total  += in_t
                tok_out_total += out_t
            except Exception:
                # Individual block failed — keep original
                result[block["id"]] = block.get("text_tagged") or block["text"]
        return result, tok_in_total, tok_out_total, last_err

    fallback = {b["id"]: b.get("text_tagged") or b["text"] for b in chunk}
    return fallback, 0, 0, last_err


# ══════════════════════════════════════════════════════════════════════════════
# PARALLEL TRANSLATION — entry point cho background thread
# ══════════════════════════════════════════════════════════════════════════════
def translate_parallel(holder: dict, client,
                       chunks: list[list[dict]],
                       target_lang: str, doc_context: str,
                       max_workers: int,
                       glossary: dict | None = None,
                       custom_rules: dict | None = None,
                       source_lang: str | None = None):
    """
    Background worker: dịch nhiều chunk song song với ThreadPoolExecutor.

    `holder` (dict shared với main thread) sẽ được update dưới lock:
    - translations  : dict {block_id -> translated_text}
    - tok_in/tok_out: cumulative tokens
    - chunk_done    : số chunk đã xong
    - chunk_log     : list per-chunk result (đẩy thêm khi mỗi chunk xong)
    - done          : True khi xong tất cả
    - error         : str nếu fatal error
    """
    lock = threading.Lock()
    try:
        with ThreadPoolExecutor(max_workers=max_workers) as ex:
            futures = {
                ex.submit(translate_chunk_with_retry,
                          client, c, target_lang, doc_context,
                          CHUNK_RETRIES, glossary, custom_rules,
                          True, source_lang): (i, c)
                for i, c in enumerate(chunks)
            }
            for future in as_completed(futures):
                idx, chunk = futures[future]
                try:
                    result, in_t, out_t, err = future.result()
                except Exception as e:
                    result, in_t, out_t, err = (
                        {b["id"]: b["text"] for b in chunk}, 0, 0, str(e)
                    )
                with lock:
                    holder["translations"].update(result)
                    holder["tok_in"]  += in_t
                    holder["tok_out"] += out_t
                    holder["chunk_log"].append({
                        "idx":   idx,
                        "size":  len(chunk),
                        "in_t":  in_t,
                        "out_t": out_t,
                        "error": err,
                    })
                    holder["chunk_done"] += 1
        holder["done"] = True
    except Exception as e:
        holder["error"] = str(e)
        holder["done"]  = True


# ══════════════════════════════════════════════════════════════════════════════
# RENDER
# ══════════════════════════════════════════════════════════════════════════════
def _resolve_translation(block: dict, translations: dict) -> str | None:
    """Resolve translation for a block, honoring _toc_mirror inheritance."""
    mirror_id = block.get("_toc_mirror")
    if mirror_id and translations.get(mirror_id):
        return translations[mirror_id]
    return translations.get(block["id"])


def apply_translations(original_bytes: bytes, blocks: list[dict],
                       translations: dict) -> bytes:
    """
    Áp dụng translations vào DOCX gốc → trả về bytes DOCX mới.
    Block nào có translation trong dict thì apply, không phân biệt role —
    để H/F translation cũng được apply khi user bấm nút "Dịch H/F".
    """
    doc         = Document(io.BytesIO(original_bytes))
    idx_to_para = {idx: para for idx, para in enumerate(iter_all_paragraphs(doc))}

    for block in blocks:
        if block.get("role") == "image_alt":
            continue
        tr   = _resolve_translation(block, translations)
        para = idx_to_para.get(block.get("para_idx"))
        if not tr or para is None:
            continue
        try:
            if block.get("has_format"):
                replace_paragraph_with_tagged(para, tr)
            else:
                replace_paragraph_text_keep_format(para, tr)
        except Exception:
            pass

    # Image alt-texts: re-iterate in same order as extract; match to blocks by sequence.
    alt_iter   = list(_iter_image_alt_texts(doc))
    alt_blocks = [b for b in blocks if b.get("role") == "image_alt"]
    for (el, attr, _), block in zip(alt_iter, alt_blocks):
        tr = _resolve_translation(block, translations)
        if tr:
            try:
                el.set(attr, tr)
            except Exception:
                pass

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def validate_docx_output(docx_bytes: bytes) -> dict:
    """
    Validate translated DOCX output. Returns:
    {
      "valid": bool,
      "block_count": int,
      "image_count": int,
      "warnings": [str, ...],
      "errors": [str, ...],
    }
    """
    import zipfile
    result = {"valid": True, "block_count": 0, "image_count": 0,
              "warnings": [], "errors": []}
    try:
        # 1. ZIP integrity
        with zipfile.ZipFile(io.BytesIO(docx_bytes)) as zf:
            bad = zf.testzip()
            if bad:
                result["errors"].append(f"ZIP corrupt: {bad}")
                result["valid"] = False
                return result
            names = zf.namelist()
            if "word/document.xml" not in names:
                result["errors"].append("Missing word/document.xml")
                result["valid"] = False
                return result
            result["image_count"] = sum(1 for n in names if n.startswith("word/media/"))
        # 2. python-docx opens cleanly
        doc = Document(io.BytesIO(docx_bytes))
        result["block_count"] = sum(1 for _ in doc.element.iter(qn("w:p")))
        # 3. XML well-formed for all .xml parts
        from lxml import etree
        with zipfile.ZipFile(io.BytesIO(docx_bytes)) as zf:
            for name in zf.namelist():
                if name.endswith(".xml") or name.endswith(".rels"):
                    try:
                        etree.fromstring(zf.read(name))
                    except etree.XMLSyntaxError as e:
                        result["errors"].append(f"{name}: {str(e)[:120]}")
                        result["valid"] = False
        # 4. Sanity checks
        if result["block_count"] == 0:
            result["warnings"].append("Document has 0 paragraphs (suspicious)")
    except Exception as e:
        result["valid"] = False
        result["errors"].append(f"Validation crashed: {str(e)[:200]}")
    return result


def _is_untranslated(b: dict, translations: dict) -> bool:
    tr = translations.get(b["id"], "")
    return (not tr) or tr == b["text"]


# ══════════════════════════════════════════════════════════════════════════════
# TRANSLATION MEMORY — hash-based cache (session-scoped)
# ══════════════════════════════════════════════════════════════════════════════
def tm_key(text: str, target_lang: str, role: str = "") -> str:
    """
    Hash key: language + (optional) role + text.
    Role-aware để giữ style consistency — heading vs body cùng text vẫn dịch riêng.
    Role rỗng = legacy key (backward compat khi caller không có role).
    """
    role_part = f"|{role}" if role else ""
    return hashlib.md5(
        f"{target_lang}{role_part}|{text}".encode("utf-8")
    ).hexdigest()[:16]


def tm_lookup(blocks: list[dict], tm: dict, target_lang: str
              ) -> tuple[dict, list[dict]]:
    """
    Trả về (cached_translations, remaining_blocks).
    Thử key role-specific trước; nếu miss thì fall back legacy key (no role)
    để dùng entry cũ trong session.
    """
    cached, remaining = {}, []
    for b in blocks:
        role = b.get("role", "")
        key_role   = tm_key(b["text"], target_lang, role)
        key_legacy = tm_key(b["text"], target_lang)
        if key_role in tm:
            cached[b["id"]] = tm[key_role]
        elif key_legacy in tm:
            cached[b["id"]] = tm[key_legacy]
        else:
            remaining.append(b)
    return cached, remaining


def tm_store(blocks: list[dict], translations: dict, tm: dict,
             target_lang: str) -> int:
    """Ghi translations mới vào TM với role-specific key. Trả về số entry thêm vào."""
    added = 0
    for b in blocks:
        tr = translations.get(b["id"])
        if not tr or tr == b["text"]:
            continue
        key = tm_key(b["text"], target_lang, b.get("role", ""))
        if key not in tm:
            tm[key] = tr
            added += 1
    return added


# ══════════════════════════════════════════════════════════════════════════════
# CHECKPOINT — persist partial translations across browser refresh
# ══════════════════════════════════════════════════════════════════════════════
import os, pickle as _pickle


def _checkpoint_path(docx_bytes: bytes, target_lang: str) -> str:
    # Hash TOÀN BỘ file để tránh collision khi 2 docx khác nhau cùng prefix 8KB.
    h = hashlib.md5(docx_bytes).hexdigest()[:16]
    slug = target_lang.replace(" ", "_")[:12]
    return f"/tmp/tr_ckpt_{h}_{slug}.pkl"


def checkpoint_save(docx_bytes: bytes, target_lang: str, translations: dict) -> None:
    try:
        with open(_checkpoint_path(docx_bytes, target_lang), "wb") as f:
            _pickle.dump(translations, f)
    except Exception:
        pass


def checkpoint_load(docx_bytes: bytes, target_lang: str) -> dict | None:
    path = _checkpoint_path(docx_bytes, target_lang)
    try:
        if os.path.exists(path):
            with open(path, "rb") as f:
                return _pickle.load(f)
    except Exception:
        pass
    return None


def checkpoint_clear(docx_bytes: bytes, target_lang: str) -> None:
    try:
        os.unlink(_checkpoint_path(docx_bytes, target_lang))
    except Exception:
        pass


def export_bilingual_docx(original_bytes: bytes, blocks: list[dict],
                          translations: dict) -> bytes:
    """
    Tạo DOCX so sánh song ngữ: bảng 2 cột (gốc | dịch).
    Chỉ bao gồm block có translation.
    """
    from docx.shared import Inches, Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    doc = Document()
    # Narrow margins to fit 2 columns
    for section in doc.sections:
        section.left_margin  = Inches(0.6)
        section.right_margin = Inches(0.6)

    table = doc.add_table(rows=1, cols=2)
    table.style = "Table Grid"
    hdr = table.rows[0].cells
    for cell, label in zip(hdr, ["Original", "Translation"]):
        run = cell.paragraphs[0].add_run(label)
        run.bold = True
        run.font.size = Pt(10)

    for b in blocks:
        tr = translations.get(b["id"])
        if not tr or tr == b["text"]:
            continue
        row = table.add_row().cells
        row[0].text = b["text"]
        row[1].text = tr
        for cell in row:
            cell.paragraphs[0].runs[0].font.size = Pt(9)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


_NUM_PAT = _re.compile(r'\b\d[\d,\.]*\b')


def quality_check(blocks: list[dict], translations: dict) -> list[dict]:
    """
    Returns list of {"id", "text", "translation", "issues"} for suspicious entries.
    Checks:
    - Numbers present in original but missing in translation
    - Translation is more than 3.5x longer or less than 0.2x shorter than original
    - Original ends with ? but translation does not
    """
    results = []
    for b in blocks:
        tr = translations.get(b["id"])
        if not tr or not b["text"].strip():
            continue
        issues = []
        orig_nums = set(_NUM_PAT.findall(b["text"]))
        tr_nums   = set(_NUM_PAT.findall(tr))
        missing   = orig_nums - tr_nums
        if missing:
            issues.append(f"Số bị mất: {', '.join(sorted(missing)[:4])}")
        ratio = len(tr) / max(len(b["text"]), 1)
        if ratio > 3.5:
            issues.append(f"Bản dịch quá dài ({ratio:.1f}x)")
        if ratio < 0.2 and len(b["text"]) > 20:
            issues.append(f"Bản dịch quá ngắn ({ratio:.1f}x)")
        orig_q = b["text"].rstrip().endswith("?")
        tr_q   = tr.rstrip().endswith("?")
        if orig_q and not tr_q:
            issues.append("Mất dấu '?'")
        if issues:
            results.append({"id": b["id"], "text": b["text"],
                            "translation": tr, "issues": issues})
    return results


def find_missed(blocks: list[dict], translations: dict,
                hf_only: bool = False) -> list[dict]:
    """
    Trả về list block chưa dịch (translation rỗng hoặc bằng text gốc).

    - hf_only=False (default): chỉ body — bỏ qua H/F, body_repeated
    - hf_only=True            : chỉ H/F + body_repeated, dùng cho nút "Dịch H/F"
    """
    if hf_only:
        in_set = lambda r: r in NO_TRANSLATE_ROLES
    else:
        in_set = lambda r: r not in NO_TRANSLATE_ROLES
    return [b for b in blocks
            if in_set(b["role"]) and _is_untranslated(b, translations)]


# ══════════════════════════════════════════════════════════════════════════════
# IMAGE OCR + TRANSLATE (Gemini Vision)
# ══════════════════════════════════════════════════════════════════════════════
def extract_images_from_docx(docx_bytes: bytes) -> list[dict]:
    """
    Returns list of {id, filename, content_type, data: bytes} for every image
    in the DOCX (word/media/* parts).
    """
    import zipfile
    images = []
    with zipfile.ZipFile(io.BytesIO(docx_bytes)) as zf:
        for name in zf.namelist():
            if not name.startswith("word/media/") or "." not in name:
                continue
            ext = name.rsplit(".", 1)[-1].lower()
            if ext in ("png", "jpg", "jpeg", "gif", "bmp", "webp"):
                mime_ext = "jpeg" if ext == "jpg" else ext
                images.append({
                    "id":           f"IMG_{len(images)}",
                    "filename":     name.rsplit("/", 1)[-1],
                    "content_type": f"image/{mime_ext}",
                    "data":         zf.read(name),
                })
    return images


def _ocr_single_image(client, img: dict, target_lang: str, model: str,
                      retries: int = CHUNK_RETRIES) -> dict:
    """
    Call Gemini Vision for one image with retry + exponential backoff.
    Returns {"ocr": str, "translation": str, "has_text": bool, "error"?: str}.
    """
    from google.genai import types as gtypes

    prompt = (
        f"You are a professional OCR and translation expert. Analyze this image carefully.\n\n"
        f"TASK:\n"
        f"1. Determine if the image contains any readable text "
        f"(printed, handwritten, tables, charts, or diagrams with labels).\n"
        f"2. If text exists:\n"
        f"   a. Extract ALL text verbatim — preserve structure using \\n for line breaks "
        f"and | to separate table columns.\n"
        f"   b. Keep numbers, units, punctuation, and special characters exactly as shown.\n"
        f"   c. Translate the extracted text to {target_lang}, maintaining the same structure.\n"
        f"   d. For mixed-language content, translate only the non-{target_lang} portions.\n\n"
        f"Respond ONLY as JSON (no extra keys):\n"
        f"  if text found : {{\"has_text\": true,  \"ocr\": \"...\", \"translation\": \"...\"}}\n"
        f"  if no text    : {{\"has_text\": false, \"ocr\": \"\",    \"translation\": \"\"}}"
    )

    last_err = None
    for attempt in range(retries):
        try:
            response = client.models.generate_content(
                model=model,
                contents=[
                    gtypes.Part.from_bytes(data=img["data"], mime_type=img["content_type"]),
                    prompt,
                ],
                config=gtypes.GenerateContentConfig(
                    response_mime_type="application/json",
                    temperature=0.1,
                ),
            )
            data = parse_json_loose(response.text) or {}
            if not isinstance(data, dict):
                data = {}
            return {
                "ocr":         str(data.get("ocr", "") or "").strip(),
                "translation": str(data.get("translation", "") or "").strip(),
                "has_text":    bool(data.get("has_text", False)),
            }
        except Exception as e:
            last_err = str(e)
            if attempt < retries - 1:
                time.sleep(2 ** attempt)

    return {"ocr": "", "translation": "", "has_text": False, "error": last_err}


def ocr_and_translate_images(
    client,
    images: list[dict],
    target_lang: str,
    progress_callback=None,
    max_workers: int = 4,
) -> dict:
    """
    OCR + translate all images in parallel using ThreadPoolExecutor.
    Skips images < 5 KB (icons/decorations).
    progress_callback(done, total) called after each image completes.
    Returns {image_id: {"ocr", "translation", "has_text", "error"?}}.
    """
    results = {}
    model = get_working_model() or WORD_MODELS[0]

    to_process = []
    for img in images:
        if len(img["data"]) < 5000:
            results[img["id"]] = {"ocr": "", "translation": "", "has_text": False}
        else:
            to_process.append(img)

    if not to_process:
        return results

    done_count = 0
    total = len(images)
    lock = threading.Lock()

    with ThreadPoolExecutor(max_workers=max_workers) as ex:
        futures = {
            ex.submit(_ocr_single_image, client, img, target_lang, model): img
            for img in to_process
        }
        for future in as_completed(futures):
            img = futures[future]
            try:
                results[img["id"]] = future.result()
            except Exception as e:
                results[img["id"]] = {
                    "ocr": "", "translation": "", "has_text": False,
                    "error": str(e),
                }
            with lock:
                done_count += 1
                current = done_count
            if progress_callback:
                progress_callback(current, total)

    return results


def insert_ocr_captions_into_docx(
    docx_bytes: bytes,
    images: list[dict],
    ocr_results: dict,
) -> bytes:
    """
    Insert OCR translation as a styled caption paragraph after each image paragraph.
    Only inserts for images where has_text=True and translation is non-empty.
    Returns modified DOCX bytes (unchanged if no captions to insert).
    """
    import zipfile
    from lxml import etree

    fname_to_translation: dict[str, str] = {}
    for img in images:
        r = ocr_results.get(img["id"], {})
        if r.get("has_text") and r.get("translation"):
            fname_to_translation[img["filename"]] = r["translation"]

    if not fname_to_translation:
        return docx_bytes

    # rId → filename from word/_rels/document.xml.rels
    with zipfile.ZipFile(io.BytesIO(docx_bytes)) as zf:
        rels_xml = zf.read("word/_rels/document.xml.rels")

    rels_root = etree.fromstring(rels_xml)
    rid_to_caption: dict[str, str] = {}
    for rel in rels_root:
        target = rel.get("Target", "")
        rid = rel.get("Id", "")
        if "media/" in target:
            fname = target.split("/")[-1]
            if fname in fname_to_translation:
                rid_to_caption[rid] = fname_to_translation[fname]

    if not rid_to_caption:
        return docx_bytes

    doc = Document(io.BytesIO(docx_bytes))
    R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    XML_NS = "http://www.w3.org/XML/1998/namespace"

    # Collect (paragraph_element, caption_text) pairs
    hits: list[tuple] = []
    for para in iter_all_paragraphs(doc):
        for elem in para._element.iter():
            embed = elem.get(f"{{{R_NS}}}embed")
            if embed and embed in rid_to_caption:
                hits.append((para._element, rid_to_caption[embed]))
                break

    # Insert captions in document order (reverse to preserve indices)
    for para_el, caption_text in reversed(hits):
        parent = para_el.getparent()
        idx = list(parent).index(para_el)

        p = etree.Element(f"{{{W_NS}}}p")

        pPr = etree.SubElement(p, f"{{{W_NS}}}pPr")
        jc = etree.SubElement(pPr, f"{{{W_NS}}}jc")
        jc.set(f"{{{W_NS}}}val", "center")

        run = etree.SubElement(p, f"{{{W_NS}}}r")
        rPr = etree.SubElement(run, f"{{{W_NS}}}rPr")
        etree.SubElement(rPr, f"{{{W_NS}}}i")
        etree.SubElement(rPr, f"{{{W_NS}}}iCs")
        color = etree.SubElement(rPr, f"{{{W_NS}}}color")
        color.set(f"{{{W_NS}}}val", "808080")
        sz = etree.SubElement(rPr, f"{{{W_NS}}}sz")
        sz.set(f"{{{W_NS}}}val", "18")
        szCs = etree.SubElement(rPr, f"{{{W_NS}}}szCs")
        szCs.set(f"{{{W_NS}}}val", "18")

        t = etree.SubElement(run, f"{{{W_NS}}}t")
        t.text = f"[OCR] {caption_text}"
        t.set(f"{{{XML_NS}}}space", "preserve")

        parent.insert(idx + 1, p)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ══════════════════════════════════════════════════════════════════════════════
# GLOSSARY SUGGEST (LLM-based, domain-aware)
# ══════════════════════════════════════════════════════════════════════════════
def suggest_glossary(client, blocks: list[dict], target_lang: str,
                     max_terms: int = 20) -> list[dict]:
    """
    Sample body text from blocks, ask LLM for domain-specific terms
    that should be in a glossary. Returns list of {"term", "suggested", "note"}.
    """
    from google.genai import types as gtypes
    body_text = "\n".join(
        b["text"] for b in blocks
        if b.get("role") in ("paragraph", "section_heading", "bullet")
    )
    body_text = body_text[:6000]
    if not body_text.strip():
        return []
    prompt = (
        f"Analyze the following document text. Identify up to {max_terms} domain-specific "
        f"terms, technical jargon, brand names, or recurring concepts that should be "
        f"translated consistently to {target_lang}.\n\n"
        f"For each term, propose a translation and a brief note explaining why it matters.\n\n"
        f"Return JSON: {{\"terms\": [{{\"term\": \"...\", \"suggested\": \"...\", \"note\": \"...\"}}]}}\n\n"
        f"Document text:\n{body_text}"
    )
    model = get_working_model() or WORD_MODELS[0]
    try:
        response = client.models.generate_content(
            model=model,
            contents=prompt,
            config=gtypes.GenerateContentConfig(
                response_mime_type="application/json",
                temperature=0.3,
            ),
        )
        data = parse_json_loose(response.text) or {}
        if not isinstance(data, dict):
            return []
        terms = data.get("terms", [])
        if not isinstance(terms, list):
            return []
        return [t for t in terms if isinstance(t, dict) and t.get("term")][:max_terms]
    except Exception:
        return []
