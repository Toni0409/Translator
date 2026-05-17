"""
Word backend: extract DOCX paragraphs, gọi Gemini (có fallback chain),
ghi lại DOCX giữ nguyên format runs.

3 nhóm logic:
1. Extract  : `extract_docx_blocks`, `iter_all_paragraphs`, `_detect_role`
2. Translate: `translate_chunk`, `call_gemini` (model fallback)
3. Render   : `apply_translations`, `replace_paragraph_text_keep_format`
"""
import io
import json

from docx import Document

from config import WORD_MODELS, MAX_WORD_TOKENS, NO_TRANSLATE_ROLES
from gemini import generate, parse_json_loose


# Module-level cache cho model đang work (persist giữa các Streamlit rerun)
_working_model: list = [None]


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


def iter_all_paragraphs(doc):
    """Lặp qua toàn bộ paragraph: body → tables → headers → footers."""
    for para in doc.paragraphs:
        yield para
    for table in doc.tables:
        yield from _iter_paragraphs_in_table(table)
    for section in doc.sections:
        try:
            for para in section.header.paragraphs:
                yield para
            for tbl in section.header.tables:
                yield from _iter_paragraphs_in_table(tbl)
        except Exception:
            pass
        try:
            for para in section.footer.paragraphs:
                yield para
            for tbl in section.footer.tables:
                yield from _iter_paragraphs_in_table(tbl)
        except Exception:
            pass


# ══════════════════════════════════════════════════════════════════════════════
# FORMAT-SAFE TEXT REPLACEMENT
# ══════════════════════════════════════════════════════════════════════════════
def replace_paragraph_text_keep_format(paragraph, new_text: str):
    """Thay text trong paragraph mà GIỮ ĐƯỢC format của run đầu tiên."""
    runs = paragraph.runs
    if not runs:
        paragraph.add_run(new_text)
        return
    first_run = next((r for r in runs if r.text), None)
    if first_run is None:
        runs[0].text = new_text
        return
    first_run.text = new_text
    after = False
    for run in runs:
        if run is first_run:
            after = True
            continue
        if after and run.text:
            run.text = ""


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
# EXTRACT
# ══════════════════════════════════════════════════════════════════════════════
def extract_docx_blocks(docx_bytes: bytes) -> list[dict]:
    """Đọc DOCX → list block {id, text, role, para_idx}."""
    doc = Document(io.BytesIO(docx_bytes))

    table_para_set:  set = set()
    header_para_set: set = set()
    footer_para_set: set = set()

    for tbl in doc.tables:
        for para in _iter_paragraphs_in_table(tbl):
            table_para_set.add(id(para._element))

    for section in doc.sections:
        try:
            for para in section.header.paragraphs:
                header_para_set.add(id(para._element))
            for tbl in section.header.tables:
                for para in _iter_paragraphs_in_table(tbl):
                    header_para_set.add(id(para._element))
        except Exception:
            pass
        try:
            for para in section.footer.paragraphs:
                footer_para_set.add(id(para._element))
            for tbl in section.footer.tables:
                for para in _iter_paragraphs_in_table(tbl):
                    footer_para_set.add(id(para._element))
        except Exception:
            pass

    blocks = []
    for idx, para in enumerate(iter_all_paragraphs(doc)):
        text = para.text.strip()
        if not text:
            continue
        pid  = id(para._element)
        role = _detect_role(
            para,
            pid in table_para_set,
            pid in header_para_set,
            pid in footer_para_set,
        )
        blocks.append({"id": f"p{idx}", "text": text, "role": role, "para_idx": idx})
    return blocks


def build_doc_context(blocks: list[dict]) -> str:
    """Tóm tắt cấu trúc tài liệu (title + headings + TOC + domain hint)."""
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
# TRANSLATE
# ══════════════════════════════════════════════════════════════════════════════
def call_gemini(client, prompt: str) -> str:
    """Gọi Gemini với fallback chain. Cache model đang work để skip retry."""
    models = ([_working_model[0]] if _working_model[0] else []) + \
             [m for m in WORD_MODELS if m != _working_model[0]]
    last_err = "no models"
    for model in models:
        try:
            resp = generate(client, model, prompt,
                            max_output_tokens=MAX_WORD_TOKENS, temperature=0.1)
            text = (resp.text or "").strip()
            if not text:
                last_err = "empty response"
                continue
            _working_model[0] = model
            return text
        except Exception as e:
            last_err = str(e)
    raise RuntimeError(f"All Word models failed: {last_err}")


def get_working_model() -> str | None:
    return _working_model[0]


def translate_chunk(client, chunk: list[dict], target_lang: str, doc_context: str) -> dict:
    """Dịch 1 chunk (list block) → dict[block_id -> translated_text]."""
    payload = json.dumps(
        [{"id": b["id"], "text": b["text"], "role": b.get("role", "paragraph")}
         for b in chunk],
        ensure_ascii=False,
    )
    prompt = f"""Translate these Word document blocks into {target_lang}.

Document context:
{doc_context}

Rules:
- Return ONLY valid JSON array. No markdown, no explanation.
- Keep "id" exactly as given.
- Translate ALL words into {target_lang}. Leave nothing in the source language.
- Preserve: numbers, units, product codes, model names, document IDs, revision codes.
- Do NOT translate company names (e.g. INVENTIO AG, Schindler, Otis).
- Preserve ALL punctuation, whitespace, tabs (\\t), line breaks (\\n).
- Match source sentence count — do NOT merge or split sentences.
- For toc: translate heading text only; preserve numbering, dot leaders (...), page numbers.
- For bullet: concise imperative, keep bullet symbol.
- For heading/section_heading: concise technical, match source brevity.
- For table_cell: concise, consistent across row.
- For note: keep WARNING/NOTE/CAUTION label verbatim.

Format:
[{{"id":"...","text":"translated text"}}]

Blocks:
{payload}"""
    try:
        raw    = call_gemini(client, prompt)
        parsed = parse_json_loose(raw) or []
        if not isinstance(parsed, list):
            parsed = []
        result = {}
        for item in parsed:
            if isinstance(item, dict) and item.get("id") and item.get("text") is not None:
                result[item["id"]] = str(item["text"])
        for b in chunk:
            result.setdefault(b["id"], b["text"])  # fallback giữ text gốc
        return result
    except Exception:
        return {b["id"]: b["text"] for b in chunk}


# ══════════════════════════════════════════════════════════════════════════════
# RENDER
# ══════════════════════════════════════════════════════════════════════════════
def apply_translations(original_bytes: bytes, blocks: list[dict],
                       translations: dict) -> bytes:
    """Áp dụng translations vào DOCX gốc → trả về bytes DOCX mới."""
    doc         = Document(io.BytesIO(original_bytes))
    idx_to_para = {idx: para for idx, para in enumerate(iter_all_paragraphs(doc))}

    for block in blocks:
        if block.get("role") in NO_TRANSLATE_ROLES:
            continue
        tr   = translations.get(block["id"])
        para = idx_to_para.get(block.get("para_idx"))
        if not tr or para is None:
            continue
        try:
            replace_paragraph_text_keep_format(para, tr)
        except Exception:
            pass

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()
