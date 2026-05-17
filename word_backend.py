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
import io
import json
import time
from collections import Counter
from concurrent.futures import ThreadPoolExecutor, as_completed

from docx import Document

from config import (
    WORD_MODELS, MAX_WORD_TOKENS, NO_TRANSLATE_ROLES,
    CHUNK_RETRIES, TARGET_CHUNK_CHARS, MIN_CHUNK_BLOCKS, MAX_CHUNK_BLOCKS,
    HF_REPEAT_THRESHOLD, HF_REPEAT_MIN_CHARS,
)
from gemini import generate, usage_tokens, parse_json_loose


# Module-level cache cho model đang work (persist giữa các Streamlit rerun + thread)
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


def extract_docx_blocks(docx_bytes: bytes) -> list[dict]:
    """
    Đọc DOCX → list block {id, text, role, para_idx}.

    Role:
    - header / footer     : nằm trong section.header / section.footer
    - body_repeated       : text lặp lại ≥ 3 lần trong body (heuristic H/F)
    - section_heading / title / bullet / table_cell / note / toc / paragraph
    """
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

    _mark_repeating_as_hf(blocks)
    return blocks


def count_by_role(blocks: list[dict]) -> dict:
    """Stat helper: đếm số block theo từng role + tổng theo nhóm."""
    counter = Counter(b["role"] for b in blocks)
    return {
        "total":       len(blocks),
        "body":        sum(c for r, c in counter.items() if r not in NO_TRANSLATE_ROLES),
        "header":      counter.get("header", 0),
        "footer":      counter.get("footer", 0),
        "body_repeated": counter.get("body_repeated", 0),
        "hf_total":    sum(counter.get(r, 0) for r in NO_TRANSLATE_ROLES),
        "by_role":     dict(counter),
    }


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
    Raise RuntimeError nếu tất cả model đều fail.
    """
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
            in_t, out_t = usage_tokens(resp)
            _working_model[0] = model
            return text, in_t, out_t
        except Exception as e:
            last_err = str(e)
    raise RuntimeError(f"All Word models failed: {last_err}")


def get_working_model() -> str | None:
    return _working_model[0]


# ══════════════════════════════════════════════════════════════════════════════
# TRANSLATE chunk (raw + retry)
# ══════════════════════════════════════════════════════════════════════════════
def _build_chunk_prompt(chunk: list[dict], target_lang: str, doc_context: str) -> str:
    payload = json.dumps(
        [{"id": b["id"], "text": b["text"], "role": b.get("role", "paragraph")}
         for b in chunk],
        ensure_ascii=False,
    )
    return f"""Translate these Word document blocks into {target_lang}.

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
- For header/footer/body_repeated: concise like a page header/footer; preserve page numbers ("Page 1 of 10"), dates, document IDs verbatim.

Format:
[{{"id":"...","text":"translated text"}}]

Blocks:
{payload}"""


def translate_chunk(client, chunk: list[dict], target_lang: str,
                    doc_context: str) -> tuple[dict, int, int]:
    """
    Dịch 1 chunk → (translations_dict, in_tokens, out_tokens).
    RAISE exception nếu API fail. Block bị model bỏ sót → fallback giữ text gốc.
    """
    prompt = _build_chunk_prompt(chunk, target_lang, doc_context)
    raw, in_t, out_t = call_gemini(client, prompt)
    parsed = parse_json_loose(raw) or []
    if not isinstance(parsed, list):
        parsed = []
    result = {}
    for item in parsed:
        if isinstance(item, dict) and item.get("id") and item.get("text") is not None:
            result[item["id"]] = str(item["text"])
    for b in chunk:
        result.setdefault(b["id"], b["text"])
    return result, in_t, out_t


def translate_chunk_with_retry(client, chunk: list[dict], target_lang: str,
                               doc_context: str,
                               retries: int = CHUNK_RETRIES
                               ) -> tuple[dict, int, int, str | None]:
    """
    Wrapper với exponential backoff. Trả về (translations, in_t, out_t, error).
    Nếu hết retries → fallback giữ text gốc + error message.
    """
    last_err = None
    for attempt in range(retries):
        try:
            t, in_t, out_t = translate_chunk(client, chunk, target_lang, doc_context)
            return t, in_t, out_t, None
        except Exception as e:
            last_err = str(e)
            if attempt < retries - 1:
                time.sleep(2 ** attempt)   # 1s, 2s, 4s
    return {b["id"]: b["text"] for b in chunk}, 0, 0, last_err


# ══════════════════════════════════════════════════════════════════════════════
# PARALLEL TRANSLATION — entry point cho background thread
# ══════════════════════════════════════════════════════════════════════════════
def translate_parallel(holder: dict, client,
                       chunks: list[list[dict]],
                       target_lang: str, doc_context: str,
                       max_workers: int):
    """
    Background worker: dịch nhiều chunk song song với ThreadPoolExecutor.

    `holder` (dict shared với main thread) sẽ được update:
    - translations  : dict {block_id -> translated_text}
    - tok_in/tok_out: cumulative tokens
    - chunk_done    : số chunk đã xong
    - chunk_log     : list per-chunk result (đẩy thêm khi mỗi chunk xong)
    - done          : True khi xong tất cả
    - error         : str nếu fatal error
    """
    try:
        with ThreadPoolExecutor(max_workers=max_workers) as ex:
            futures = {
                ex.submit(translate_chunk_with_retry,
                          client, c, target_lang, doc_context): (i, c)
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


def _is_untranslated(b: dict, translations: dict) -> bool:
    tr = translations.get(b["id"], "")
    return (not tr) or tr == b["text"]


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
