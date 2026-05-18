"""
PDF backend: extract text spans, gọi Gemini, ghi lại PDF giữ nguyên layout.

3 nhóm logic chính:
1. Extract  : `extract_line_groups`, `parse_page_range`, `extract_pdf_images`,
              `build_pdf_glossary`
2. Translate: `translate_page` (sequential + live UI), `translate_page_pure`,
              `translate_pages_parallel`, `ocr_pdf_images`
3. Render   : `write_translated_pdf` (with text wrap), `insert_ocr_captions_into_pdf`
"""
import hashlib
import io
import os
import pickle
import re
import threading
import time
from collections import Counter
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime

import fitz

from config import (
    PDF_MODEL, MAX_RETRIES, RETRY_CODES,
    UNICODE_FONTS, BOLD_FONT_PAIRS,
)
from gemini import generate, usage_tokens, parse_json_loose, call_in_thread
from ui_common import timer_box_html


# ══════════════════════════════════════════════════════════════════════════════
# FONT
# ══════════════════════════════════════════════════════════════════════════════
def find_font() -> str | None:
    for p in UNICODE_FONTS:
        if os.path.isfile(p):
            return p
    return None


def get_bold_font_path(font_path: str | None) -> str | None:
    if not font_path:
        return None
    for regular, bold in BOLD_FONT_PAIRS:
        candidate = font_path.replace(regular, bold)
        if os.path.isfile(candidate):
            return candidate
    return None


def _int_to_rgb(c: int) -> tuple[float, float, float]:
    return ((c >> 16) & 0xFF) / 255, ((c >> 8) & 0xFF) / 255, (c & 0xFF) / 255


# ══════════════════════════════════════════════════════════════════════════════
# TABLE DETECTION — qua PyMuPDF page.find_tables()
# ══════════════════════════════════════════════════════════════════════════════
def detect_tables_on_page(page) -> list[dict]:
    """
    Phát hiện bảng trên page bằng `page.find_tables()`.
    Trả về list of dicts:
        [{'bbox': (x0,y0,x1,y1),
          'cells': [{'bbox': (...), 'row': r, 'col': c}, ...],
          'rows': N, 'cols': M}]
    Trả [] nếu PyMuPDF version cũ hoặc không có bảng.
    """
    try:
        finder = page.find_tables()
    except Exception:
        return []

    try:
        table_list = list(finder)
    except Exception:
        table_list = []

    out = []
    for tab in table_list:
        try:
            tbbox = tuple(tab.bbox)
        except Exception:
            continue
        cells: list[dict] = []
        rows = cols = 0
        # Newer PyMuPDF: tab.rows[i].cells[j] = bbox tuple
        try:
            for r_idx, row in enumerate(tab.rows):
                for c_idx, cbb in enumerate(row.cells):
                    if cbb:
                        cells.append({"bbox": tuple(cbb), "row": r_idx, "col": c_idx})
                rows = max(rows, r_idx + 1)
                cols = max(cols, len(row.cells))
        except (AttributeError, TypeError):
            # Fallback: tab.cells = flat list of bbox tuples
            try:
                for cbb in tab.cells:
                    if cbb:
                        cells.append({"bbox": tuple(cbb), "row": -1, "col": -1})
            except Exception:
                pass
        if cells:
            out.append({"bbox": tbbox, "cells": cells, "rows": rows, "cols": cols})
    return out


def find_cell_info(line_bbox: tuple, tables: list[dict]) -> dict | None:
    """
    Map 1 line bbox → cell info nó nằm trong (nếu có).
    Trả về {'table': N, 'row': R, 'col': C} (1-based) hoặc None.
    Dùng tâm dòng để check, tolerance với bbox overlap.
    """
    cx = (line_bbox[0] + line_bbox[2]) / 2
    cy = (line_bbox[1] + line_bbox[3]) / 2
    for ti, table in enumerate(tables):
        tx0, ty0, tx1, ty1 = table["bbox"]
        if not (tx0 <= cx <= tx1 and ty0 <= cy <= ty1):
            continue
        for cell in table["cells"]:
            cx0, cy0, cx1, cy1 = cell["bbox"]
            if cx0 <= cx <= cx1 and cy0 <= cy <= cy1:
                return {
                    "table": ti + 1,
                    "row":   cell["row"] + 1 if cell["row"] >= 0 else 0,
                    "col":   cell["col"] + 1 if cell["col"] >= 0 else 0,
                }
        # Trong table nhưng không khớp cell → vẫn đánh dấu là cell mơ hồ
        return {"table": ti + 1, "row": 0, "col": 0}
    return None


# ══════════════════════════════════════════════════════════════════════════════
# EXTRACT
# ══════════════════════════════════════════════════════════════════════════════
def extract_line_groups(pdf_path: str, page_nums: list[int] | None = None):
    """
    Trích các "line group" từ PDF — mỗi group là 1 dòng text với bbox, font size,
    color, bold flag, và optional `cell` info {table, row, col}.

    Trả về (groups_per_page, total_pages, table_stats):
      - groups_per_page : dict[page_idx -> list[group]]
      - total_pages     : int
      - table_stats     : dict[page_idx -> {'tables': N, 'cell_lines': M}]
    """
    result:      dict = {}
    table_stats: dict = {}
    doc          = fitz.open(pdf_path)
    total        = len(doc)
    targets      = page_nums if page_nums else list(range(total))

    for pi in targets:
        if not (0 <= pi < total):
            continue
        page   = doc[pi]
        tables = detect_tables_on_page(page)
        data   = page.get_text(
            "dict",
            flags=fitz.TEXT_PRESERVE_WHITESPACE | fitz.TEXT_MEDIABOX_CLIP,
        )
        groups = []
        cell_lines = 0
        for block in data["blocks"]:
            if block.get("type") != 0:
                continue
            for line in block["lines"]:
                spans = [sp for sp in line["spans"] if sp["text"].strip()]
                if not spans:
                    continue
                merged = "".join(sp["text"] for sp in spans).strip()
                if not merged:
                    continue
                bbox = (
                    min(sp["bbox"][0] for sp in spans),
                    min(sp["bbox"][1] for sp in spans),
                    max(sp["bbox"][2] for sp in spans),
                    max(sp["bbox"][3] for sp in spans),
                )
                cell = find_cell_info(bbox, tables) if tables else None
                if cell:
                    cell_lines += 1
                groups.append({
                    "bbox": bbox,
                    "text": merged,
                    "size": max(sp["size"] for sp in spans),
                    "rgb":  _int_to_rgb(spans[0]["color"]),
                    "bold": any(bool(sp["flags"] & (1 << 4)) for sp in spans),
                    "cell": cell,
                })

        # Sort theo (y0, x0) — reading order tự nhiên: top-to-bottom, left-to-right.
        # Trong table, các cell cùng row có y0 gần nhau → sắp xếp tốt theo cột.
        groups.sort(key=lambda g: (round(g["bbox"][1], 1), g["bbox"][0]))

        result[pi]      = groups
        table_stats[pi] = {"tables": len(tables), "cell_lines": cell_lines}

    doc.close()
    return result, total, table_stats


def parse_page_range(s: str, total: int) -> list[int]:
    """Parse "1-5,8,10-12" thành list page index (0-based)."""
    pages = []
    for part in s.split(","):
        part = part.strip()
        if "-" in part:
            a, b = part.split("-", 1)
            pages.extend(range(int(a) - 1, int(b)))
        elif part.isdigit():
            pages.append(int(part) - 1)
    return sorted(p for p in set(pages) if 0 <= p < total)


def extract_pdf_blocks(pdf_bytes: bytes) -> list[dict]:
    """
    Extract PDF as block list compatible with DOCX schema for review/comparison.
    Each line group → 1 block; role inferred from font size and table membership.
    Output schema matches `extract_docx_blocks`:
      {id, text, text_tagged, has_format, role, para_idx, table_cell}
    """
    import tempfile

    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(pdf_bytes)
        tmp_path = tmp.name

    try:
        groups_per_page, _, _ = extract_line_groups(tmp_path)
    finally:
        try:
            os.unlink(tmp_path)
        except Exception:
            pass

    # Body font size = median of all sizes (cheap heuristic)
    all_sizes: list[float] = []
    for groups in groups_per_page.values():
        all_sizes.extend(g["size"] for g in groups)
    if all_sizes:
        all_sizes.sort()
        body_size = all_sizes[len(all_sizes) // 2]
    else:
        body_size = 10.0

    blocks: list[dict] = []
    para_idx = 0
    for page_idx in sorted(groups_per_page.keys()):
        for g in groups_per_page[page_idx]:
            text = g["text"].strip()
            if not text:
                continue

            cell = g.get("cell")
            if cell:
                role = "table_cell"
                table_cell = (cell["table"], cell["row"], cell["col"])
            elif g["size"] >= body_size * 1.4 and g["bold"]:
                role = "title"
                table_cell = None
            elif g["size"] >= body_size * 1.15 or (g["bold"] and g["size"] >= body_size):
                role = "section_heading"
                table_cell = None
            else:
                role = "paragraph"
                table_cell = None

            blocks.append({
                "id":          f"pdf_p{page_idx}_{para_idx}",
                "text":        text,
                "text_tagged": text,
                "has_format":  False,
                "role":        role,
                "para_idx":    para_idx,
                "table_cell":  table_cell,
                "page":        page_idx,
            })
            para_idx += 1
    return blocks


# ══════════════════════════════════════════════════════════════════════════════
# RENDER
# ══════════════════════════════════════════════════════════════════════════════
def _insert_line(page, font_path, font_bold_path, bbox, text, fontsize, color, bold):
    """
    Insert translated text. Strategy:
    1. Try to fit at original font size by expanding vertical area (wrap to multi-line)
    2. Only scale down font (min 6pt) if multi-line wrap still doesn't fit
    """
    x0, y0, x1, y1 = bbox
    pw, ph         = page.rect.width, page.rect.height
    x1_use         = max(x1, pw - 30)
    line_h         = max(y1 - y0, fontsize * 1.2)

    if font_path:
        fontfile = font_bold_path if bold and font_bold_path else font_path
        fontname = "FBold" if bold else "FReg"
        try:
            page.insert_font(fontname=fontname, fontfile=fontfile)
        except Exception:
            pass
    else:
        fontname = "hebo" if bold else "helv"

    size = max(fontsize, 5.0)

    # Step 1: try multi-line wrap at original size (1, 2, 3, 4, 5 lines)
    for n_lines in (1, 2, 3, 4, 5):
        y1_try = min(y0 + line_h * n_lines, ph - 10)
        rc = page.insert_textbox(
            fitz.Rect(x0, y0, x1_use, y1_try),
            text, fontsize=size, fontname=fontname, color=color, align=0,
        )
        if rc >= 0:
            return

    # Step 2: extra-tall area + scale font down (min 6pt — still readable)
    y1_max = ph - 10
    while size >= 6.0:
        rc = page.insert_textbox(
            fitz.Rect(x0, y0, x1_use, y1_max),
            text, fontsize=size, fontname=fontname, color=color, align=0,
        )
        if rc >= 0:
            return
        size -= 0.5


def write_translated_pdf(src: str, dst: str, all_groups: dict,
                         all_trans: dict, font_path: str | None):
    """Redact text gốc và in lại bản dịch vào cùng vị trí."""
    doc = fitz.open(src)
    font_bold_path = get_bold_font_path(font_path)

    for pi, groups in all_groups.items():
        trans = all_trans.get(pi, [])
        if not groups or not trans:
            continue
        page = doc[pi]
        for g in groups:
            r        = fitz.Rect(g["bbox"])
            expanded = fitz.Rect(r.x0 - 10, r.y0 - 5, r.x1 + 10, r.y1 + 5)
            page.add_redact_annot(expanded.intersect(page.rect))
        page.apply_redactions(images=0)
        for g, t in zip(groups, trans):
            text = t.strip() if t and t.strip() else g["text"]
            try:
                _insert_line(page, font_path, font_bold_path,
                             g["bbox"], text, g["size"], g["rgb"], g["bold"])
            except Exception:
                pass

    doc.save(dst, garbage=4, deflate=True)
    doc.close()


# ══════════════════════════════════════════════════════════════════════════════
# TRANSLATE — call Gemini với live timer & retry/backoff
# ══════════════════════════════════════════════════════════════════════════════
def _worker(holder: dict, client, prompt: str):
    try:
        resp = generate(client, PDF_MODEL, prompt, max_output_tokens=65_536, temperature=0.1)
        in_t, out_t = usage_tokens(resp)
        holder["result"] = (resp.text.strip(), in_t, out_t)
    except Exception as e:
        holder["error"] = e


def call_gemini_live(client, prompt: str, timer_ph, t0: float, page_start: float,
                     label: str, log_lines: list, log_ph):
    """
    Gọi Gemini trong thread, main loop update timer mỗi giây.
    Retry với exponential backoff khi gặp rate limit.
    """
    for attempt in range(MAX_RETRIES):
        task = call_in_thread(_worker, client, prompt)
        t, holder = task["thread"], task["holder"]

        dot = 0
        while t.is_alive():
            elapsed      = time.time() - t0
            page_elapsed = time.time() - page_start
            dots         = "." * (dot % 4)
            timer_ph.markdown(
                timer_box_html(
                    elapsed,
                    f"🔄 {label} — chờ Gemini API{dots} ({page_elapsed:.0f}s)",
                ),
                unsafe_allow_html=True,
            )
            dot += 1
            time.sleep(1)

        if "result" in holder:
            return holder["result"]

        err     = holder.get("error", Exception("Unknown error"))
        err_str = str(err).lower()
        is_rate = any(c in err_str for c in RETRY_CODES)

        if is_rate and attempt < MAX_RETRIES - 1:
            wait = (2 ** attempt) * 5
            ts   = datetime.now().strftime("%H:%M:%S")
            log_lines.append(f"[{ts}] ⚠️  Rate limit! Chờ {wait}s rồi thử lại ({attempt+1}/{MAX_RETRIES})...")
            log_ph.markdown(
                f"<div class='log-box'>{'<br>'.join(log_lines[-40:])}</div>",
                unsafe_allow_html=True,
            )
            for remaining in range(wait, 0, -1):
                elapsed = time.time() - t0
                timer_ph.markdown(
                    timer_box_html(
                        elapsed,
                        f"⚠️ Rate limit — thử lại sau {remaining}s",
                        border="#f59e0b", val="#f59e0b", status_color="#f59e0b",
                    ),
                    unsafe_allow_html=True,
                )
                time.sleep(1)
        else:
            raise err


def _format_group_for_prompt(i: int, g: dict) -> str:
    """Format 1 line group cho prompt — thêm (T# R# C#) nếu là cell trong bảng."""
    cell = g.get("cell")
    if cell and cell["row"] > 0 and cell["col"] > 0:
        return f"[{i}] (T{cell['table']} R{cell['row']} C{cell['col']}) {g['text']}"
    elif cell:
        return f"[{i}] (T{cell['table']}) {g['text']}"
    return f"[{i}] {g['text']}"


def _build_page_prompt(groups: list, target_lang: str,
                       glossary: list[str] | None = None) -> str:
    """Build prompt for translating 1 page of line groups."""
    numbered   = "\n".join(_format_group_for_prompt(i, g) for i, g in enumerate(groups))
    has_tables = any(g.get("cell") for g in groups)
    table_hint = (
        "\nMột số dòng có dạng (T# R# C#) — đó là cell trong bảng "
        "(T = số bảng, R = row, C = column).\n"
        "Khi dịch các cell bảng:\n"
        "- Cell cùng cột → dịch consistent, dùng cùng thuật ngữ\n"
        "- Header row (R1) → dịch ngắn gọn, súc tích như tiêu đề cột\n"
        "- Giá trị số / ngày tháng / đơn vị (mm, kg, V, %) → GIỮ NGUYÊN\n"
        "- Mã / serial / model code → GIỮ NGUYÊN\n"
        if has_tables else ""
    )
    glossary_hint = ""
    if glossary:
        terms = ", ".join(glossary[:30])
        glossary_hint = (
            f"\nThuật ngữ lặp lại nhiều lần trong tài liệu — dịch consistent "
            f"(dùng cùng 1 cách dịch xuyên suốt):\n{terms}\n"
        )
    return (
        f"Dịch sang {target_lang}. Giữ nguyên số thứ tự [0]...[{len(groups)-1}].\n"
        f"{table_hint}"
        f"{glossary_hint}"
        f"Trả về JSON object, ĐẦY ĐỦ từ \"0\" đến \"{len(groups)-1}\", không bỏ sót.\n"
        f"KHÔNG kèm prefix (T# R# C#) trong bản dịch — chỉ trả về text đã dịch:\n"
        f"{{\"0\": \"bản dịch\", \"1\": \"bản dịch\", ...}}\n"
        f"Chỉ JSON, không giải thích.\n\n"
        f"{numbered}"
    )


def _parse_page_response(raw: str, groups: list) -> list[str]:
    """Parse Gemini JSON response → list[str] same length as groups."""
    parsed = parse_json_loose(raw)
    if isinstance(parsed, dict):
        return [str(parsed.get(str(i), groups[i]["text"])) for i in range(len(groups))]
    if isinstance(parsed, list) and len(parsed) == len(groups):
        return [str(x) for x in parsed]
    return [g["text"] for g in groups]


def translate_page(client, groups: list, target_lang: str, page_idx: int,
                   timer_ph, t0: float, page_start: float,
                   log_lines: list, log_ph,
                   glossary: list[str] | None = None,
                   tm: dict | None = None):
    """Dịch 1 trang PDF (sequential mode — có live timer UI). Returns (trans, in_t, out_t, tm_hits)."""
    from word_backend import tm_key

    if tm is not None:
        cached = [tm.get(tm_key(g["text"], target_lang)) for g in groups]
    else:
        cached = [None] * len(groups)
    tm_hits = sum(1 for c in cached if c)
    to_translate_idx = [i for i, c in enumerate(cached) if not c]

    if not to_translate_idx:
        return cached, 0, 0, tm_hits

    sub_groups = [groups[i] for i in to_translate_idx]
    prompt = _build_page_prompt(sub_groups, target_lang, glossary)
    label  = f"Đang dịch trang {page_idx + 1}"
    raw, in_t, out_t = call_gemini_live(
        client, prompt, timer_ph, t0, page_start, label, log_lines, log_ph
    )
    new_trans = _parse_page_response(raw, sub_groups)

    final = list(cached)
    for idx, tr in zip(to_translate_idx, new_trans):
        final[idx] = tr

    if tm is not None:
        for g, tr in zip(sub_groups, new_trans):
            if tr and tr != g["text"]:
                tm[tm_key(g["text"], target_lang)] = tr

    return final, in_t, out_t, tm_hits


# ══════════════════════════════════════════════════════════════════════════════
# PARALLEL PAGE TRANSLATION — không có live timer, dùng ThreadPoolExecutor
# ══════════════════════════════════════════════════════════════════════════════
def _call_gemini_with_retry(client, prompt: str) -> tuple[str, int, int]:
    """Sync call with exponential backoff on rate limit. Returns (text, in_tok, out_tok)."""
    last_err = None
    for attempt in range(MAX_RETRIES):
        try:
            resp = generate(client, PDF_MODEL, prompt,
                            max_output_tokens=65_536, temperature=0.1)
            in_t, out_t = usage_tokens(resp)
            return resp.text.strip(), in_t, out_t
        except Exception as e:
            last_err = e
            err_str = str(e).lower()
            is_rate = any(c in err_str for c in RETRY_CODES)
            if is_rate and attempt < MAX_RETRIES - 1:
                time.sleep((2 ** attempt) * 5)
            elif attempt < MAX_RETRIES - 1:
                time.sleep(2 ** attempt)
            else:
                raise last_err
    raise last_err


def translate_page_pure(client, groups: list, target_lang: str,
                        glossary: list[str] | None = None,
                        tm: dict | None = None,
                        fallback_individually: bool = True,
                        custom_glossary: dict[str, str] | None = None,
                        custom_rules: str | None = None,
                        ) -> tuple[list[str], int, int, str | None, int]:
    """
    Pure translate (no UI). Supports:
    - TM lookup: skip lines already cached, only translate new ones
    - Smart fallback: if batch fails, retry each line individually
    - Custom glossary + custom rules: injected into prompt
    Returns (translations, in_tok, out_tok, error_or_none, tm_hits).
    """
    from word_backend import tm_key

    if not groups:
        return [], 0, 0, None, 0

    # TM lookup
    if tm is not None:
        cached_per_line = [tm.get(tm_key(g["text"], target_lang)) for g in groups]
    else:
        cached_per_line = [None] * len(groups)

    tm_hits = sum(1 for c in cached_per_line if c)
    to_translate_idx = [i for i, c in enumerate(cached_per_line) if not c]

    if not to_translate_idx:
        return cached_per_line, 0, 0, None, tm_hits

    sub_groups = [groups[i] for i in to_translate_idx]
    in_t = out_t = 0
    last_err: str | None = None

    try:
        prompt = _build_page_prompt_v2(sub_groups, target_lang, glossary,
                                       custom_glossary, custom_rules)
        raw, in_t, out_t = _call_gemini_with_retry(client, prompt)
        new_trans = _parse_page_response(raw, sub_groups)
    except Exception as e:
        last_err = str(e)
        if not fallback_individually:
            return [g["text"] for g in groups], 0, 0, last_err, tm_hits
        # Smart fallback: try each line individually
        new_trans = []
        for g in sub_groups:
            try:
                single_prompt = _build_page_prompt_v2([g], target_lang, glossary,
                                                      custom_glossary, custom_rules)
                raw_s, in_s, out_s = _call_gemini_with_retry(client, single_prompt)
                parsed = _parse_page_response(raw_s, [g])
                new_trans.append(parsed[0])
                in_t  += in_s
                out_t += out_s
            except Exception:
                new_trans.append(g["text"])

    # Merge cached + new translations back to full list
    final: list[str] = list(cached_per_line)  # type: ignore[arg-type]
    for idx, tr in zip(to_translate_idx, new_trans):
        final[idx] = tr

    # Store new translations into TM
    if tm is not None:
        for g, tr in zip(sub_groups, new_trans):
            if tr and tr != g["text"]:
                tm[tm_key(g["text"], target_lang)] = tr

    return final, in_t, out_t, last_err, tm_hits


def translate_pages_parallel(client, all_groups: dict, target_lang: str,
                             targets: list[int],
                             max_workers: int = 4,
                             glossary: list[str] | None = None,
                             tm: dict | None = None,
                             progress_callback=None,
                             custom_glossary: dict[str, str] | None = None,
                             custom_rules: str | None = None,
                             skip_pages: dict[int, str] | None = None,
                             resume_trans: dict | None = None):
    """
    Translate multiple pages concurrently with ThreadPoolExecutor.
    - skip_pages: {page_idx: reason} for TOC/References pages → copy original text
    - resume_trans: {page_idx: [...]} pre-existing translations to skip (from checkpoint)
    progress_callback(page_idx, done_count, total, in_tok_delta, out_tok_delta, error, tm_hits).
    Returns (all_trans: dict, total_in_tok, total_out_tok, errors: dict, total_tm_hits: int).
    """
    all_trans: dict = {}
    errors: dict = {}
    total_in = total_out = 0
    done_count = 0
    total_tm_hits = 0
    total = len(targets)
    lock = threading.Lock()

    skip_pages = skip_pages or {}
    resume_trans = resume_trans or {}

    # Pre-fill skipped pages (TOC/References) with original text
    for pi, reason in skip_pages.items():
        if pi in targets:
            groups = all_groups.get(pi, [])
            all_trans[pi] = [g["text"] for g in groups]

    # Pre-fill resumed pages from checkpoint
    for pi, trans in resume_trans.items():
        if pi in targets and pi not in all_trans:
            all_trans[pi] = trans

    pages_to_translate = [pi for pi in targets
                          if pi not in skip_pages and pi not in resume_trans]

    with ThreadPoolExecutor(max_workers=max_workers) as ex:
        futures = {
            ex.submit(translate_page_pure, client, all_groups.get(pi, []),
                      target_lang, glossary, tm, True,
                      custom_glossary, custom_rules): pi
            for pi in pages_to_translate
        }
        for future in as_completed(futures):
            pi = futures[future]
            try:
                trans, in_t, out_t, err, tm_hits = future.result()
            except Exception as e:
                trans, in_t, out_t, err, tm_hits = (
                    [g["text"] for g in all_groups.get(pi, [])], 0, 0, str(e), 0,
                )
            all_trans[pi] = trans
            if err:
                errors[pi] = err
            with lock:
                total_in  += in_t
                total_out += out_t
                total_tm_hits += tm_hits
                done_count += 1
                snapshot = (done_count, total_in, total_out)
            if progress_callback:
                progress_callback(pi, snapshot[0], len(pages_to_translate),
                                  in_t, out_t, err, tm_hits)

    return all_trans, total_in, total_out, errors, total_tm_hits


# ══════════════════════════════════════════════════════════════════════════════
# GLOSSARY — frequency-based proper noun / acronym detection
# ══════════════════════════════════════════════════════════════════════════════
def build_pdf_glossary(all_groups: dict, min_count: int = 2,
                       max_terms: int = 30) -> list[str]:
    """
    Scan all extracted text for capitalized terms (proper nouns, acronyms, CamelCase)
    that appear >= min_count times. Returns top N by frequency for prompt consistency.
    """
    counter: Counter = Counter()
    # Match: ACRONYM (2+ caps), CamelCase, Proper noun (Cap + 2+ letters)
    pattern = re.compile(r"\b[A-Z][A-Za-z0-9]{2,}\b")
    for groups in all_groups.values():
        for g in groups:
            for term in pattern.findall(g["text"]):
                # Skip common English words that just happen to be capitalized
                if term.lower() in {"the", "and", "for", "with", "from", "this",
                                    "that", "page", "table", "figure"}:
                    continue
                counter[term] += 1
    return [term for term, c in counter.most_common(max_terms) if c >= min_count]


# ══════════════════════════════════════════════════════════════════════════════
# OCR — extract embedded images from PDF and translate text in them
# ══════════════════════════════════════════════════════════════════════════════
def extract_pdf_images(pdf_path: str, page_nums: list[int] | None = None) -> list[dict]:
    """
    Extract embedded images from PDF pages.
    Returns list of {id, page_idx, xref, bbox, content_type, data}.
    bbox = position on page (None if image not placed via standard refs).
    """
    out: list[dict] = []
    doc = fitz.open(pdf_path)
    total = len(doc)
    targets = page_nums if page_nums else list(range(total))
    seen_xrefs: set = set()

    for pi in targets:
        if not (0 <= pi < total):
            continue
        page = doc[pi]
        for img_info in page.get_images(full=True):
            xref = img_info[0]
            if xref in seen_xrefs:
                continue
            seen_xrefs.add(xref)
            try:
                base = doc.extract_image(xref)
            except Exception:
                continue
            data = base.get("image")
            ext  = base.get("ext", "png")
            if not data or len(data) < 5000:
                continue
            # Find bbox on page (first occurrence)
            bbox = None
            try:
                rects = page.get_image_rects(xref)
                if rects:
                    bbox = tuple(rects[0])
            except Exception:
                pass
            out.append({
                "id":           f"PDFIMG_{len(out)}",
                "page_idx":     pi,
                "xref":         xref,
                "bbox":         bbox,
                "content_type": f"image/{'jpeg' if ext == 'jpg' else ext}",
                "data":         data,
            })
    doc.close()
    return out


def ocr_pdf_images(client, images: list[dict], target_lang: str,
                   progress_callback=None, max_workers: int = 4) -> dict:
    """
    OCR + translate PDF images in parallel. Reuses word_backend's _ocr_single_image.
    Returns {image_id: {ocr, translation, has_text, error?}}.
    """
    from word_backend import _ocr_single_image, get_working_model, WORD_MODELS

    results: dict = {}
    model = get_working_model() or WORD_MODELS[0]
    if not images:
        return results

    done = 0
    total = len(images)
    lock = threading.Lock()

    with ThreadPoolExecutor(max_workers=max_workers) as ex:
        futures = {
            ex.submit(_ocr_single_image, client, img, target_lang, model): img
            for img in images
        }
        for future in as_completed(futures):
            img = futures[future]
            try:
                results[img["id"]] = future.result()
            except Exception as e:
                results[img["id"]] = {
                    "ocr": "", "translation": "", "has_text": False, "error": str(e),
                }
            with lock:
                done += 1
                current = done
            if progress_callback:
                progress_callback(current, total)
    return results


def insert_ocr_captions_into_pdf(pdf_path: str, dst_path: str,
                                 images: list[dict], ocr_results: dict,
                                 font_path: str | None = None) -> int:
    """
    Annotate each image (that has translated text) with a yellow text annotation
    containing the OCR translation. Saves modified PDF to dst_path.
    Returns number of captions inserted.
    """
    doc = fitz.open(pdf_path)
    inserted = 0
    for img in images:
        r = ocr_results.get(img["id"], {})
        if not (r.get("has_text") and r.get("translation")):
            continue
        if img.get("bbox") is None:
            continue
        page = doc[img["page_idx"]]
        bbox = fitz.Rect(img["bbox"])
        # Anchor annotation at top-left of image
        try:
            annot = page.add_text_annot(
                (bbox.x0, bbox.y0),
                f"[OCR] {r['translation']}",
                icon="Comment",
            )
            annot.set_colors(stroke=(1.0, 0.85, 0.2))
            annot.update()
            inserted += 1
        except Exception:
            pass
    doc.save(dst_path, garbage=4, deflate=True)
    doc.close()
    return inserted


# ══════════════════════════════════════════════════════════════════════════════
# QUALITY CHECK — số bị mất, length ratio, dòng bỏ sót
# ══════════════════════════════════════════════════════════════════════════════
_PDF_NUM_PAT = re.compile(r"\b\d[\d,\.]*\b")


def quality_check_pdf(all_groups: dict, all_trans: dict) -> list[dict]:
    """
    Returns list of {page, line_idx, text, translation, issues} for suspicious lines.
    Checks per line: missing numbers, abnormal length ratio, untranslated (text == translation).
    """
    issues_out: list[dict] = []
    for pi in sorted(all_groups.keys()):
        groups = all_groups[pi]
        trans  = all_trans.get(pi, [])
        for li, g in enumerate(groups):
            if li >= len(trans):
                break
            orig = g["text"].strip()
            tr   = (trans[li] or "").strip()
            if not orig or not tr:
                continue
            issues: list[str] = []

            orig_nums = set(_PDF_NUM_PAT.findall(orig))
            tr_nums   = set(_PDF_NUM_PAT.findall(tr))
            missing   = orig_nums - tr_nums
            if missing:
                issues.append(f"Số bị mất: {', '.join(sorted(missing)[:4])}")

            ratio = len(tr) / max(len(orig), 1)
            if ratio > 3.5:
                issues.append(f"Dịch quá dài ({ratio:.1f}x)")
            elif ratio < 0.2 and len(orig) > 20:
                issues.append(f"Dịch quá ngắn ({ratio:.1f}x)")

            if tr == orig and len(orig) > 15 and any(c.isalpha() for c in orig):
                issues.append("Chưa dịch (giữ nguyên text gốc)")

            if issues:
                issues_out.append({
                    "page": pi + 1, "line_idx": li,
                    "text": orig, "translation": tr, "issues": issues,
                })
    return issues_out


# ══════════════════════════════════════════════════════════════════════════════
# BILINGUAL PDF — interleave original + translated pages (page 1: orig, page 2: trans, ...)
# ══════════════════════════════════════════════════════════════════════════════
def build_bilingual_pdf(src_path: str, translated_path: str, dst_path: str,
                        targets: list[int]):
    """
    Create a bilingual PDF: pages from `src_path` (original) interleaved with
    corresponding pages from `translated_path` (translated). Only targets pages.
    Layout: [orig p1] [trans p1] [orig p2] [trans p2] ...
    """
    src   = fitz.open(src_path)
    trans = fitz.open(translated_path)
    out   = fitz.open()
    for pi in targets:
        if 0 <= pi < len(src):
            out.insert_pdf(src,   from_page=pi, to_page=pi)
        if 0 <= pi < len(trans):
            out.insert_pdf(trans, from_page=pi, to_page=pi)
    out.save(dst_path, garbage=4, deflate=True)
    out.close()
    src.close()
    trans.close()


# ══════════════════════════════════════════════════════════════════════════════
# CHECKPOINT — persist partial translations, resume if interrupted
# ══════════════════════════════════════════════════════════════════════════════
def _pdf_checkpoint_path(pdf_bytes: bytes, target_lang: str) -> str:
    h    = hashlib.md5(pdf_bytes[:8192]).hexdigest()[:12]
    slug = target_lang.replace(" ", "_")[:12]
    return f"/tmp/pdf_ckpt_{h}_{slug}.pkl"


def pdf_checkpoint_save(pdf_bytes: bytes, target_lang: str,
                        all_groups: dict, all_trans: dict,
                        glossary: list[str] | None = None) -> None:
    """Save current translation state. Called per-page during translation."""
    try:
        data = {
            "groups":   all_groups,
            "trans":    all_trans,
            "glossary": glossary or [],
            "ts":       time.time(),
        }
        with open(_pdf_checkpoint_path(pdf_bytes, target_lang), "wb") as f:
            pickle.dump(data, f)
    except Exception:
        pass


def pdf_checkpoint_load(pdf_bytes: bytes, target_lang: str) -> dict | None:
    path = _pdf_checkpoint_path(pdf_bytes, target_lang)
    try:
        if os.path.exists(path):
            with open(path, "rb") as f:
                return pickle.load(f)
    except Exception:
        pass
    return None


def pdf_checkpoint_clear(pdf_bytes: bytes, target_lang: str) -> None:
    try:
        os.unlink(_pdf_checkpoint_path(pdf_bytes, target_lang))
    except Exception:
        pass


# ══════════════════════════════════════════════════════════════════════════════
# TOC / REFERENCES detection — skip pages that aren't worth translating
# ══════════════════════════════════════════════════════════════════════════════
_SKIP_TITLES = (
    "table of contents", "contents", "mục lục",
    "references", "bibliography", "tài liệu tham khảo",
    "index", "appendix", "phụ lục",
)
_TOC_LINE_PAT = re.compile(r"\.{3,}\s*\d+\s*$|\s\d+\s*$")


def detect_skip_pages(all_groups: dict) -> dict[int, str]:
    """
    Heuristic detection of TOC / References / Index / Appendix pages.
    Returns {page_idx: reason} for pages that should be skipped.

    Detection rules:
    - First few lines contain known title (Contents, References, Index, ...)
    - >50% of lines end with "...." + number (typical TOC entry pattern)
    """
    skip_pages: dict[int, str] = {}
    for pi, groups in all_groups.items():
        if not groups:
            continue
        # Check first 3 non-empty lines for a known section title
        first_texts = [g["text"].strip().lower() for g in groups[:5] if g["text"].strip()]
        title_hit = next(
            (kw for line in first_texts for kw in _SKIP_TITLES
             if kw in line and len(line) < 60),
            None,
        )
        if title_hit:
            skip_pages[pi] = f"Detected: '{title_hit}'"
            continue
        # Check if >50% of lines look like TOC entries
        toc_like = sum(1 for g in groups if _TOC_LINE_PAT.search(g["text"]))
        if len(groups) >= 6 and toc_like / len(groups) > 0.5:
            skip_pages[pi] = f"TOC pattern ({toc_like}/{len(groups)} lines)"
    return skip_pages


# ══════════════════════════════════════════════════════════════════════════════
# CUSTOM GLOSSARY IMPORT — parse user-provided source→target term pairs
# ══════════════════════════════════════════════════════════════════════════════
def parse_custom_glossary(text: str) -> dict[str, str]:
    """
    Parse glossary from text. Each line: 'source = target' or 'source -> target'
    or 'source,target' (CSV). Returns {source: target} dict.
    """
    out: dict[str, str] = {}
    for line in text.splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        for sep in (" = ", "=", " -> ", "->", "\t", ","):
            if sep in line:
                src, _, tgt = line.partition(sep)
                src, tgt = src.strip(), tgt.strip()
                if src and tgt:
                    out[src] = tgt
                break
    return out


def _build_page_prompt_v2(groups: list, target_lang: str,
                          glossary: list[str] | None = None,
                          custom_glossary: dict[str, str] | None = None,
                          custom_rules: str | None = None) -> str:
    """Extended prompt with custom rules and custom glossary support."""
    base = _build_page_prompt(groups, target_lang, glossary)
    extras = []
    if custom_glossary:
        pairs = "\n".join(f"  • {s}  →  {t}" for s, t in list(custom_glossary.items())[:50])
        extras.append(
            f"\nGLOSSARY BẮT BUỘC (dùng đúng bản dịch cho mỗi term):\n{pairs}\n"
        )
    if custom_rules and custom_rules.strip():
        extras.append(f"\nHƯỚNG DẪN DỊCH BỔ SUNG (tuân thủ nghiêm):\n{custom_rules.strip()}\n")
    if not extras:
        return base
    insert_at = base.find("Trả về JSON object")
    if insert_at < 0:
        return base + "\n" + "".join(extras)
    return base[:insert_at] + "".join(extras) + "\n" + base[insert_at:]
