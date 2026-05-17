"""
PDF backend: extract text spans, gọi Gemini, ghi lại PDF giữ nguyên layout.

3 nhóm logic chính:
1. Extract  : `extract_line_groups`, `parse_page_range`
2. Translate: `translate_page` + helpers (`call_gemini_live` xử lý retry & live timer)
3. Render   : `write_translated_pdf` + font helpers
"""
import os
import time
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


# ══════════════════════════════════════════════════════════════════════════════
# RENDER
# ══════════════════════════════════════════════════════════════════════════════
def _insert_line(page, font_path, font_bold_path, bbox, text, fontsize, color, bold):
    x0, y0, _, y1 = bbox
    pw, ph        = page.rect.width, page.rect.height
    x1_use        = pw - 30
    line_h        = max(y1 - y0, fontsize * 1.5)
    y1_use        = min(y0 + line_h * 3, ph - 10)

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
    while size >= 4.0:
        rc = page.insert_textbox(
            fitz.Rect(x0, y0, x1_use, y1_use),
            text, fontsize=size, fontname=fontname, color=color, align=0,
        )
        if rc >= 0:
            break
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


def translate_page(client, groups: list, target_lang: str, page_idx: int,
                   timer_ph, t0: float, page_start: float,
                   log_lines: list, log_ph):
    """Dịch 1 trang PDF (1 list line groups) → list[str] cùng độ dài."""
    numbered     = "\n".join(_format_group_for_prompt(i, g) for i, g in enumerate(groups))
    has_tables   = any(g.get("cell") for g in groups)
    table_hint   = (
        "\nMột số dòng có dạng (T# R# C#) — đó là cell trong bảng "
        "(T = số bảng, R = row, C = column).\n"
        "Khi dịch các cell bảng:\n"
        "- Cell cùng cột → dịch consistent, dùng cùng thuật ngữ\n"
        "- Header row (R1) → dịch ngắn gọn, súc tích như tiêu đề cột\n"
        "- Giá trị số / ngày tháng / đơn vị (mm, kg, V, %) → GIỮ NGUYÊN\n"
        "- Mã / serial / model code → GIỮ NGUYÊN\n"
        if has_tables else ""
    )
    prompt = (
        f"Dịch sang {target_lang}. Giữ nguyên số thứ tự [0]...[{len(groups)-1}].\n"
        f"{table_hint}"
        f"Trả về JSON object, ĐẦY ĐỦ từ \"0\" đến \"{len(groups)-1}\", không bỏ sót.\n"
        f"KHÔNG kèm prefix (T# R# C#) trong bản dịch — chỉ trả về text đã dịch:\n"
        f"{{\"0\": \"bản dịch\", \"1\": \"bản dịch\", ...}}\n"
        f"Chỉ JSON, không giải thích.\n\n"
        f"{numbered}"
    )
    label = f"Đang dịch trang {page_idx + 1}"
    raw, in_t, out_t = call_gemini_live(
        client, prompt, timer_ph, t0, page_start, label, log_lines, log_ph
    )
    parsed = parse_json_loose(raw)
    if isinstance(parsed, dict):
        return [str(parsed.get(str(i), groups[i]["text"])) for i in range(len(groups))], in_t, out_t
    if isinstance(parsed, list) and len(parsed) == len(groups):
        return [str(x) for x in parsed], in_t, out_t
    return [g["text"] for g in groups], in_t, out_t
