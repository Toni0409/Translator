"""
Streamlit UI cho tab Word — 2-phase flow.

Phase 1 (Phân tích):
  extract → glossary build → show stats + glossary editor + TM preview

Phase 2 (Dịch):
  TM lookup → chunk còn lại → translate parallel → merge TM + new → apply
"""
import io
import time
import threading
import zipfile

import pandas as pd
import streamlit as st

from config import (
    LANGUAGES, LANG_EN, TRANSLATION_DIRECTIONS,
    MAX_WORD_WORKERS, NO_TRANSLATE_ROLES, AUTO_RESCAN_PASSES,
)
from gemini import get_client
from ui_common import (
    timer_box_html, timer_done_html, timer_error_html,
    stat_box_html, make_log_adder, calc_cost,
)
from word_backend import (
    extract_docx_blocks, build_doc_context, build_glossary, chunk_blocks,
    translate_parallel, apply_translations, find_missed,
    get_working_model, count_by_role,
    checkpoint_save, checkpoint_load, checkpoint_clear,
    validate_docx_output,
)
from domain_glossary import detect_subdomain, seed_for_direction


# Keys session_state cho tab Word — tập trung 1 chỗ để dễ clear
SS_KEYS = (
    "word_blocks", "word_translations", "word_docx_bytes",
    "word_filename", "word_lang", "word_doc_context",
    "word_tok_in", "word_tok_out", "word_elapsed", "word_num_chunks",
    "word_summary", "word_glossary", "word_analysis",
)


# ══════════════════════════════════════════════════════════════════════════════
# UI HELPERS
# ══════════════════════════════════════════════════════════════════════════════
def _safe_zip_name(raw: str, lang_word: str, used: set[str]) -> str:
    """Sanitize tên file cho batch ZIP (P3.7).

    - basename only (loại path traversal `..` và separator `/` `\\`)
    - giữ extension `.docx`; thêm suffix `_<2chars lang>` trước `.docx`
    - đảm bảo unique trong `used` (thêm `_1`, `_2`... nếu trùng)
    """
    import os as _os, re as _re_local
    base = _os.path.basename(raw or "file.docx")
    base = _re_local.sub(r"[\\/]+", "_", base)
    base = base.replace("..", "_")
    if not base.lower().endswith(".docx"):
        base += ".docx"
    suffix = f"_{lang_word[:2]}"
    stem, ext = base[:-5], base[-5:]   # ".docx"
    candidate = f"{stem}{suffix}{ext}"
    if candidate not in used:
        return candidate
    # Đụng tên → thêm counter
    i = 1
    while f"{stem}{suffix}_{i}{ext}" in used:
        i += 1
    return f"{stem}{suffix}_{i}{ext}"


def _resolve_langs(direction_label: str | None = None) -> tuple[str, str, str]:
    """Resolve (label, source_lang, target_lang) từ direction label hoặc session_state.

    Fallback hợp lệ khi state thiếu: dùng hướng đầu tiên trong TRANSLATION_DIRECTIONS.
    """
    label = direction_label or st.session_state.get("word_direction")
    for d_label, src, tgt in TRANSLATION_DIRECTIONS:
        if d_label == label:
            return d_label, src, tgt
    d_label, src, tgt = TRANSLATION_DIRECTIONS[0]
    return d_label, src, tgt


def _make_job_ui(heading: str, stat_col1_label: str,
                 log_heading: str = "### 📋 Nhật ký"):
    """Build standard job UI. Returns (timer_ph, prog, render_stats, add_log)."""
    st.markdown(heading)
    timer_ph = st.empty()
    c1, c2, c3, c4 = st.columns(4)
    p1, p2, p3, p4 = (c1.empty(), c2.empty(), c3.empty(), c4.empty())

    def render_stats(total_bl, num_chunks, done_chunks, tok_in, tok_out):
        usd, vnd = calc_cost(tok_in, tok_out)
        p1.markdown(stat_box_html(f"{total_bl:,}", stat_col1_label), unsafe_allow_html=True)
        p2.markdown(stat_box_html(f"{done_chunks}/{num_chunks}", "Chunk"), unsafe_allow_html=True)
        p3.markdown(stat_box_html(f"${usd:.4f}", "USD"), unsafe_allow_html=True)
        p4.markdown(stat_box_html(f"{vnd:,.0f}₫", "VND"), unsafe_allow_html=True)

    render_stats(0, 0, 0, 0, 0)
    prog = st.progress(0, text="Đang chuẩn bị...")
    st.markdown(log_heading)
    log_ph    = st.empty()
    log_lines: list = []
    add_log   = make_log_adder(log_lines, log_ph)
    return timer_ph, prog, render_stats, add_log


def _finalize_job(timer_ph, prog, add_log,
                  elapsed: float, count: int, label: str,
                  tok_in: int, tok_out: int):
    usd, vnd = calc_cost(tok_in, tok_out)
    prog.progress(100, text=f"✅ {label} xong!")
    timer_ph.markdown(
        timer_done_html(elapsed, f"{label} {count:,} đoạn xong!"),
        unsafe_allow_html=True,
    )
    add_log("─" * 44)
    add_log(f"🎉 Xong {count:,} đoạn trong {elapsed:.1f}s")
    add_log(f"💵 Chi phí: ${usd:.4f} USD ≈ {vnd:,.0f} VND")


# ══════════════════════════════════════════════════════════════════════════════
# CORE TRANSLATION LOOP
# ══════════════════════════════════════════════════════════════════════════════
def _run_translation(chunks, target_lang, doc_context,
                     timer_ph, prog_ph, render_stats, add_log,
                     total_blocks, prefix_label="Dịch",
                     glossary: dict | None = None,
                     custom_rules: dict | None = None,
                     on_chunk_done=None,
                     source_lang: str | None = None):
    """
    Chạy `translate_parallel` trong thread, main loop poll & render UI.
    Trả về (translations, tok_in, tok_out, elapsed).
    """
    holder = {
        "translations": {},
        "tok_in":       0,
        "tok_out":      0,
        "chunk_done":   0,
        "chunk_log":    [],
        "done":         False,
        "error":        None,
    }
    client     = get_client()
    num_chunks = len(chunks)
    t0         = time.time()

    threading.Thread(
        target=translate_parallel,
        args=(holder, client, chunks, target_lang, doc_context,
              MAX_WORD_WORKERS, glossary, custom_rules, source_lang),
        daemon=True,
    ).start()

    last_logged = 0
    dot = 0
    while not holder["done"]:
        while last_logged < len(holder["chunk_log"]):
            entry = holder["chunk_log"][last_logged]
            if entry["error"]:
                add_log(f"   ❌ Chunk {entry['idx']+1}: {entry['error'][:80]}")
            else:
                add_log(f"   ✅ Chunk {entry['idx']+1}: {entry['size']} đoạn "
                        f"({entry['in_t']:,}in/{entry['out_t']:,}out tok)")
            last_logged += 1

        if on_chunk_done and last_logged > 0:
            on_chunk_done(holder["translations"])

        elapsed = time.time() - t0
        done_c  = holder["chunk_done"]
        pct     = int(done_c / num_chunks * 90) if num_chunks > 0 else 0
        dots    = "." * (dot % 4)
        timer_ph.markdown(
            timer_box_html(
                elapsed,
                f"🔄 {prefix_label} song song — {done_c}/{num_chunks} chunk{dots}",
            ),
            unsafe_allow_html=True,
        )
        prog_ph.progress(pct, text=f"{prefix_label}: {done_c}/{num_chunks} chunk...")
        render_stats(total_blocks, num_chunks, done_c, holder["tok_in"], holder["tok_out"])
        dot += 1
        time.sleep(0.5)

    while last_logged < len(holder["chunk_log"]):
        entry = holder["chunk_log"][last_logged]
        if entry["error"]:
            add_log(f"   ❌ Chunk {entry['idx']+1}: {entry['error'][:80]}")
        else:
            add_log(f"   ✅ Chunk {entry['idx']+1}: {entry['size']} đoạn "
                    f"({entry['in_t']:,}in/{entry['out_t']:,}out tok)")
        last_logged += 1

    if holder["error"]:
        raise RuntimeError(holder["error"])

    return holder["translations"], holder["tok_in"], holder["tok_out"], time.time() - t0


def _auto_rescan_missed(blocks: list[dict], translations: dict,
                        source_lang: str, target_lang: str,
                        doc_context: str, glossary: dict | None,
                        custom_rules: dict | None,
                        timer_ph, prog_ph, render_stats, add_log) -> tuple[int, int, float, list]:
    """Automatically translate blocks still detected as missed after the main pass."""
    total_in = total_out = 0
    total_elapsed = 0.0
    remaining_missed: list = []

    for pass_no in range(1, AUTO_RESCAN_PASSES + 1):
        missed = find_missed(
            blocks, translations,
            source_lang=source_lang,
            target_lang=target_lang,
        )
        if not missed:
            if pass_no == 1:
                add_log("✅ Auto scan: không còn đoạn bỏ sót")
            break

        remaining_missed = missed
        add_log(f"🔍 Auto scan {pass_no}/{AUTO_RESCAN_PASSES}: "
                f"phát hiện {len(missed):,} đoạn cần dịch bù")
        chunks = chunk_blocks(missed)
        add_log(f"📦 Quét bù: {len(chunks)} chunk")

        new_tr, tok_in, tok_out, elapsed = _run_translation(
            chunks, target_lang, doc_context,
            timer_ph, prog_ph, render_stats, add_log,
            len(missed), prefix_label=f"Quét {pass_no}",
            glossary=glossary,
            custom_rules=custom_rules or None,
            source_lang=source_lang,
        )
        translations.update(new_tr)
        total_in += tok_in
        total_out += tok_out
        total_elapsed += elapsed

    remaining_missed = find_missed(
        blocks, translations,
        source_lang=source_lang,
        target_lang=target_lang,
    )
    if remaining_missed:
        add_log(f"⚠️ Auto scan vẫn còn {len(remaining_missed):,} đoạn nghi chưa sạch "
                "sau giới hạn quét bù")
    elif total_in or total_out:
        add_log("✅ Auto scan: đã dịch bù xong, không còn đoạn nghi bỏ sót")

    return total_in, total_out, total_elapsed, remaining_missed


def _find_output_coverage_missed(docx_bytes: bytes, blocks: list[dict],
                                 translations: dict,
                                 source_lang: str, target_lang: str) -> list:
    """Scan applied DOCX bytes so silent paragraph-apply failures are visible."""
    output_blocks = extract_docx_blocks(docx_bytes)
    return find_missed(
        blocks, translations,
        source_lang=source_lang,
        target_lang=target_lang,
        output_blocks=output_blocks,
    )


# ══════════════════════════════════════════════════════════════════════════════
# PHASE 1 — PHÂN TÍCH (extract + glossary build)
# ══════════════════════════════════════════════════════════════════════════════
def _run_analysis(uploaded_docx, lang_word, source_lang: str, target_lang: str):
    """
    Phase 1: extract + build glossary + TM preview.
    Lưu kết quả vào session_state["word_analysis"], xong rerun để show editor.
    """
    st.markdown("### 🔬 Phân tích tài liệu")
    log_ph    = st.empty()
    log_lines: list = []
    add_log   = make_log_adder(log_lines, log_ph)

    try:
        docx_bytes = uploaded_docx.read()
        add_log(f"📄 File: {uploaded_docx.name}")

        blocks       = extract_docx_blocks(docx_bytes)
        stats        = count_by_role(blocks)
        translatable = [b for b in blocks if b["role"] not in NO_TRANSLATE_ROLES]
        total_chars  = sum(len(b["text"]) for b in translatable)
        textbox_cnt  = sum(1 for b in blocks if b["role"] == "textbox")
        table_cnt    = sum(1 for b in blocks if b.get("table_cell"))

        footnote_cnt = stats.get("footnote", 0)
        endnote_cnt  = stats.get("endnote", 0)

        add_log(f"✅ {stats['total']:,} đoạn ({total_chars:,} ký tự body)")
        add_log(f"   • Body cần dịch:       {stats['body']:,}")
        hf_cnt = stats.get("header", 0) + stats.get("footer", 0) + stats.get("body_repeated", 0)
        if hf_cnt > 0:
            add_log(f"   • Header/Footer/lặp:   {hf_cnt:,} (đã gộp vào body để dịch)")
        if textbox_cnt > 0:
            add_log(f"   • Text-box / shape:    {textbox_cnt:,} (sẽ dịch)")
        if table_cnt > 0:
            add_log(f"   • Cell bảng:           {table_cnt:,} → table-aware (T#R#C# context)")
        if footnote_cnt > 0:
            add_log(f"   • Footnotes:           {footnote_cnt:,} (sẽ dịch)")
        if endnote_cnt > 0:
            add_log(f"   • Endnotes:            {endnote_cnt:,} (sẽ dịch)")
        comment_cnt = stats.get("comment", 0)
        if comment_cnt > 0:
            add_log(f"   • Comments:             {comment_cnt:,} (sẽ dịch)")
        image_alt_cnt = stats.get("by_role", {}).get("image_alt", 0)
        if image_alt_cnt > 0:
            add_log(f"   • Image alt-texts:      {image_alt_cnt:,} (sẽ dịch)")

        # target_lang/source_lang nhận từ caller (xem _resolve_langs).

        # Domain detection + seed glossary (P4.4 + P4.5)
        subdomains = detect_subdomain(blocks)
        seed = seed_for_direction(subdomains, source_lang, target_lang)
        if subdomains:
            add_log(f"🏗 Domain: {', '.join(sorted(subdomains))} "
                    f"→ nạp {len(seed)} thuật ngữ chuyên ngành làm seed")

        # Glossary build (seed → AI-extract; seed không bị override)
        add_log("📚 Build glossary từ thuật ngữ lặp lại...")
        client   = get_client()
        glossary = build_glossary(client, blocks, target_lang,
                                  source_lang=source_lang, seed=seed)
        new_from_ai = len(glossary) - len(seed)
        if glossary:
            if seed and new_from_ai > 0:
                add_log(f"📖 Glossary: {len(seed)} seed + {new_from_ai} AI = {len(glossary)} term")
            else:
                add_log(f"📖 Tìm thấy {len(glossary)} thuật ngữ — review/sửa ở dưới rồi dịch")
        else:
            add_log("📖 Không có thuật ngữ lặp đủ ngưỡng — bỏ qua glossary")

        st.session_state["word_analysis"] = {
            "blocks":        blocks,
            "translatable":  translatable,
            "stats":         stats,
            "total_chars":   total_chars,
            "textbox_cnt":   textbox_cnt,
            "table_cnt":     table_cnt,
            "footnote_cnt":  footnote_cnt,
            "endnote_cnt":   endnote_cnt,
            "comment_cnt":   comment_cnt,
            "doc_context":   build_doc_context(blocks, source_lang=source_lang,
                                               subdomains=subdomains),
            "glossary":      glossary,
            # Snapshot để restore khi user lỡ xoá entries trong editor
            "_glossary_initial": dict(glossary),
            "docx_bytes":    docx_bytes,
            "filename":      uploaded_docx.name,
            "lang":          lang_word,
            "source_lang":   source_lang,
            "target_lang":   target_lang,
            "subdomains":    subdomains,
            "seed_glossary": seed,
        }
        st.session_state["word_subdomains"]    = subdomains
        st.session_state["word_seed_glossary"] = seed
        # Clear old kết quả từ lần dịch trước
        for k in ("word_blocks", "word_translations", "word_summary"):
            st.session_state.pop(k, None)
        st.session_state.pop("word_translated_bytes_cache", None)
        add_log("✅ Phân tích xong — kéo xuống review glossary + bấm **Dịch**")
        time.sleep(0.3)
        st.rerun()

    except Exception as e:
        add_log(f"❌ Lỗi: {e}")
        st.error(f"❌ Phân tích fail: {e}")



# ══════════════════════════════════════════════════════════════════════════════
# PHASE 2 — DỊCH (consume analysis, run translation, merge TM + new)
# ══════════════════════════════════════════════════════════════════════════════
def _run_full_translation():
    """Phase 2: dùng analysis đã extract + glossary đã edit → translate."""
    a = st.session_state["word_analysis"]
    blocks       = a["blocks"]
    target_lang  = a["target_lang"]
    source_lang  = a.get("source_lang") or _resolve_langs()[1]
    doc_context  = a["doc_context"]
    glossary     = a["glossary"]
    custom_rules = a.get("custom_rules") or {}
    docx_bytes   = a["docx_bytes"]

    # Dynamic skip_roles from per-role toggles
    toggles = a.get("role_toggles", {})
    if toggles:
        skip_roles = {role for role, on in toggles.items() if not on}
    else:
        skip_roles = set(NO_TRANSLATE_ROLES)
    # Recompute translatable list from blocks (excluding TOC mirrors — they inherit headings)
    translatable = [b for b in blocks
                    if b["role"] not in skip_roles
                    and not b.get("_toc_mirror")]
    coverage_blocks = [b for b in blocks if b["role"] not in skip_roles]

    timer_ph, prog, render_stats, add_log = _make_job_ui(
        "### 📊 Tiến độ dịch", "Đoạn văn", "### 📋 Nhật ký",
    )

    try:
        toc_mirror_cnt = sum(1 for b in blocks if b.get("_toc_mirror"))
        if toc_mirror_cnt:
            add_log(f"🔗 TOC: {toc_mirror_cnt} entries sẽ dùng lại translation từ heading")

        add_log(f"📝 Sẽ gọi API cho {len(translatable):,} đoạn")

        translations: dict = {}

        # Checkpoint recovery: restore partial translations from previous run
        remaining = list(translatable)
        ckpt = checkpoint_load(docx_bytes, target_lang)
        if ckpt:
            ckpt_hits = sum(1 for b in remaining if b["id"] in ckpt)
            if ckpt_hits > 0:
                add_log(f"🔄 Checkpoint: {ckpt_hits:,} đoạn từ lần dịch trước")
                translations.update(ckpt)
                remaining = [b for b in remaining if b["id"] not in translations]

        tok_in = tok_out = 0
        elapsed = 0.0
        num_chunks = 0

        if remaining:
            chunks     = chunk_blocks(remaining)
            num_chunks = len(chunks)
            avg_chars  = sum(len(b["text"]) for b in remaining) // max(num_chunks, 1)
            add_log(f"📦 Chia thành {num_chunks} chunk (~{avg_chars:,} chars/chunk)")
            if glossary:
                add_log(f"📖 Áp dụng glossary {len(glossary)} thuật ngữ")
            add_log(f"⚡ Dịch song song {MAX_WORD_WORKERS} luồng")

            def _save_ckpt(tr_so_far):
                checkpoint_save(docx_bytes, target_lang, dict(tr_so_far))

            new_translations, tok_in, tok_out, elapsed = _run_translation(
                chunks, target_lang, doc_context,
                timer_ph, prog, render_stats, add_log,
                len(remaining), prefix_label="Dịch",
                glossary=glossary,
                custom_rules=custom_rules or None,
                on_chunk_done=_save_ckpt,
                source_lang=source_lang,
            )
            translations.update(new_translations)
            checkpoint_clear(docx_bytes, target_lang)
            add_log(f"🤖 Model: {get_working_model()}")
        else:
            timer_ph.markdown(
                timer_done_html(0, "Khôi phục từ checkpoint — không gọi API!"),
                unsafe_allow_html=True,
            )
            prog.progress(100, text="✅ Done!")

        extra_in, extra_out, extra_elapsed, remaining_missed = _auto_rescan_missed(
            translatable, translations, source_lang, target_lang,
            doc_context, glossary, custom_rules,
            timer_ph, prog, render_stats, add_log,
        )
        tok_in += extra_in
        tok_out += extra_out
        elapsed += extra_elapsed

        usd, vnd = calc_cost(tok_in, tok_out)
        render_stats(len(translatable), num_chunks, num_chunks, tok_in, tok_out)
        prog.progress(100, text="✅ Hoàn thành!")
        timer_ph.markdown(
            timer_done_html(elapsed, f"Dịch xong {len(translatable):,} đoạn!"),
            unsafe_allow_html=True,
        )
        add_log("─" * 44)
        add_log(f"🎉 Xong {len(translatable):,} đoạn — {elapsed:.1f}s")
        if tok_in or tok_out:
            add_log(f"💰 Token: {tok_in:,} in + {tok_out:,} out")
        add_log(f"💵 Chi phí: ${usd:.4f} USD ≈ {vnd:,.0f} VND")

        st.session_state["word_blocks"]       = blocks
        st.session_state["word_translations"] = translations
        st.session_state["word_docx_bytes"]   = docx_bytes
        st.session_state["word_filename"]     = a["filename"]
        st.session_state["word_lang"]         = a["lang"]
        st.session_state["word_source_lang"]  = source_lang
        st.session_state["word_target_lang"]  = target_lang
        st.session_state["word_doc_context"]  = doc_context
        st.session_state["word_tok_in"]       = tok_in
        st.session_state["word_tok_out"]      = tok_out
        st.session_state["word_elapsed"]      = elapsed
        st.session_state["word_num_chunks"]   = num_chunks
        st.session_state["word_glossary"]     = glossary
        st.session_state["word_custom_rules"] = custom_rules
        st.session_state["word_auto_missed_remaining"] = len(remaining_missed)
        st.session_state["word_summary"]      = (
            f"✅ Dịch xong {len(translatable):,} đoạn  "
            f"|  {elapsed:.1f}s  |  ${usd:.4f} USD ≈ {vnd:,.0f} VND"
        )
        st.session_state["word_translations_version"] = 1
        # Persist role toggles (for post-translation editor filter)
        st.session_state["word_role_toggles"] = toggles
        st.session_state["word_skip_roles"]   = skip_roles

        # Validate output DOCX (so sánh media với bản gốc — P2.7)
        try:
            out_bytes = apply_translations(docx_bytes, blocks, translations)
            val = validate_docx_output(out_bytes, original_bytes=docx_bytes)
            coverage_missed = _find_output_coverage_missed(
                out_bytes, coverage_blocks, translations, source_lang, target_lang,
            )
            if coverage_missed:
                val["warnings"].append(
                    f"Coverage scan: còn {len(coverage_missed):,} đoạn nghi chưa dịch sạch"
                )
            st.session_state["word_validation"] = val
            st.session_state["word_output_coverage_missed"] = len(coverage_missed)
            st.session_state["word_translated_bytes_cache"] = {
                "version": 1, "bytes": out_bytes,
            }
            if not val["valid"]:
                add_log(f"❌ Validate: {len(val['errors'])} lỗi")
                for err in val["errors"][:3]:
                    add_log(f"   • {err}")
            else:
                add_log(f"✅ Validate: {val['block_count']:,} paragraph, "
                        f"{val['image_count']} ảnh — OK")
            if coverage_missed:
                add_log(f"⚠️ Coverage scan: {len(coverage_missed):,} đoạn output "
                        "vẫn có dấu hiệu còn tiếng nguồn")
            for warn in val["warnings"]:
                add_log(f"⚠️ {warn}")
        except Exception as e:
            add_log(f"⚠️ Validate fail: {e}")

        st.session_state.pop("word_analysis", None)  # đã consume

    except Exception as e:
        add_log(f"❌ Lỗi: {e}")
        st.error(f"❌ Lỗi dịch: {e}")
        timer_ph.markdown(timer_error_html(str(e)), unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# RESCAN + H/F — share TM logic
# ══════════════════════════════════════════════════════════════════════════════
def _run_partial(missed: list, label: str, heading: str, log_heading: str,
                 stat_label: str, success_msg: str):
    """Generic partial-translation runner (rescan / H-F)."""
    if not missed:
        st.info("✨ Không có đoạn nào cần xử lý!")
        return

    doc_context  = st.session_state["word_doc_context"]
    target_lang  = st.session_state.get("word_target_lang") or LANG_EN[st.session_state["word_lang"]]
    source_lang  = st.session_state.get("word_source_lang") or _resolve_langs()[1]
    glossary     = st.session_state.get("word_glossary")
    custom_rules = st.session_state.get("word_custom_rules")
    translations = st.session_state["word_translations"]

    timer_ph, prog, render_stats, add_log = _make_job_ui(
        heading, stat_label, log_heading,
    )
    add_log(f"📝 Cần gọi API cho {len(missed):,} đoạn")

    chunks = chunk_blocks(missed)
    add_log(f"📦 Chia thành {len(chunks)} chunk — {MAX_WORD_WORKERS} luồng")

    try:
        new_translations, tok_in, tok_out, elapsed = _run_translation(
            chunks, target_lang, doc_context,
            timer_ph, prog, render_stats, add_log,
            len(missed), prefix_label=label,
            glossary=glossary,
            custom_rules=custom_rules or None,
            source_lang=source_lang,
        )
        translations.update(new_translations)
        st.session_state["word_translations"] = translations
        st.session_state["word_tok_in"]      += tok_in
        st.session_state["word_tok_out"]     += tok_out
        st.session_state["word_translations_version"] = (
            st.session_state.get("word_translations_version", 0) + 1
        )
        _finalize_job(timer_ph, prog, add_log, elapsed,
                      len(missed), label, tok_in, tok_out)

        # Validate updated output (so sánh media với bản gốc — P2.7)
        try:
            out_bytes = apply_translations(
                st.session_state["word_docx_bytes"],
                st.session_state["word_blocks"],
                translations,
            )
            val = validate_docx_output(
                out_bytes,
                original_bytes=st.session_state["word_docx_bytes"],
            )
            skip_roles = st.session_state.get("word_skip_roles", set(NO_TRANSLATE_ROLES))
            coverage_blocks = [
                b for b in st.session_state["word_blocks"]
                if b["role"] not in skip_roles
            ]
            coverage_missed = _find_output_coverage_missed(
                out_bytes,
                coverage_blocks,
                translations,
                source_lang,
                target_lang,
            )
            if coverage_missed:
                val["warnings"].append(
                    f"Coverage scan: còn {len(coverage_missed):,} đoạn nghi chưa dịch sạch"
                )
            st.session_state["word_validation"] = val
            st.session_state["word_output_coverage_missed"] = len(coverage_missed)
            new_version = st.session_state["word_translations_version"]
            st.session_state["word_translated_bytes_cache"] = {
                "version": new_version, "bytes": out_bytes,
            }
            if not val["valid"]:
                add_log(f"❌ Validate: {len(val['errors'])} lỗi")
                for err in val["errors"][:3]:
                    add_log(f"   • {err}")
            else:
                add_log(f"✅ Validate: {val['block_count']:,} paragraph, "
                        f"{val['image_count']} ảnh — OK")
            if coverage_missed:
                add_log(f"⚠️ Coverage scan: {len(coverage_missed):,} đoạn output "
                        "vẫn có dấu hiệu còn tiếng nguồn")
            for warn in val["warnings"]:
                add_log(f"⚠️ {warn}")
        except Exception as e:
            add_log(f"⚠️ Validate fail: {e}")

        st.success(success_msg.format(n=len(missed)))

    except Exception as e:
        add_log(f"❌ Lỗi: {e}")
        st.error(f"❌ Lỗi: {e}")
        timer_ph.markdown(timer_error_html(str(e)), unsafe_allow_html=True)


def _run_rescan():
    blocks       = st.session_state["word_blocks"]
    translations = st.session_state["word_translations"]
    source_lang  = st.session_state.get("word_source_lang") or _resolve_langs()[1]
    target_lang  = st.session_state.get("word_target_lang") or LANG_EN[st.session_state["word_lang"]]
    skip_roles   = st.session_state.get("word_skip_roles", set(NO_TRANSLATE_ROLES))
    candidates   = [b for b in blocks if b["role"] not in skip_roles]
    missed       = find_missed(
        candidates, translations,
        source_lang=source_lang,
        target_lang=target_lang,
    )
    _run_partial(
        missed,
        label="Quét",
        heading=f"### 🔍 Quét bỏ sót — {len(missed):,} đoạn",
        log_heading="### 📋 Nhật ký quét",
        stat_label="Đoạn còn sót",
        success_msg="✅ Đã dịch thêm {n:,} đoạn bị bỏ sót!",
    )


# ══════════════════════════════════════════════════════════════════════════════
# CACHED TRANSLATED BYTES
# ══════════════════════════════════════════════════════════════════════════════
def _get_cached_translated_bytes() -> bytes:
    version = st.session_state.get("word_translations_version", 0)
    cached  = st.session_state.get("word_translated_bytes_cache")
    if cached and cached.get("version") == version:
        return cached["bytes"]
    result = apply_translations(
        st.session_state["word_docx_bytes"],
        st.session_state["word_blocks"],
        st.session_state["word_translations"],
    )
    st.session_state["word_translated_bytes_cache"] = {"version": version, "bytes": result}
    return result


# ══════════════════════════════════════════════════════════════════════════════
# MAIN UI
# ══════════════════════════════════════════════════════════════════════════════
def _clear_state():
    for k in SS_KEYS:
        st.session_state.pop(k, None)
    for k in ("word_translated_bytes_cache", "word_translations_version",
              "word_validation", "word_role_toggles", "word_skip_roles",
              "word_auto_missed_remaining", "word_output_coverage_missed",
              "word_glossary_editor_ver",
              "word_image_ocr", "word_ocr_state",
              "word_batch_result"):
        st.session_state.pop(k, None)
    for k in list(st.session_state.keys()):
        if k.startswith("word_glossary_editor"):
            st.session_state.pop(k, None)


def _run_batch(uploaded_files, lang_word, source_lang: str, target_lang: str):
    """Batch-translate multiple DOCX files."""
    st.markdown("### 📚 Dịch theo lô")

    file_status = {f.name: "⏳ Chờ" for f in uploaded_files}
    status_ph   = st.empty()
    results     = {}  # name → bytes

    for uploaded in uploaded_files:
        file_status[uploaded.name] = "🔄 Đang dịch..."
        status_ph.table({
            "File":       list(file_status.keys()),
            "Trạng thái": list(file_status.values()),
        })
        try:
            docx_bytes  = uploaded.read()
            blocks      = extract_docx_blocks(docx_bytes)
            doc_context = build_doc_context(blocks)
            client      = get_client()
            glossary    = build_glossary(client, blocks, target_lang, source_lang)
            translatable = [b for b in blocks if b["role"] not in NO_TRANSLATE_ROLES]
            translations: dict = {}

            if translatable:
                chunks = chunk_blocks(translatable)
                timer_ph  = st.empty()
                prog_ph   = st.empty()

                def _noop_stats(*_a, **_kw): pass
                def _noop_log(_msg): pass

                new_tr, _, _, _ = _run_translation(
                    chunks, target_lang, doc_context,
                    timer_ph, prog_ph, _noop_stats, _noop_log,
                    len(translatable), prefix_label=uploaded.name[:20],
                    glossary=glossary,
                    source_lang=source_lang,
                )
                translations.update(new_tr)
                timer_ph.empty()
                prog_ph.empty()

            for _pass_no in range(1, AUTO_RESCAN_PASSES + 1):
                missed = find_missed(
                    blocks, translations,
                    source_lang=source_lang,
                    target_lang=target_lang,
                )
                if not missed:
                    break
                chunks = chunk_blocks(missed)
                timer_ph = st.empty()
                prog_ph = st.empty()

                def _noop_stats2(*_a, **_kw): pass
                def _noop_log2(_msg): pass

                new_tr, _, _, _ = _run_translation(
                    chunks, target_lang, doc_context,
                    timer_ph, prog_ph, _noop_stats2, _noop_log2,
                    len(missed), prefix_label=f"{uploaded.name[:14]} scan",
                    glossary=glossary,
                    source_lang=source_lang,
                )
                translations.update(new_tr)
                timer_ph.empty()
                prog_ph.empty()

            out_bytes = apply_translations(docx_bytes, blocks, translations)
            results[uploaded.name] = out_bytes
            file_status[uploaded.name] = "✅ Xong"
        except Exception as e:
            file_status[uploaded.name] = f"❌ {str(e)[:40]}"

        status_ph.table({
            "File":       list(file_status.keys()),
            "Trạng thái": list(file_status.values()),
        })

    # Lưu kết quả vào session_state để hiện summary + download persist qua rerun.
    # Sanitize tên file (P3.7): basename only, không cho `..`/`/`/`\`, đảm bảo unique.
    if results:
        buf = io.BytesIO()
        used: set[str] = set()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for fname, data in results.items():
                safe = _safe_zip_name(fname, lang_word, used)
                used.add(safe)
                zf.writestr(safe, data)
        buf.seek(0)
        # zip_name: chỉ chứa ascii an toàn, slug từ lang_word
        import re as _re_slug
        lang_slug = _re_slug.sub(r"[^A-Za-z0-9_-]+", "", lang_word)[:8] or "out"
        st.session_state["word_batch_result"] = {
            "zip_bytes": buf.getvalue(),
            "zip_name":  f"translated_{lang_slug}.zip",
            "status":    dict(file_status),
            "lang":      lang_word,
            "ok_count":  sum(1 for v in file_status.values() if v.startswith("✅")),
            "fail_count": sum(1 for v in file_status.values() if v.startswith("❌")),
        }
    else:
        # Mọi file đều fail — vẫn lưu status để hiển thị
        st.session_state["word_batch_result"] = {
            "zip_bytes": None, "zip_name": None,
            "status":    dict(file_status),
            "lang":      lang_word,
            "ok_count":  0,
            "fail_count": sum(1 for v in file_status.values() if v.startswith("❌")),
        }


def _render_batch_result():
    """Render summary + download cho batch result đã chạy xong (persist qua rerun)."""
    res = st.session_state.get("word_batch_result")
    if not res:
        return

    n = len(res["status"])
    if res["fail_count"] == 0:
        st.success(f"✅ Batch xong: {res['ok_count']}/{n} file dịch thành công.")
    elif res["ok_count"] == 0:
        st.error(f"❌ Batch fail: 0/{n} file thành công.")
    else:
        st.warning(f"⚠️ Batch xong: {res['ok_count']}/{n} thành công, "
                   f"{res['fail_count']}/{n} fail.")

    # Status table
    st.table({
        "File":       list(res["status"].keys()),
        "Trạng thái": list(res["status"].values()),
    })

    cols = st.columns([3, 1])
    if res["zip_bytes"]:
        cols[0].download_button(
            "⬇️ Tải ZIP tất cả file đã dịch",
            data=res["zip_bytes"],
            file_name=res["zip_name"],
            mime="application/zip",
            use_container_width=True,
            key="word_batch_zip_dl",
        )
    if cols[1].button("🗑 Xoá kết quả batch", use_container_width=True,
                      key="word_batch_clear"):
        st.session_state.pop("word_batch_result", None)
        st.rerun()


def render():
    """Entry point — gọi từ `streamlit_app.py`."""
    uploaded_files = st.file_uploader(
        "📁 Tải lên file Word (.docx)",
        type=["docx"],
        accept_multiple_files=True,
        key="word_upload",
    )
    # Backward compat: treat single file the same as before
    uploaded_docx = uploaded_files[0] if uploaded_files and len(uploaded_files) == 1 else None

    direction_label = st.radio(
        "🌐 Hướng dịch",
        [d[0] for d in TRANSLATION_DIRECTIONS],
        horizontal=True,
        key="word_direction",
    )
    _, source_lang, target_lang = _resolve_langs(direction_label)
    # Backward-compat label cho session_state["word_lang"]:
    # giữ "Tiếng Anh"/"Tiếng Việt" để code download/filename không phải đổi nhiều.
    lang_word = "Tiếng Anh" if target_lang == "English" else "Tiếng Việt"

    st.divider()

    # ── BATCH mode: multiple files ────────────────────────────────────────
    if uploaded_files and len(uploaded_files) > 1:
        if st.button("📚  Dịch tất cả file (batch)", use_container_width=True,
                     type="primary", key="word_batch_btn"):
            st.session_state.pop("word_batch_result", None)
            _run_batch(uploaded_files, lang_word, source_lang, target_lang)
            st.rerun()
        elif "word_batch_result" not in st.session_state:
            st.info(f"📁 Đã chọn {len(uploaded_files)} file — bấm **Dịch tất cả file (batch)** để dịch.")
        if "word_batch_result" in st.session_state:
            st.divider()
            _render_batch_result()
        return  # don't show single-file UI when batch

    # Single-file mode: nếu batch result còn → clear (user vừa giảm xuống 1 file)
    if uploaded_files and len(uploaded_files) <= 1 and "word_batch_result" in st.session_state:
        st.session_state.pop("word_batch_result", None)

    # ── NÚT DỊCH (1 click → phân tích + dịch luôn) ───────────────────────
    if st.button(
        "🚀  Dịch",
        disabled=(uploaded_docx is None),
        use_container_width=True, type="primary", key="word_translate_btn",
        help="Phân tích + dịch ngay với cài đặt mặc định",
    ):
        _clear_state()
        _run_analysis(uploaded_docx, lang_word, source_lang, target_lang)

    # ── PHASE 1 → 2: phân tích xong thì dịch luôn ──────────────────────
    if "word_analysis" in st.session_state:
        st.divider()
        _run_full_translation()
        st.rerun()

    # ── PHASE 2 RESULT: download + rescan + OCR ─────────────────────────
    if "word_translations" in st.session_state:
        st.divider()

        val          = st.session_state.get("word_validation")
        blocks       = st.session_state["word_blocks"]
        translations = st.session_state["word_translations"]
        source_lang  = st.session_state.get("word_source_lang") or _resolve_langs()[1]
        target_lang  = st.session_state.get("word_target_lang") or LANG_EN[st.session_state["word_lang"]]
        skip_roles   = st.session_state.get("word_skip_roles", set(NO_TRANSLATE_ROLES))
        candidates   = [b for b in blocks if b["role"] not in skip_roles]
        missed       = find_missed(
            candidates, translations,
            source_lang=source_lang,
            target_lang=target_lang,
        )

        # 1) Validation errors (chỉ show khi LỖI — user không nên download)
        if val and not val["valid"]:
            st.error("❌ Output có lỗi — KHÔNG nên download:")
            for err in val["errors"]:
                st.code(err)
            for warn in val.get("warnings", []):
                st.warning(warn)

        # 2) DOWNLOAD — top, tách biệt, primary
        translated_bytes = _get_cached_translated_bytes()
        out_name = st.session_state["word_filename"].replace(
            ".docx", f"_translated_{st.session_state['word_lang'][:2]}.docx"
        )
        st.download_button(
            label="⬇️  Tải Word đã dịch",
            data=translated_bytes,
            file_name=out_name,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
            type="primary",
        )

        # 3) Rescan button — header/footer giờ dịch chung với body, không tách
        rescan_clicked = st.button(
            f"🔍  Quét bỏ sót ({len(missed)})",
            disabled=(len(missed) == 0),
            use_container_width=True, key="word_rescan_btn",
            help="Dịch lại các đoạn bị API bỏ sót / fail",
        )

        # 4) OCR — section riêng, tách biệt (user dùng thường xuyên)
        _render_ocr_section()

        # 5) Tiny status — 1 dòng caption, không expander
        summary_line = st.session_state.get("word_summary", "")
        if summary_line:
            st.caption(summary_line)

        if rescan_clicked:
            _run_rescan()
            st.rerun()

    elif not uploaded_files:
        st.info("👆 Vui lòng upload file Word (.docx) để bắt đầu")


# ══════════════════════════════════════════════════════════════════════════════
# OCR SECTION (P2.1, P2.2, P3.1, U1-U4)
# ══════════════════════════════════════════════════════════════════════════════
def _ocr_state() -> dict:
    """Lazy-init `word_ocr_state` (P0.2 — single source of truth)."""
    if "word_ocr_state" not in st.session_state:
        st.session_state["word_ocr_state"] = {
            "occurrences":   [],
            "results":       {},
            "selection":     {},   # occ_id → bool (đưa vào output)
            "keep_original": {},   # occ_id → bool (giữ ảnh gốc)
            "edited":        {},   # occ_id → str (bản dịch user sửa)
            "estimate":      None,
            "phase":         "idle",     # idle | done
            "export":        None,       # {"bytes": ..., "name": ...} sau khi xuất
            "export_error":  None,       # lỗi xuất (nếu có) — show ở UI
        }
    return st.session_state["word_ocr_state"]


def _reset_ocr_state():
    st.session_state.pop("word_ocr_state", None)


def _render_ocr_section():
    from word_backend import extract_image_occurrences, estimate_ocr_cost

    state = _ocr_state()
    phase = state["phase"]

    with st.expander("🖼  OCR & dịch text trong ảnh", expanded=False):
        # ── Phase: idle → quét + OCR luôn (1 click, không confirm) ───
        if phase == "idle":
            st.caption(
                "Quét tất cả ảnh trong DOCX, OCR chữ và dịch sang ngôn ngữ đích. "
                "Bấm nút bên dưới để chạy ngay — chi phí thực tế hiển thị sau khi xong."
            )
            if st.button("🚀 Dịch OCR ảnh trong file",
                         key="word_ocr_run_btn",
                         type="primary",
                         use_container_width=True):
                occs = extract_image_occurrences(st.session_state["word_docx_bytes"])
                state["occurrences"]   = occs
                state["estimate"]      = estimate_ocr_cost(occs)
                state["selection"]     = {o["id"]: (len(o["data"]) >= 5_000) for o in occs}
                state["keep_original"] = {o["id"]: True for o in occs}
                if not occs:
                    st.info("Không có ảnh trong DOCX.")
                    state["phase"] = "done"
                else:
                    _run_ocr(state)
                st.rerun()

        # ── Phase: done → review + export ─────────────────────────────
        elif phase == "done":
            _render_ocr_review(state)


def _run_ocr(state: dict):
    """Chạy OCR cho occurrences trong state, lưu results + actual cost."""
    from word_backend import ocr_and_translate_images

    occs = state["occurrences"]
    if not occs:
        state["phase"] = "done"
        return

    n_to_ocr = state["estimate"]["n_to_ocr"]
    progress_bar = st.progress(0, text=f"Đang OCR 0/{n_to_ocr} ảnh...")

    def _on_progress(done, total):
        progress_bar.progress(min(done / max(total, 1), 1.0),
                              text=f"Đang OCR {done}/{total} ảnh...")

    source_lang = st.session_state.get("word_source_lang") or _resolve_langs()[1]
    target_lang = (st.session_state.get("word_target_lang")
                   or LANG_EN[st.session_state["word_lang"]])
    glossary    = st.session_state.get("word_glossary")
    subdomains  = st.session_state.get("word_subdomains") or set()

    results = ocr_and_translate_images(
        get_client(), occs,
        target_lang=target_lang,
        source_lang=source_lang,
        glossary=glossary,
        subdomains=subdomains,
        progress_callback=_on_progress,
    )
    progress_bar.empty()

    state["results"] = results
    # Init `edited` từ translation; selection tick những ảnh có text.
    state["edited"]    = {o["id"]: results.get(o["id"], {}).get("translation", "")
                          for o in occs}
    state["selection"] = {
        o["id"]: bool(results.get(o["id"], {}).get("has_text"))
        for o in occs
    }
    state["phase"] = "done"


def _render_ocr_review(state: dict):
    """P2.1+P2.2+U2+U3+U4: review từng occurrence + chọn mode + download."""
    occs    = state["occurrences"]
    results = state["results"]
    total   = results.get("_total", {})

    # Summary actual cost
    sc1, sc2, sc3, sc4 = st.columns(4)
    sc1.metric("Ảnh OCR", f"{total.get('n_called', 0):,}")
    sc2.metric("Tokens", f"{total.get('tok_in', 0):,} in / {total.get('tok_out', 0):,} out")
    sc3.metric("USD", f"${total.get('usd', 0):.4f}")
    sc4.metric("VND", f"{total.get('vnd', 0):,.0f}")

    # Group: có text / không text + lỗi
    with_text = [o for o in occs if results.get(o["id"], {}).get("has_text")]
    no_text   = [o for o in occs
                 if not results.get(o["id"], {}).get("has_text")
                 and not results.get(o["id"], {}).get("error")]
    errors    = [o for o in occs if results.get(o["id"], {}).get("error")]

    # Bulk select buttons
    if with_text:
        b1, b2, _ = st.columns([1, 1, 2])
        if b1.button("✅ Chọn tất cả ảnh có chữ", key="word_ocr_select_all"):
            for o in with_text:
                state["selection"][o["id"]] = True
            st.rerun()
        if b2.button("⬜ Bỏ chọn tất cả", key="word_ocr_deselect_all"):
            for o in with_text:
                state["selection"][o["id"]] = False
            st.rerun()

    # Output mode (U3)
    st.divider()

    # Per-image review (P2.1)
    if with_text:
        st.markdown(f"##### 📷 Ảnh có chữ — {len(with_text)}")
        for o in with_text:
            r = results.get(o["id"], {})
            with st.container(border=True):
                col_img, col_info = st.columns([1, 2])
                with col_img:
                    st.image(o["data"], width=220,
                             caption=f"{o['filename']}  ·  occ #{o['occurrence_index']}")
                    st.caption(f"Part: `{o['doc_part'].rsplit('/', 1)[-1]}`  ·  "
                               f"rId: `{o['rId']}`")
                    st.caption(f"💵 ${r.get('usd', 0):.4f}  ·  "
                               f"{r.get('tok_in', 0)}+{r.get('tok_out', 0)} tok"
                               + (f"  ·  conf {r.get('confidence', 0):.2f}"
                                  if r.get('confidence') else ""))
                with col_info:
                    state["selection"][o["id"]] = st.checkbox(
                        "📎 Đưa ảnh này vào file xuất",
                        value=state["selection"].get(o["id"], True),
                        key=f"word_ocr_pick_{o['id']}",
                    )
                    state["keep_original"][o["id"]] = st.checkbox(
                        "🖼 Giữ ảnh gốc",
                        value=state["keep_original"].get(o["id"], True),
                        key=f"word_ocr_keep_{o['id']}",
                    )
                    with st.expander("📜 OCR (gốc)", expanded=False):
                        st.text(r.get("ocr", ""))
                    state["edited"][o["id"]] = st.text_area(
                        "✏️ Bản dịch — chỉnh trước khi xuất:",
                        value=state["edited"].get(o["id"], r.get("translation", "")),
                        height=120,
                        key=f"word_ocr_edit_{o['id']}",
                    )

    # No-text group + errors (collapsed by default — P2.1: nhóm riêng, không chọn)
    if no_text or errors:
        with st.expander(
            f"⚠️ Không phát hiện chữ / lỗi — {len(no_text)} không text · {len(errors)} lỗi",
            expanded=False,
        ):
            for o in no_text[:30]:
                r = results.get(o["id"], {})
                with st.container(border=True):
                    c1, c2 = st.columns([1, 3])
                    c1.image(o["data"], width=120, caption=o["filename"])
                    c2.caption(f"Không phát hiện chữ  ·  ${r.get('usd', 0):.4f}")
            for o in errors[:30]:
                r = results.get(o["id"], {})
                with st.container(border=True):
                    c1, c2 = st.columns([1, 3])
                    c1.image(o["data"], width=120, caption=o["filename"])
                    c2.warning(f"Lỗi OCR: {r.get('error', '?')[:150]}")

    # Export (U4)
    st.divider()
    selected = [o for o in occs if state["selection"].get(o["id"])]
    st.caption(
        f"📤 Sẽ xuất: **{len(selected)} ảnh** — chèn bản dịch dưới mỗi ảnh "
        f"· chi phí actual: **${total.get('usd', 0):.4f}** ≈ {total.get('vnd', 0):,.0f} VND"
    )

    if not selected:
        st.info("Tick chọn ít nhất 1 ảnh để xuất.")
    else:
        if st.button("⬇️ Xuất DOCX với OCR", type="primary",
                     use_container_width=True, key="word_ocr_export_btn"):
            state["export"]       = None
            state["export_error"] = None
            try:
                out_bytes, suffix = _build_ocr_export_bytes(state, selected)
                out_name = st.session_state["word_filename"].replace(
                    ".docx",
                    f"_translated_{st.session_state['word_lang'][:2]}{suffix}.docx",
                )
                state["export"] = {"bytes": out_bytes, "name": out_name}
            except Exception as e:
                import traceback
                state["export_error"] = f"{type(e).__name__}: {e}\n\n{traceback.format_exc()}"
            st.rerun()

    # Persist download / error qua rerun
    if state.get("export_error"):
        st.error("❌ Xuất DOCX thất bại:")
        st.code(state["export_error"])
    elif state.get("export"):
        exp = state["export"]
        st.success(f"✅ Đã tạo file `{exp['name']}` — bấm để tải.")
        st.download_button(
            label=f"📥 Tải {exp['name']}",
            data=exp["bytes"],
            file_name=exp["name"],
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
            key="word_ocr_download_final",
        )

    if st.button("🔁 Quét lại OCR", key="word_ocr_restart"):
        _reset_ocr_state()
        st.rerun()


def _build_ocr_export_bytes(state: dict, selected: list[dict]) -> tuple[bytes, str]:
    """Build DOCX bytes cho OCR export — caption mode (chèn bản dịch dưới ảnh).
    Trả (bytes, suffix). Raise Exception nếu fail — caller hiển thị lỗi.
    """
    from word_backend import insert_ocr_captions_into_docx

    base_bytes   = _get_cached_translated_bytes()
    selected_ids = [o["id"] for o in selected]
    edited       = {oid: state["edited"].get(oid, "") for oid in selected_ids}
    remove_ids   = [o["id"] for o in selected
                    if not state["keep_original"].get(o["id"], True)]

    out_bytes = insert_ocr_captions_into_docx(
        base_bytes, state["occurrences"], state["results"],
        selected_ids=selected_ids,
        edited_translations=edited,
        remove_original_ids=remove_ids,
    )
    return out_bytes, "_ocr_caption"
