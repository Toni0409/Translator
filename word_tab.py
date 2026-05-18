"""
Streamlit UI cho tab Word — 2-phase flow.

Phase 1 (Phân tích):
  extract → glossary build → show stats + glossary editor + TM preview

Phase 2 (Dịch):
  TM lookup → chunk còn lại → translate parallel → merge TM + new → apply
"""
import time
import threading

import pandas as pd
import streamlit as st

from config import (
    LANGUAGES, LANG_EN, MAX_WORD_WORKERS, NO_TRANSLATE_ROLES,
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
    tm_lookup, tm_store,
    checkpoint_save, checkpoint_load, checkpoint_clear,
    export_bilingual_docx,
)


# Keys session_state cho tab Word — tập trung 1 chỗ để dễ clear
SS_KEYS = (
    "word_blocks", "word_translations", "word_docx_bytes",
    "word_filename", "word_lang", "word_doc_context",
    "word_tok_in", "word_tok_out", "word_elapsed", "word_num_chunks",
    "word_summary", "word_glossary", "word_analysis",
)
# Note: "word_tm" intentionally KHÔNG ở đây — persist xuyên suốt nhiều docs/session


# ══════════════════════════════════════════════════════════════════════════════
# UI HELPERS
# ══════════════════════════════════════════════════════════════════════════════
def _ensure_tm() -> dict:
    if "word_tm" not in st.session_state:
        st.session_state["word_tm"] = {}
    return st.session_state["word_tm"]


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
                  tok_in: int, tok_out: int, tm_added: int = 0):
    usd, vnd = calc_cost(tok_in, tok_out)
    prog.progress(100, text=f"✅ {label} xong!")
    timer_ph.markdown(
        timer_done_html(elapsed, f"{label} {count:,} đoạn xong!"),
        unsafe_allow_html=True,
    )
    add_log("─" * 44)
    add_log(f"🎉 Xong {count:,} đoạn trong {elapsed:.1f}s")
    if tm_added:
        add_log(f"💾 TM: lưu thêm {tm_added} entry")
    add_log(f"💵 Chi phí: ${usd:.4f} USD ≈ {vnd:,.0f} VND")


# ══════════════════════════════════════════════════════════════════════════════
# CORE TRANSLATION LOOP
# ══════════════════════════════════════════════════════════════════════════════
def _run_translation(chunks, target_lang, doc_context,
                     timer_ph, prog_ph, render_stats, add_log,
                     total_blocks, prefix_label="Dịch",
                     glossary: dict | None = None,
                     custom_rules: dict | None = None,
                     on_chunk_done=None):
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
              MAX_WORD_WORKERS, glossary, custom_rules),
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


# ══════════════════════════════════════════════════════════════════════════════
# PHASE 1 — PHÂN TÍCH (extract + glossary build)
# ══════════════════════════════════════════════════════════════════════════════
def _run_analysis(uploaded_docx, lang_word):
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
        add_log(f"   • Header/Footer (skip):{stats['hf_total']:,}")
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

        target_lang = LANG_EN[lang_word]

        # TM preview
        tm           = _ensure_tm()
        tm_cached, _ = tm_lookup(translatable, tm, target_lang)
        if tm_cached:
            add_log(f"💾 TM: {len(tm_cached):,}/{len(translatable):,} đoạn dùng cache (skip API)")
        else:
            add_log(f"💾 TM: 0 hit (TM hiện có {len(tm):,} entry)")

        # Glossary build
        add_log("📚 Build glossary từ thuật ngữ lặp lại...")
        client   = get_client()
        glossary = build_glossary(client, blocks, target_lang)
        if glossary:
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
            "tm_hits":       len(tm_cached),
            "doc_context":   build_doc_context(blocks),
            "glossary":      glossary,
            "docx_bytes":    docx_bytes,
            "filename":      uploaded_docx.name,
            "lang":          lang_word,
            "target_lang":   target_lang,
        }
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
# PHASE 1 — RENDER (stats + glossary editor + Dịch button)
# ══════════════════════════════════════════════════════════════════════════════
def _render_analysis_panel():
    a = st.session_state["word_analysis"]

    st.markdown("### 📊 Kết quả phân tích")
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.markdown(stat_box_html(f"{a['stats']['body']:,}", "Body cần dịch"),
                unsafe_allow_html=True)
    c2.markdown(stat_box_html(f"{a['stats']['hf_total']:,}", "Header/Footer"),
                unsafe_allow_html=True)
    c3.markdown(stat_box_html(f"{a['textbox_cnt']:,}", "Text-box"),
                unsafe_allow_html=True)
    c4.markdown(stat_box_html(f"{a['table_cnt']:,}", "Cell bảng"),
                unsafe_allow_html=True)
    fn_total = a.get("footnote_cnt", 0) + a.get("endnote_cnt", 0)
    c5.markdown(stat_box_html(f"{fn_total:,}", "Footnote/Endnote"),
                unsafe_allow_html=True)

    comment_cnt = a.get("comment_cnt", 0)
    if comment_cnt > 0:
        st.info(f"💬 **{comment_cnt:,} comment** trong tài liệu — sẽ được dịch.")

    if a["tm_hits"] > 0:
        pct = 100 * a["tm_hits"] / max(len(a["translatable"]), 1)
        st.success(
            f"💾 **Translation Memory**: {a['tm_hits']:,}/{len(a['translatable']):,} đoạn "
            f"({pct:.1f}%) sẽ dùng cache cũ — không gọi API."
        )

    # Glossary editor
    with st.expander(
        f"📚 Glossary ({len(a['glossary'])} thuật ngữ) — bấm để review / sửa",
        expanded=False,
    ):
        if not a["glossary"]:
            st.info("Không có thuật ngữ lặp lại đủ ngưỡng (≥3 lần). Vẫn dịch được — chỉ là không có guarantee consistency.")
        else:
            st.caption(
                "💡 Sửa bản dịch để bắt buộc terminology cụ thể. "
                "Xóa cell **Bản dịch** (để trống) để loại entry khỏi glossary. "
                "Có thể thêm dòng mới."
            )
            df = pd.DataFrame([
                {"Thuật ngữ gốc": k, "Bản dịch": v}
                for k, v in a["glossary"].items()
            ])
            edited = st.data_editor(
                df,
                column_config={
                    "Thuật ngữ gốc": st.column_config.TextColumn(width="medium"),
                    "Bản dịch":      st.column_config.TextColumn(width="medium"),
                },
                use_container_width=True, hide_index=True,
                num_rows="dynamic",
                key="word_glossary_editor",
            )
            # Sync edits back
            new_gloss = {}
            for _, row in edited.iterrows():
                en = (row.get("Thuật ngữ gốc") or "").strip() if row.get("Thuật ngữ gốc") else ""
                vi = (row.get("Bản dịch") or "").strip() if row.get("Bản dịch") else ""
                if en and vi:
                    new_gloss[en] = vi
            a["glossary"] = new_gloss
            st.session_state["word_analysis"] = a

    with st.expander("⚙️ Custom rules per role (tuỳ chọn)", expanded=False):
        st.caption(
            "Định nghĩa rule riêng cho từng role — sẽ được inject vào mọi chunk prompt. "
            "Để trống = dùng rule mặc định."
        )
        custom_roles = ["section_heading", "paragraph", "bullet", "table_cell",
                        "note", "textbox", "footnote", "endnote", "comment"]
        stored = a.get("custom_rules", {})
        new_rules = {}
        for role in custom_roles:
            val = st.text_input(
                f"{ROLE_LABEL.get(role, role)}:",
                value=stored.get(role, ""),
                key=f"word_custom_rule_{role}",
                placeholder="vd: keep concise, max 8 words",
            )
            if val.strip():
                new_rules[role] = val.strip()
        a["custom_rules"] = new_rules
        st.session_state["word_analysis"] = a

    # Dịch button — nói rõ còn bao nhiêu sau TM
    n_remain = len(a["translatable"]) - a["tm_hits"]
    label = (f"▶  Dịch {n_remain:,} đoạn"
             if n_remain > 0
             else f"▶  Apply {a['tm_hits']:,} TM hit (không gọi API)")
    if st.button(label, use_container_width=True, type="primary",
                 key="word_translate_phase2_btn"):
        _run_full_translation()
        st.rerun()


# ══════════════════════════════════════════════════════════════════════════════
# PHASE 2 — DỊCH (consume analysis, run translation, merge TM + new)
# ══════════════════════════════════════════════════════════════════════════════
def _run_full_translation():
    """Phase 2: dùng analysis đã extract + glossary đã edit → translate."""
    a = st.session_state["word_analysis"]
    blocks       = a["blocks"]
    translatable = a["translatable"]
    target_lang  = a["target_lang"]
    doc_context  = a["doc_context"]
    glossary     = a["glossary"]
    custom_rules = a.get("custom_rules") or {}
    docx_bytes   = a["docx_bytes"]
    tm           = _ensure_tm()

    timer_ph, prog, render_stats, add_log = _make_job_ui(
        "### 📊 Tiến độ dịch", "Đoạn văn", "### 📋 Nhật ký",
    )

    try:
        cached, remaining = tm_lookup(translatable, tm, target_lang)
        if cached:
            add_log(f"💾 TM hit: {len(cached):,}/{len(translatable):,} đoạn — skip API")
        add_log(f"📝 Sẽ gọi API cho {len(remaining):,} đoạn")

        translations = dict(cached)   # khởi đầu với TM hits

        # Checkpoint recovery: restore partial translations from previous run
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
        tm_added = 0

        if remaining:
            chunks     = chunk_blocks(remaining)
            num_chunks = len(chunks)
            avg_chars  = sum(len(b["text"]) for b in remaining) // max(num_chunks, 1)
            add_log(f"📦 Chia thành {num_chunks} chunk (~{avg_chars:,} chars/chunk)")
            if glossary:
                add_log(f"📖 Áp dụng glossary {len(glossary)} thuật ngữ")
            add_log(f"⚡ Dịch song song {MAX_WORD_WORKERS} luồng")

            def _save_ckpt(tr_so_far):
                checkpoint_save(docx_bytes, target_lang, {**cached, **tr_so_far})

            new_translations, tok_in, tok_out, elapsed = _run_translation(
                chunks, target_lang, doc_context,
                timer_ph, prog, render_stats, add_log,
                len(remaining), prefix_label="Dịch",
                glossary=glossary,
                custom_rules=custom_rules or None,
                on_chunk_done=_save_ckpt,
            )
            translations.update(new_translations)
            tm_added = tm_store(remaining, new_translations, tm, target_lang)
            checkpoint_clear(docx_bytes, target_lang)
            add_log(f"🤖 Model: {get_working_model()}")
        else:
            timer_ph.markdown(
                timer_done_html(0, f"Apply {len(cached):,} TM hit — không gọi API!"),
                unsafe_allow_html=True,
            )
            prog.progress(100, text="✅ Full TM hit!")

        usd, vnd = calc_cost(tok_in, tok_out)
        render_stats(len(translatable), num_chunks, num_chunks, tok_in, tok_out)
        prog.progress(100, text="✅ Hoàn thành!")
        timer_ph.markdown(
            timer_done_html(elapsed, f"Dịch xong {len(translatable):,} đoạn!"),
            unsafe_allow_html=True,
        )
        add_log("─" * 44)
        add_log(f"🎉 Xong {len(translatable):,} đoạn "
                f"(💾 {len(cached):,} TM + 🔄 {len(remaining):,} API) — {elapsed:.1f}s")
        if tm_added:
            add_log(f"💾 TM: lưu thêm {tm_added} entry — tổng {len(tm):,}")
        if tok_in or tok_out:
            add_log(f"💰 Token: {tok_in:,} in + {tok_out:,} out")
        add_log(f"💵 Chi phí: ${usd:.4f} USD ≈ {vnd:,.0f} VND")

        st.session_state["word_blocks"]       = blocks
        st.session_state["word_translations"] = translations
        st.session_state["word_docx_bytes"]   = docx_bytes
        st.session_state["word_filename"]     = a["filename"]
        st.session_state["word_lang"]         = a["lang"]
        st.session_state["word_doc_context"]  = doc_context
        st.session_state["word_tok_in"]       = tok_in
        st.session_state["word_tok_out"]      = tok_out
        st.session_state["word_elapsed"]      = elapsed
        st.session_state["word_num_chunks"]   = num_chunks
        st.session_state["word_glossary"]     = glossary
        st.session_state["word_custom_rules"] = custom_rules
        st.session_state["word_summary"]      = (
            f"✅ Dịch xong {len(translatable):,} đoạn  "
            f"(💾 {len(cached):,} TM + 🔄 {len(remaining):,} API)  "
            f"|  {elapsed:.1f}s  |  ${usd:.4f} USD ≈ {vnd:,.0f} VND"
        )
        st.session_state["word_translations_version"] = 1
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
    """Generic partial-translation runner (rescan / H-F). Có TM lookup."""
    if not missed:
        st.info("✨ Không có đoạn nào cần xử lý!")
        return

    doc_context  = st.session_state["word_doc_context"]
    target_lang  = LANG_EN[st.session_state["word_lang"]]
    glossary     = st.session_state.get("word_glossary")
    custom_rules = st.session_state.get("word_custom_rules")
    tm           = _ensure_tm()
    translations = st.session_state["word_translations"]

    cached, remaining = tm_lookup(missed, tm, target_lang)

    timer_ph, prog, render_stats, add_log = _make_job_ui(
        heading, stat_label, log_heading,
    )
    if cached:
        add_log(f"💾 TM hit: {len(cached):,}/{len(missed):,} đoạn dùng cache")
    add_log(f"📝 Cần gọi API cho {len(remaining):,} đoạn")

    # Apply TM hits ngay
    if cached:
        translations.update(cached)

    if not remaining:
        # Full TM hit — không cần API
        st.session_state["word_translations"] = translations
        st.session_state["word_translations_version"] = (
            st.session_state.get("word_translations_version", 0) + 1
        )
        prog.progress(100, text="✅ Full TM hit!")
        timer_ph.markdown(
            timer_done_html(0, f"Apply {len(cached):,} TM hit!"),
            unsafe_allow_html=True,
        )
        add_log(f"🎉 Toàn bộ {len(missed):,} đoạn dùng TM — không gọi API")
        st.success(f"✅ {len(missed):,} đoạn từ TM cache — chi phí $0!")
        return

    chunks = chunk_blocks(remaining)
    add_log(f"📦 Chia thành {len(chunks)} chunk — {MAX_WORD_WORKERS} luồng")

    try:
        new_translations, tok_in, tok_out, elapsed = _run_translation(
            chunks, target_lang, doc_context,
            timer_ph, prog, render_stats, add_log,
            len(remaining), prefix_label=label,
            glossary=glossary,
            custom_rules=custom_rules or None,
        )
        translations.update(new_translations)
        tm_added = tm_store(remaining, new_translations, tm, target_lang)
        st.session_state["word_translations"] = translations
        st.session_state["word_tok_in"]      += tok_in
        st.session_state["word_tok_out"]     += tok_out
        st.session_state["word_translations_version"] = (
            st.session_state.get("word_translations_version", 0) + 1
        )
        _finalize_job(timer_ph, prog, add_log, elapsed,
                      len(missed), label, tok_in, tok_out, tm_added=tm_added)
        st.success(success_msg.format(n=len(missed)))

    except Exception as e:
        add_log(f"❌ Lỗi: {e}")
        st.error(f"❌ Lỗi: {e}")
        timer_ph.markdown(timer_error_html(str(e)), unsafe_allow_html=True)


def _run_rescan():
    blocks       = st.session_state["word_blocks"]
    translations = st.session_state["word_translations"]
    missed       = find_missed(blocks, translations)
    _run_partial(
        missed,
        label="Quét",
        heading=f"### 🔍 Quét bỏ sót — {len(missed):,} đoạn",
        log_heading="### 📋 Nhật ký quét",
        stat_label="Đoạn còn sót",
        success_msg="✅ Đã dịch thêm {n:,} đoạn bị bỏ sót!",
    )


def _run_hf_translation():
    blocks       = st.session_state["word_blocks"]
    translations = st.session_state["word_translations"]
    hf_missed    = find_missed(blocks, translations, hf_only=True)
    _run_partial(
        hf_missed,
        label="Dịch H/F",
        heading=f"### 🌐 Dịch Header & Footer — {len(hf_missed):,} đoạn",
        log_heading="### 📋 Nhật ký dịch H/F",
        stat_label="H/F đoạn",
        success_msg="✅ Đã dịch {n:,} đoạn Header/Footer!",
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
# PREVIEW & INLINE EDIT
# ══════════════════════════════════════════════════════════════════════════════
ROLE_LABEL = {
    "title":           "Title",
    "section_heading": "Heading",
    "paragraph":       "Paragraph",
    "bullet":          "Bullet",
    "table_cell":      "Cell",
    "note":            "Note",
    "toc":             "TOC",
    "textbox":         "🔲 Text-box",
    "header":          "📌 Header",
    "footer":          "📌 Footer",
    "body_repeated":   "🔁 Lặp lại",
    "footnote":        "📝 Footnote",
    "endnote":         "📝 Endnote",
    "comment":         "💬 Comment",
}


def _render_editor():
    blocks       = st.session_state["word_blocks"]
    translations = st.session_state["word_translations"]

    col_f1, col_f2 = st.columns(2)
    with col_f1:
        only_missed = st.checkbox("📍 Chỉ hiển thị đoạn chưa dịch",
                                  key="word_only_missed_filter")
    with col_f2:
        show_hf = st.checkbox("📌 Hiện cả Header/Footer",
                              value=True, key="word_show_hf_filter")

    display_blocks = []
    for b in blocks:
        is_hf = b["role"] in NO_TRANSLATE_ROLES
        if is_hf and not show_hf:
            continue
        tr = translations.get(b["id"], "")
        is_missed = (not tr) or tr == b["text"]
        if only_missed and not is_missed:
            continue
        display_blocks.append(b)

    if not display_blocks:
        st.info("✨ Tất cả đoạn đã dịch — không có gì để hiển thị.")
        return

    def _label(b):
        base = ROLE_LABEL.get(b["role"], b["role"])
        tc   = b.get("table_cell")
        if tc:
            return f"{base} (T{tc[0]}R{tc[1]}C{tc[2]})"
        return base

    df = pd.DataFrame([
        {
            "ID":       b["id"],
            "Vai trò":  _label(b),
            "Gốc":      b["text"],
            "Bản dịch": translations.get(b["id"], ""),
        }
        for b in display_blocks
    ])

    parts      = ["missed" if only_missed else "all",
                  "hf" if show_hf else "nohf"]
    version    = st.session_state.get("word_editor_version", 0)
    editor_key = f"word_editor_v{version}_{'_'.join(parts)}"

    edited = st.data_editor(
        df,
        column_config={
            "ID":       st.column_config.TextColumn("ID", disabled=True, width="small"),
            "Vai trò":  st.column_config.TextColumn("Vai trò", disabled=True, width="small"),
            "Gốc":      st.column_config.TextColumn("Tiếng gốc", disabled=True, width="large"),
            "Bản dịch": st.column_config.TextColumn("Bản dịch", width="large"),
        },
        use_container_width=True,
        hide_index=True,
        num_rows="fixed",
        height=480,
        key=editor_key,
    )

    changes = 0
    for _, row in edited.iterrows():
        bid     = row["ID"]
        new_val = row["Bản dịch"]
        if translations.get(bid) != new_val:
            translations[bid] = new_val
            changes += 1
    if changes:
        st.session_state["word_translations"] = translations
        st.session_state["word_translations_version"] = (
            st.session_state.get("word_translations_version", 0) + 1
        )
        st.caption(f"💾 Đã ghi nhận {changes} thay đổi — bấm **Tải Word đã dịch** ở trên.")


# ══════════════════════════════════════════════════════════════════════════════
# MAIN UI
# ══════════════════════════════════════════════════════════════════════════════
def _clear_state():
    for k in SS_KEYS:
        st.session_state.pop(k, None)
    st.session_state.pop("word_translated_bytes_cache", None)
    st.session_state.pop("word_translations_version", None)
    for k in list(st.session_state.keys()):
        if k.startswith("word_editor") or k in ("word_only_missed_filter",
                                                  "word_show_hf_filter",
                                                  "word_glossary_editor"):
            st.session_state.pop(k, None)
    st.session_state["word_editor_version"] = 0
    # KHÔNG xóa "word_tm" — TM persist xuyên suốt session


def render():
    """Entry point — gọi từ `streamlit_app.py`."""
    uploaded_docx = st.file_uploader("📝 Chọn file Word cần dịch",
                                      type=["docx"], key="word_uploader")
    lang_word = st.selectbox("🌐 Ngôn ngữ đích", LANGUAGES, key="word_lang_select")

    # TM status sidebar info
    tm_size = len(st.session_state.get("word_tm", {}))
    if tm_size > 0:
        cols = st.columns([3, 1])
        cols[0].caption(f"💾 Translation Memory hiện có **{tm_size:,}** entries (persist trong session)")
        if cols[1].button("🗑 Xóa TM", help="Reset Translation Memory",
                          key="word_tm_clear"):
            st.session_state.pop("word_tm", None)
            st.rerun()

    st.divider()

    # ── PHASE 1: Phân tích ────────────────────────────────────────────────
    col_a, col_b = st.columns(2)
    with col_a:
        analyze_clicked = st.button(
            "🔬  Phân tích + Glossary",
            disabled=(uploaded_docx is None),
            use_container_width=True, key="word_analyze",
            help="Extract paragraphs + build glossary từ thuật ngữ lặp lại",
        )
    with col_b:
        # Quick mode: phân tích + dịch luôn (skip glossary review)
        quick_clicked = st.button(
            "⚡  Phân tích & dịch luôn",
            disabled=(uploaded_docx is None),
            use_container_width=True, key="word_quick",
            help="Dùng glossary auto-build, không cần review",
        )

    if analyze_clicked:
        _clear_state()
        _run_analysis(uploaded_docx, lang_word)

    if quick_clicked:
        _clear_state()
        _run_analysis(uploaded_docx, lang_word)
        # Sau analysis xong, session_state["word_analysis"] đã có
        # Trigger Phase 2 ngay — nhưng cần rerun trước để UI hiển thị analysis
        # Quick mode set flag → Phase 1 done → render() check flag → run Phase 2
        st.session_state["word_quick_mode"] = True

    # ── PHASE 1 RESULT: stats + glossary editor + Dịch button ───────────
    if "word_analysis" in st.session_state:
        st.divider()
        # Quick mode: trigger Phase 2 ngay
        if st.session_state.pop("word_quick_mode", False):
            _run_full_translation()
            st.rerun()
        else:
            _render_analysis_panel()

    # ── PHASE 2 RESULT: download + rescan + H/F + editor ────────────────
    if "word_translations" in st.session_state:
        st.divider()
        st.success(st.session_state["word_summary"])

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
        )

        bilingual_bytes = export_bilingual_docx(
            st.session_state["word_docx_bytes"],
            st.session_state["word_blocks"],
            st.session_state["word_translations"],
        )
        bi_name = st.session_state["word_filename"].replace(
            ".docx", f"_bilingual_{st.session_state['word_lang'][:2]}.docx"
        )
        st.download_button(
            label="📊  Tải DOCX so sánh song ngữ",
            data=bilingual_bytes,
            file_name=bi_name,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
            key="word_dl_bilingual",
        )

        blocks       = st.session_state["word_blocks"]
        translations = st.session_state["word_translations"]
        missed_body  = find_missed(blocks, translations, hf_only=False)
        missed_hf    = find_missed(blocks, translations, hf_only=True)
        total_hf     = sum(1 for b in blocks if b["role"] in NO_TRANSLATE_ROLES)

        rescan_clicked = hf_clicked = False
        if total_hf > 0:
            col_rs, col_hf = st.columns(2)
            with col_rs:
                rescan_clicked = st.button(
                    f"🔍  Quét bỏ sót ({len(missed_body)})",
                    disabled=(len(missed_body) == 0),
                    use_container_width=True, key="word_rescan_btn",
                    help="Dịch lại các đoạn body API fail (TM auto)",
                )
            with col_hf:
                hf_clicked = st.button(
                    f"🌐  Dịch Header / Footer ({len(missed_hf)}/{total_hf})",
                    disabled=(len(missed_hf) == 0),
                    use_container_width=True, key="word_hf_btn",
                    help="Dịch Header, Footer và đoạn lặp trong body (TM auto)",
                )
        else:
            rescan_clicked = st.button(
                f"🔍  Quét bỏ sót ({len(missed_body)})",
                disabled=(len(missed_body) == 0),
                use_container_width=True, key="word_rescan_btn",
            )

        if len(missed_body) > 0:
            st.warning(f"⚠️ Còn {len(missed_body):,} đoạn body chưa dịch. "
                       f"Bấm **Quét bỏ sót** để dịch lại.")
        if total_hf > 0 and len(missed_hf) > 0:
            st.info(f"📌 {len(missed_hf):,}/{total_hf:,} Header/Footer chưa dịch. "
                    f"Bấm **Dịch Header/Footer** nếu muốn dịch.")

        with st.expander("📋  Xem & sửa bản dịch inline", expanded=False):
            _render_editor()

        if rescan_clicked:
            _run_rescan()
            st.session_state["word_editor_version"] = (
                st.session_state.get("word_editor_version", 0) + 1
            )
            st.rerun()
        if hf_clicked:
            _run_hf_translation()
            st.session_state["word_editor_version"] = (
                st.session_state.get("word_editor_version", 0) + 1
            )
            st.rerun()

    elif not uploaded_docx:
        st.info("👆 Vui lòng upload file Word (.docx) để bắt đầu")
