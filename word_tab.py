"""
Streamlit UI cho tab Word.

Flow:
1. Upload .docx → extract → chunk adaptive
2. Translate song song (ThreadPoolExecutor) + per-chunk retry
3. Save state vào session_state (blocks, translations, docx_bytes, ...)
4. Hiển thị: success summary + download + rescan missed + edit inline
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
    extract_docx_blocks, build_doc_context, chunk_blocks,
    translate_parallel, apply_translations, find_missed,
    get_working_model, count_by_role,
)


# Keys session_state cho tab Word — tập trung 1 chỗ để dễ clear
SS_KEYS = (
    "word_blocks", "word_translations", "word_docx_bytes",
    "word_filename", "word_lang", "word_doc_context",
    "word_tok_in", "word_tok_out", "word_elapsed", "word_num_chunks",
    "word_summary",
)


# ══════════════════════════════════════════════════════════════════════════════
# CORE TRANSLATION LOOP
# ══════════════════════════════════════════════════════════════════════════════
def _run_translation(chunks, target_lang, doc_context,
                     timer_ph, prog_ph, render_stats, add_log,
                     total_blocks, prefix_label="Dịch"):
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
        args=(holder, client, chunks, target_lang, doc_context, MAX_WORD_WORKERS),
        daemon=True,
    ).start()

    last_logged = 0
    dot = 0
    while not holder["done"]:
        # Flush chunk log entries thread đã append
        while last_logged < len(holder["chunk_log"]):
            entry = holder["chunk_log"][last_logged]
            if entry["error"]:
                add_log(f"   ❌ Chunk {entry['idx']+1}: {entry['error'][:80]}")
            else:
                add_log(f"   ✅ Chunk {entry['idx']+1}: {entry['size']} đoạn "
                        f"({entry['in_t']:,}in/{entry['out_t']:,}out tok)")
            last_logged += 1

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

    # Flush remaining log entries
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
# FULL TRANSLATION (lần đầu)
# ══════════════════════════════════════════════════════════════════════════════
def _run_full_translation(uploaded_docx, lang_word):
    st.markdown("### 📊 Tiến độ")
    timer_ph = st.empty()
    col_bl, col_ck, col_usd, col_vnd = st.columns(4)
    ph_bl, ph_ck, ph_usd, ph_vnd = (col_bl.empty(), col_ck.empty(),
                                     col_usd.empty(), col_vnd.empty())

    def render_stats(total_bl, num_chunks, done_chunks, tok_in, tok_out):
        usd, vnd = calc_cost(tok_in, tok_out)
        ph_bl.markdown(stat_box_html(f"{total_bl:,}", "Đoạn văn"), unsafe_allow_html=True)
        ph_ck.markdown(stat_box_html(f"{done_chunks}/{num_chunks}", "Chunk"), unsafe_allow_html=True)
        ph_usd.markdown(stat_box_html(f"${usd:.4f}", "USD"), unsafe_allow_html=True)
        ph_vnd.markdown(stat_box_html(f"{vnd:,.0f}₫", "VND"), unsafe_allow_html=True)

    render_stats(0, 0, 0, 0, 0)
    prog = st.progress(0, text="Đang chuẩn bị...")

    st.markdown("### 📋 Nhật ký hoạt động")
    log_ph    = st.empty()
    log_lines = []
    add_log   = make_log_adder(log_lines, log_ph)

    try:
        docx_bytes = uploaded_docx.read()
        add_log(f"📄 Đã nhận file: {uploaded_docx.name}")

        add_log("🔍 Phân tích cấu trúc tài liệu...")
        blocks       = extract_docx_blocks(docx_bytes)
        stats        = count_by_role(blocks)
        translatable = [b for b in blocks if b["role"] not in NO_TRANSLATE_ROLES]
        total_chars  = sum(len(b["text"]) for b in translatable)

        add_log(f"✅ {stats['total']:,} đoạn ({total_chars:,} ký tự body)")
        add_log(f"   • Body cần dịch:    {stats['body']:,}")
        add_log(f"   • Header (real):    {stats['header']:,}")
        add_log(f"   • Footer (real):    {stats['footer']:,}")
        add_log(f"   • Body lặp (H/F-like): {stats['body_repeated']:,} "
                f"(phát hiện qua text lặp ≥3 lần)")
        if stats['hf_total'] > 0:
            add_log(f"⏭  Tạm thời bỏ qua {stats['hf_total']:,} H/F — bấm nút riêng sau khi xong")

        doc_context = build_doc_context(blocks)
        target_lang = LANG_EN[lang_word]
        chunks      = chunk_blocks(translatable)
        num_chunks  = len(chunks)
        add_log(f"📦 Chia thành {num_chunks} chunk (adaptive ~{total_chars // max(num_chunks,1):,} chars/chunk)")
        add_log(f"⚡ Dịch song song {MAX_WORD_WORKERS} luồng")

        translations, tok_in, tok_out, elapsed = _run_translation(
            chunks, target_lang, doc_context,
            timer_ph, prog, render_stats, add_log,
            len(translatable), prefix_label="Dịch",
        )

        add_log(f"✅ Dịch xong {len(translatable):,} đoạn trong {elapsed:.1f}s")
        add_log(f"🤖 Model: {get_working_model()}")

        prog.progress(95, text="Tạo file DOCX...")
        timer_ph.markdown(
            timer_box_html(elapsed, "💾 Đang ghi file DOCX..."),
            unsafe_allow_html=True,
        )

        usd, vnd = calc_cost(tok_in, tok_out)
        render_stats(len(translatable), num_chunks, num_chunks, tok_in, tok_out)
        prog.progress(100, text="✅ Hoàn thành!")
        timer_ph.markdown(
            timer_done_html(elapsed, f"Dịch xong {len(translatable):,} đoạn!"),
            unsafe_allow_html=True,
        )
        add_log("─" * 44)
        add_log(f"🎉 Xong {len(translatable):,} đoạn / {num_chunks} chunk trong {elapsed:.1f}s")
        add_log(f"💰 Token: {tok_in:,} in + {tok_out:,} out")
        add_log(f"💵 Chi phí: ${usd:.4f} USD ≈ {vnd:,.0f} VND")

        # Save toàn bộ vào session_state
        st.session_state["word_blocks"]       = blocks
        st.session_state["word_translations"] = translations
        st.session_state["word_docx_bytes"]   = docx_bytes
        st.session_state["word_filename"]     = uploaded_docx.name
        st.session_state["word_lang"]         = lang_word
        st.session_state["word_doc_context"]  = doc_context
        st.session_state["word_tok_in"]       = tok_in
        st.session_state["word_tok_out"]      = tok_out
        st.session_state["word_elapsed"]      = elapsed
        st.session_state["word_num_chunks"]   = num_chunks
        st.session_state["word_summary"]      = (
            f"✅ Dịch xong {len(translatable):,} đoạn trong {elapsed:.1f}s  "
            f"|  ${usd:.4f} USD ≈ {vnd:,.0f} VND"
        )

    except Exception as e:
        add_log(f"❌ Lỗi: {e}")
        st.error(f"❌ Có lỗi xảy ra: {e}")
        timer_ph.markdown(timer_error_html(str(e)), unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# RESCAN — chỉ dịch lại các đoạn bị bỏ sót
# ══════════════════════════════════════════════════════════════════════════════
def _run_rescan():
    blocks       = st.session_state["word_blocks"]
    translations = st.session_state["word_translations"]
    doc_context  = st.session_state["word_doc_context"]
    target_lang  = LANG_EN[st.session_state["word_lang"]]

    missed = find_missed(blocks, translations)
    if not missed:
        st.info("✨ Không có đoạn nào còn sót — tất cả đã dịch!")
        return

    st.markdown(f"### 🔍 Quét bỏ sót — {len(missed):,} đoạn")
    timer_ph = st.empty()
    col_ms, col_ck, col_usd, col_vnd = st.columns(4)
    ph_ms, ph_ck, ph_usd, ph_vnd = (col_ms.empty(), col_ck.empty(),
                                     col_usd.empty(), col_vnd.empty())

    def render_stats(total_bl, num_chunks, done_chunks, tok_in, tok_out):
        usd, vnd = calc_cost(tok_in, tok_out)
        ph_ms.markdown(stat_box_html(f"{total_bl:,}", "Đoạn còn sót"), unsafe_allow_html=True)
        ph_ck.markdown(stat_box_html(f"{done_chunks}/{num_chunks}", "Chunk"), unsafe_allow_html=True)
        ph_usd.markdown(stat_box_html(f"${usd:.4f}", "USD thêm"), unsafe_allow_html=True)
        ph_vnd.markdown(stat_box_html(f"{vnd:,.0f}₫", "VND thêm"), unsafe_allow_html=True)

    render_stats(0, 0, 0, 0, 0)
    prog = st.progress(0, text="Đang chuẩn bị quét...")

    st.markdown("### 📋 Nhật ký quét")
    log_ph    = st.empty()
    log_lines = []
    add_log   = make_log_adder(log_lines, log_ph)
    add_log(f"🔍 Tìm thấy {len(missed):,} đoạn chưa dịch / dịch fail")

    chunks     = chunk_blocks(missed)
    num_chunks = len(chunks)
    add_log(f"📦 Chia thành {num_chunks} chunk — dịch lại với {MAX_WORD_WORKERS} luồng")

    try:
        new_translations, tok_in, tok_out, elapsed = _run_translation(
            chunks, target_lang, doc_context,
            timer_ph, prog, render_stats, add_log,
            len(missed), prefix_label="Quét",
        )

        # Merge vào translations cũ
        translations.update(new_translations)
        st.session_state["word_translations"] = translations
        st.session_state["word_tok_in"]      += tok_in
        st.session_state["word_tok_out"]     += tok_out

        usd, vnd = calc_cost(tok_in, tok_out)
        prog.progress(100, text="✅ Quét xong!")
        timer_ph.markdown(
            timer_done_html(elapsed, f"Quét lại {len(missed):,} đoạn xong!"),
            unsafe_allow_html=True,
        )
        add_log("─" * 44)
        add_log(f"🎉 Quét xong {len(missed):,} đoạn trong {elapsed:.1f}s")
        add_log(f"💵 Chi phí thêm: ${usd:.4f} USD ≈ {vnd:,.0f} VND")
        st.success(f"✅ Đã dịch thêm {len(missed):,} đoạn bị bỏ sót!")

    except Exception as e:
        add_log(f"❌ Lỗi: {e}")
        st.error(f"❌ Lỗi rescan: {e}")
        timer_ph.markdown(timer_error_html(str(e)), unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# HEADER/FOOTER TRANSLATION — chạy khi user bấm nút "Dịch H/F"
# ══════════════════════════════════════════════════════════════════════════════
def _run_hf_translation():
    """Dịch tất cả H/F + body_repeated block chưa được dịch."""
    blocks       = st.session_state["word_blocks"]
    translations = st.session_state["word_translations"]
    doc_context  = st.session_state["word_doc_context"]
    target_lang  = LANG_EN[st.session_state["word_lang"]]

    hf_missed = find_missed(blocks, translations, hf_only=True)
    if not hf_missed:
        st.info("✨ Tất cả Header/Footer đã được dịch!")
        return

    st.markdown(f"### 🌐 Dịch Header & Footer — {len(hf_missed):,} đoạn")
    timer_ph = st.empty()
    col_hf, col_ck, col_usd, col_vnd = st.columns(4)
    ph_hf, ph_ck, ph_usd, ph_vnd = (col_hf.empty(), col_ck.empty(),
                                     col_usd.empty(), col_vnd.empty())

    def render_stats(total_bl, num_chunks, done_chunks, tok_in, tok_out):
        usd, vnd = calc_cost(tok_in, tok_out)
        ph_hf.markdown(stat_box_html(f"{total_bl:,}", "H/F đoạn"), unsafe_allow_html=True)
        ph_ck.markdown(stat_box_html(f"{done_chunks}/{num_chunks}", "Chunk"), unsafe_allow_html=True)
        ph_usd.markdown(stat_box_html(f"${usd:.4f}", "USD thêm"), unsafe_allow_html=True)
        ph_vnd.markdown(stat_box_html(f"{vnd:,.0f}₫", "VND thêm"), unsafe_allow_html=True)

    render_stats(0, 0, 0, 0, 0)
    prog = st.progress(0, text="Đang chuẩn bị...")

    st.markdown("### 📋 Nhật ký dịch H/F")
    log_ph    = st.empty()
    log_lines = []
    add_log   = make_log_adder(log_lines, log_ph)
    add_log(f"🌐 Bắt đầu dịch {len(hf_missed):,} đoạn Header/Footer")

    chunks     = chunk_blocks(hf_missed)
    num_chunks = len(chunks)
    add_log(f"📦 Chia thành {num_chunks} chunk — {MAX_WORD_WORKERS} luồng")

    try:
        new_translations, tok_in, tok_out, elapsed = _run_translation(
            chunks, target_lang, doc_context,
            timer_ph, prog, render_stats, add_log,
            len(hf_missed), prefix_label="Dịch H/F",
        )

        translations.update(new_translations)
        st.session_state["word_translations"] = translations
        st.session_state["word_tok_in"]      += tok_in
        st.session_state["word_tok_out"]     += tok_out

        usd, vnd = calc_cost(tok_in, tok_out)
        prog.progress(100, text="✅ Dịch H/F xong!")
        timer_ph.markdown(
            timer_done_html(elapsed, f"Dịch {len(hf_missed):,} H/F xong!"),
            unsafe_allow_html=True,
        )
        add_log("─" * 44)
        add_log(f"🎉 Xong {len(hf_missed):,} H/F trong {elapsed:.1f}s")
        add_log(f"💵 Chi phí thêm: ${usd:.4f} USD ≈ {vnd:,.0f} VND")
        st.success(f"✅ Đã dịch {len(hf_missed):,} đoạn Header/Footer!")

    except Exception as e:
        add_log(f"❌ Lỗi: {e}")
        st.error(f"❌ Lỗi dịch H/F: {e}")
        timer_ph.markdown(timer_error_html(str(e)), unsafe_allow_html=True)


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
    "header":          "📌 Header",
    "footer":          "📌 Footer",
    "body_repeated":   "🔁 Lặp lại",
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

    df = pd.DataFrame([
        {
            "ID":       b["id"],
            "Vai trò":  ROLE_LABEL.get(b["role"], b["role"]),
            "Gốc":      b["text"],
            "Bản dịch": translations.get(b["id"], ""),
        }
        for b in display_blocks
    ])

    # Editor key thay đổi khi filter / rescan / HF translate để widget state không bị lệch
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

    # Sync edits back vào session_state mỗi rerun
    changes = 0
    for _, row in edited.iterrows():
        bid     = row["ID"]
        new_val = row["Bản dịch"]
        if translations.get(bid) != new_val:
            translations[bid] = new_val
            changes += 1
    if changes:
        st.session_state["word_translations"] = translations
        st.caption(f"💾 Đã ghi nhận {changes} thay đổi — bấm **Tải Word đã dịch** ở trên.")


# ══════════════════════════════════════════════════════════════════════════════
# MAIN UI
# ══════════════════════════════════════════════════════════════════════════════
def _clear_state():
    for k in SS_KEYS:
        st.session_state.pop(k, None)
    # Reset editor widget state — clear cả filter + version
    for k in list(st.session_state.keys()):
        if k.startswith("word_editor") or k in ("word_only_missed_filter",
                                                  "word_show_hf_filter"):
            st.session_state.pop(k, None)
    st.session_state["word_editor_version"] = 0


def render():
    """Entry point — gọi từ `streamlit_app.py`."""
    uploaded_docx = st.file_uploader("📝 Chọn file Word cần dịch",
                                      type=["docx"], key="word_uploader")
    lang_word = st.selectbox("🌐 Ngôn ngữ đích", LANGUAGES, key="word_lang_select")
    st.divider()

    if st.button("▶  Bắt đầu dịch Word",
                 disabled=(uploaded_docx is None), key="word_run"):
        _clear_state()
        _run_full_translation(uploaded_docx, lang_word)

    # ── Kết quả + thao tác sau dịch ─────────────────────────────────────────
    if "word_translations" in st.session_state:
        st.divider()
        st.success(st.session_state["word_summary"])

        # Rebuild DOCX bytes mỗi rerun từ translations hiện tại (có thể đã edit)
        translated_bytes = apply_translations(
            st.session_state["word_docx_bytes"],
            st.session_state["word_blocks"],
            st.session_state["word_translations"],
        )
        out_name = st.session_state["word_filename"].replace(
            ".docx", f"_translated_{st.session_state['word_lang'][:2]}.docx"
        )

        # Download — full width
        st.download_button(
            label="⬇️  Tải Word đã dịch",
            data=translated_bytes,
            file_name=out_name,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )

        # Đếm còn sót body + H/F để hiển thị label button + warning
        blocks       = st.session_state["word_blocks"]
        translations = st.session_state["word_translations"]
        missed_body  = find_missed(blocks, translations, hf_only=False)
        missed_hf    = find_missed(blocks, translations, hf_only=True)
        total_hf     = sum(1 for b in blocks if b["role"] in NO_TRANSLATE_ROLES)

        # Row 2 buttons: rescan + translate H/F (chỉ show nếu có H/F)
        rescan_clicked = hf_clicked = False
        if total_hf > 0:
            col_rs, col_hf = st.columns(2)
            with col_rs:
                rescan_clicked = st.button(
                    f"🔍  Quét bỏ sót ({len(missed_body)})",
                    disabled=(len(missed_body) == 0),
                    use_container_width=True, key="word_rescan_btn",
                    help="Dịch lại các đoạn body API fail",
                )
            with col_hf:
                hf_clicked = st.button(
                    f"🌐  Dịch Header / Footer ({len(missed_hf)}/{total_hf})",
                    disabled=(len(missed_hf) == 0),
                    use_container_width=True, key="word_hf_btn",
                    help="Dịch tất cả Header, Footer và đoạn lặp trong body",
                )
        else:
            rescan_clicked = st.button(
                f"🔍  Quét bỏ sót ({len(missed_body)})",
                disabled=(len(missed_body) == 0),
                use_container_width=True, key="word_rescan_btn",
            )

        # Warnings
        if len(missed_body) > 0:
            st.warning(f"⚠️ Còn {len(missed_body):,} đoạn body chưa dịch. "
                       f"Bấm **Quét bỏ sót** để dịch lại.")
        if total_hf > 0 and len(missed_hf) > 0:
            st.info(f"📌 {len(missed_hf):,}/{total_hf:,} Header/Footer chưa dịch "
                    f"(mặc định bỏ qua). Bấm **Dịch Header/Footer** nếu muốn dịch.")

        with st.expander("📋  Xem & sửa bản dịch inline", expanded=False):
            _render_editor()

        # Run heavy ops cuối hàm → progress UI full-width
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
