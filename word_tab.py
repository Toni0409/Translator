"""Streamlit UI cho tab Word — gọi backend `word_backend`."""
import time
import threading

import streamlit as st

from config import LANGUAGES, LANG_EN, CHUNK_SIZE, WORD_MODELS, NO_TRANSLATE_ROLES
from gemini import get_client
from ui_common import (
    timer_box_html, timer_done_html, timer_error_html,
    stat_box_html, make_log_adder,
)
from word_backend import (
    extract_docx_blocks, build_doc_context, translate_chunk,
    apply_translations, get_working_model,
)


def _worker_translate_all(holder, client, chunks, target_lang, doc_context):
    """Background worker: dịch tuần tự từng chunk, update progress vào holder."""
    try:
        for i, chunk in enumerate(chunks):
            holder["current_chunk"] = i + 1
            r = translate_chunk(client, chunk, target_lang, doc_context)
            holder["translations"].update(r)
            holder["chunk_done"] = i + 1
        holder["done"] = True
    except Exception as e:
        holder["error"] = str(e)
        holder["done"]  = True


def _run_word_translation(uploaded_docx, lang_word):
    """Chạy pipeline dịch Word và lưu kết quả vào session_state."""
    st.markdown("### 📊 Tiến độ")
    timer_ph = st.empty()
    col_bl, col_ck, col_pct = st.columns(3)
    ph_blocks, ph_chunks, ph_pct = col_bl.empty(), col_ck.empty(), col_pct.empty()

    def render_stats(total_bl, num_chunks, done_chunks):
        pct = int(done_chunks / num_chunks * 100) if num_chunks > 0 else 0
        ph_blocks.markdown(stat_box_html(str(total_bl), "Đoạn văn"), unsafe_allow_html=True)
        ph_chunks.markdown(stat_box_html(f"{done_chunks}/{num_chunks}", "Chunk"), unsafe_allow_html=True)
        ph_pct.markdown(stat_box_html(f"{pct}%", "Hoàn thành"), unsafe_allow_html=True)

    render_stats(0, 0, 0)
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
        translatable = [b for b in blocks if b["role"] not in NO_TRANSLATE_ROLES]
        hf_count     = len(blocks) - len(translatable)
        toc_count    = sum(1 for b in blocks if b["role"] == "toc")
        add_log(f"✅ {len(blocks)} đoạn: {len(translatable)} cần dịch, "
                f"{toc_count} TOC, {hf_count} header/footer")

        doc_context = build_doc_context(blocks)
        target_lang = LANG_EN[lang_word]
        chunks      = [translatable[i:i + CHUNK_SIZE]
                       for i in range(0, len(translatable), CHUNK_SIZE)]
        num_chunks  = len(chunks)
        render_stats(len(translatable), num_chunks, 0)
        add_log(f"🤖 Bắt đầu dịch — {len(translatable)} đoạn / {num_chunks} chunk")
        add_log(f"📡 Model ưu tiên: {WORD_MODELS[0]} (fallback: {', '.join(WORD_MODELS[1:])})")

        holder = {
            "translations":  {},
            "done":          False,
            "error":         None,
            "chunk_done":    0,
            "current_chunk": 0,
        }
        client = get_client()
        t0 = time.time()
        threading.Thread(
            target=_worker_translate_all,
            args=(holder, client, chunks, target_lang, doc_context),
            daemon=True,
        ).start()

        dot = 0
        while not holder["done"]:
            elapsed   = time.time() - t0
            done_c    = holder["chunk_done"]
            current_c = holder["current_chunk"]
            pct       = int(done_c / num_chunks * 90) if num_chunks > 0 else 0
            dots      = "." * (dot % 4)
            timer_ph.markdown(
                timer_box_html(elapsed, f"🔄 Đang dịch chunk {current_c}/{num_chunks}{dots}"),
                unsafe_allow_html=True,
            )
            prog.progress(pct, text=f"Chunk {done_c}/{num_chunks}...")
            render_stats(len(translatable), num_chunks, done_c)
            dot += 1
            time.sleep(1)

        if holder["error"]:
            raise RuntimeError(holder["error"])

        translations = holder["translations"]
        elapsed_w    = time.time() - t0

        add_log(f"✅ Dịch xong {len(translatable)} đoạn trong {elapsed_w:.1f}s")
        add_log(f"🤖 Model đã dùng: {get_working_model() or WORD_MODELS[0]}")

        prog.progress(92, text="Tạo file DOCX...")
        add_log("💾 Đang tạo file DOCX...")
        timer_ph.markdown(
            timer_box_html(elapsed_w, "💾 Đang ghi file DOCX..."),
            unsafe_allow_html=True,
        )

        translated_bytes = apply_translations(docx_bytes, blocks, translations)
        render_stats(len(translatable), num_chunks, num_chunks)
        prog.progress(100, text="✅ Hoàn thành!")
        timer_ph.markdown(
            timer_done_html(elapsed_w, f"Dịch xong {len(translatable)} đoạn!"),
            unsafe_allow_html=True,
        )
        add_log("─" * 44)
        add_log(f"🎉 Xong {len(translatable)} đoạn trong {elapsed_w:.1f}s")

        st.session_state["word_bytes"]    = translated_bytes
        st.session_state["word_out_name"] = uploaded_docx.name.replace(
            ".docx", f"_translated_{lang_word[:2]}.docx"
        )
        st.session_state["word_summary"]  = (
            f"✅ Dịch xong {len(translatable)} đoạn trong {elapsed_w:.1f}s"
        )

    except Exception as e:
        add_log(f"❌ Lỗi: {e}")
        st.error(f"❌ Có lỗi xảy ra: {e}")
        timer_ph.markdown(timer_error_html(str(e)), unsafe_allow_html=True)


def render():
    """Entry point — gọi từ `streamlit_app.py`."""
    uploaded_docx = st.file_uploader("📝 Chọn file Word cần dịch",
                                      type=["docx"], key="word_uploader")
    lang_word = st.selectbox("🌐 Ngôn ngữ đích", LANGUAGES, key="word_lang")
    st.divider()

    if st.button("▶  Bắt đầu dịch Word",
                 disabled=(uploaded_docx is None), key="word_run"):
        for k in ("word_bytes", "word_out_name", "word_summary"):
            st.session_state.pop(k, None)
        _run_word_translation(uploaded_docx, lang_word)

    if "word_bytes" in st.session_state:
        st.divider()
        st.success(st.session_state["word_summary"])
        st.download_button(
            label="⬇️  Tải Word đã dịch",
            data=st.session_state["word_bytes"],
            file_name=st.session_state["word_out_name"],
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
    elif not uploaded_docx:
        st.info("👆 Vui lòng upload file Word (.docx) để bắt đầu")
