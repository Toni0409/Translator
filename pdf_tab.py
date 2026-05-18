"""Streamlit UI cho tab PDF — gọi backend `pdf_backend`."""
import os
import time
import tempfile

import streamlit as st
import fitz

from config import LANGUAGES, PDF_DELAY, PDF_MODEL
from gemini import get_client
from ui_common import (
    timer_box_html, timer_done_html, timer_error_html,
    stat_box_html, make_log_adder, calc_cost,
)
from pdf_backend import (
    extract_line_groups, parse_page_range, find_font,
    translate_page, write_translated_pdf,
    translate_pages_parallel, build_pdf_glossary,
    extract_pdf_images, ocr_pdf_images, insert_ocr_captions_into_pdf,
)


def _run_pdf_translation(uploaded_pdf, lang_pdf, pages_s,
                         use_parallel: bool = False,
                         use_glossary: bool = False,
                         use_ocr: bool = False,
                         parallel_workers: int = 4):
    """Chạy pipeline dịch PDF và lưu kết quả vào session_state."""
    st.markdown("### 📊 Tiến độ")
    timer_ph = st.empty()
    col_pg, col_ln, col_usd, col_vnd = st.columns(4)
    ph_pg, ph_ln, ph_usd, ph_vnd = (col_pg.empty(), col_ln.empty(),
                                     col_usd.empty(), col_vnd.empty())

    def render_stats(pages_done, total_pg, total_lines, tok_in, tok_out):
        usd, vnd = calc_cost(tok_in, tok_out)
        ph_pg.markdown(stat_box_html(f"{pages_done}/{total_pg}", "Trang"), unsafe_allow_html=True)
        ph_ln.markdown(stat_box_html(f"{total_lines:,}", "Dòng text"), unsafe_allow_html=True)
        ph_usd.markdown(stat_box_html(f"${usd:.4f}", "USD"), unsafe_allow_html=True)
        ph_vnd.markdown(stat_box_html(f"{vnd:,.0f}₫", "VND"), unsafe_allow_html=True)

    render_stats(0, 0, 0, 0, 0)
    prog = st.progress(0, text="Đang chuẩn bị...")

    st.markdown("### 📋 Nhật ký hoạt động")
    log_ph    = st.empty()
    log_lines = []
    add_log   = make_log_adder(log_lines, log_ph)

    src_path = dst_path = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_in:
            tmp_in.write(uploaded_pdf.read())
            src_path = tmp_in.name
        tmp_out  = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
        dst_path = tmp_out.name
        tmp_out.close()

        add_log(f"📄 Đã nhận file: {uploaded_pdf.name}")

        probe    = fitz.open(src_path)
        total_pg = len(probe)
        probe.close()

        targets = (parse_page_range(pages_s, total_pg)
                   if pages_s.strip() else list(range(total_pg)))
        add_log(f"📄 {total_pg} trang tổng, sẽ dịch {len(targets)} trang")

        prog.progress(5, text="Trích xuất text + phát hiện bảng...")
        add_log("🔍 Trích xuất text + phát hiện bảng từ PDF...")
        all_groups, _, table_stats = extract_line_groups(src_path, targets)
        total_lines    = sum(len(v) for v in all_groups.values())
        total_tables   = sum(s["tables"] for s in table_stats.values())
        total_cells    = sum(s["cell_lines"] for s in table_stats.values())
        pages_w_tables = sum(1 for s in table_stats.values() if s["tables"] > 0)
        add_log(f"✅ {total_lines:,} dòng text")
        if total_tables > 0:
            add_log(f"📊 Phát hiện {total_tables} bảng trên {pages_w_tables} trang "
                    f"({total_cells:,} dòng nằm trong cell) — sẽ kèm context (T# R# C#) khi dịch")
        else:
            add_log("📊 Không phát hiện bảng — dịch text thường")
        render_stats(0, total_pg, total_lines, 0, 0)
        add_log(f"🤖 Kết nối {PDF_MODEL}")

        client    = get_client()
        all_trans = {}
        tok_in = tok_out = 0
        t0 = time.time()

        # Optional: build glossary
        glossary = None
        if use_glossary:
            glossary = build_pdf_glossary(all_groups)
            if glossary:
                preview = ", ".join(glossary[:10])
                more = f" (+{len(glossary)-10})" if len(glossary) > 10 else ""
                add_log(f"📚 Glossary: {len(glossary)} thuật ngữ lặp → {preview}{more}")
            else:
                add_log("📚 Không tìm thấy thuật ngữ lặp")

        if use_parallel and len(targets) > 1:
            add_log(f"⚡ Parallel mode: {parallel_workers} workers")
            page_results: dict[int, dict] = {}

            def on_page_done(pi, done, total, in_t, out_t, err):
                nonlocal tok_in, tok_out
                tok_in  += in_t
                tok_out += out_t
                page_results[pi] = {"in_t": in_t, "out_t": out_t, "err": err}
                pct = int(10 + (done / total) * 80)
                prog.progress(pct, text=f"Dịch {done}/{total} trang xong...")
                status = "❌" if err else "✅"
                msg = f"   {status} Trang {pi+1} ({in_t:,}in/{out_t:,}out)"
                if err:
                    msg += f" — Lỗi: {err[:120]}"
                add_log(msg)
                render_stats(done, total_pg, total_lines, tok_in, tok_out)
                timer_ph.markdown(
                    timer_box_html(time.time() - t0,
                                   f"⚡ Đang dịch song song — {done}/{total} trang xong"),
                    unsafe_allow_html=True,
                )

            all_trans, tok_in, tok_out, errors = translate_pages_parallel(
                client, all_groups, lang_pdf, targets,
                max_workers=parallel_workers,
                glossary=glossary,
                progress_callback=on_page_done,
            )
        else:
            for idx, pi in enumerate(targets):
                groups     = all_groups.get(pi, [])
                page_start = time.time()
                pct        = int(10 + (idx / len(targets)) * 80)
                prog.progress(pct, text=f"Dịch trang {pi + 1}/{total_pg}...")
                add_log(f"📄 Trang {pi + 1}/{total_pg}: {len(groups)} dòng — gửi API...")

                if not groups:
                    all_trans[pi] = []
                    timer_ph.markdown(
                        timer_box_html(time.time() - t0, f"⏭ Trang {pi+1} trống, bỏ qua"),
                        unsafe_allow_html=True,
                    )
                    continue

                try:
                    trans, in_t, out_t = translate_page(
                        client, groups, lang_pdf, pi,
                        timer_ph, t0, page_start, log_lines, log_ph,
                        glossary=glossary,
                    )
                    tok_in  += in_t
                    tok_out += out_t
                    add_log(f"   ✅ {len(trans)} dòng ({in_t:,}in/{out_t:,}out tok) — {time.time()-page_start:.1f}s")
                except Exception as e:
                    add_log(f"   ❌ Lỗi trang {pi + 1}: {e}")
                    trans = [g["text"] for g in groups]

                all_trans[pi] = trans
                render_stats(idx + 1, total_pg, total_lines, tok_in, tok_out)
                time.sleep(PDF_DELAY)

        prog.progress(92, text="Tạo PDF...")
        add_log("💾 Đang tạo PDF...")
        timer_ph.markdown(
            timer_box_html(time.time() - t0, "💾 Đang ghi file PDF..."),
            unsafe_allow_html=True,
        )
        font_path = find_font()
        add_log(f"🔤 Font: {os.path.basename(font_path) if font_path else 'built-in'}")
        write_translated_pdf(src_path, dst_path, all_groups, all_trans, font_path)

        # Optional: OCR embedded images, add as PDF annotations on the translated file
        if use_ocr:
            prog.progress(94, text="OCR ảnh trong PDF...")
            add_log("🖼 Đang trích ảnh từ PDF...")
            imgs = extract_pdf_images(src_path, targets)
            if not imgs:
                add_log("🖼 Không có ảnh phù hợp để OCR (≥ 5KB)")
            else:
                add_log(f"🖼 Tìm thấy {len(imgs)} ảnh — đang OCR + dịch song song...")

                def on_ocr_progress(done, total):
                    pct = int(94 + (done / total) * 4)
                    prog.progress(pct, text=f"OCR ảnh {done}/{total}...")
                    timer_ph.markdown(
                        timer_box_html(time.time() - t0, f"🖼 OCR ảnh {done}/{total}"),
                        unsafe_allow_html=True,
                    )

                ocr_results = ocr_pdf_images(
                    client, imgs, lang_pdf,
                    progress_callback=on_ocr_progress,
                )
                with_text = sum(1 for r in ocr_results.values() if r.get("has_text"))
                add_log(f"🖼 OCR xong: {with_text}/{len(imgs)} ảnh có text")
                if with_text > 0:
                    inserted = insert_ocr_captions_into_pdf(
                        dst_path, dst_path, imgs, ocr_results, font_path,
                    )
                    add_log(f"📎 Đã chèn {inserted} OCR caption (annotation màu vàng) vào PDF")

        elapsed  = time.time() - t0
        usd, vnd = calc_cost(tok_in, tok_out)
        render_stats(len(targets), total_pg, total_lines, tok_in, tok_out)
        prog.progress(100, text="✅ Hoàn thành!")
        timer_ph.markdown(
            timer_done_html(elapsed, f"Dịch xong {len(targets)} trang!"),
            unsafe_allow_html=True,
        )
        add_log("─" * 44)
        add_log(f"🎉 Xong {len(targets)} trang trong {elapsed:.1f}s")
        add_log(f"💰 Token: {tok_in:,} in + {tok_out:,} out")
        add_log(f"💵 Chi phí: ${usd:.4f} USD ≈ {vnd:,.0f} VND")

        with open(dst_path, "rb") as f:
            st.session_state["pdf_bytes"] = f.read()
        st.session_state["pdf_out_name"] = uploaded_pdf.name.replace(".pdf", f"_translated_{lang_pdf[:2]}.pdf")
        st.session_state["pdf_summary"]  = (
            f"✅ Dịch xong {len(targets)} trang trong {elapsed:.1f}s  "
            f"|  ${usd:.4f} USD ≈ {vnd:,.0f} VND"
        )

    except Exception as e:
        add_log(f"❌ Lỗi: {e}")
        st.error(f"❌ Có lỗi xảy ra: {e}")
        timer_ph.markdown(timer_error_html(str(e)), unsafe_allow_html=True)
    finally:
        for p in (src_path, dst_path):
            try:
                if p:
                    os.unlink(p)
            except Exception:
                pass


def render():
    """Entry point — gọi từ `streamlit_app.py`."""
    uploaded_pdf = st.file_uploader("📄 Chọn file PDF cần dịch",
                                     type=["pdf"], key="pdf_uploader")
    col1, col2 = st.columns(2)
    with col1:
        lang_pdf = st.selectbox("🌐 Ngôn ngữ đích", LANGUAGES, key="pdf_lang")
    with col2:
        pages_s = st.text_input("📑 Trang cụ thể (tuỳ chọn)",
                                placeholder="Vd: 1-5,8  •  Để trống = tất cả",
                                key="pdf_pages")

    with st.expander("⚙️ Tuỳ chọn nâng cao", expanded=False):
        c1, c2 = st.columns(2)
        use_parallel = c1.checkbox(
            "⚡ Dịch song song nhiều trang",
            value=True, key="pdf_parallel",
            help="Tăng tốc đáng kể với PDF nhiều trang. Tắt nếu muốn xem live timer từng trang.",
        )
        parallel_workers = c2.slider(
            "Số worker song song", 2, 8, 4, key="pdf_workers",
            disabled=not use_parallel,
        )
        use_glossary = st.checkbox(
            "📚 Tự build glossary để dịch consistent",
            value=True, key="pdf_glossary",
            help="Phát hiện thuật ngữ lặp lại và đảm bảo dịch nhất quán xuyên suốt tài liệu.",
        )
        use_ocr = st.checkbox(
            "🖼 OCR text trong ảnh embedded",
            value=False, key="pdf_ocr",
            help="Trích ảnh từ PDF, OCR + dịch text trong ảnh, chèn vào PDF dưới dạng annotation màu vàng.",
        )

    st.divider()

    if st.button("▶  Bắt đầu dịch PDF", disabled=(uploaded_pdf is None), key="pdf_run"):
        for k in ("pdf_bytes", "pdf_out_name", "pdf_summary"):
            st.session_state.pop(k, None)
        _run_pdf_translation(
            uploaded_pdf, lang_pdf, pages_s,
            use_parallel=use_parallel,
            use_glossary=use_glossary,
            use_ocr=use_ocr,
            parallel_workers=parallel_workers,
        )

    if "pdf_bytes" in st.session_state:
        st.divider()
        st.success(st.session_state["pdf_summary"])
        st.download_button(
            label="⬇️  Tải PDF đã dịch",
            data=st.session_state["pdf_bytes"],
            file_name=st.session_state["pdf_out_name"],
            mime="application/pdf",
            use_container_width=True,
        )
    elif not uploaded_pdf:
        st.info("👆 Vui lòng upload file PDF để bắt đầu")
