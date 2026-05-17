"""Streamlit UI cho tab PDF — gọi backend `pdf_backend`."""
import os
import time
import tempfile

import streamlit as st
import fitz

from config import (
    LANGUAGES, PRICE_INPUT, PRICE_OUTPUT, USD_TO_VND, PDF_DELAY, PDF_MODEL,
)
from gemini import get_client
from ui_common import (
    timer_box_html, timer_done_html, timer_error_html,
    stat_box_html, make_log_adder,
)
from pdf_backend import (
    extract_line_groups, parse_page_range, find_font,
    translate_page, write_translated_pdf,
)


def _calc_cost(tok_in: int, tok_out: int) -> tuple[float, float]:
    usd = ((tok_in / 1e6) * PRICE_INPUT + (tok_out / 1e6) * PRICE_OUTPUT) * 10
    vnd = usd * USD_TO_VND
    return usd, vnd


def _run_pdf_translation(uploaded_pdf, lang_pdf, pages_s):
    """Chạy pipeline dịch PDF và lưu kết quả vào session_state."""
    st.markdown("### 📊 Tiến độ")
    timer_ph = st.empty()
    col_pg, col_ln, col_usd, col_vnd = st.columns(4)
    ph_pg, ph_ln, ph_usd, ph_vnd = (col_pg.empty(), col_ln.empty(),
                                     col_usd.empty(), col_vnd.empty())

    def render_stats(pages_done, total_pg, total_lines, tok_in, tok_out):
        usd, vnd = _calc_cost(tok_in, tok_out)
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

        prog.progress(5, text="Trích xuất text...")
        add_log("🔍 Trích xuất text từ PDF...")
        all_groups, _ = extract_line_groups(src_path, targets)
        total_lines   = sum(len(v) for v in all_groups.values())
        add_log(f"✅ {total_lines} dòng text")
        render_stats(0, total_pg, total_lines, 0, 0)
        add_log(f"🤖 Kết nối {PDF_MODEL}")

        client    = get_client()
        all_trans = {}
        tok_in = tok_out = 0
        t0 = time.time()

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

        elapsed  = time.time() - t0
        usd, vnd = _calc_cost(tok_in, tok_out)
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
    st.divider()

    if st.button("▶  Bắt đầu dịch PDF", disabled=(uploaded_pdf is None), key="pdf_run"):
        for k in ("pdf_bytes", "pdf_out_name", "pdf_summary"):
            st.session_state.pop(k, None)
        _run_pdf_translation(uploaded_pdf, lang_pdf, pages_s)

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
