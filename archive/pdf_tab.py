"""Streamlit UI cho tab PDF — gọi backend `pdf_backend`."""
import io
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
    quality_check_pdf, build_bilingual_pdf,
    pdf_checkpoint_save, pdf_checkpoint_load, pdf_checkpoint_clear,
    detect_skip_pages, parse_custom_glossary,
    detect_scanned_pages, ocr_scanned_pages_parallel,
    insert_scanned_translation_pages,
)


def _run_batch_pdf_translation(uploaded_pdfs, lang_pdf, pages_s, **kwargs):
    """Translate multiple PDFs sequentially, share TM, zip outputs."""
    import zipfile
    n = len(uploaded_pdfs)
    st.markdown(f"### 📦 Batch mode: {n} file PDF")
    overall = st.progress(0, text=f"0/{n} file...")
    summary_lines = []
    zip_buf = io.BytesIO()

    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for i, pdf in enumerate(uploaded_pdfs):
            st.markdown(f"---\n#### 📄 File {i+1}/{n}: `{pdf.name}`")
            _run_pdf_translation(pdf, lang_pdf, pages_s, **kwargs)
            if "pdf_bytes" in st.session_state:
                out_name = st.session_state["pdf_out_name"]
                zf.writestr(out_name, st.session_state["pdf_bytes"])
                summary_lines.append(f"✅ {pdf.name} → {out_name}")
                if "pdf_bilingual_bytes" in st.session_state:
                    biling_name = out_name.replace(".pdf", "_bilingual.pdf")
                    zf.writestr(biling_name, st.session_state["pdf_bilingual_bytes"])
                # Clear single-file state so next iteration is fresh
                for k in ("pdf_bytes", "pdf_out_name", "pdf_summary",
                          "pdf_bilingual_bytes", "pdf_quality_issues"):
                    st.session_state.pop(k, None)
            else:
                summary_lines.append(f"❌ {pdf.name} — không có output")
            overall.progress((i + 1) / n, text=f"{i+1}/{n} file xong")

    st.session_state["pdf_batch_zip"]  = zip_buf.getvalue()
    st.session_state["pdf_batch_name"] = f"translated_pdfs_{lang_pdf[:2]}.zip"
    st.session_state["pdf_batch_summary"] = "\n".join(summary_lines)
    st.success(f"✅ Hoàn thành {n} file. Tải ZIP bên dưới.")


def _run_pdf_translation(uploaded_pdf, lang_pdf, pages_s,
                         use_parallel: bool = False,
                         use_glossary: bool = False,
                         use_ocr: bool = False,
                         use_tm: bool = True,
                         use_bilingual: bool = False,
                         use_quality_check: bool = True,
                         use_skip_toc: bool = False,
                         use_checkpoint: bool = True,
                         use_scanned_ocr: bool = True,
                         custom_rules: str = "",
                         custom_glossary_text: str = "",
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
        pdf_bytes = uploaded_pdf.read()
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_in:
            tmp_in.write(pdf_bytes)
            src_path = tmp_in.name
        tmp_out  = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
        dst_path = tmp_out.name
        tmp_out.close()

        add_log(f"📄 Đã nhận file: {uploaded_pdf.name}")

        # Resume from checkpoint if available
        resume_trans = {}
        if use_checkpoint:
            ckpt = pdf_checkpoint_load(pdf_bytes, lang_pdf)
            if ckpt and ckpt.get("trans"):
                resume_trans = ckpt["trans"]
                add_log(f"♻️ Resume từ checkpoint: {len(resume_trans)} trang đã dịch trước đó")

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

        # Custom user glossary (parsed from text input)
        custom_glossary = parse_custom_glossary(custom_glossary_text) if custom_glossary_text.strip() else None
        if custom_glossary:
            add_log(f"📝 Custom glossary: {len(custom_glossary)} cặp thuật ngữ bắt buộc")
        if custom_rules.strip():
            add_log(f"📜 Custom rules: {len(custom_rules)} ký tự hướng dẫn dịch riêng")

        # Detect scanned (image-only) pages — no extractable text layer
        scanned_pages: list[int] = []
        if use_scanned_ocr:
            scanned_pages = detect_scanned_pages(src_path, targets)
            if scanned_pages:
                pages_str = ", ".join(str(p+1) for p in scanned_pages[:10])
                more = f" (+{len(scanned_pages)-10})" if len(scanned_pages) > 10 else ""
                add_log(f"🔎 Phát hiện {len(scanned_pages)} trang scan (không có text layer): {pages_str}{more}")

        # Skip TOC / References pages
        skip_pages = {}
        if use_skip_toc:
            skip_pages = detect_skip_pages(all_groups)
            if skip_pages:
                pages_str = ", ".join(str(p+1) for p in sorted(skip_pages.keys())[:10])
                more = f" (+{len(skip_pages)-10})" if len(skip_pages) > 10 else ""
                add_log(f"⏭ Skip {len(skip_pages)} trang TOC/References: {pages_str}{more}")
            else:
                add_log("⏭ Không phát hiện trang TOC/References để skip")

        # Translation Memory: persistent across runs in session_state
        tm = None
        if use_tm:
            if "pdf_tm" not in st.session_state:
                st.session_state["pdf_tm"] = {}
            tm = st.session_state["pdf_tm"]
            add_log(f"💾 TM: {len(tm):,} entry sẵn có")

        total_tm_hits = 0
        errors: dict = {}

        if use_parallel and len(targets) > 1:
            add_log(f"⚡ Parallel mode: {parallel_workers} workers")

            def on_page_done(pi, done, total, in_t, out_t, err, tm_hits):
                nonlocal tok_in, tok_out
                tok_in  += in_t
                tok_out += out_t
                pct = int(10 + (done / total) * 80) if total else 90
                prog.progress(pct, text=f"Dịch {done}/{total} trang xong...")
                status = "❌" if err else "✅"
                tm_note = f" [TM {tm_hits}]" if tm_hits else ""
                msg = f"   {status} Trang {pi+1} ({in_t:,}in/{out_t:,}out){tm_note}"
                if err:
                    msg += f" — Lỗi: {err[:120]}"
                add_log(msg)
                render_stats(done, total_pg, total_lines, tok_in, tok_out)
                timer_ph.markdown(
                    timer_box_html(time.time() - t0,
                                   f"⚡ Đang dịch song song — {done}/{total} trang xong"),
                    unsafe_allow_html=True,
                )
                if use_checkpoint:
                    pdf_checkpoint_save(pdf_bytes, lang_pdf, all_groups, all_trans, glossary)

            all_trans, tok_in, tok_out, errors, total_tm_hits = translate_pages_parallel(
                client, all_groups, lang_pdf, targets,
                max_workers=parallel_workers,
                glossary=glossary, tm=tm,
                custom_glossary=custom_glossary,
                custom_rules=custom_rules,
                skip_pages=skip_pages,
                resume_trans=resume_trans,
                progress_callback=on_page_done,
            )
        else:
            for idx, pi in enumerate(targets):
                groups     = all_groups.get(pi, [])
                page_start = time.time()
                pct        = int(10 + (idx / len(targets)) * 80)
                prog.progress(pct, text=f"Dịch trang {pi + 1}/{total_pg}...")

                # Skip TOC/References
                if pi in skip_pages:
                    all_trans[pi] = [g["text"] for g in groups]
                    add_log(f"⏭ Trang {pi+1}: skip ({skip_pages[pi]}) — giữ nguyên text")
                    continue
                # Resume from checkpoint
                if pi in resume_trans:
                    all_trans[pi] = resume_trans[pi]
                    add_log(f"♻️ Trang {pi+1}: dùng kết quả từ checkpoint")
                    continue

                add_log(f"📄 Trang {pi + 1}/{total_pg}: {len(groups)} dòng — gửi API...")

                if not groups:
                    all_trans[pi] = []
                    timer_ph.markdown(
                        timer_box_html(time.time() - t0, f"⏭ Trang {pi+1} trống, bỏ qua"),
                        unsafe_allow_html=True,
                    )
                    continue

                try:
                    trans, in_t, out_t, tm_hits = translate_page(
                        client, groups, lang_pdf, pi,
                        timer_ph, t0, page_start, log_lines, log_ph,
                        glossary=glossary, tm=tm,
                    )
                    tok_in  += in_t
                    tok_out += out_t
                    total_tm_hits += tm_hits
                    tm_note = f" [TM hit: {tm_hits}]" if tm_hits else ""
                    add_log(f"   ✅ {len(trans)} dòng ({in_t:,}in/{out_t:,}out tok){tm_note} — {time.time()-page_start:.1f}s")
                except Exception as e:
                    add_log(f"   ❌ Lỗi trang {pi + 1}: {e}")
                    trans = [g["text"] for g in groups]

                all_trans[pi] = trans
                render_stats(idx + 1, total_pg, total_lines, tok_in, tok_out)
                if use_checkpoint:
                    pdf_checkpoint_save(pdf_bytes, lang_pdf, all_groups, all_trans, glossary)
                time.sleep(PDF_DELAY)

        if use_tm and total_tm_hits > 0:
            add_log(f"💾 TM: tổng {total_tm_hits:,} dòng tái dùng — tiết kiệm chi phí")
        if use_tm:
            add_log(f"💾 TM: {len(tm):,} entry sau khi xong")

        # Clear checkpoint on success
        if use_checkpoint and not errors:
            pdf_checkpoint_clear(pdf_bytes, lang_pdf)

        prog.progress(92, text="Tạo PDF...")
        add_log("💾 Đang tạo PDF...")
        timer_ph.markdown(
            timer_box_html(time.time() - t0, "💾 Đang ghi file PDF..."),
            unsafe_allow_html=True,
        )
        font_path = find_font()
        add_log(f"🔤 Font: {os.path.basename(font_path) if font_path else 'built-in'}")
        write_translated_pdf(src_path, dst_path, all_groups, all_trans, font_path)

        # OCR scanned pages → insert translation as new pages after originals
        if use_scanned_ocr and scanned_pages:
            prog.progress(93, text="OCR trang scan...")
            add_log(f"🔎 Đang OCR + dịch {len(scanned_pages)} trang scan...")

            def on_scan_progress(done, total):
                pct = int(93 + (done / total) * 1)
                prog.progress(pct, text=f"OCR trang scan {done}/{total}...")
                timer_ph.markdown(
                    timer_box_html(time.time() - t0, f"🔎 OCR trang scan {done}/{total}"),
                    unsafe_allow_html=True,
                )

            scan_results = ocr_scanned_pages_parallel(
                client, src_path, scanned_pages, lang_pdf,
                progress_callback=on_scan_progress,
            )
            scan_in  = sum(r.get("in_t", 0)  for r in scan_results.values())
            scan_out = sum(r.get("out_t", 0) for r in scan_results.values())
            tok_in  += scan_in
            tok_out += scan_out
            with_text = sum(1 for r in scan_results.values() if r.get("translation"))
            add_log(f"🔎 OCR scan xong: {with_text}/{len(scanned_pages)} trang có text dịch")
            if with_text > 0:
                added = insert_scanned_translation_pages(
                    dst_path, dst_path, scan_results, font_path,
                )
                add_log(f"📄 Đã chèn {added} trang dịch (sau mỗi trang scan tương ứng)")

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

        # State for inline editor (re-render with user edits)
        with open(src_path, "rb") as f:
            st.session_state["pdf_src_bytes"] = f.read()
        st.session_state["pdf_all_groups"] = all_groups
        st.session_state["pdf_all_trans"]  = all_trans
        st.session_state["pdf_lang_used"]  = lang_pdf

        # Bilingual PDF: interleave original + translated pages
        if use_bilingual:
            add_log("📑 Tạo PDF song ngữ (gốc + dịch xen kẽ)...")
            biling_path = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf").name
            try:
                build_bilingual_pdf(src_path, dst_path, biling_path, targets)
                with open(biling_path, "rb") as f:
                    st.session_state["pdf_bilingual_bytes"] = f.read()
                add_log("✅ PDF song ngữ sẵn sàng để tải")
            except Exception as e:
                add_log(f"⚠️ Không tạo được PDF song ngữ: {e}")
            finally:
                try:
                    os.unlink(biling_path)
                except Exception:
                    pass

        # Quality check
        if use_quality_check:
            issues = quality_check_pdf(all_groups, all_trans)
            st.session_state["pdf_quality_issues"] = issues
            if issues:
                add_log(f"⚠️ Quality check: {len(issues)} dòng có vấn đề (xem chi tiết dưới)")
            else:
                add_log("✅ Quality check: không phát hiện vấn đề")

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


def _render_editor():
    """Inline editor for translated lines + re-render button."""
    import pandas as pd

    all_groups = st.session_state["pdf_all_groups"]
    all_trans  = st.session_state["pdf_all_trans"]
    issues     = st.session_state.get("pdf_quality_issues") or []

    # Build editable DataFrame
    rows = []
    for pi in sorted(all_groups.keys()):
        groups = all_groups[pi]
        trans  = all_trans.get(pi, [])
        for li, g in enumerate(groups):
            if li >= len(trans):
                break
            rows.append({
                "key":         f"{pi}_{li}",
                "page":        pi + 1,
                "original":    g["text"],
                "translation": trans[li] or "",
            })

    if not rows:
        return

    suspicious_keys = {f"{it['page']-1}_{it['line_idx']}" for it in issues}

    with st.expander(f"✏️ Edit bản dịch ({len(rows):,} dòng) — sửa rồi re-render PDF",
                     expanded=False):
        c1, c2, c3 = st.columns([2, 1, 1])
        filter_mode = c1.selectbox(
            "Hiển thị",
            ["Tất cả", "Chỉ dòng có vấn đề (quality check)", "Theo trang"],
            key="pdf_edit_filter",
        )
        page_filter = None
        if filter_mode == "Theo trang":
            pages_avail = sorted({r["page"] for r in rows})
            page_filter = c2.selectbox("Trang", pages_avail, key="pdf_edit_page")
        search = c3.text_input("🔍 Tìm trong gốc", key="pdf_edit_search")

        visible_rows = rows
        if filter_mode == "Chỉ dòng có vấn đề (quality check)":
            visible_rows = [r for r in rows if r["key"] in suspicious_keys]
        elif page_filter is not None:
            visible_rows = [r for r in rows if r["page"] == page_filter]
        if search.strip():
            s = search.strip().lower()
            visible_rows = [r for r in visible_rows if s in r["original"].lower()]

        st.caption(f"Hiển thị {len(visible_rows):,}/{len(rows):,} dòng")

        df = pd.DataFrame(visible_rows)
        edited = st.data_editor(
            df,
            column_config={
                "key":         None,
                "page":        st.column_config.NumberColumn("Trang", disabled=True, width="small"),
                "original":    st.column_config.TextColumn("Gốc",  disabled=True, width="large"),
                "translation": st.column_config.TextColumn("Bản dịch (sửa được)", width="large"),
            },
            hide_index=True,
            use_container_width=True,
            height=min(600, max(200, len(visible_rows) * 35)),
            key="pdf_edit_table",
        )

        cc1, cc2 = st.columns(2)
        if cc1.button("💾 Áp dụng sửa + Re-render PDF", type="primary",
                      use_container_width=True, key="pdf_edit_apply"):
            # Update all_trans with edits from the editor
            edit_map = {row["key"]: row["translation"] for _, row in edited.iterrows()}
            for pi in all_trans:
                for li in range(len(all_trans[pi])):
                    k = f"{pi}_{li}"
                    if k in edit_map:
                        all_trans[pi][li] = edit_map[k]
            st.session_state["pdf_all_trans"] = all_trans

            with st.spinner("Đang re-render PDF với bản sửa..."):
                src_path = dst_path = None
                try:
                    src_path = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf").name
                    with open(src_path, "wb") as f:
                        f.write(st.session_state["pdf_src_bytes"])
                    dst_path = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf").name
                    font_path = find_font()
                    write_translated_pdf(src_path, dst_path, all_groups, all_trans, font_path)
                    with open(dst_path, "rb") as f:
                        st.session_state["pdf_bytes"] = f.read()
                    st.success("✅ PDF đã re-render với bản sửa. Tải lại ở nút bên trên.")
                finally:
                    for p in (src_path, dst_path):
                        try:
                            if p:
                                os.unlink(p)
                        except Exception:
                            pass
            st.rerun()

        if cc2.button("🔄 Reset về bản dịch gốc", use_container_width=True,
                      key="pdf_edit_reset"):
            st.warning("Tải lại từ run mới để reset — bản sửa hiện không thể undo riêng phần.")


def render():
    """Entry point — gọi từ `streamlit_app.py`."""
    uploaded_pdfs = st.file_uploader(
        "📄 Chọn 1 hoặc nhiều file PDF cần dịch",
        type=["pdf"], key="pdf_uploader",
        accept_multiple_files=True,
    )
    col1, col2 = st.columns(2)
    with col1:
        lang_pdf = st.selectbox("🌐 Ngôn ngữ đích", LANGUAGES, key="pdf_lang")
    with col2:
        pages_s = st.text_input("📑 Trang cụ thể (tuỳ chọn)",
                                placeholder="Vd: 1-5,8  •  Để trống = tất cả",
                                key="pdf_pages")

    st.divider()
    n_files = len(uploaded_pdfs or [])

    # ── 2 NÚT: Dịch cơ bản / Dịch nâng cao ───────────────────────────────────
    show_adv = st.session_state.get("pdf_show_advanced", False)
    col_basic, col_adv = st.columns(2)
    with col_basic:
        basic_label = (f"🚀  Dịch cơ bản ({n_files} file)" if n_files > 1
                       else "🚀  Dịch cơ bản")
        basic_clicked = st.button(
            basic_label,
            disabled=(n_files == 0),
            use_container_width=True, type="primary", key="pdf_basic_btn",
            help="Dịch ngay với cài đặt mặc định — không cần chọn gì thêm",
        )
    with col_adv:
        adv_label = "🔽  Ẩn tuỳ chọn nâng cao" if show_adv else "⚙️  Dịch nâng cao"
        adv_toggle_clicked = st.button(
            adv_label,
            disabled=(n_files == 0),
            use_container_width=True, key="pdf_adv_toggle",
            help="Tuỳ chỉnh glossary, TM, OCR ảnh, custom rules, song ngữ... trước khi dịch",
        )
    if adv_toggle_clicked:
        st.session_state["pdf_show_advanced"] = not show_adv
        st.rerun()

    # ── Defaults cho cả hai mode ─────────────────────────────────────────────
    opts = dict(
        use_parallel=True,
        use_glossary=True,
        use_ocr=False,
        use_tm=True,
        use_bilingual=False,
        use_quality_check=True,
        use_skip_toc=False,
        use_checkpoint=True,
        use_scanned_ocr=True,
        custom_rules="",
        custom_glossary_text="",
        parallel_workers=4,
    )

    start_advanced_clicked = False

    # ── Advanced mode: hiện expander với options + nút Bắt đầu riêng ─────────
    if show_adv:
        with st.container(border=True):
            st.markdown("##### ⚙️ Tuỳ chọn nâng cao")
            c1, c2 = st.columns(2)
            opts["use_parallel"] = c1.checkbox(
                "⚡ Dịch song song nhiều trang",
                value=True, key="pdf_parallel",
                help="Tăng tốc đáng kể với PDF nhiều trang.",
            )
            opts["parallel_workers"] = c2.slider(
                "Số worker song song", 2, 8, 4, key="pdf_workers",
                disabled=not opts["use_parallel"],
            )
            opts["use_glossary"] = st.checkbox(
                "📚 Tự build glossary để dịch consistent",
                value=True, key="pdf_glossary",
                help="Phát hiện thuật ngữ lặp lại và đảm bảo dịch nhất quán xuyên suốt tài liệu.",
            )
            opts["use_tm"] = st.checkbox(
                "💾 Translation Memory (cache giữa các lần dịch)",
                value=True, key="pdf_use_tm",
                help="Cache bản dịch trong session — dòng đã dịch lần sau sẽ reuse, không tốn API call.",
            )
            opts["use_checkpoint"] = st.checkbox(
                "♻️ Auto-checkpoint (resume nếu bị gián đoạn)",
                value=True, key="pdf_checkpoint",
                help="Lưu state sau mỗi trang để resume khi crash/disconnect.",
            )
            opts["use_skip_toc"] = st.checkbox(
                "⏭ Skip TOC / References / Index",
                value=False, key="pdf_skip_toc",
                help="Tự phát hiện trang mục lục/tham khảo/phụ lục, bỏ qua để tiết kiệm token.",
            )
            opts["use_scanned_ocr"] = st.checkbox(
                "🔎 Tự OCR trang scan (không có text layer)",
                value=True, key="pdf_scanned_ocr",
                help="OCR + dịch trang scan bằng Gemini Vision, chèn 1 trang dịch sau mỗi trang scan.",
            )
            opts["use_ocr"] = st.checkbox(
                "🖼 OCR text trong ảnh embedded",
                value=False, key="pdf_ocr",
                help="Trích ảnh từ PDF, OCR + dịch text trong ảnh, chèn vào PDF dưới dạng annotation.",
            )
            opts["use_bilingual"] = st.checkbox(
                "📑 Xuất thêm PDF song ngữ (gốc + dịch xen kẽ)",
                value=False, key="pdf_bilingual",
            )
            opts["use_quality_check"] = st.checkbox(
                "🔍 Quality check sau khi dịch",
                value=True, key="pdf_qc",
                help="Tự phát hiện số bị mất, dòng quá dài/ngắn, dòng chưa dịch.",
            )

            st.markdown("---")
            st.markdown("**📝 Custom translation rules** *(tuỳ chọn)*")
            opts["custom_rules"] = st.text_area(
                "Hướng dẫn dịch riêng cho tài liệu này",
                value="", key="pdf_custom_rules", height=80,
                placeholder="Vd: Dịch formal, giữ thuật ngữ IT bằng tiếng Anh, dùng đại từ 'chúng tôi'...",
            )
            st.markdown("**📖 Custom glossary** *(tuỳ chọn)*")
            opts["custom_glossary_text"] = st.text_area(
                "Mỗi dòng: source = target  (hoặc source,target  /  source -> target)",
                value="", key="pdf_custom_glossary", height=100,
                placeholder="API = giao diện lập trình\nendpoint = điểm cuối\nrequest -> yêu cầu",
            )

            tm_count = len(st.session_state.get("pdf_tm", {}))
            if tm_count > 0:
                cc1, cc2 = st.columns([3, 1])
                cc1.caption(f"💾 TM hiện có **{tm_count:,}** entry")
                if cc2.button("🗑 Xoá TM", key="pdf_clear_tm"):
                    st.session_state["pdf_tm"] = {}
                    st.rerun()

            adv_btn_label = (f"▶  Bắt đầu dịch {n_files} file PDF" if n_files > 1
                             else "▶  Bắt đầu dịch (nâng cao)")
            start_advanced_clicked = st.button(
                adv_btn_label, disabled=(n_files == 0), type="primary",
                use_container_width=True, key="pdf_advanced_run",
            )

    # ── Trigger dịch (basic dùng defaults, advanced dùng opts đã chỉnh) ──────
    if basic_clicked or start_advanced_clicked:
        for k in ("pdf_bytes", "pdf_out_name", "pdf_summary",
                  "pdf_bilingual_bytes", "pdf_quality_issues",
                  "pdf_batch_zip", "pdf_batch_name",
                  "pdf_all_groups", "pdf_all_trans", "pdf_src_bytes",
                  "pdf_lang_used"):
            st.session_state.pop(k, None)

        if n_files == 1:
            _run_pdf_translation(uploaded_pdfs[0], lang_pdf, pages_s, **opts)
        else:
            _run_batch_pdf_translation(uploaded_pdfs, lang_pdf, pages_s, **opts)

    if "pdf_bytes" in st.session_state:
        st.divider()
        st.success(st.session_state["pdf_summary"])

        if "pdf_bilingual_bytes" in st.session_state:
            dc1, dc2 = st.columns(2)
            dc1.download_button(
                label="⬇️  Tải PDF đã dịch",
                data=st.session_state["pdf_bytes"],
                file_name=st.session_state["pdf_out_name"],
                mime="application/pdf",
                use_container_width=True,
            )
            biling_name = st.session_state["pdf_out_name"].replace(".pdf", "_bilingual.pdf")
            dc2.download_button(
                label="⬇️  Tải PDF song ngữ",
                data=st.session_state["pdf_bilingual_bytes"],
                file_name=biling_name,
                mime="application/pdf",
                use_container_width=True,
            )
        else:
            st.download_button(
                label="⬇️  Tải PDF đã dịch",
                data=st.session_state["pdf_bytes"],
                file_name=st.session_state["pdf_out_name"],
                mime="application/pdf",
                use_container_width=True,
            )

        issues = st.session_state.get("pdf_quality_issues") or []
        if issues:
            with st.expander(f"⚠️ Quality check: {len(issues)} dòng có vấn đề", expanded=False):
                for it in issues[:100]:
                    with st.container(border=True):
                        st.markdown(f"**Trang {it['page']}** — *{'; '.join(it['issues'])}*")
                        cc1, cc2 = st.columns(2)
                        cc1.text_area("Gốc", it["text"], height=80,
                                      key=f"pdf_qc_orig_{it['page']}_{it['line_idx']}",
                                      disabled=True, label_visibility="collapsed")
                        cc2.text_area("Dịch", it["translation"], height=80,
                                      key=f"pdf_qc_tr_{it['page']}_{it['line_idx']}",
                                      disabled=True, label_visibility="collapsed")
                if len(issues) > 100:
                    st.caption(f"Chỉ hiển thị 100/{len(issues)} dòng đầu.")

        # Inline editor: edit translations and re-render PDF
        if "pdf_all_groups" in st.session_state and "pdf_all_trans" in st.session_state:
            _render_editor()

    if "pdf_batch_zip" in st.session_state:
        st.divider()
        st.success("📦 Batch hoàn thành — tải ZIP chứa tất cả PDF đã dịch")
        st.code(st.session_state.get("pdf_batch_summary", ""), language=None)
        st.download_button(
            label="⬇️  Tải ZIP toàn bộ batch",
            data=st.session_state["pdf_batch_zip"],
            file_name=st.session_state["pdf_batch_name"],
            mime="application/zip",
            use_container_width=True,
        )
    elif not uploaded_pdfs:
        st.info("👆 Vui lòng upload file PDF để bắt đầu")
