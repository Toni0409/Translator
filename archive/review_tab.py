"""Streamlit UI cho tab So sánh / Đánh giá bản dịch Word."""
import time
import threading

import pandas as pd
import streamlit as st

from config import LANGUAGES
from gemini import get_client
from ui_common import (
    timer_box_html, timer_done_html, stat_box_html,
    make_log_adder, calc_cost,
)
from word_backend import extract_docx_blocks, build_doc_context
from pdf_backend import extract_pdf_blocks
from review_backend import (
    align_blocks, chunk_pairs, review_parallel, summarize_reviews,
    build_review_report_docx, REVIEW_CHUNK_SIZE,
    word_diff_html, compute_quality_score,
    apply_edits_to_docx, apply_edits_to_pdf,
)
from review_vision_backend import (
    render_doc_as_page_images, compare_pages_parallel,
    vision_quality_score, build_vision_report_docx,
    parse_page_spec, synthesize_overall_summary,
    aggregate_issues_by_category, flatten_issues,
)


SS_KEYS = ("rv_reviews", "rv_pairs", "rv_target_lang", "rv_orig_name",
           "rv_trans_name", "rv_tok_in", "rv_tok_out", "rv_elapsed",
           "rv_trans_bytes", "rv_orig_bytes", "rv_trans_is_pdf", "rv_orig_is_pdf",
           "rv_edits", "rv_smart_align",
           "rv_v_pages", "rv_v_orig_images", "rv_v_trans_images",
           "rv_v_target_lang", "rv_v_orig_name", "rv_v_trans_name",
           "rv_v_tok_in", "rv_v_tok_out", "rv_v_elapsed",
           "rv_v_overall_summary", "rv_v_page_spec", "rv_v_done_pages",
           "rv_v_focus_areas")


def _clear_state():
    for k in SS_KEYS:
        st.session_state.pop(k, None)


def _extract_blocks_any(file_obj) -> tuple[list[dict], str, bytes]:
    """
    Detect file format from name and extract blocks.
    Returns (blocks, format_label, raw_bytes). format_label is "DOCX" or "PDF".
    """
    name = (file_obj.name or "").lower()
    data = file_obj.read()
    if name.endswith(".pdf"):
        return extract_pdf_blocks(data), "PDF", data
    return extract_docx_blocks(data), "DOCX", data


def render():
    st.markdown("## 🔍 So sánh / Đánh giá bản dịch")
    st.caption(
        "Tải lên bản gốc và bản dịch (DOCX hoặc PDF — có thể trộn). "
        "Chọn **Text mode** để đối chiếu từng đoạn, hoặc **Vision mode** để "
        "Gemini xem ảnh từng trang và báo cáo chỉ những khác biệt thực sự quan trọng."
    )

    mode = st.radio(
        "Chế độ so sánh:",
        ["📝 Text mode — đối chiếu từng đoạn (chi tiết)",
         "🖼 Vision mode — so sánh ảnh từng trang (gọn, theo trang)"],
        key="rv_mode",
        horizontal=False,
    )
    is_vision = mode.startswith("🖼")

    col1, col2 = st.columns(2)
    with col1:
        orig_file = st.file_uploader("📄 Bản GỐC (.docx / .pdf)",
                                     type=["docx", "pdf"],
                                     key="rv_orig_upload")
    with col2:
        trans_file = st.file_uploader("📝 Bản DỊCH (.docx / .pdf)",
                                      type=["docx", "pdf"],
                                      key="rv_trans_upload")

    if is_vision:
        c1, c2, c3 = st.columns([2, 2, 1])
        with c1:
            target_lang = st.selectbox("Ngôn ngữ bản dịch:", LANGUAGES,
                                       index=0, key="rv_lang_v")
        with c2:
            page_spec = st.text_input(
                "Trang cần so sánh:", value="",
                placeholder="vd: 1-5,7  (để trống = tất cả)",
                key="rv_v_page_spec_input",
                help="1-indexed. VD: '1-3,5,8-10'. Trống = so sánh toàn bộ.",
            )
        with c3:
            st.markdown("&nbsp;", unsafe_allow_html=True)
            if st.button("🗑 Xoá kết quả", use_container_width=True, key="rv_clear_v"):
                _clear_state()
                st.rerun()

        focus_areas = st.text_area(
            "🎯 Trọng tâm review (tuỳ chọn):",
            value="",
            placeholder="vd: Tôi đặc biệt quan tâm thuật ngữ kỹ thuật và các bảng số liệu. "
                        "Bỏ qua phần mục lục và phụ lục.",
            height=70,
            key="rv_v_focus_input",
            help="Gemini sẽ ưu tiên kiểm tra những điểm bạn nêu ở đây cho từng trang.",
        )

        if st.button("🚀 So sánh theo trang (Vision)", type="primary",
                     disabled=not (orig_file and trans_file),
                     use_container_width=True):
            _run_vision_review(orig_file, trans_file, target_lang,
                               page_spec, focus_areas)

        if "rv_v_pages" in st.session_state:
            _render_vision_results()
        return

    c1, c2, c3, c4 = st.columns([2, 1, 1, 1])
    with c1:
        target_lang = st.selectbox("Ngôn ngữ bản dịch:", LANGUAGES,
                                   index=0, key="rv_lang")
    with c2:
        strictness = st.selectbox(
            "Độ nghiêm:",
            ["balanced", "strict", "permissive"],
            format_func=lambda s: {"balanced": "Cân bằng",
                                   "strict": "Nghiêm",
                                   "permissive": "Nới"}.get(s, s),
            key="rv_strict",
        )
    with c3:
        smart_align = st.checkbox(
            "🧠 Smart align", value=True, key="rv_smart_align_cb",
            help="Align theo nội dung (length/số/code) thay vì chỉ vị trí — fix lệch nhịp khi số đoạn 2 file khác nhau.",
        )
    with c4:
        st.markdown("&nbsp;", unsafe_allow_html=True)
        if st.button("🗑 Xoá kết quả", use_container_width=True):
            _clear_state()
            st.rerun()

    if st.button("🚀 Bắt đầu đánh giá", type="primary",
                 disabled=not (orig_file and trans_file),
                 use_container_width=True):
        _run_review(orig_file, trans_file, target_lang, strictness, smart_align)

    if "rv_reviews" in st.session_state:
        _render_results()


def _run_review(orig_file, trans_file, target_lang, strictness, smart_align: bool = True):
    log_ph = st.empty()
    log_lines: list = []
    add_log = make_log_adder(log_lines, log_ph)

    try:
        add_log(f"📄 Gốc: {orig_file.name}")
        add_log(f"📝 Dịch: {trans_file.name}")

        orig_blocks,  orig_fmt,  orig_bytes  = _extract_blocks_any(orig_file)
        trans_blocks, trans_fmt, trans_bytes = _extract_blocks_any(trans_file)
        add_log(f"   • Gốc ({orig_fmt}): {len(orig_blocks):,} block")
        add_log(f"   • Dịch ({trans_fmt}): {len(trans_blocks):,} block")
        if orig_fmt != trans_fmt:
            add_log(f"ℹ️ So sánh chéo định dạng: {orig_fmt} ↔ {trans_fmt}")

        pairs, warnings = align_blocks(orig_blocks, trans_blocks, smart=smart_align)
        for w in warnings[:8]:
            add_log(w)
        add_log(f"🔗 Aligned {len(pairs):,} pair ({'smart' if smart_align else 'positional'})")

        if not pairs:
            st.error("Không có pair nào để đánh giá.")
            return

        doc_context = build_doc_context(orig_blocks)
        chunks = chunk_pairs(pairs, REVIEW_CHUNK_SIZE)
        add_log(f"📦 Chia thành {len(chunks)} chunk (mỗi chunk ≤ {REVIEW_CHUNK_SIZE} pair)")

        timer_ph = st.empty()
        c1, c2, c3, c4 = st.columns(4)
        p1, p2, p3, p4 = c1.empty(), c2.empty(), c3.empty(), c4.empty()

        def render_stats(done_c, tok_in, tok_out):
            usd, vnd = calc_cost(tok_in, tok_out)
            p1.markdown(stat_box_html(f"{len(pairs):,}", "Pair"), unsafe_allow_html=True)
            p2.markdown(stat_box_html(f"{done_c}/{len(chunks)}", "Chunk"), unsafe_allow_html=True)
            p3.markdown(stat_box_html(f"${usd:.4f}", "USD"), unsafe_allow_html=True)
            p4.markdown(stat_box_html(f"{vnd:,.0f}₫", "VND"), unsafe_allow_html=True)
        render_stats(0, 0, 0)
        prog = st.progress(0, text="Đang chuẩn bị...")

        holder = {
            "reviews": [], "tok_in": 0, "tok_out": 0,
            "chunk_done": 0, "chunk_log": [], "done": False, "error": None,
        }
        t0 = time.time()
        threading.Thread(
            target=review_parallel,
            args=(holder, get_client(), chunks, target_lang, doc_context, strictness),
            daemon=True,
        ).start()

        last_logged = 0
        dot = 0
        while not holder["done"]:
            while last_logged < len(holder["chunk_log"]):
                e = holder["chunk_log"][last_logged]
                if e["error"]:
                    add_log(f"   ❌ Chunk {e['idx']+1}: {e['error'][:80]}")
                else:
                    add_log(f"   ✅ Chunk {e['idx']+1}: {e['size']} pair "
                            f"({e['in_t']:,}in/{e['out_t']:,}out tok)")
                last_logged += 1
            elapsed = time.time() - t0
            dots = "." * (dot % 4)
            timer_ph.markdown(
                timer_box_html(elapsed, f"🔄 Đánh giá song song{dots}"),
                unsafe_allow_html=True,
            )
            pct = int(holder["chunk_done"] / len(chunks) * 95) if chunks else 0
            prog.progress(pct, text=f"{holder['chunk_done']}/{len(chunks)} chunk...")
            render_stats(holder["chunk_done"], holder["tok_in"], holder["tok_out"])
            dot += 1
            time.sleep(0.5)

        while last_logged < len(holder["chunk_log"]):
            e = holder["chunk_log"][last_logged]
            if e["error"]:
                add_log(f"   ❌ Chunk {e['idx']+1}: {e['error'][:80]}")
            else:
                add_log(f"   ✅ Chunk {e['idx']+1}: {e['size']} pair "
                        f"({e['in_t']:,}in/{e['out_t']:,}out tok)")
            last_logged += 1

        if holder["error"]:
            st.error(holder["error"])
            return

        elapsed = time.time() - t0
        prog.progress(100, text="✅ Xong!")
        timer_ph.markdown(
            timer_done_html(elapsed, f"Đánh giá {len(pairs):,} pair xong!"),
            unsafe_allow_html=True,
        )

        st.session_state["rv_reviews"]   = holder["reviews"]
        st.session_state["rv_pairs"]     = {ob["id"]: (ob, tb) for ob, tb in pairs}
        st.session_state["rv_target_lang"] = target_lang
        st.session_state["rv_orig_name"]   = orig_file.name
        st.session_state["rv_trans_name"]  = trans_file.name
        st.session_state["rv_tok_in"]      = holder["tok_in"]
        st.session_state["rv_tok_out"]     = holder["tok_out"]
        st.session_state["rv_elapsed"]     = elapsed
        st.session_state["rv_orig_bytes"]    = orig_bytes
        st.session_state["rv_trans_bytes"]   = trans_bytes
        st.session_state["rv_orig_is_pdf"]   = (orig_fmt == "PDF")
        st.session_state["rv_trans_is_pdf"]  = (trans_fmt == "PDF")
        st.session_state["rv_edits"]         = {}

        usd, vnd = calc_cost(holder["tok_in"], holder["tok_out"])
        add_log("─" * 44)
        add_log(f"🎉 Xong trong {elapsed:.1f}s")
        add_log(f"💵 Chi phí: ${usd:.4f} USD ≈ {vnd:,.0f} VND")
        st.rerun()

    except Exception as e:
        st.error(f"Lỗi: {e}")


def _render_results():
    reviews = st.session_state["rv_reviews"]
    pairs   = st.session_state["rv_pairs"]

    st.markdown("---")
    st.markdown("### 📊 Kết quả đánh giá")

    # Quality score + breakdown
    metrics = compute_quality_score(reviews, pairs)
    score   = metrics["overall"]
    total   = metrics["total"]
    score_color = "#16a34a" if score >= 85 else ("#d97706" if score >= 60 else "#dc2626")

    sc_col, br_col = st.columns([1, 2])
    sc_col.markdown(
        f"""<div style="text-align:center;padding:18px;border:2px solid {score_color};
        border-radius:12px;background:#fafafa">
        <div style="font-size:14px;color:#666">Quality score</div>
        <div style="font-size:48px;font-weight:700;color:{score_color}">{score}</div>
        <div style="font-size:12px;color:#888">{metrics['total']} pair đã đánh giá</div>
        </div>""",
        unsafe_allow_html=True,
    )
    counts = metrics["counts"]
    pct = metrics["severity_pct"]
    br_col.markdown(
        f"""<div style="padding:12px">
        <div style="display:flex;gap:10px;margin-bottom:6px">
          <span style="color:#dc2626">🔴 Critical: <b>{counts['critical']}</b> ({pct['critical']}%)</span>
          <span style="color:#d97706">🟠 Major: <b>{counts['major']}</b> ({pct['major']}%)</span>
          <span style="color:#ca8a04">🟡 Minor: <b>{counts['minor']}</b> ({pct['minor']}%)</span>
          <span style="color:#16a34a">✅ OK: <b>{counts['ok']}</b> ({pct['ok']}%)</span>
        </div>
        <div style="height:24px;border-radius:6px;overflow:hidden;display:flex;border:1px solid #ddd">
          <div style="flex:{counts['critical']};background:#dc2626"></div>
          <div style="flex:{counts['major']};background:#f59e0b"></div>
          <div style="flex:{counts['minor']};background:#fcd34d"></div>
          <div style="flex:{counts['ok']};background:#22c55e"></div>
        </div>
        </div>""",
        unsafe_allow_html=True,
    )

    # Per-page score (only for PDF blocks that carry "page" attr)
    per_page = metrics.get("per_page") or {}
    if per_page:
        with st.expander(f"📑 Score theo trang ({len(per_page)} trang)", expanded=False):
            rows = []
            for page in sorted(per_page.keys()):
                s = per_page[page]
                rows.append({
                    "Trang":    page,
                    "Score":    s["score"],
                    "🔴":       s["critical"],
                    "🟠":       s["major"],
                    "🟡":       s["minor"],
                    "✅":       s["ok"],
                    "Total":    s["total"],
                })
            st.dataframe(pd.DataFrame(rows), hide_index=True, use_container_width=True)

    # Export
    cA, cB = st.columns(2)
    with cA:
        report_bytes = build_review_report_docx(pairs, reviews)
        st.download_button(
            "⬇️ Tải báo cáo DOCX",
            data=report_bytes,
            file_name=f"review_{st.session_state['rv_orig_name'].rsplit('.', 1)[0]}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
    with cB:
        rows = []
        for r in reviews:
            pair = pairs.get(r.get("id"))
            if not pair:
                continue
            ob, tb = pair
            rows.append({
                "id": r.get("id", ""), "severity": r.get("severity", "ok"),
                "issues": "; ".join(r.get("issues") or []),
                "original": ob["text"], "translated": tb["text"],
                "suggested": r.get("suggested", ""), "reason": r.get("reason", ""),
            })
        df_csv = pd.DataFrame(rows).to_csv(index=False).encode("utf-8")
        st.download_button(
            "⬇️ Tải CSV",
            data=df_csv,
            file_name=f"review_{st.session_state['rv_orig_name'].rsplit('.', 1)[0]}.csv",
            mime="text/csv",
            use_container_width=True,
        )

    # Filter
    st.markdown("#### 🔎 Lọc đoạn có vấn đề")
    sev_filter = st.multiselect(
        "Hiển thị severity:",
        ["critical", "major", "minor", "ok"],
        default=["critical", "major", "minor"],
    )

    filtered = [r for r in reviews if r.get("severity", "ok") in sev_filter]
    sev_order = {"critical": 0, "major": 1, "minor": 2, "ok": 3}
    filtered.sort(key=lambda r: sev_order.get(r.get("severity"), 4))

    st.caption(f"Hiển thị {len(filtered)}/{total} đoạn")

    edits = st.session_state.get("rv_edits", {}) or {}

    for r in filtered[:200]:
        rid = r.get("id", "")
        pair = pairs.get(rid)
        if not pair:
            continue
        ob, tb = pair
        sev = r.get("severity", "ok")
        sev_emoji = {"critical": "🔴", "major": "🟠",
                     "minor": "🟡", "ok": "✅"}.get(sev, "⚪")
        suggested = (r.get("suggested") or "").strip()
        current_trans = edits.get(rid, tb["text"])

        with st.container(border=True):
            top = st.columns([3, 1])
            edited_badge = " ✏️" if rid in edits else ""
            top[0].markdown(f"**{sev_emoji} {sev.upper()}** — `{rid}`{edited_badge}")
            issues = r.get("issues") or []
            if issues:
                top[1].markdown(f"*{', '.join(issues)}*")

            cc1, cc2 = st.columns(2)
            cc1.markdown("**Gốc**")
            cc1.text_area(" ", ob["text"], height=80, disabled=True,
                          key=f"rv_orig_{rid}", label_visibility="collapsed")
            cc2.markdown("**Bản dịch (chỉnh sửa được):**")
            new_val = cc2.text_area(
                " ", current_trans, height=80,
                key=f"rv_tr_edit_{rid}", label_visibility="collapsed",
            )
            if new_val != tb["text"]:
                edits[rid] = new_val
            elif rid in edits:
                edits.pop(rid)

            if suggested and suggested != tb["text"]:
                old_html, new_html = word_diff_html(current_trans, suggested)
                st.markdown("**💡 Đề xuất (diff với bản hiện tại):**")
                st.markdown(
                    f"""<div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;
                    font-size:15px;line-height:1.7;margin-bottom:8px">
                    <div style="padding:12px;background:#f9f9f9;border:2px solid #ffaaaa;
                    border-radius:8px;color:#1a1a1a">
                    <div style="font-weight:700;color:#cc0000;margin-bottom:6px;font-size:13px">
                    ✂️ Hiện tại</div>
                    <span style="color:#1a1a1a">{old_html}</span></div>
                    <div style="padding:12px;background:#f9f9f9;border:2px solid #88dd88;
                    border-radius:8px;color:#1a1a1a">
                    <div style="font-weight:700;color:#006600;margin-bottom:6px;font-size:13px">
                    ✅ Đề xuất</div>
                    <span style="color:#1a1a1a">{new_html}</span></div></div>""",
                    unsafe_allow_html=True,
                )
                if st.button("⬆️ Áp dụng bản đề xuất", key=f"rv_apply_{rid}"):
                    edits[rid] = suggested
                    st.session_state["rv_edits"] = edits
                    st.rerun()

            if r.get("reason"):
                st.markdown(f"**📝 Lý do:** {r['reason']}")

    st.session_state["rv_edits"] = edits

    if len(filtered) > 200:
        st.info("Chỉ hiển thị 200 đoạn đầu. Tải báo cáo DOCX/CSV để xem đầy đủ.")

    # Apply edits & re-render translated file
    if edits:
        st.markdown("---")
        st.markdown(f"### ✏️ Áp dụng {len(edits)} chỉnh sửa")
        is_pdf = st.session_state.get("rv_trans_is_pdf", False)

        if st.button(f"💾 Xuất {'PDF' if is_pdf else 'DOCX'} với {len(edits)} bản sửa",
                     type="primary", use_container_width=True, key="rv_export_edits"):
            with st.spinner("Đang áp dụng + render file mới..."):
                try:
                    if is_pdf:
                        orig_bytes = st.session_state.get("rv_orig_bytes")
                        trans_bytes = st.session_state.get("rv_trans_bytes")
                        out_bytes = apply_edits_to_pdf(trans_bytes, orig_bytes, edits)
                        out_name = st.session_state["rv_trans_name"].rsplit(".", 1)[0] + "_edited.pdf"
                        mime = "application/pdf"
                    else:
                        trans_bytes = st.session_state.get("rv_trans_bytes")
                        out_bytes = apply_edits_to_docx(trans_bytes, edits)
                        out_name = st.session_state["rv_trans_name"].rsplit(".", 1)[0] + "_edited.docx"
                        mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    st.session_state["rv_edited_bytes"] = out_bytes
                    st.session_state["rv_edited_name"]  = out_name
                    st.session_state["rv_edited_mime"]  = mime
                except Exception as e:
                    st.error(f"Lỗi khi áp dụng edits: {e}")

        if "rv_edited_bytes" in st.session_state:
            st.download_button(
                f"⬇️ Tải file đã sửa: {st.session_state['rv_edited_name']}",
                data=st.session_state["rv_edited_bytes"],
                file_name=st.session_state["rv_edited_name"],
                mime=st.session_state["rv_edited_mime"],
                use_container_width=True,
            )


# ══════════════════════════════════════════════════════════════════════════════
# VISION MODE — render pages → Gemini Vision per-page comparison
# ══════════════════════════════════════════════════════════════════════════════
def _run_vision_review(orig_file, trans_file, target_lang,
                       page_spec: str = "", focus_areas: str = ""):
    log_ph = st.empty()
    log_lines: list = []
    add_log = make_log_adder(log_lines, log_ph)

    try:
        orig_bytes  = orig_file.read()
        trans_bytes = trans_file.read()
        add_log(f"📄 Gốc: {orig_file.name}")
        add_log(f"📝 Dịch: {trans_file.name}")

        render_ph = st.empty()

        def _render_log(label):
            def cb(done, total, stage):
                if stage == "converting":
                    render_ph.info(f"🔄 {label}: đang chuyển DOCX → PDF...")
                else:
                    render_ph.info(f"🖼 {label}: render {done}/{total} trang...")
            return cb

        orig_images  = render_doc_as_page_images(orig_bytes,  orig_file.name,  _render_log("Gốc"))
        trans_images = render_doc_as_page_images(trans_bytes, trans_file.name, _render_log("Dịch"))
        render_ph.empty()
        add_log(f"   • Gốc: {len(orig_images)} trang")
        add_log(f"   • Dịch: {len(trans_images)} trang")

        if not orig_images or not trans_images:
            st.error("Không render được trang nào. Kiểm tra file đầu vào.")
            return

        n_pairs = min(len(orig_images), len(trans_images))
        if len(orig_images) != len(trans_images):
            add_log(f"⚠️ Số trang khác: gốc={len(orig_images)} vs dịch={len(trans_images)}. "
                    f"Ghép {n_pairs} cặp đầu.")

        selected_pages = parse_page_spec(page_spec, n_pairs)
        if not selected_pages:
            st.error(f"Không có trang hợp lệ trong '{page_spec}'. Tài liệu có {n_pairs} trang.")
            return
        if len(selected_pages) < n_pairs:
            add_log(f"📌 Chỉ so sánh {len(selected_pages)}/{n_pairs} trang theo yêu cầu: {page_spec}")

        n_compare = len(selected_pages)
        timer_ph = st.empty()
        c1, c2, c3, c4 = st.columns(4)
        p1, p2, p3, p4 = c1.empty(), c2.empty(), c3.empty(), c4.empty()

        def render_stats(done_c, tok_in, tok_out):
            usd, vnd = calc_cost(tok_in, tok_out)
            p1.markdown(stat_box_html(f"{n_compare}", "Trang gửi"), unsafe_allow_html=True)
            p2.markdown(stat_box_html(f"{done_c}/{n_compare}", "Đã so sánh"), unsafe_allow_html=True)
            p3.markdown(stat_box_html(f"${usd:.4f}", "USD"), unsafe_allow_html=True)
            p4.markdown(stat_box_html(f"{vnd:,.0f}₫", "VND"), unsafe_allow_html=True)
        render_stats(0, 0, 0)
        prog = st.progress(0, text="Đang chuẩn bị...")

        holder = {
            "pages": [], "tok_in": 0, "tok_out": 0,
            "page_done": 0, "page_log": [], "warnings": [],
            "done": False, "error": None,
        }
        t0 = time.time()
        threading.Thread(
            target=compare_pages_parallel,
            args=(holder, get_client(), orig_images, trans_images,
                  target_lang, selected_pages, focus_areas),
            daemon=True,
        ).start()

        last_logged = 0
        dot = 0
        while not holder["done"]:
            while last_logged < len(holder["page_log"]):
                e = holder["page_log"][last_logged]
                if e["error"]:
                    add_log(f"   ❌ Trang {e['page']}: {e['error'][:80]}")
                else:
                    sev_emoji = {"critical": "🔴", "major": "🟠",
                                 "minor": "🟡", "ok": "✅"}.get(e["severity"], "⚪")
                    add_log(f"   {sev_emoji} Trang {e['page']}: {e['issues']} vấn đề "
                            f"({e['in_t']:,}in/{e['out_t']:,}out tok)")
                last_logged += 1
            elapsed = time.time() - t0
            dots = "." * (dot % 4)
            timer_ph.markdown(
                timer_box_html(elapsed, f"🔄 So sánh trang song song{dots}"),
                unsafe_allow_html=True,
            )
            pct = int(holder["page_done"] / n_compare * 95) if n_compare else 0
            prog.progress(pct, text=f"{holder['page_done']}/{n_compare} trang...")
            render_stats(holder["page_done"], holder["tok_in"], holder["tok_out"])
            dot += 1
            time.sleep(0.5)

        while last_logged < len(holder["page_log"]):
            e = holder["page_log"][last_logged]
            if e["error"]:
                add_log(f"   ❌ Trang {e['page']}: {e['error'][:80]}")
            else:
                sev_emoji = {"critical": "🔴", "major": "🟠",
                             "minor": "🟡", "ok": "✅"}.get(e["severity"], "⚪")
                add_log(f"   {sev_emoji} Trang {e['page']}: {e['issues']} vấn đề "
                        f"({e['in_t']:,}in/{e['out_t']:,}out tok)")
            last_logged += 1

        if holder["error"]:
            st.error(holder["error"])
            return

        pages = sorted(holder["pages"], key=lambda x: x.get("page", 0))

        # Executive summary across all pages
        prog.progress(96, text="✍️ Tổng hợp báo cáo chung...")
        add_log("✍️ Tổng hợp executive summary...")
        overall, in_t, out_t = synthesize_overall_summary(get_client(), pages, target_lang)
        holder["tok_in"]  += in_t
        holder["tok_out"] += out_t
        if overall:
            add_log(f"   ✅ Summary ({in_t:,}in/{out_t:,}out tok)")
        else:
            add_log("   ⚠️ Không tạo được summary (bỏ qua)")

        elapsed = time.time() - t0
        prog.progress(100, text="✅ Xong!")
        timer_ph.markdown(
            timer_done_html(elapsed, f"So sánh {n_compare} trang xong!"),
            unsafe_allow_html=True,
        )

        st.session_state["rv_v_pages"]            = pages
        st.session_state["rv_v_orig_images"]      = orig_images
        st.session_state["rv_v_trans_images"]     = trans_images
        st.session_state["rv_v_target_lang"]      = target_lang
        st.session_state["rv_v_orig_name"]        = orig_file.name
        st.session_state["rv_v_trans_name"]       = trans_file.name
        st.session_state["rv_v_tok_in"]           = holder["tok_in"]
        st.session_state["rv_v_tok_out"]          = holder["tok_out"]
        st.session_state["rv_v_elapsed"]          = elapsed
        st.session_state["rv_v_overall_summary"]  = overall
        st.session_state["rv_v_page_spec"]        = page_spec
        st.session_state["rv_v_done_pages"]       = sorted({p.get("page") for p in pages})
        st.session_state["rv_v_focus_areas"]      = focus_areas

        usd, vnd = calc_cost(holder["tok_in"], holder["tok_out"])
        add_log("─" * 44)
        add_log(f"🎉 Xong trong {elapsed:.1f}s")
        add_log(f"💵 Chi phí: ${usd:.4f} USD ≈ {vnd:,.0f} VND")
        st.rerun()

    except Exception as e:
        st.error(f"Lỗi: {e}")


def _retry_failed_vision_pages():
    """Re-run only error pages from previous result."""
    pages = st.session_state.get("rv_v_pages") or []
    error_pages = [p.get("page") for p in pages
                   if p.get("severity") == "error" or not p.get("page")]
    if not error_pages:
        return

    orig_images  = st.session_state.get("rv_v_orig_images")  or []
    trans_images = st.session_state.get("rv_v_trans_images") or []
    target_lang  = st.session_state.get("rv_v_target_lang", "")
    focus_areas  = st.session_state.get("rv_v_focus_areas", "")

    holder = {
        "pages": [], "tok_in": 0, "tok_out": 0,
        "page_done": 0, "page_log": [], "warnings": [],
        "done": False, "error": None,
    }
    with st.spinner(f"🔄 Chạy lại {len(error_pages)} trang lỗi..."):
        t = threading.Thread(
            target=compare_pages_parallel,
            args=(holder, get_client(), orig_images, trans_images,
                  target_lang, error_pages, focus_areas),
            daemon=True,
        )
        t.start()
        t.join(timeout=600)

    if holder["error"]:
        st.error(holder["error"])
        return

    # Merge: replace error pages with retried pages
    retried_by_page = {p.get("page"): p for p in holder["pages"]}
    merged = []
    for p in pages:
        pn = p.get("page")
        if pn in retried_by_page:
            merged.append(retried_by_page[pn])
        else:
            merged.append(p)
    st.session_state["rv_v_pages"] = sorted(merged, key=lambda x: x.get("page", 0))
    st.session_state["rv_v_tok_in"]  = (st.session_state.get("rv_v_tok_in")  or 0) + holder["tok_in"]
    st.session_state["rv_v_tok_out"] = (st.session_state.get("rv_v_tok_out") or 0) + holder["tok_out"]


def _render_vision_results():
    pages = st.session_state["rv_v_pages"]
    orig_images  = st.session_state.get("rv_v_orig_images")  or []
    trans_images = st.session_state.get("rv_v_trans_images") or []
    target_lang  = st.session_state.get("rv_v_target_lang", "")
    orig_name    = st.session_state.get("rv_v_orig_name",  "orig")
    trans_name   = st.session_state.get("rv_v_trans_name", "trans")
    overall_sum  = st.session_state.get("rv_v_overall_summary", "")

    st.markdown("---")
    st.markdown("### 📊 Kết quả so sánh (Vision mode)")

    if overall_sum:
        with st.container(border=True):
            st.markdown("#### 📋 Tóm tắt tổng quan")
            st.markdown(overall_sum)

    metrics = vision_quality_score(pages)
    score = metrics["overall"]
    ic    = metrics["issue_counts"]
    ps    = metrics["page_severity_counts"]
    score_color = "#16a34a" if score >= 85 else ("#d97706" if score >= 60 else "#dc2626")

    sc_col, br_col = st.columns([1, 2])
    sc_col.markdown(
        f"""<div style="text-align:center;padding:18px;border:2px solid {score_color};
        border-radius:12px;background:#fafafa">
        <div style="font-size:14px;color:#666">Quality score</div>
        <div style="font-size:48px;font-weight:700;color:{score_color}">{score}</div>
        <div style="font-size:12px;color:#888">{metrics['total_pages']} trang •
        {metrics['total_issues']} vấn đề</div>
        </div>""",
        unsafe_allow_html=True,
    )
    br_col.markdown(
        f"""<div style="padding:12px">
        <div style="font-weight:600;margin-bottom:4px">Theo issue:</div>
        <div style="display:flex;gap:14px;margin-bottom:10px">
          <span style="color:#dc2626">🔴 Critical: <b>{ic['critical']}</b></span>
          <span style="color:#d97706">🟠 Major: <b>{ic['major']}</b></span>
          <span style="color:#ca8a04">🟡 Minor: <b>{ic['minor']}</b></span>
        </div>
        <div style="font-weight:600;margin-bottom:4px">Theo trang:</div>
        <div style="display:flex;gap:14px">
          <span style="color:#dc2626">🔴 {ps['critical']}</span>
          <span style="color:#d97706">🟠 {ps['major']}</span>
          <span style="color:#ca8a04">🟡 {ps['minor']}</span>
          <span style="color:#16a34a">✅ {ps['ok']}</span>
          <span style="color:#666">⚠️ Lỗi: {ps['error']}</span>
        </div>
        </div>""",
        unsafe_allow_html=True,
    )

    # Export DOCX + CSV
    cA, cB = st.columns(2)
    with cA:
        report = build_vision_report_docx(pages, orig_name, trans_name,
                                          target_lang, overall_sum)
        st.download_button(
            "⬇️ Tải báo cáo DOCX",
            data=report,
            file_name=f"vision_review_{orig_name.rsplit('.', 1)[0]}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
            key="rv_v_dl_docx",
        )
    with cB:
        rows = []
        for p in pages:
            page_num = p.get("page")
            for iss in p.get("issues") or []:
                rows.append({
                    "page":        page_num,
                    "severity":    iss.get("severity", ""),
                    "category":    iss.get("category", ""),
                    "location":    iss.get("location", ""),
                    "original":    iss.get("original", ""),
                    "translated":  iss.get("translated", ""),
                    "description": iss.get("description", ""),
                    "suggested":   iss.get("suggested", ""),
                })
        df_csv = pd.DataFrame(rows).to_csv(index=False).encode("utf-8")
        st.download_button(
            "⬇️ Tải CSV",
            data=df_csv,
            file_name=f"vision_review_{orig_name.rsplit('.', 1)[0]}.csv",
            mime="text/csv",
            use_container_width=True,
            key="rv_v_dl_csv",
        )

    # Retry failed pages
    error_pages = [p.get("page") for p in pages if p.get("severity") == "error"]
    if error_pages:
        rc1, rc2 = st.columns([3, 1])
        rc1.warning(f"⚠️ Có {len(error_pages)} trang lỗi: {sorted(error_pages)}")
        with rc2:
            if st.button("🔄 Thử lại trang lỗi", use_container_width=True,
                         key="rv_v_retry"):
                _retry_failed_vision_pages()
                st.rerun()

    # Category aggregation
    by_cat = aggregate_issues_by_category(pages)
    if by_cat:
        with st.expander(f"📊 Tổng hợp theo loại vấn đề ({len(by_cat)} loại)",
                         expanded=False):
            cat_rows = [{
                "Loại":      r["category"],
                "Tổng":      r["total"],
                "🔴":        r["critical"],
                "🟠":        r["major"],
                "🟡":        r["minor"],
                "Trang xuất hiện": ", ".join(str(p) for p in r["pages"][:12])
                                   + ("..." if len(r["pages"]) > 12 else ""),
            } for r in by_cat]
            st.dataframe(pd.DataFrame(cat_rows), hide_index=True,
                         use_container_width=True)

    # View mode toggle + filter
    st.markdown("#### 🔎 Hiển thị")
    vc1, vc2 = st.columns([1, 2])
    with vc1:
        view_mode = st.radio(
            "Sắp xếp:", ["Theo trang", "Theo mức độ"],
            horizontal=True, key="rv_v_view_mode",
        )
    with vc2:
        sev_filter = st.multiselect(
            "Severity:",
            ["critical", "major", "minor", "ok", "error"],
            default=["critical", "major", "minor"],
            key="rv_v_filter",
        )

    if view_mode == "Theo trang":
        _render_vision_by_page(pages, sev_filter, orig_images, trans_images)
    else:
        _render_vision_by_severity(pages, sev_filter)


def _render_issue_card(iss: dict, page_num=None):
    isev = iss.get("severity", "minor")
    isev_emoji = {"critical": "🔴", "major": "🟠", "minor": "🟡"}.get(isev, "⚪")
    with st.container(border=True):
        head = st.columns([2, 2, 1])
        head[0].markdown(f"**{isev_emoji} {isev.upper()}** "
                         f"· `{iss.get('category', '')}`")
        head[1].markdown(f"*{iss.get('location', '')}*")
        if page_num is not None:
            head[2].markdown(f"📄 Trang {page_num}")

        if iss.get("original") or iss.get("translated"):
            cc1, cc2 = st.columns(2)
            cc1.markdown("**Gốc:**")
            cc1.markdown(
                f"<div style='padding:8px;background:#fff7f7;"
                f"border:1px solid #ffd0d0;border-radius:6px;"
                f"color:#1a1a1a'>{iss.get('original','') or '—'}</div>",
                unsafe_allow_html=True,
            )
            cc2.markdown("**Dịch:**")
            cc2.markdown(
                f"<div style='padding:8px;background:#f7fff7;"
                f"border:1px solid #d0ffd0;border-radius:6px;"
                f"color:#1a1a1a'>{iss.get('translated','') or '—'}</div>",
                unsafe_allow_html=True,
            )

        if iss.get("description"):
            st.markdown(f"**📝 Mô tả:** {iss['description']}")
        if iss.get("suggested"):
            st.markdown(f"**💡 Đề xuất:** {iss['suggested']}")


def _render_vision_by_page(pages, sev_filter, orig_images, trans_images):
    sev_order = {"critical": 0, "major": 1, "minor": 2, "ok": 3, "error": 4}
    visible = [p for p in pages if p.get("severity", "ok") in sev_filter]
    visible.sort(key=lambda p: (sev_order.get(p.get("severity"), 5),
                                p.get("page", 0)))
    st.caption(f"Hiển thị {len(visible)}/{len(pages)} trang")

    show_pages_btn = st.checkbox(
        "🖼 Hiện ảnh 2 trang khi mở chi tiết", value=False,
        key="rv_v_show_imgs",
        help="Mặc định tắt để trang nhẹ. Bật để xem ảnh gốc + dịch cạnh nhau."
    )

    for page in visible:
        page_num = page.get("page", "?")
        sev = page.get("severity", "ok")
        sev_emoji = {"critical": "🔴", "major": "🟠",
                     "minor": "🟡", "ok": "✅",
                     "error": "⚠️"}.get(sev, "⚪")
        n_iss = len(page.get("issues") or [])

        with st.expander(
            f"{sev_emoji} Trang {page_num} — {sev.upper()} — {n_iss} vấn đề",
            expanded=(sev in ("critical", "major")),
        ):
            if page.get("summary"):
                st.markdown(f"**📝 Tổng quan:** {page['summary']}")

            if show_pages_btn and isinstance(page_num, int):
                idx = page_num - 1
                if 0 <= idx < len(orig_images) and 0 <= idx < len(trans_images):
                    img_c1, img_c2 = st.columns(2)
                    img_c1.markdown(f"**Gốc — trang {page_num}**")
                    img_c1.image(orig_images[idx], use_container_width=True)
                    img_c2.markdown(f"**Dịch — trang {page_num}**")
                    img_c2.image(trans_images[idx], use_container_width=True)

            for iss in page.get("issues") or []:
                _render_issue_card(iss)


def _render_vision_by_severity(pages, sev_filter):
    issues = flatten_issues(pages)
    issues = [i for i in issues if i.get("severity", "minor") in sev_filter]
    st.caption(f"Hiển thị {len(issues)} vấn đề (đã sort theo mức độ)")

    if not issues:
        st.info("Không có vấn đề nào khớp với bộ lọc.")
        return

    LIMIT = 200
    for iss in issues[:LIMIT]:
        _render_issue_card(iss, page_num=iss.get("page"))

    if len(issues) > LIMIT:
        st.info(f"Chỉ hiển thị {LIMIT}/{len(issues)} vấn đề. "
                f"Tải DOCX/CSV để xem đầy đủ.")
