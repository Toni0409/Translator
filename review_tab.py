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


SS_KEYS = ("rv_reviews", "rv_pairs", "rv_target_lang", "rv_orig_name",
           "rv_trans_name", "rv_tok_in", "rv_tok_out", "rv_elapsed",
           "rv_trans_bytes", "rv_orig_bytes", "rv_trans_is_pdf", "rv_orig_is_pdf",
           "rv_edits", "rv_smart_align")


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
        "LLM đối chiếu từng đoạn, đánh giá độ chính xác + văn phong, "
        "liệt kê đoạn có vấn đề kèm đề xuất + lý do."
    )

    col1, col2 = st.columns(2)
    with col1:
        orig_file = st.file_uploader("📄 Bản GỐC (.docx / .pdf)",
                                     type=["docx", "pdf"],
                                     key="rv_orig_upload")
    with col2:
        trans_file = st.file_uploader("📝 Bản DỊCH (.docx / .pdf)",
                                      type=["docx", "pdf"],
                                      key="rv_trans_upload")

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
