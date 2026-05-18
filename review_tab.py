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
)


SS_KEYS = ("rv_reviews", "rv_pairs", "rv_target_lang", "rv_orig_name",
           "rv_trans_name", "rv_tok_in", "rv_tok_out", "rv_elapsed")


def _clear_state():
    for k in SS_KEYS:
        st.session_state.pop(k, None)


def _extract_blocks_any(file_obj) -> tuple[list[dict], str]:
    """
    Detect file format from name and extract blocks.
    Returns (blocks, format_label) where format_label is "DOCX" or "PDF".
    """
    name = (file_obj.name or "").lower()
    data = file_obj.read()
    if name.endswith(".pdf"):
        return extract_pdf_blocks(data), "PDF"
    return extract_docx_blocks(data), "DOCX"


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

    c1, c2, c3 = st.columns([2, 1, 1])
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
        st.markdown("&nbsp;", unsafe_allow_html=True)
        if st.button("🗑 Xoá kết quả", use_container_width=True):
            _clear_state()
            st.rerun()

    if st.button("🚀 Bắt đầu đánh giá", type="primary",
                 disabled=not (orig_file and trans_file),
                 use_container_width=True):
        _run_review(orig_file, trans_file, target_lang, strictness)

    if "rv_reviews" in st.session_state:
        _render_results()


def _run_review(orig_file, trans_file, target_lang, strictness):
    log_ph = st.empty()
    log_lines: list = []
    add_log = make_log_adder(log_lines, log_ph)

    try:
        add_log(f"📄 Gốc: {orig_file.name}")
        add_log(f"📝 Dịch: {trans_file.name}")

        orig_blocks,  orig_fmt  = _extract_blocks_any(orig_file)
        trans_blocks, trans_fmt = _extract_blocks_any(trans_file)
        add_log(f"   • Gốc ({orig_fmt}): {len(orig_blocks):,} block")
        add_log(f"   • Dịch ({trans_fmt}): {len(trans_blocks):,} block")
        if orig_fmt != trans_fmt:
            add_log(f"ℹ️ So sánh chéo định dạng: {orig_fmt} ↔ {trans_fmt}")

        pairs, warnings = align_blocks(orig_blocks, trans_blocks)
        for w in warnings[:5]:
            add_log(w)
        add_log(f"🔗 Aligned {len(pairs):,} pair")

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

    counts = summarize_reviews(reviews)
    total  = len(reviews)
    c1, c2, c3, c4 = st.columns(4)
    c1.markdown(stat_box_html(f"{counts['critical']}", "🔴 Critical"), unsafe_allow_html=True)
    c2.markdown(stat_box_html(f"{counts['major']}", "🟠 Major"), unsafe_allow_html=True)
    c3.markdown(stat_box_html(f"{counts['minor']}", "🟡 Minor"), unsafe_allow_html=True)
    c4.markdown(stat_box_html(f"{counts['ok']}", "✅ OK"), unsafe_allow_html=True)

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

    for r in filtered[:200]:
        pair = pairs.get(r.get("id"))
        if not pair:
            continue
        ob, tb = pair
        sev = r.get("severity", "ok")
        sev_emoji = {"critical": "🔴", "major": "🟠",
                     "minor": "🟡", "ok": "✅"}.get(sev, "⚪")

        with st.container(border=True):
            top = st.columns([3, 1])
            top[0].markdown(f"**{sev_emoji} {sev.upper()}** — `{r.get('id', '')}`")
            issues = r.get("issues") or []
            if issues:
                top[1].markdown(f"*{', '.join(issues)}*")

            cc1, cc2 = st.columns(2)
            cc1.markdown("**Gốc**")
            cc1.text_area(" ", ob["text"], height=80, disabled=True,
                          key=f"rv_orig_{r.get('id', '')}", label_visibility="collapsed")
            cc2.markdown("**Bản dịch hiện tại**")
            cc2.text_area(" ", tb["text"], height=80, disabled=True,
                          key=f"rv_tr_{r.get('id', '')}", label_visibility="collapsed")

            if r.get("suggested"):
                st.markdown("**💡 Đề xuất:**")
                st.text_area(" ", r["suggested"], height=80, disabled=True,
                             key=f"rv_sug_{r.get('id', '')}", label_visibility="collapsed")
            if r.get("reason"):
                st.markdown(f"**📝 Lý do:** {r['reason']}")

    if len(filtered) > 200:
        st.info("Chỉ hiển thị 200 đoạn đầu. Tải báo cáo DOCX/CSV để xem đầy đủ.")
