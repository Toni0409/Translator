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
    MAX_WORD_WORKERS, NO_TRANSLATE_ROLES,
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
    tm_lookup, tm_store, tm_key,
    checkpoint_save, checkpoint_load, checkpoint_clear,
    export_bilingual_docx, quality_check, validate_docx_output,
)
from domain_glossary import detect_subdomain, seed_for_direction


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


def _current_translatable(a: dict) -> list:
    """Translatable blocks dựa trên `a["role_toggles"]` hiện tại (P3.3).

    Nếu user chưa toggle gì, fallback về list ban đầu từ extract.
    """
    toggles = a.get("role_toggles")
    if not toggles:
        return a["translatable"]
    skip = {r for r, on in toggles.items() if not on}
    return [b for b in a["blocks"]
            if b["role"] not in skip
            and not b.get("_toc_mirror")]


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
        image_alt_cnt = stats.get("by_role", {}).get("image_alt", 0)
        if image_alt_cnt > 0:
            add_log(f"   • Image alt-texts:      {image_alt_cnt:,} (sẽ dịch)")

        # target_lang/source_lang nhận từ caller (xem _resolve_langs).
        # TM preview — kèm breakdown per role
        tm           = _ensure_tm()
        tm_cached, _ = tm_lookup(translatable, tm, target_lang)
        if tm_cached:
            add_log(f"💾 TM: {len(tm_cached):,}/{len(translatable):,} đoạn dùng cache (skip API)")
            # Per-role hit rate breakdown
            id_to_role = {b["id"]: b.get("role", "?") for b in translatable}
            role_total: dict = {}
            role_hits:  dict = {}
            for b in translatable:
                role_total[b["role"]] = role_total.get(b["role"], 0) + 1
            for bid in tm_cached:
                r = id_to_role.get(bid, "?")
                role_hits[r] = role_hits.get(r, 0) + 1
            role_lines = [
                f"      {r}: {role_hits[r]}/{role_total[r]}"
                for r in sorted(role_hits, key=lambda x: -role_hits[x])
            ]
            if role_lines:
                add_log("   TM hit theo role:")
                for line in role_lines[:6]:
                    add_log(line)
        else:
            add_log(f"💾 TM: 0 hit (TM hiện có {len(tm):,} entry)")

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
            "tm_hits":       len(tm_cached),
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
# PHASE 1 — RENDER (stats + glossary editor + Dịch button)
# ══════════════════════════════════════════════════════════════════════════════
def _render_analysis_panel():
    a = st.session_state["word_analysis"]

    # ── TIER 1 — Quick mode: stats + primary action ──────────────────────
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

    # ── TIER 2 — Single advanced expander (no nested expanders!) ─────────
    with st.expander("⚙️ Tuỳ chỉnh nâng cao", expanded=False):

        # ── 1. Role toggles ──────────────────────────────────────────────
        st.markdown("##### 🎚 Chọn loại nội dung dịch")
        st.caption("Mặc định bỏ qua header/footer (thường chứa page#, tên file...). "
                   "Tick để dịch loại đó.")
        role_options = [
            ("header",       "📄 Header (đầu trang)"),
            ("footer",       "📄 Footer (chân trang)"),
            ("body_repeated","🔁 Body lặp (watermark, banner)"),
            ("comment",      "💬 Comment"),
            ("footnote",     "📌 Footnote"),
            ("endnote",      "📌 Endnote"),
            ("image_alt",    "🖼 Image alt-text"),
            ("toc",          "📑 TOC (mục lục)"),
        ]
        enabled = {}
        rcols = st.columns(2)
        for i, (role, label) in enumerate(role_options):
            default_on = role not in {"header", "footer", "body_repeated"}
            enabled[role] = rcols[i % 2].checkbox(
                label, value=a.get("role_toggles", {}).get(role, default_on),
                key=f"word_role_toggle_{role}",
            )
        a["role_toggles"] = enabled
        st.session_state["word_analysis"] = a
        skip_roles_panel = {r for r, on in enabled.items() if not on}
        translatable_now = [b for b in a["blocks"]
                            if b["role"] not in skip_roles_panel
                            and not b.get("_toc_mirror")]
        chars_now = sum(len(b["text"]) for b in translatable_now)
        st.caption(f"➡️ Sẽ dịch: **{len(translatable_now):,} đoạn** ({chars_now:,} ký tự) "
                   f"sau khi áp dụng toggle.")

        st.divider()

        # ── 2. Cost cap warning ──────────────────────────────────────────
        st.markdown("##### 💰 Cảnh báo chi phí")
        cost_cap = st.number_input(
            "Ngưỡng cảnh báo (USD) — 0 = tắt",
            min_value=0.0, max_value=10.0, value=0.5, step=0.1,
            key="word_cost_cap",
        )

        st.divider()

        # ── 3. Glossary (editor + LLM suggest + export/import) ───────────
        st.markdown(f"##### 📖 Glossary ({len(a['glossary'])} thuật ngữ)")

        # LLM glossary suggest — works whether or not heuristic glossary is empty
        if st.button("🪄 Gợi ý glossary từ LLM", key="word_glossary_suggest"):
            from word_backend import suggest_glossary
            with st.spinner("Đang phân tích doc..."):
                suggestions = suggest_glossary(
                    get_client(), a["blocks"], a["target_lang"],
                )
                st.session_state["word_glossary_suggestions"] = suggestions

        if "word_glossary_suggestions" in st.session_state:
            sugg = st.session_state["word_glossary_suggestions"]
            if sugg:
                st.markdown(f"**🪄 {len(sugg)} gợi ý** — chọn để thêm vào glossary:")
                df_sug = pd.DataFrame([{
                    "✓":         False,
                    "Term":      s.get("term", ""),
                    "Suggested": s.get("suggested", ""),
                    "Note":      s.get("note", ""),
                } for s in sugg])
                edited_sug = st.data_editor(
                    df_sug, hide_index=True, use_container_width=True,
                    key="word_glossary_sugg_edit",
                    column_config={
                        "✓": st.column_config.CheckboxColumn(required=True),
                    },
                )
                if st.button("➕ Thêm các từ đã chọn", key="word_glossary_sugg_add"):
                    added = 0
                    for _, row in edited_sug.iterrows():
                        if row["✓"] and row["Term"]:
                            a["glossary"][row["Term"]] = row["Suggested"]
                            added += 1
                    st.session_state["word_analysis"] = a
                    st.success(f"Đã thêm {added} thuật ngữ vào glossary")
                    st.session_state.pop("word_glossary_suggestions", None)
                    st.rerun()
            else:
                st.info("LLM không tìm được thuật ngữ nào (hoặc API fail).")

        if not a["glossary"]:
            st.info("Không có thuật ngữ lặp lại đủ ngưỡng (≥3 lần). Vẫn dịch được — chỉ là không có guarantee consistency.")
        else:
            st.caption(
                "💡 Sửa bản dịch để bắt buộc terminology cụ thể. "
                "Xóa cell **Bản dịch** (để trống) để loại entry khỏi glossary. "
                "Có thể thêm dòng mới."
            )
            # Versioned key — đổi khi restore để force re-render editor
            gver = st.session_state.get("word_glossary_editor_ver", 0)
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
                key=f"word_glossary_editor_v{gver}",
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

            # Diff so với glossary ban đầu — cảnh báo nếu user lỡ xoá entries
            initial = a.get("_glossary_initial") or {}
            removed = set(initial) - set(new_gloss)
            added   = set(new_gloss) - set(initial)
            if removed or added:
                cols = st.columns([4, 1])
                msg = f"📊 Glossary: {len(initial)} → {len(new_gloss)}"
                if removed:
                    msg += f"  •  ⚠️ Đã loại **{len(removed)}** entry"
                if added:
                    msg += f"  •  ➕ Thêm **{len(added)}**"
                cols[0].caption(msg)
                if removed:
                    if cols[1].button("↩️ Khôi phục", key="word_glossary_restore",
                                      help="Reset glossary về kết quả phân tích ban đầu"):
                        a["glossary"] = dict(initial)
                        st.session_state["word_analysis"] = a
                        st.session_state["word_glossary_editor_ver"] = gver + 1
                        st.rerun()

        # Glossary export/import — always shown so user can import even when empty
        gcolA, gcolB = st.columns(2)
        with gcolA:
            if a["glossary"]:
                import json as _json
                st.download_button(
                    "⬇️ Export Glossary (.json)",
                    data=_json.dumps(a["glossary"], ensure_ascii=False, indent=2),
                    file_name="word_glossary.json", mime="application/json",
                    use_container_width=True, key="word_glossary_export",
                )
            else:
                st.caption("(Glossary rỗng — chưa có gì để export)")
        with gcolB:
            gup = st.file_uploader("⬆️ Import Glossary (.json)", type=["json"],
                                   key="word_glossary_import",
                                   label_visibility="collapsed")
            if gup:
                try:
                    import json as _json
                    gloaded = _json.loads(gup.read())
                    if isinstance(gloaded, dict):
                        a["glossary"].update(gloaded)
                        st.session_state["word_analysis"] = a
                        st.success(f"Đã import {len(gloaded)} glossary entries")
                except Exception as e:
                    st.error(str(e))

        # Seed restore (P4.8) — merge lại seed mà KHÔNG xoá user edits.
        seed_gloss = a.get("seed_glossary") or {}
        if seed_gloss:
            sub_label = ", ".join(sorted(a.get("subdomains") or set())) or "—"
            if st.button(
                f"🏗 Khôi phục seed thuật ngữ ngành ({sub_label}: {len(seed_gloss)} term)",
                key="word_glossary_seed_restore",
                help="Merge seed glossary chuyên ngành thang máy/thang cuốn vào "
                     "glossary hiện tại. User edits được giữ — chỉ seed bị thiếu sẽ thêm lại.",
                use_container_width=True,
            ):
                cur = a.get("glossary") or {}
                added_back = 0
                for k, v in seed_gloss.items():
                    if k not in cur:
                        cur[k] = v
                        added_back += 1
                a["glossary"] = cur
                st.session_state["word_analysis"] = a
                gver_now = st.session_state.get("word_glossary_editor_ver", 0)
                st.session_state["word_glossary_editor_ver"] = gver_now + 1
                st.success(f"Đã merge lại {added_back} thuật ngữ seed "
                           f"(tổng glossary hiện có: {len(cur)})")
                st.rerun()

        st.divider()

        # ── 4. Translation Memory ────────────────────────────────────────
        tm = _ensure_tm()
        st.markdown(f"##### 💾 Translation Memory")
        st.caption(f"Hiện có **{len(tm):,} entry** trong TM.")
        tcolA, tcolB = st.columns(2)
        with tcolA:
            if tm:
                import json as _json
                st.download_button(
                    "⬇️ Export TM (.json)",
                    data=_json.dumps(tm, ensure_ascii=False, indent=2),
                    file_name="word_tm.json", mime="application/json",
                    use_container_width=True, key="word_tm_export",
                )
            else:
                st.caption("(TM hiện tại rỗng)")
        with tcolB:
            up = st.file_uploader("⬆️ Import TM (.json)", type=["json"], key="word_tm_import",
                                  label_visibility="collapsed")
            if up:
                try:
                    import json as _json
                    loaded = _json.loads(up.read())
                    if isinstance(loaded, dict):
                        _ensure_tm().update(loaded)
                        st.success(f"Đã import {len(loaded)} TM entries")
                except Exception as e:
                    st.error(str(e))

        st.divider()

        # ── 5. Custom rules per role ─────────────────────────────────────
        st.markdown("##### ⚙️ Custom rules per role")
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

    # ── Action buttons (after advanced expander, still inside Phase 1) ───
    # Cost cap value — read from session_state since the widget lives inside
    # the (possibly collapsed) advanced expander.
    cost_cap = st.session_state.get("word_cost_cap", 0.5)

    # Dịch button — nói rõ còn bao nhiêu sau TM
    n_remain = len(a["translatable"]) - a["tm_hits"]
    label = (f"🚀 Dịch ngay ({n_remain:,} đoạn)"
             if n_remain > 0
             else f"🚀 Dịch ngay — apply {a['tm_hits']:,} TM hit (không gọi API)")

    # Estimate cost for remaining blocks — dùng translatable CURRENT theo role
    # toggles (P3.3), không dùng `a["translatable"]` stale từ lúc extract.
    if n_remain > 0:
        translatable = _current_translatable(a)
        tm = _ensure_tm()
        target_lang = a["target_lang"]
        _, remaining_blocks = tm_lookup(translatable, tm, target_lang)
        est_chars = sum(len(b["text"]) for b in remaining_blocks)
        est_in_tok = int(est_chars * 0.4)
        est_out_tok = int(est_in_tok * 1.5)
        est_usd, _ = calc_cost(est_in_tok, est_out_tok)
        st.caption(f"💰 Ước tính chi phí: ~${est_usd:.3f} USD cho {len(remaining_blocks):,} đoạn còn cần API")
        if cost_cap > 0 and est_usd > cost_cap:
            st.warning(
                f"⚠️ Ước tính ~${est_usd:.3f} USD vượt ngưỡng cảnh báo ${cost_cap:.2f} USD. "
                f"Xác nhận để tiếp tục."
            )
            cost_confirmed = st.checkbox(
                f"✅ Tôi xác nhận dịch với chi phí ước tính ~${est_usd:.3f} USD",
                key="word_cost_confirm",
            )
        else:
            cost_confirmed = True
    else:
        cost_confirmed = True

    btn_primary, btn_secondary = st.columns([3, 1])
    with btn_primary:
        if st.button(label, use_container_width=True, type="primary",
                     key="word_translate_phase2_btn",
                     disabled=(not cost_confirmed)):
            _run_full_translation()
            st.rerun()
    with btn_secondary:
        if st.button("🔄 Phân tích lại", use_container_width=True,
                     key="word_reanalyze_btn",
                     help="Xoá kết quả phân tích để upload lại file"):
            st.session_state.pop("word_analysis", None)
            st.session_state.pop("word_glossary_suggestions", None)
            st.rerun()


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
    tm           = _ensure_tm()

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

    timer_ph, prog, render_stats, add_log = _make_job_ui(
        "### 📊 Tiến độ dịch", "Đoạn văn", "### 📋 Nhật ký",
    )

    try:
        toc_mirror_cnt = sum(1 for b in blocks if b.get("_toc_mirror"))
        if toc_mirror_cnt:
            add_log(f"🔗 TOC: {toc_mirror_cnt} entries sẽ dùng lại translation từ heading")

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
                source_lang=source_lang,
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
        st.session_state["word_source_lang"]  = source_lang
        st.session_state["word_target_lang"]  = target_lang
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
        # Persist role toggles (for post-translation editor filter)
        st.session_state["word_role_toggles"] = toggles
        st.session_state["word_skip_roles"]   = skip_roles

        # Validate output DOCX (so sánh media với bản gốc — P2.7)
        try:
            out_bytes = apply_translations(docx_bytes, blocks, translations)
            val = validate_docx_output(out_bytes, original_bytes=docx_bytes)
            st.session_state["word_validation"] = val
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
    """Generic partial-translation runner (rescan / H-F). Có TM lookup."""
    if not missed:
        st.info("✨ Không có đoạn nào cần xử lý!")
        return

    doc_context  = st.session_state["word_doc_context"]
    target_lang  = st.session_state.get("word_target_lang") or LANG_EN[st.session_state["word_lang"]]
    source_lang  = st.session_state.get("word_source_lang") or _resolve_langs()[1]
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
            source_lang=source_lang,
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
            st.session_state["word_validation"] = val
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
    "image_alt":       "🖼 Image alt",
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

    skip_roles = st.session_state.get("word_skip_roles") or set(NO_TRANSLATE_ROLES)
    display_blocks = []
    for b in blocks:
        is_hf = b["role"] in skip_roles
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
    for k in ("word_translated_bytes_cache", "word_translations_version",
              "word_validation", "word_role_toggles", "word_skip_roles",
              "word_glossary_editor_ver", "word_image_ocr",
              "word_review_df", "word_review_df_version",
              "word_batch_result"):
        st.session_state.pop(k, None)
    for k in list(st.session_state.keys()):
        if k.startswith("word_editor") or k.startswith("word_glossary_editor") \
                or k in ("word_only_missed_filter", "word_show_hf_filter"):
            st.session_state.pop(k, None)
    st.session_state["word_editor_version"] = 0
    # KHÔNG xóa "word_tm" — TM persist xuyên suốt session


def _run_batch(uploaded_files, lang_word, source_lang: str, target_lang: str):
    """Batch-translate multiple DOCX files, sharing TM across all files."""
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
            tm          = _ensure_tm()
            translatable = [b for b in blocks if b["role"] not in NO_TRANSLATE_ROLES]
            cached, remaining = tm_lookup(translatable, tm, target_lang)
            translations = dict(cached)

            if remaining:
                chunks = chunk_blocks(remaining)
                timer_ph  = st.empty()
                prog_ph   = st.empty()

                def _noop_stats(*_a, **_kw): pass
                def _noop_log(_msg): pass

                new_tr, _, _, _ = _run_translation(
                    chunks, target_lang, doc_context,
                    timer_ph, prog_ph, _noop_stats, _noop_log,
                    len(remaining), prefix_label=uploaded.name[:20],
                    glossary=glossary,
                    source_lang=source_lang,
                )
                translations.update(new_tr)
                tm_store(remaining, new_tr, tm, target_lang)
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

    # ── 2 NÚT: Dịch cơ bản / Dịch nâng cao ───────────────────────────────
    col_basic, col_adv = st.columns(2)
    with col_basic:
        basic_clicked = st.button(
            "🚀  Dịch cơ bản",
            disabled=(uploaded_docx is None),
            use_container_width=True, type="primary", key="word_basic",
            help="Dịch nhanh với cài đặt mặc định — không cần chọn gì thêm",
        )
    with col_adv:
        advanced_clicked = st.button(
            "⚙️  Dịch nâng cao",
            disabled=(uploaded_docx is None),
            use_container_width=True, key="word_advanced",
            help="Xem stats + chỉnh glossary, role, custom rules, TM... trước khi dịch",
        )

    if basic_clicked:
        _clear_state()
        # Cơ bản = phân tích + dịch luôn. Set quick-mode TRƯỚC khi
        # _run_analysis() trigger st.rerun() (nếu set sau, flag bị mất).
        st.session_state["word_quick_mode"] = True
        _run_analysis(uploaded_docx, lang_word, source_lang, target_lang)

    if advanced_clicked:
        _clear_state()
        _run_analysis(uploaded_docx, lang_word, source_lang, target_lang)

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

        val          = st.session_state.get("word_validation")
        blocks       = st.session_state["word_blocks"]
        translations = st.session_state["word_translations"]
        missed_body  = find_missed(blocks, translations, hf_only=False)
        missed_hf    = find_missed(blocks, translations, hf_only=True)
        total_hf     = sum(1 for b in blocks if b["role"] in NO_TRANSLATE_ROLES)

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

        # 3) ACTION buttons — combined label+count (bỏ st.warning/st.info trùng)
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

        # 4) OCR — section riêng, tách biệt (user dùng thường xuyên)
        with st.expander("🖼  OCR & dịch text trong ảnh", expanded=False):
            st.caption(
                "Quét tất cả ảnh trong DOCX, OCR chữ và dịch sang ngôn ngữ đích. "
                "Xử lý song song — có thể tốn thêm chi phí Gemini Vision."
            )
            if st.button("🔍 Quét và dịch ảnh", key="word_ocr_btn"):
                from word_backend import extract_images_from_docx, ocr_and_translate_images
                imgs = extract_images_from_docx(st.session_state["word_docx_bytes"])
                if not imgs:
                    st.info("Không có ảnh trong DOCX.")
                else:
                    progress_bar = st.progress(0, text=f"Đang OCR 0/{len(imgs)} ảnh...")
                    _ocr_counter = [0]

                    def _on_progress(done, total):
                        _ocr_counter[0] = done
                        progress_bar.progress(done / total,
                                              text=f"Đang OCR {done}/{total} ảnh...")

                    ocr_target = st.session_state.get("word_target_lang") \
                                 or LANG_EN[st.session_state["word_lang"]]
                    results = ocr_and_translate_images(
                        get_client(), imgs,
                        ocr_target,
                        progress_callback=_on_progress,
                    )
                    progress_bar.empty()
                    st.session_state["word_image_ocr"] = (imgs, results)

            if "word_image_ocr" in st.session_state:
                imgs, results = st.session_state["word_image_ocr"]
                shown = 0
                for img in imgs:
                    r = results.get(img["id"], {})
                    if not r.get("has_text"):
                        continue
                    shown += 1
                    with st.container(border=True):
                        c1, c2 = st.columns([1, 2])
                        c1.image(img["data"], width=200, caption=img["filename"])
                        c2.markdown(f"**OCR (gốc):**\n```\n{r['ocr']}\n```")
                        c2.markdown(f"**Dịch:**\n```\n{r['translation']}\n```")
                        if r.get("error"):
                            c2.warning(f"⚠️ Lỗi: {r['error']}")

                errors = [r for r in results.values() if r.get("error") and not r.get("has_text")]
                if shown == 0 and not errors:
                    st.info(f"Đã quét {len(imgs)} ảnh — không phát hiện text.")
                if errors:
                    st.warning(f"⚠️ {len(errors)} ảnh gặp lỗi khi OCR.")

                if shown > 0:
                    if st.checkbox("📎 Chèn bản dịch OCR vào file DOCX output", key="word_ocr_insert"):
                        from word_backend import insert_ocr_captions_into_docx
                        base_bytes = _get_cached_translated_bytes()
                        ocr_docx = insert_ocr_captions_into_docx(base_bytes, imgs, results)
                        out_name_ocr = st.session_state["word_filename"].replace(
                            ".docx",
                            f"_translated_{st.session_state['word_lang'][:2]}_ocr.docx",
                        )
                        st.download_button(
                            label="⬇️  Tải Word đã dịch + OCR caption",
                            data=ocr_docx,
                            file_name=out_name_ocr,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True,
                            key="word_ocr_download",
                        )

        # 5) Chi tiết & review — gom tất cả phần phụ vào 1 expander gập
        issues = quality_check(blocks, translations)
        summary_line = st.session_state.get("word_summary", "")
        details_label = "ℹ️  Chi tiết kết quả & sửa thủ công"
        if issues:
            details_label += f"  ·  ⚠️ {len(issues)} đoạn nghi vấn"
        with st.expander(details_label, expanded=False):
            if summary_line:
                st.success(summary_line)
            if val and val["valid"]:
                st.caption(f"✅ Output validated: {val['block_count']:,} paragraph, "
                           f"{val['image_count']} ảnh")
                for warn in val.get("warnings", []):
                    st.warning(warn)

            # Quality check (sub-section)
            if issues:
                st.markdown(f"##### ⚠️ Quality check — {len(issues)} đoạn nghi vấn")
                for item in issues[:30]:
                    st.markdown(f"**🔴 {', '.join(item['issues'])}**")
                    col1, col2 = st.columns(2)
                    col1.text_area("Gốc", item["text"], height=60, disabled=True,
                                   key=f"qc_orig_{item['id']}")
                    col2.text_area("Dịch", item["translation"], height=60, disabled=True,
                                   key=f"qc_tr_{item['id']}")
                st.divider()

            # Review & lưu vào TM (sub-section)
            st.markdown("##### ✏️ Review & lưu vào TM")
            review_version = st.session_state.get("word_translations_version", 0)
            stored_rv = st.session_state.get("word_review_df_version")
            if stored_rv != review_version or "word_review_df" not in st.session_state:
                skip_roles_v = st.session_state.get("word_skip_roles") or set(NO_TRANSLATE_ROLES)
                translatable_blocks = [b for b in blocks if b["role"] not in skip_roles_v]
                st.session_state["word_review_df"] = pd.DataFrame([
                    {
                        "id":          b["id"],
                        "role":        b["role"],
                        "original":    b["text"],
                        "translation": translations.get(b["id"], ""),
                    }
                    for b in translatable_blocks
                ])
                st.session_state["word_review_df_version"] = review_version

            review_df = st.session_state["word_review_df"]
            edited_review = st.data_editor(
                review_df,
                column_config={
                    "id":          st.column_config.TextColumn("ID", disabled=True, width="small"),
                    "role":        st.column_config.TextColumn("Vai trò", disabled=True, width="small"),
                    "original":    st.column_config.TextColumn("Gốc", disabled=True, width="large"),
                    "translation": st.column_config.TextColumn("Bản dịch", width="large"),
                },
                use_container_width=True,
                hide_index=True,
                num_rows="fixed",
                height=400,
                key="word_review_editor",
            )

            st.markdown("**🔄 Dịch lại đoạn cụ thể** (tuỳ chọn — dùng prompt bổ sung):")
            skip_roles_v2 = st.session_state.get("word_skip_roles") or set(NO_TRANSLATE_ROLES)
            all_block_ids = [
                b["id"] for b in st.session_state["word_blocks"]
                if b["role"] not in skip_roles_v2
                and st.session_state["word_translations"].get(b["id"])
            ]
            blocks_by_id_lookup = {b["id"]: b for b in st.session_state["word_blocks"]}
            selected_ids = st.multiselect(
                "Chọn đoạn cần dịch lại:",
                options=all_block_ids,
                format_func=lambda bid: (
                    f"{bid} — "
                    + (blocks_by_id_lookup.get(bid, {}).get("text", "")[:60])
                ),
                key="word_regen_select",
            )
            extra_inst = st.text_input(
                "Hướng dẫn bổ sung (vd: 'dịch trang trọng hơn'):",
                key="word_regen_instruction",
            )
            if st.button("🔄 Dịch lại các đoạn đã chọn", key="word_regen_btn",
                         disabled=not selected_ids):
                from word_backend import translate_chunk
                chunk = [blocks_by_id_lookup[bid] for bid in selected_ids]
                rules = {"_extra": extra_inst} if extra_inst.strip() else None
                regen_tgt = st.session_state.get("word_target_lang") \
                            or LANG_EN[st.session_state["word_lang"]]
                regen_src = st.session_state.get("word_source_lang") \
                            or _resolve_langs()[1]
                with st.spinner(f"Dịch lại {len(chunk)} đoạn..."):
                    try:
                        new_tr, _, _ = translate_chunk(
                            get_client(), chunk,
                            regen_tgt,
                            st.session_state["word_doc_context"],
                            glossary=st.session_state.get("word_glossary"),
                            custom_rules=rules,
                            source_lang=regen_src,
                        )
                        st.session_state["word_translations"].update(new_tr)
                        st.session_state["word_translations_version"] = (
                            st.session_state.get("word_translations_version", 0) + 1
                        )
                        st.session_state.pop("word_review_df", None)
                        st.success(f"Đã dịch lại {len(new_tr)} đoạn")
                        st.rerun()
                    except Exception as e:
                        st.error(str(e))

            if st.button("✅ Xác nhận & lưu vào TM", key="word_review_confirm_btn",
                         use_container_width=True):
                tm = _ensure_tm()
                target_lang = st.session_state.get("word_target_lang") \
                              or LANG_EN[st.session_state["word_lang"]]
                changed = 0
                for _, row in edited_review.iterrows():
                    orig_tr = translations.get(row["id"], "")
                    new_tr  = row["translation"]
                    if new_tr and new_tr != orig_tr:
                        translations[row["id"]] = new_tr
                        # Also save to TM với role-specific key
                        tm[tm_key(row["original"], target_lang,
                                  row.get("role", ""))] = new_tr
                        changed += 1
                if changed:
                    st.session_state["word_translations"] = translations
                    st.session_state["word_translations_version"] = (
                        st.session_state.get("word_translations_version", 0) + 1
                    )
                    st.session_state.pop("word_review_df", None)
                    st.success(f"✅ Đã cập nhật {changed} bản dịch và lưu vào TM")
                else:
                    st.info("Không có thay đổi nào.")

            # Inline editor (sub-section cuối)
            st.divider()
            st.markdown("##### 📋 Xem & sửa bản dịch inline")
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

    elif not uploaded_files:
        st.info("👆 Vui lòng upload file Word (.docx) để bắt đầu")
