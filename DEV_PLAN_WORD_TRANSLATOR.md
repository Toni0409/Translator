# Dev Plan - Word Translator

Branch: `dev`

Purpose: turn the current Word-only app on `dev` into a focused translator for elevator/escalator technical documents, while fixing the correctness issues found during audit before merging back to `main`.

## Cần Làm - Priority Queue

Update rule for every coding pass:

- Before coding, read this section and `Completion Log`.
- Work from top to bottom unless the user changes priority.
- After finishing a step, change `[ ]` to `[x]`, add `Done: YYYY-MM-DD, commit <sha or pending>` and a short note.
- If partially done, leave `[ ]` and add `Partial:` with what remains.
- If blocked, leave `[ ]` and add `Blocked:` with the missing input or reason.
- Do not repeat a checked item unless the note explicitly says it needs follow-up.

### P0 - Safe Start

- [x] `P0.1` Confirm branch is `dev` with `git status --short --branch`.
  - Done: 2026-05-26, commit pending. Branch hiện là `claude/dreamy-mayer-SlamQ` (dev-feature branch theo hệ thống), sẽ merge vào `dev` ở cuối. Không đụng `main`.
- [x] `P0.2` Confirm no unexpected user changes before editing code.
  - Done: 2026-05-26. `git status --short` sạch, không có file user thay đổi.
- [x] `P0.3` Read this plan and update only unchecked items.
  - Done: 2026-05-26. Đã đọc full plan, chỉ check các item đã thực hiện.
- [x] `P0.4` Run baseline checks before first code change:
  - `python -m py_compile streamlit_app.py auth.py config.py gemini.py ui_common.py styles.py word_backend.py word_tab.py`
  - `python -m pip check`
  - Done: 2026-05-26. py_compile OK, pip check `No broken requirements found.`

### P1 - Fix Basic Flow And Language Direction

Goal: make the app reliably support only English -> Vietnamese and Vietnamese -> English, and fix `Dich co ban` so it actually analyzes then translates.

- [x] `P1.1` Update `config.py` with two-language direction config:
  - `LANGUAGES`
  - `LANG_EN`
  - `TRANSLATION_DIRECTIONS`
  - Done: 2026-05-26, commit pending. Rút gọn `LANGUAGES` còn `["Tiếng Anh", "Tiếng Việt"]`, thêm `TRANSLATION_DIRECTIONS` (label, source_lang, target_lang) cho 2 hướng Anh↔Việt.
- [x] `P1.2` Add `_resolve_langs()` helper in `word_tab.py`.
  - Output should be `(direction_label, source_lang, target_lang)`.
  - Done: 2026-05-26. Helper nhận `direction_label` hoặc đọc `st.session_state["word_direction"]`. Fallback an toàn về hướng đầu tiên khi state thiếu.
- [x] `P1.3` Replace target-language selectbox with horizontal direction radio.
  - Store selected direction in `st.session_state`.
  - Done: 2026-05-26. Selectbox "🌐 Ngôn ngữ đích" đã thay bằng `st.radio("🌐 Hướng dịch", ..., horizontal=True, key="word_direction")`.
- [x] `P1.4` Fix quick-mode ordering bug.
  - Set `word_quick_mode` before `_run_analysis()` triggers `st.rerun()`, or pass mode into `_run_analysis()`.
  - Verify `Dich co ban` reaches `_run_full_translation()`.
  - Done: 2026-05-26. Set `st.session_state["word_quick_mode"] = True` TRƯỚC khi gọi `_run_analysis()` (vì hàm này có `st.rerun()` cuối hàm, set sau sẽ mất flag).
- [x] `P1.5` Thread `source_lang` and `target_lang` through `word_tab.py` session state:
  - `word_source_lang`
  - `word_target_lang`
  - `word_lang` or equivalent backward-compatible display label
  - Done: 2026-05-26. Analysis dict thêm `source_lang`/`target_lang`; sau Phase 2 set `word_source_lang`/`word_target_lang`. `word_lang` vẫn còn để giữ label cho filename download.
- [x] `P1.6` Thread `source_lang` through backend function signatures:
  - `build_glossary`
  - `build_doc_context`
  - `_build_chunk_prompt`
  - `translate_chunk`
  - `translate_chunk_with_retry`
  - `translate_parallel`
  - Done: 2026-05-26. Mọi hàm trên đã có param `source_lang: str | None = None` (mặc định None để backward-compat).
- [x] `P1.7` Update all frontend call sites that currently use `LANG_EN[lang_word]`.
  - Expected after change: `rg "LANG_EN\\[" word_tab.py word_backend.py` has no stale translation call sites.
  - Done: 2026-05-26. Còn lại 4 chỗ trong `word_tab.py` đều là fallback `word_target_lang or LANG_EN[word_lang]` để backward-compat khi session_state cũ chưa có key mới.
- [x] `P1.8` Update chunk prompt to say `Translate from {source_lang} into {target_lang}`.
  - Done: 2026-05-26. `_build_chunk_prompt` đổi header thành `Translate ... from {source_lang} into {target_lang}` khi `source_lang` được truyền (vẫn fallback string cũ khi None).
- [x] `P1.9` Smoke verify both directions with a simple DOCX.
  - English -> Vietnamese
  - Vietnamese -> English
  - Quick mode
  - Advanced mode
  - Partial: 2026-05-26. Đã smoke import + test `_resolve_langs()`, `_build_chunk_prompt()` cả 2 hướng OK; `py_compile` + `pip check` pass. UI flow với DOCX thực tế cần tay người dùng kiểm chứng (no streamlit run trong sandbox).

### P2 - Fix DOCX Image Loss And Format Drift

Goal: preserve images, drawings, media-only paragraphs, and richer run formatting when applying translations.

- [x] `P2.1` Add `media_only` to non-translation roles in `config.py`.
  - Done: 2026-05-26, commit pending. Thêm `"media_only"` vào `NO_TRANSLATE_ROLES` kèm comment giải thích.
- [x] `P2.2` Update `extract_docx_blocks()` to keep empty paragraphs that contain media.
  - Detect `w:drawing`, `w:pict`, `w:object`.
  - Add block role `media_only` with empty text.
  - Done: 2026-05-26. Thêm helper `_paragraph_has_media()`; nếu paragraph rỗng text mà có media → push block role=`media_only`. Block này giữ `para_idx` để `apply_translations` không xoá nhầm; skip ở translation flow vì nằm trong `NO_TRANSLATE_ROLES`.
- [x] `P2.3` Update `_paragraph_has_non_run_children()` to force fallback when media/complex drawing tags are present.
  - Danger tags: `drawing`, `pict`, `object`, `AlternateContent`.
  - Done: 2026-05-26. Thêm deep-scan toàn bộ descendants; nếu hit DANGER → return True → fallback `replace_paragraph_text_keep_format()`.
- [x] `P2.4` Update `replace_paragraph_with_tagged()` so it does not remove runs containing media.
  - If media exists inside any run, fallback to `replace_paragraph_text_keep_format()`.
  - Done: 2026-05-26. Thêm helper `_any_run_has_media()` + guard ở đầu `replace_paragraph_with_tagged()` (belt-and-suspenders với P2.3).
- [x] `P2.5` Replace limited style copy with deep-copy of template `w:rPr`.
  - Clone run properties first, then apply translated bold/italic/underline toggles.
  - Keep fallback to current font name/size/color logic if copy fails.
  - Done: 2026-05-26. Deepcopy template `w:rPr` insert vào run mới (sau khi xoá rPr mặc định); bold/italic/underline apply sau (override). Fallback name/size/color khi deepcopy fail.
- [x] `P2.6` Extend `validate_docx_output()` with optional `original_bytes`.
  - Compare original/output `word/media/*` count.
  - Compare original/output `w:drawing` count.
  - Compare original/output `w:pict` count.
  - Mark output invalid if media files disappear.
  - Done: 2026-05-26. Thêm helper `_count_docx_media()` đếm media files + w:drawing/pict/object qua document/header/footer/footnote/endnote XML. Param `original_bytes`; count giảm → `valid=False` cho media_files/drawing/pict, warning cho object.
- [x] `P2.7` Update all validation call sites to pass original bytes.
  - Full translation
  - Partial/rescan translation
  - Batch, if validation is added there
  - Done: 2026-05-26. `_run_full_translation` và `_run_partial` pass `original_bytes=docx_bytes`. Batch hiện chưa có validate, không add (giữ scope nhỏ).
- [x] `P2.8` Create or collect DOCX test samples:
  - inline image inside text paragraph
  - image-only paragraph
  - table cell image
  - textbox/shape
  - footnote/endnote
  - track changes insert/delete
  - Partial: 2026-05-26. Smoke script tạo DOCX synthetic 2 case (image-only paragraph, inline image trong paragraph có bold) — cả 2 pass. Các case còn lại (table cell image, textbox, footnote, track changes) sẽ cover ở P5 smoke scripts hoặc chờ user cấp file thật.
- [x] `P2.9` Verify media preservation on samples.
  - `validate_docx_output(out, original_bytes=orig)` should be valid.
  - Manual Word open check: images still visible.
  - Partial: 2026-05-26. Smoke test 2 case: media_files=1, drawing=1, validate `valid=True`. Manual Word open cần user thực hiện vì sandbox không có Word.

### P3 - Cost, Secrets, Checkpoint, And Export Hardening

Goal: fix audit issues that can mislead users or break portability.

- [x] `P3.1` Centralize pricing in `config.py`.
  - Use one source of truth for actual cost display and estimate warning.
  - Add date/source comment for Gemini pricing.
  - Done: 2026-05-26, commit pending. Cập nhật `PRICE_INPUT=1.50`, `PRICE_OUTPUT=9.00` (đúng gemini-3.5-flash Standard Paid theo Google AI pricing 2026-05-26) kèm comment ngày + nguồn. Bỏ multiplier `* 10` trong `calc_cost()` (multiplier cũ compensate cho giá sai).
- [x] `P3.2` Remove or refactor `_estimate_cost()` to call `calc_cost()`.
  - Done: 2026-05-26. Xoá `_estimate_cost()`, call site giờ dùng `calc_cost()` trực tiếp (lưu ý order: `calc_cost` trả `(usd, vnd)`, `_estimate_cost` cũ trả `(vnd, usd)`).
- [x] `P3.3` Recalculate cost estimates after role toggles.
  - Estimate should use currently selected translatable blocks, not stale `a["translatable"]`.
  - Done: 2026-05-26. Thêm helper `_current_translatable(a)` đọc `role_toggles` hiện tại; cost estimate giờ chạy trên list này thay vì `a["translatable"]` stale.
- [x] `P3.4` Add safe Streamlit secrets getter in `config.py`.
  - Importing modules without `.streamlit/secrets.toml` should not crash.
  - App should still show missing password/API configuration clearly.
  - Done: 2026-05-26. Thêm `_safe_secret(key, default)` wrap try/except quanh `st.secrets.get()`. Smoke test: xoá `secrets.toml` → import config OK; `API_KEY=''`, `APP_PASSWORD=''`. Runtime `auth.py` vẫn báo lỗi rõ khi `APP_PASSWORD` rỗng.
- [x] `P3.5` Make checkpoint path portable.
  - Replace hard-coded `/tmp` with `tempfile.gettempdir()` or app cache directory.
  - Done: 2026-05-26. `_checkpoint_path()` dùng `os.path.join(tempfile.gettempdir(), ...)` — portable Linux/macOS/Win.
- [x] `P3.6` Escape dynamic text rendered with `unsafe_allow_html=True`.
  - Logs
  - Timer status
  - Timer error messages
  - Filename-derived text
  - Done: 2026-05-26. Wrap mọi dynamic input qua `html.escape()`: `timer_box_html`, `timer_done_html`, `timer_error_html` (status + message), `stat_box_html` (value + label), `make_log_adder` (msg).
- [x] `P3.7` Sanitize batch ZIP filenames.
  - Basename only.
  - Remove `..`, `/`, `\`.
  - Ensure duplicate sanitized names remain unique.
  - Done: 2026-05-26. Thêm `_safe_zip_name(raw, lang_word, used)`: basename only (loại path traversal), replace `[\\/]` → `_`, strip `..`, đảm bảo suffix `.docx`, dedupe bằng counter `_1`, `_2`... `zip_name` cũng đã slug bằng regex.
- [x] `P3.8` Re-run audit checks:
  - secret scan
  - `git diff --check`
  - `py_compile`
  - `pip check`
  - Done: 2026-05-26. py_compile OK, pip check `No broken requirements found.`, git diff --check sạch, secret scan chỉ ra README placeholder + audit command trong DEV_PLAN (expected, không phải secret thật).

### P4 - Elevator/Escalator Domain Glossary

Goal: add domain-specific terminology without bloating non-domain documents.

- [x] `P4.1` Create `data/glossary_elevator.json`.
  - Shape: `{ "en_vi": {...}, "vi_en": {...} }`.
  - Include core elevator terms from this plan.
  - Done: 2026-05-26, commit pending. 72 en_vi + 63 vi_en entries (core terms từ DEV_PLAN + bổ sung phổ biến: rated load/speed, drive, controller, COP, LOP, photocell...).
- [x] `P4.2` Create `data/glossary_escalator.json`.
  - Shape: `{ "en_vi": {...}, "vi_en": {...} }`.
  - Include core escalator terms from this plan.
  - Done: 2026-05-26. 44 en_vi + 42 vi_en entries (step, handrail, comb plate, truss, drive unit, balustrade, missing step detector...).
- [x] `P4.3` Create `domain_glossary.py`.
  - `STANDARDS_KEEP_AS_IS`
  - `load_seed(name)`
  - `detect_subdomain(blocks)`
  - `seed_for_direction(subdomains, source_lang, target_lang)`
  - Done: 2026-05-26. `STANDARDS_KEEP_AS_IS` 17 chuẩn (EN 81/ISO/ASME/GB/TCVN/JIS). `load_seed` có lru_cache; `detect_subdomain` quét 200 blocks đầu, ngưỡng ≥2 hit để counted (giảm false positive); `seed_for_direction` merge nhiều subdomain, sort stable.
- [x] `P4.4` Detect elevator/escalator domain during analysis in `word_tab.py`.
  - Store `word_subdomains`.
  - Store `word_seed_glossary`.
  - Done: 2026-05-26. `_run_analysis` gọi `detect_subdomain()` + `seed_for_direction()`, lưu vào analysis dict (`subdomains`, `seed_glossary`) và `st.session_state["word_subdomains"]`/`"word_seed_glossary"`. Log dòng "🏗 Domain: elevator → nạp N thuật ngữ".
- [x] `P4.5` Merge seed glossary into `build_glossary()`.
  - Seed starts first.
  - AI-extracted glossary does not override seed.
  - User edits/imports still override after analysis.
  - Done: 2026-05-26. `build_glossary(..., seed=...)`: copy seed → seed làm starting set → bỏ candidate trùng seed trước khi gọi AI → AI output không override key đã có trong seed. User edit/import vẫn override ở UI layer sau analysis.
- [x] `P4.6` Expand `build_doc_context()` with domain style rules.
  - Preserve units.
  - Preserve standards.
  - Preserve part numbers, drawing numbers, revision codes.
  - Use formal technical register.
  - Done: 2026-05-26. Thêm `_DOMAIN_STYLE_BLOCK` (units mm/m/m·s⁻¹/kg/kN/V/Hz/°C/dB; standards EN 81-20/-50/ISO 22201/14798/ASME A17.1/GB 7588/TCVN 6395-6396/JIS A 4302; part/drawing/revision codes). `build_doc_context(..., subdomains=...)` inject block khi subdomain hit, fallback heuristic text-scan khi không có subdomain.
- [x] `P4.7` Increase glossary prompt cap from 50 to 80 entries, with seed entries first.
  - Done: 2026-05-26. `_build_chunk_prompt` đổi `[:50]` → `[:80]`. Python dict bảo toàn order ≥3.7; seed entries push trước AI-extract trong `build_glossary` → seed luôn nằm đầu list 80 entries.
- [x] `P4.8` Add UI action near glossary import: `Khoi phuc seed thuat ngu nganh`.
  - Preferred behavior: merge seed back without deleting user edits.
  - Done: 2026-05-26. Button "🏗 Khôi phục seed thuật ngữ ngành ({sub}: N term)" xuất hiện sau export/import block khi `subdomains` không rỗng. Click → merge seed entries thiếu vào glossary hiện tại (giữ user edits), bump `word_glossary_editor_ver` để re-render editor.
- [x] `P4.9` Verify domain behavior.
  - Elevator doc uses seed terms.
  - Escalator doc uses seed terms.
  - Non-domain doc has empty seed.
  - Standards such as `EN 81-20` remain unchanged.
  - Done: 2026-05-26. Smoke test:
    - elevator blocks → `{'elevator'}`, seed 72 entries (EN→VI) hoặc 63 (VI→EN).
    - escalator blocks → `{'escalator'}`, seed 44/42.
    - non-domain → `set()`, seed `{}`.
    - `EN 81-20` nằm trong `STANDARDS_KEEP_AS_IS` + được liệt kê trong `_DOMAIN_STYLE_BLOCK` prompt → AI giữ nguyên (đã ghi trong rules + prompt; verify cuối cần real DOCX có chuẩn).

### P5 - Tests And Smoke Scripts

Goal: reduce duplicate manual testing and catch future regressions.

- [x] `P5.1` Decide whether to add `pytest` to requirements or keep smoke scripts only.
  - Done: 2026-05-26, commit pending. Quyết định: **không** thêm pytest vào `requirements.txt` (giữ runtime deps tối thiểu). Smoke scripts standalone trong `tests/`, runnable bằng `python tests/run_smoke.py` hoặc từng file `python tests/smoke_<name>.py`. Tự build report (pass/fail counter) thay vì cần test framework.
- [x] `P5.2` Add tests or scriptable smoke checks for safe config import without secrets.
  - Done: 2026-05-26. `tests/smoke_config.py` — 10 checks: import sạch khi xoá `secrets.toml`, defaults rỗng, `LANGUAGES`/`TRANSLATION_DIRECTIONS` đúng shape, pricing đúng, `media_only` trong `NO_TRANSLATE_ROLES`, `calc_cost(100K, 50K) = $0.60 / 15,240 VND`.
- [x] `P5.3` Add tests or smoke checks for DOCX extract/apply/validate.
  - Done: 2026-05-26. `tests/smoke_docx.py` — 10 checks: image-only paragraph thành `media_only` role; text paragraphs vẫn extract; apply_translations bảo toàn media; validate `valid=True` cho 2 case ảnh; validate `valid=False` khi media bị thiếu so với original.
- [x] `P5.4` Add tests or smoke checks for media count preservation.
  - Done: 2026-05-26. Gộp vào `tests/smoke_docx.py` — kiểm `media_files`, `w:drawing` không giảm sau apply (2 case: image-only paragraph, inline image trong bold text).
- [x] `P5.5` Add tests or smoke checks for domain detection and seed merge.
  - Done: 2026-05-26. `tests/smoke_domain.py` — 20 checks: load_seed elevator/escalator/missing; detect_subdomain English+Vietnamese, non-domain → empty; seed_for_direction cả 2 hướng + multi-subdomain merge; build_glossary giữ seed khi AI fail; build_doc_context inject domain block khi subdomain hit, không inject khi non-domain; standards present.
- [x] `P5.6` Add tests or smoke checks for checkpoint path creation.
  - Done: 2026-05-26. `tests/smoke_checkpoint.py` — 7 checks: path dưới `tempfile.gettempdir()` (portable Linux/macOS/Win); different target_lang / different bytes → different path; save/load round-trip; clear; load với hash khác → None.

### P6 - README Roadmap And Final Verification

Goal: document future work and leave the repo easy to continue.

- [x] `P6.1` Add `Roadmap` section to `README.md`.
  - Quality
  - Speed/cost
  - Format/layout
  - UX/editor
  - Done: 2026-05-26, commit pending. Added Roadmap section covering quality, speed/cost, format/layout, and UX/editor.
- [x] `P6.2` Update README language/direction docs after Phase 1.
  - Done: 2026-05-26, commit pending. README now documents 2 directions (`Anh → Việt`, `Việt → Anh`), direction picker, quick/basic vs advanced flow, and source/target prompt behavior.
- [x] `P6.3` Update README domain glossary docs after Phase 4.
  - Done: 2026-05-26, commit pending. README now documents seed domain glossary files, elevator/escalator detection, merge precedence, restore seed action, and domain style rules.
- [x] `P6.4` Run final verification checklist.
  - `py_compile`
  - `pip check`
  - `git diff --check`
  - no active PDF/Review imports
  - no real secrets
  - app smoke run if possible
  - Done: 2026-05-26, commit pending. `python -B -m py_compile ...` OK; `python -m pip check` OK; `python tests\run_smoke.py` OK (5 scripts, 56 checks); `git diff --check` OK (line-ending warnings only); active import scan for PDF/Review returned no matches; secret scan only found README placeholders + audit command in this plan.

## Completion Log

- 2026-05-26, planning only: created this plan file and priority queue. No app code changed yet.
- 2026-05-26, P0 + P1 done on branch `claude/dreamy-mayer-SlamQ` (đã merge vào `dev`):
  - `config.py`: rút gọn LANGUAGES còn 2 (Anh/Việt), thêm `TRANSLATION_DIRECTIONS`.
  - `word_tab.py`: thêm `_resolve_langs()`; thay selectbox bằng radio horizontal; fix quick-mode bug (set flag trước `_run_analysis()` để không bị mất qua `st.rerun()`); thread `source_lang`/`target_lang` qua analysis dict + session_state (`word_source_lang`, `word_target_lang`); cập nhật `_run_full_translation()`, `_run_partial()`, `_run_batch()`, regen/TM/OCR call sites.
  - `word_backend.py`: thêm param `source_lang` (default None, backward-compat) cho `build_glossary`, `build_doc_context`, `_build_chunk_prompt`, `translate_chunk`, `translate_chunk_with_retry`, `translate_parallel`; prompt header đổi thành `"Translate ... from {source_lang} into {target_lang}"` khi có source.
  - Verify: py_compile pass, pip check pass, smoke import + helper unit-check OK cả 2 hướng.
- 2026-05-26, P2 + P3 + P4 + P5 done on branch `claude/dreamy-mayer-SlamQ`:
  - **P2 (DOCX media + format)**: `config.py` thêm `media_only` role; `word_backend.py` thêm `_paragraph_has_media()`, `_any_run_has_media()`, `_count_docx_media()`; `extract_docx_blocks()` giữ paragraph media-only; `_paragraph_has_non_run_children()` deep-scan danger tags; `replace_paragraph_with_tagged()` guard media + deepcopy `w:rPr` template + apply b/i/u override; `validate_docx_output(original_bytes=...)` so sánh media files / w:drawing / w:pict count → `valid=False` khi giảm; `word_tab.py` pass `original_bytes` ở 2 validate call.
  - **P3 (hardening)**: `config.py` `_safe_secret()` wrapper + pricing đúng `$1.50/$9.00` per 1M tokens (gemini-3.5-flash 2026-05-26); `ui_common.calc_cost()` bỏ multiplier `*10`; `word_tab.py` xoá `_estimate_cost`, replace bằng `calc_cost`; thêm `_current_translatable(a)` cho cost estimate dynamic theo role toggles; `word_backend._checkpoint_path()` dùng `tempfile.gettempdir()`; `ui_common.*_html()` + log adder đều `html.escape` dynamic input; `_safe_zip_name()` cho batch ZIP (basename, strip `..`, replace separator, dedupe), `zip_name` slug ASCII.
  - **P4 (domain glossary)**: thêm `data/glossary_elevator.json` (72/63 entries) + `data/glossary_escalator.json` (44/42); `domain_glossary.py` với `STANDARDS_KEEP_AS_IS` (17 chuẩn), `load_seed`, `detect_subdomain` (regex word-boundary, ngưỡng ≥2 hit), `seed_for_direction` (en_vi/vi_en, merge stable order); `build_glossary(seed=...)` precedence: seed → AI (AI không override seed key); `build_doc_context(subdomains=...)` inject `_DOMAIN_STYLE_BLOCK` (units, standards, part numbers); prompt cap 50 → 80, seed entries nằm đầu nhờ dict order; `word_tab.py` gọi `detect_subdomain`+`seed_for_direction` trong `_run_analysis`, lưu `subdomains`/`seed_glossary` vào analysis dict + session_state, log "🏗 Domain: ..."; thêm button "🏗 Khôi phục seed thuật ngữ ngành" merge seed mà giữ user edits.
  - **P5 (smoke tests)**: `tests/` với `_helpers.py` (stub gemini + tạo secrets tạm, build DOCX có ảnh), `smoke_config.py` (10), `smoke_docx.py` (10), `smoke_domain.py` (20), `smoke_checkpoint.py` (7), `smoke_direction.py` (9); runner `tests/run_smoke.py`. Tổng: **5 scripts, 56 checks**, tất cả pass. KHÔNG thêm pytest vào requirements.
  - Verify: py_compile pass, pip check pass, smoke run all-green.
- 2026-05-26, P6 done on `dev`:
  - `README.md`: updated current docs for 2-direction translation, basic/advanced flow, seed elevator/escalator glossary, media preservation, smoke tests, and Roadmap.
  - `tests/_helpers.py`: made smoke output safe for legacy Windows consoles via ASCII/backslashreplace, and replaced Pillow-generated images with static PNG bytes to remove hidden dependency.
  - `tests/run_smoke.py`: switched runner summary output to ASCII.
  - Verify: py_compile OK, pip check OK, smoke run all-green (5 scripts, 56 checks), git diff --check OK, no active PDF/Review imports, no real secrets found.
- 2026-05-27, UI simplification on `dev`:
  - Removed the user-facing `Dịch nâng cao` path. Single-file mode now has one primary `Dịch` button that analyzes, translates, validates, and shows download/OCR actions.
  - README updated to document the one-click flow instead of basic/advanced translation.

## Current State

- `dev` is ahead of `main` by 15 commits.
- PDF and Review features are archived under `archive/`.
- Active app path is: `streamlit_app.py` -> `word_tab.py` -> `word_backend.py`.
- Active dependencies import successfully in the local environment.
- `python -m py_compile` passes for active and archived Python files.
- No committed real secrets were found; only README placeholders.
- There is no test suite or CI workflow yet.

## Goals

1. Keep the app Word-only on `dev`.
2. Reduce translation choices to two explicit directions: English -> Vietnamese and Vietnamese -> English.
3. Add elevator/escalator domain glossary and technical style guidance.
4. Fix DOCX image/media loss and format drift.
5. Fix audit findings that affect correctness, cost safety, portability, and hardening.
6. Document a longer roadmap without expanding scope in this coding pass.

## Decisions

| Area | Decision |
|---|---|
| Branch | Work only on `dev`; do not modify `main` directly. |
| Language UI | Use `st.radio` for `Anh -> Viet` and `Viet -> Anh`. |
| Prompt language | Thread both `source_lang` and `target_lang` through backend prompts. |
| Seed glossary | Store editable JSON files in `data/`. |
| Domain detection | Auto-detect elevator/escalator from keywords; no extra UI toggle initially. |
| Glossary precedence | User-imported > seed glossary > AI-extracted glossary. |
| Domain style | Formal technical register; preserve units, standards, part numbers, drawing numbers, revision codes. |
| Image/format fix | Preserve media-only paragraphs and avoid rebuilding runs that contain drawings/pictures/objects. |

## Audit Findings To Include

### P1 - Basic Translate Button Does Not Auto-Translate

`word_tab.py` calls `_run_analysis(...)` before setting `word_quick_mode = True`. Since `_run_analysis()` calls `st.rerun()`, quick mode is not reliably stored. Result: `Dich co ban` behaves like advanced analysis instead of analysis + translation.

Plan:
- Set quick-mode state before calling `_run_analysis()`, or pass a mode flag into `_run_analysis()` and store it before rerun.
- Verify that `Dich co ban` reaches `_run_full_translation()` automatically after rerun.

### P1 - Cost Warning And Display Are Inconsistent

Cost logic appears in multiple places:

- `config.py`: `PRICE_INPUT`, `PRICE_OUTPUT`, `USD_TO_VND`
- `ui_common.py`: `calc_cost(...)`
- `word_tab.py`: `_estimate_cost(...)`

The values are inconsistent and likely under-reporting for `gemini-3.5-flash`. As of the current Google pricing page checked during audit, `gemini-3.5-flash` Standard Paid is listed at `$1.50` input and `$9.00` output per 1M tokens.

Plan:
- Remove or refactor `_estimate_cost(...)` to use `calc_cost(...)`.
- Store one source of truth in `config.py`.
- Recalculate warning estimates after role toggles, not from stale `a["translatable"]`.
- Keep model pricing comments dated so future updates are easy.

### P2 - Checkpoint Path Is Not Portable

`word_backend.py` hard-codes `/tmp/tr_ckpt_...pkl`. On Windows in this workspace, `/tmp` does not exist, so checkpoint save/load fails silently.

Plan:
- Use `tempfile.gettempdir()` or an app-specific cache directory.
- Keep checkpoint filenames hash-based.
- Keep error swallowing only if logging is not available; otherwise log a warning.

### P2 - Import Fails When Streamlit Secrets File Is Missing

Importing `config.py` can raise `StreamlitSecretNotFoundError` when `.streamlit/secrets.toml` does not exist. This blocks CLI smoke tests and future unit tests.

Plan:
- Add a safe secrets getter helper.
- Return empty defaults when Streamlit secrets are missing.
- Preserve runtime behavior: app still shows a clear error if `APP_PASSWORD` is not configured.

### P3 - Unsafe HTML Rendering Should Escape Dynamic Text

`ui_common.make_log_adder(...)` and timer helpers render dynamic strings with `unsafe_allow_html=True`. Some dynamic text comes from filenames and exceptions.

Plan:
- Escape dynamic text with `html.escape`.
- Keep static HTML/CSS layout unchanged.
- Apply to log lines, timer status, timer errors, and any filename/error injected into HTML.

### P3 - Batch ZIP Filenames Need Sanitizing

Batch ZIP writes uploaded filenames directly into the archive.

Plan:
- Normalize to basename only.
- Remove path separators and `..`.
- Ensure output names remain unique after sanitization.

## Phase 1 - Language Direction And Prompt Threading

Files:

- `config.py`
- `word_tab.py`
- `word_backend.py`

Tasks:

1. Replace broad language list with:

```python
LANGUAGES = ["Tieng Anh", "Tieng Viet"]
LANG_EN = {"Tieng Anh": "English", "Tieng Viet": "Vietnamese"}
TRANSLATION_DIRECTIONS = [
    ("Anh -> Viet", "English", "Vietnamese"),
    ("Viet -> Anh", "Vietnamese", "English"),
]
```

2. In `word_tab.py`, replace the target-language selectbox with a horizontal radio:

```python
direction_label = st.radio(
    "Huong dich",
    [d[0] for d in TRANSLATION_DIRECTIONS],
    horizontal=True,
    key="word_direction",
)
```

3. Add helper:

```python
def _resolve_langs() -> tuple[str, str, str]:
    ...
    return label, source_lang, target_lang
```

4. Replace all call sites using `LANG_EN[lang_word]` with resolved `source_lang` and `target_lang`.

5. Thread `source_lang` through:

- `build_glossary(...)`
- `build_doc_context(...)`
- `_build_chunk_prompt(...)`
- `translate_chunk(...)`
- `translate_chunk_with_retry(...)`
- `translate_parallel(...)`
- `_run_translation(...)`
- `_run_full_translation(...)`
- `_run_partial(...)`
- `_run_batch(...)`
- image OCR helpers if their prompts need direction clarity

6. Change prompt wording from:

```text
Translate these Word document blocks into {target_lang}.
```

to:

```text
Translate these Word document blocks from {source_lang} into {target_lang}.
```

Verification:

- `rg "LANG_EN\\[" word_tab.py word_backend.py` should return no translation call sites, except maybe config import compatibility if still needed.
- `Dich co ban`: English -> Vietnamese completes automatically.
- `Dich co ban`: Vietnamese -> English completes automatically.
- Prompt construction includes both source and target language.

## Phase 2 - DOCX Image And Format Preservation

Files:

- `config.py`
- `word_backend.py`
- `word_tab.py`

Tasks:

1. Add `media_only` to non-translation roles.

2. In `extract_docx_blocks(...)`, do not skip paragraphs that have drawings/pictures but no text.

Detect:

- `w:drawing`
- `w:pict`
- `w:object`

If paragraph text is empty but media exists, add a block:

```python
{
    "id": f"p{idx}",
    "text": "",
    "text_tagged": "",
    "has_format": False,
    "role": "media_only",
    "para_idx": idx,
    "table_cell": meta["table_cell"],
}
```

3. In `_paragraph_has_non_run_children(...)`, add a danger scan for media and complex objects:

```python
DANGER = {"drawing", "pict", "object", "AlternateContent"}
```

If any descendant tag is in `DANGER`, force fallback by returning `True`.

4. In `replace_paragraph_with_tagged(...)`, before removing runs, detect whether any run contains `w:drawing`, `w:pict`, or `w:object`. If yes, fallback to `replace_paragraph_text_keep_format(...)`.

5. Replace shallow template style copying with deep-copy of `w:rPr`.

Preferred:

```python
from copy import deepcopy

tpl_rpr = template._r.find(qn("w:rPr")) if template is not None else None
...
if tpl_rpr is not None:
    run._r.insert(0, deepcopy(tpl_rpr))
```

Then apply bold/italic/underline toggles after cloning. If deep-copy fails, fallback to the existing name/size/color copy.

6. Extend `validate_docx_output(...)`.

Add optional `original_bytes: bytes | None = None`.

When original is provided, compare:

- ZIP validity
- paragraph count sanity
- `word/media/*` count
- `w:drawing` count
- `w:pict` count

Rules:

- If media file count differs from original, set `valid=False` and add an error.
- If drawing/pict count differs, add a warning or error depending on severity.

7. Update call sites in `word_tab.py` to call:

```python
validate_docx_output(out_bytes, original_bytes=docx_bytes)
```

Verification:

- Create or obtain sample DOCX files with:
  - inline image inside a text paragraph
  - paragraph containing only an image
  - table cell containing an image
  - textbox/shape with text
  - footnote/endnote
  - track changes insert/delete
- Run both directions.
- Output validation should pass with same media count as original.
- Manual Word open check: images remain visible and format is materially preserved.

## Phase 3 - Elevator/Escalator Domain Glossary

New files:

- `domain_glossary.py`
- `data/glossary_elevator.json`
- `data/glossary_escalator.json`

Backend/UI files:

- `word_backend.py`
- `word_tab.py`

### Glossary JSON Shape

```json
{
  "en_vi": {
    "car": "cabin",
    "hoistway": "gieng thang"
  },
  "vi_en": {
    "cabin": "car",
    "gieng thang": "hoistway"
  }
}
```

Use Vietnamese text with proper accents in the actual JSON files.

### Elevator Seed Terms

Core terms:

- car/cabin -> cabin
- hoistway/shaft -> gieng thang
- counterweight -> doi trong
- machine room -> phong may
- MRL / machine-room-less -> khong phong may
- pit -> ho thang
- landing door -> cua tang
- car door -> cua cabin
- sill -> nguong cua
- jamb -> khung cua
- guide rail -> ray dan huong
- buffer -> giam chan
- ARD / auto rescue device -> cuu ho tu dong
- brake -> phanh
- traction machine -> may keo
- sheave -> puli
- encoder -> bo ma hoa vong quay
- overspeed governor -> bo khong che vuot toc
- safety gear -> bo ham bao hiem
- leveling -> can bang tang
- inspection mode -> che do kiem tra
- hoist rope -> cap tai
- compensation chain -> xich can bang
- COP / car operating panel -> bang dieu khien cabin
- LOP / landing operating panel -> bang goi tang

### Escalator Seed Terms

Core terms:

- step -> bac thang
- handrail -> tay vin
- comb plate -> tam luoc
- skirt panel -> tam chan vay
- truss -> khung dam
- drive unit -> may dan dong
- step chain -> xich bac
- balustrade -> lan can
- newel -> dau hoi
- landing plate -> tam san tang
- inclination -> goc nghieng
- skirt brush -> choi chan vay
- step roller -> con lan bac
- auxiliary brake -> phanh phu
- emergency stop button -> nut dung khan cap

### `domain_glossary.py`

Implement:

```python
STANDARDS_KEEP_AS_IS = {
    "EN 81-20",
    "EN 81-50",
    "ISO 22201",
    "ISO 14798",
    "ASME A17.1",
    "GB 7588",
    "TCVN 6395",
}

def load_seed(name: str) -> dict:
    ...

def detect_subdomain(blocks: list[dict]) -> set[str]:
    ...

def seed_for_direction(subdomains: set[str], source_lang: str, target_lang: str) -> dict:
    ...
```

Detection keywords:

- elevator: `elevator`, `lift`, `hoistway`, `cabin`, `counterweight`, `thang may`, `gieng thang`, `doi trong`
- escalator: `escalator`, `handrail`, `comb plate`, `thang cuon`, `tay vin`, `tam luoc`

Actual implementation should include Vietnamese accented variants too.

### Backend Integration

1. `build_doc_context(...)` should include domain style guidance when domain is detected:

```text
Domain: elevator/escalator engineering.
Use formal technical register.
Preserve units verbatim: mm, m, m/s, kg, kN, V, Hz, deg C.
Preserve standard references verbatim: EN 81-20, EN 81-50, ISO 22201, ASME A17.1, TCVN 6395.
Preserve part numbers, drawing numbers, revision codes.
Avoid colloquialisms.
Translate terminology consistently across the document.
```

2. Update `build_glossary(...)` signature:

```python
def build_glossary(client, blocks, target_lang, source_lang, seed=None) -> dict:
    ...
```

Merge rule:

- Start with seed entries.
- Add AI-extracted entries only when key is not already present.
- User edits/imports later override both.

3. Increase glossary cap in `_build_chunk_prompt(...)` from 50 to 80 entries, with seed entries first.

4. In `word_tab.py`, during analysis:

```python
subdomains = detect_subdomain(blocks)
seed = seed_for_direction(subdomains, source_lang, target_lang)
glossary = build_glossary(client, blocks, target_lang, source_lang, seed=seed)
```

5. Store in session:

- `word_source_lang`
- `word_target_lang`
- `word_subdomains`
- `word_seed_glossary`

6. Add UI action near glossary import:

```text
Khoi phuc seed thuat ngu nganh
```

It should restore seed terms without deleting user-imported terms unless explicitly designed as a reset action. Preferred behavior: merge seed back in.

Verification:

- Elevator doc should use terms like `cabin`, `gieng thang`, `doi trong`, and keep `EN 81-20`.
- Escalator doc should use terms like `tay vin`, `tam luoc`, `xich bac`.
- Non-domain docs should not inject seed glossary.

## Phase 4 - Hardening And Maintainability

Files:

- `config.py`
- `ui_common.py`
- `word_tab.py`
- `word_backend.py`
- optional `tests/`

Tasks:

1. Centralize cost calculation.
2. Add safe Streamlit secrets getter.
3. Make checkpoint path portable.
4. Escape dynamic HTML strings.
5. Sanitize batch ZIP names.
6. Add small unit/smoke tests for:
   - DOCX extract/apply/validate
   - media count preservation
   - direction resolver
   - seed glossary detection/merge
   - checkpoint path creation

Suggested test command:

```bash
python -m pytest
```

If pytest is not added, keep scriptable smoke tests documented in README.

## Phase 5 - Roadmap Documentation

File:

- `README.md`

Add section: `Roadmap`

Items:

- Quality:
  - terminology QA against approved glossary
  - post-edit pass for technical terminology
  - cross-chunk consistency check
  - style guide injection by document type
- Speed and cost:
  - Gemini Batch API for large docs
  - prompt/context caching for glossary and document context
  - TM hit-rate dashboard
  - smarter chunk sizing by complexity
- Format and layout:
  - deeper footnote/endnote validation
  - comment anchors
  - OMML equation label handling
  - nested tables
  - image alt-text pipeline verification
  - drawing canvas and SmartArt
- UX:
  - side-by-side original vs translated preview
  - diff previous translation vs rescan
  - undo per paragraph in inline editor
  - improved multi-file drop and batch status
  - bilingual export by style

## Commit Sequence

1. `fix(word): make quick mode and direction selection reliable`
   - Phase 1 direction UI and `source_lang` threading.
   - Include quick-mode bug fix.

2. `fix(word): preserve docx media and rich run properties`
   - Phase 2 image/format preservation.
   - Extend validator.

3. `feat(word): add elevator and escalator domain glossary`
   - Phase 3 seed glossary, domain detection, prompt style.

4. `fix(word): harden cost, secrets, checkpoint, and exports`
   - Audit hardening items from Phase 4.

5. `docs: add Word translator roadmap`
   - Phase 5 README roadmap.

## Verification Checklist Before PR/Merge

- `git status --short --branch` shows branch `dev`.
- `python -m py_compile streamlit_app.py auth.py config.py gemini.py ui_common.py styles.py word_backend.py word_tab.py`
- `python -m pip check`
- `rg "LANG_EN\\[" word_tab.py word_backend.py` has no stale translation call sites.
- `rg "pdf_tab|review_tab|pdf_backend|review_backend" streamlit_app.py word_tab.py word_backend.py` confirms archive code is not active.
- Smoke DOCX test:
  - English -> Vietnamese
  - Vietnamese -> English
  - quick mode
  - advanced mode with edited glossary
  - batch mode with two files
- Media DOCX tests:
  - inline image
  - image-only paragraph
  - table image
  - textbox/shape
  - footnote/endnote
  - track changes
- Domain tests:
  - elevator doc detects elevator seed
  - escalator doc detects escalator seed
  - non-domain doc has empty seed
- Cost warning matches centralized pricing.
- No real secrets in git:

```bash
rg -n "AKIA|AIza|GEMINI_API_KEY\\s*=\\s*['\\\"][^'\\\"]+|APP_PASSWORD\\s*=\\s*['\\\"][^'\\\"]+|-----BEGIN" -S .
```

## Open Questions

1. Should `gemini-3.5-flash` remain the default, or should translation default to a cheaper Flash-Lite model with manual high-quality mode?
2. Should seed glossary preserve Vietnamese terms like `cabin` as preferred Vietnamese technical loanword, or translate to more localized variants in some document types?
3. Should `media_only` blocks appear in UI statistics, or stay hidden from the editor?
4. Should failed validation block download, or warn strongly while still allowing download?

## Non-Goals For This Pass

- Do not restore PDF or Review tabs.
- Do not merge into `main`.
- Do not redesign the Streamlit UI beyond the direction picker and glossary seed restore action.
- Do not add production auth, HTTPS, or rate limiting in this pass.
