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

- [ ] `P0.1` Confirm branch is `dev` with `git status --short --branch`.
  - Done:
- [ ] `P0.2` Confirm no unexpected user changes before editing code.
  - Done:
- [ ] `P0.3` Read this plan and update only unchecked items.
  - Done:
- [ ] `P0.4` Run baseline checks before first code change:
  - `python -m py_compile streamlit_app.py auth.py config.py gemini.py ui_common.py styles.py word_backend.py word_tab.py`
  - `python -m pip check`
  - Done:

### P1 - Fix Basic Flow And Language Direction

Goal: make the app reliably support only English -> Vietnamese and Vietnamese -> English, and fix `Dich co ban` so it actually analyzes then translates.

- [ ] `P1.1` Update `config.py` with two-language direction config:
  - `LANGUAGES`
  - `LANG_EN`
  - `TRANSLATION_DIRECTIONS`
  - Done:
- [ ] `P1.2` Add `_resolve_langs()` helper in `word_tab.py`.
  - Output should be `(direction_label, source_lang, target_lang)`.
  - Done:
- [ ] `P1.3` Replace target-language selectbox with horizontal direction radio.
  - Store selected direction in `st.session_state`.
  - Done:
- [ ] `P1.4` Fix quick-mode ordering bug.
  - Set `word_quick_mode` before `_run_analysis()` triggers `st.rerun()`, or pass mode into `_run_analysis()`.
  - Verify `Dich co ban` reaches `_run_full_translation()`.
  - Done:
- [ ] `P1.5` Thread `source_lang` and `target_lang` through `word_tab.py` session state:
  - `word_source_lang`
  - `word_target_lang`
  - `word_lang` or equivalent backward-compatible display label
  - Done:
- [ ] `P1.6` Thread `source_lang` through backend function signatures:
  - `build_glossary`
  - `build_doc_context`
  - `_build_chunk_prompt`
  - `translate_chunk`
  - `translate_chunk_with_retry`
  - `translate_parallel`
  - Done:
- [ ] `P1.7` Update all frontend call sites that currently use `LANG_EN[lang_word]`.
  - Expected after change: `rg "LANG_EN\\[" word_tab.py word_backend.py` has no stale translation call sites.
  - Done:
- [ ] `P1.8` Update chunk prompt to say `Translate from {source_lang} into {target_lang}`.
  - Done:
- [ ] `P1.9` Smoke verify both directions with a simple DOCX.
  - English -> Vietnamese
  - Vietnamese -> English
  - Quick mode
  - Advanced mode
  - Done:

### P2 - Fix DOCX Image Loss And Format Drift

Goal: preserve images, drawings, media-only paragraphs, and richer run formatting when applying translations.

- [ ] `P2.1` Add `media_only` to non-translation roles in `config.py`.
  - Done:
- [ ] `P2.2` Update `extract_docx_blocks()` to keep empty paragraphs that contain media.
  - Detect `w:drawing`, `w:pict`, `w:object`.
  - Add block role `media_only` with empty text.
  - Done:
- [ ] `P2.3` Update `_paragraph_has_non_run_children()` to force fallback when media/complex drawing tags are present.
  - Danger tags: `drawing`, `pict`, `object`, `AlternateContent`.
  - Done:
- [ ] `P2.4` Update `replace_paragraph_with_tagged()` so it does not remove runs containing media.
  - If media exists inside any run, fallback to `replace_paragraph_text_keep_format()`.
  - Done:
- [ ] `P2.5` Replace limited style copy with deep-copy of template `w:rPr`.
  - Clone run properties first, then apply translated bold/italic/underline toggles.
  - Keep fallback to current font name/size/color logic if copy fails.
  - Done:
- [ ] `P2.6` Extend `validate_docx_output()` with optional `original_bytes`.
  - Compare original/output `word/media/*` count.
  - Compare original/output `w:drawing` count.
  - Compare original/output `w:pict` count.
  - Mark output invalid if media files disappear.
  - Done:
- [ ] `P2.7` Update all validation call sites to pass original bytes.
  - Full translation
  - Partial/rescan translation
  - Batch, if validation is added there
  - Done:
- [ ] `P2.8` Create or collect DOCX test samples:
  - inline image inside text paragraph
  - image-only paragraph
  - table cell image
  - textbox/shape
  - footnote/endnote
  - track changes insert/delete
  - Done:
- [ ] `P2.9` Verify media preservation on samples.
  - `validate_docx_output(out, original_bytes=orig)` should be valid.
  - Manual Word open check: images still visible.
  - Done:

### P3 - Cost, Secrets, Checkpoint, And Export Hardening

Goal: fix audit issues that can mislead users or break portability.

- [ ] `P3.1` Centralize pricing in `config.py`.
  - Use one source of truth for actual cost display and estimate warning.
  - Add date/source comment for Gemini pricing.
  - Done:
- [ ] `P3.2` Remove or refactor `_estimate_cost()` to call `calc_cost()`.
  - Done:
- [ ] `P3.3` Recalculate cost estimates after role toggles.
  - Estimate should use currently selected translatable blocks, not stale `a["translatable"]`.
  - Done:
- [ ] `P3.4` Add safe Streamlit secrets getter in `config.py`.
  - Importing modules without `.streamlit/secrets.toml` should not crash.
  - App should still show missing password/API configuration clearly.
  - Done:
- [ ] `P3.5` Make checkpoint path portable.
  - Replace hard-coded `/tmp` with `tempfile.gettempdir()` or app cache directory.
  - Done:
- [ ] `P3.6` Escape dynamic text rendered with `unsafe_allow_html=True`.
  - Logs
  - Timer status
  - Timer error messages
  - Filename-derived text
  - Done:
- [ ] `P3.7` Sanitize batch ZIP filenames.
  - Basename only.
  - Remove `..`, `/`, `\`.
  - Ensure duplicate sanitized names remain unique.
  - Done:
- [ ] `P3.8` Re-run audit checks:
  - secret scan
  - `git diff --check`
  - `py_compile`
  - `pip check`
  - Done:

### P4 - Elevator/Escalator Domain Glossary

Goal: add domain-specific terminology without bloating non-domain documents.

- [ ] `P4.1` Create `data/glossary_elevator.json`.
  - Shape: `{ "en_vi": {...}, "vi_en": {...} }`.
  - Include core elevator terms from this plan.
  - Done:
- [ ] `P4.2` Create `data/glossary_escalator.json`.
  - Shape: `{ "en_vi": {...}, "vi_en": {...} }`.
  - Include core escalator terms from this plan.
  - Done:
- [ ] `P4.3` Create `domain_glossary.py`.
  - `STANDARDS_KEEP_AS_IS`
  - `load_seed(name)`
  - `detect_subdomain(blocks)`
  - `seed_for_direction(subdomains, source_lang, target_lang)`
  - Done:
- [ ] `P4.4` Detect elevator/escalator domain during analysis in `word_tab.py`.
  - Store `word_subdomains`.
  - Store `word_seed_glossary`.
  - Done:
- [ ] `P4.5` Merge seed glossary into `build_glossary()`.
  - Seed starts first.
  - AI-extracted glossary does not override seed.
  - User edits/imports still override after analysis.
  - Done:
- [ ] `P4.6` Expand `build_doc_context()` with domain style rules.
  - Preserve units.
  - Preserve standards.
  - Preserve part numbers, drawing numbers, revision codes.
  - Use formal technical register.
  - Done:
- [ ] `P4.7` Increase glossary prompt cap from 50 to 80 entries, with seed entries first.
  - Done:
- [ ] `P4.8` Add UI action near glossary import: `Khoi phuc seed thuat ngu nganh`.
  - Preferred behavior: merge seed back without deleting user edits.
  - Done:
- [ ] `P4.9` Verify domain behavior.
  - Elevator doc uses seed terms.
  - Escalator doc uses seed terms.
  - Non-domain doc has empty seed.
  - Standards such as `EN 81-20` remain unchanged.
  - Done:

### P5 - Tests And Smoke Scripts

Goal: reduce duplicate manual testing and catch future regressions.

- [ ] `P5.1` Decide whether to add `pytest` to requirements or keep smoke scripts only.
  - Done:
- [ ] `P5.2` Add tests or scriptable smoke checks for safe config import without secrets.
  - Done:
- [ ] `P5.3` Add tests or smoke checks for DOCX extract/apply/validate.
  - Done:
- [ ] `P5.4` Add tests or smoke checks for media count preservation.
  - Done:
- [ ] `P5.5` Add tests or smoke checks for domain detection and seed merge.
  - Done:
- [ ] `P5.6` Add tests or smoke checks for checkpoint path creation.
  - Done:

### P6 - README Roadmap And Final Verification

Goal: document future work and leave the repo easy to continue.

- [ ] `P6.1` Add `Roadmap` section to `README.md`.
  - Quality
  - Speed/cost
  - Format/layout
  - UX/editor
  - Done:
- [ ] `P6.2` Update README language/direction docs after Phase 1.
  - Done:
- [ ] `P6.3` Update README domain glossary docs after Phase 4.
  - Done:
- [ ] `P6.4` Run final verification checklist.
  - `py_compile`
  - `pip check`
  - `git diff --check`
  - no active PDF/Review imports
  - no real secrets
  - app smoke run if possible
  - Done:

## Completion Log

- 2026-05-26, planning only: created this plan file and priority queue. No app code changed yet.

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
