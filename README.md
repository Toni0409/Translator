# ⬡ Translator — Word

App Streamlit dịch file **Microsoft Word (.docx)** theo 2 hướng **Anh → Việt**
và **Việt → Anh** bằng **Gemini**, tập trung cho tài liệu kỹ thuật
**thang máy / thang cuốn**, giữ nguyên layout, hình ảnh và format gốc.

- 📝 **Word** — extract paragraph → detect H/F + heuristic repeating → auto-detect
  domain thang máy/thang cuốn → nạp seed glossary → dịch song song (4 luồng)
  → validate DOCX output → download + OCR tuỳ chọn
- 🏗 **Domain glossary** — seed thuật ngữ ngành từ `data/glossary_elevator.json`
  và `data/glossary_escalator.json`, merge với glossary AI và glossary user import
- 🖼 **Media preservation** — giữ paragraph chỉ có ảnh/drawing, tránh rebuild run
  có media, validate media/drawing count so với file gốc
- 🔐 Password protection · ⏱ Live timer · 🔁 Exponential backoff · 💵 Cost tracking

> Các tính năng **PDF** và **So sánh / Đánh giá** đã chuyển sang `archive/` —
> không build, không có trong active app. Giữ lại để tham khảo.

---

## 🚀 Chạy nhanh

```bash
pip install -r requirements.txt
cp .streamlit/secrets.toml.example .streamlit/secrets.toml
# Mở .streamlit/secrets.toml điền GEMINI_API_KEY và APP_PASSWORD

streamlit run streamlit_app.py
```

App mở tại http://localhost:8501. Nhập password → upload file → chọn hướng dịch
`Anh → Việt` hoặc `Việt → Anh` → dịch.

---

## ⚙️ Secrets

File `.streamlit/secrets.toml`:

```toml
GEMINI_API_KEY = "your-gemini-api-key-here"
APP_PASSWORD   = "your-app-password"
```

Lấy API key tại https://aistudio.google.com/apikey.

---

## 🧩 Cấu trúc module

```
streamlit_app.py     # Entry — set config, inject CSS, render tab Word
config.py            # Hằng số: API key, model, prices, directions, thresholds
styles.py            # CSS dark theme
auth.py              # Password gate + logout
gemini.py            # Shared Gemini client, JSON parser, threaded call helper
ui_common.py         # timer_box / stat_box / log_adder / calc_cost
domain_glossary.py   # Seed glossary + domain detection elevator/escalator
word_backend.py      # Word: extract/apply/validate DOCX, translate_parallel
word_tab.py          # Word: Streamlit UI tab
data/
  glossary_elevator.json
  glossary_escalator.json
tests/               # Standalone smoke scripts, no pytest dependency

archive/             # Tính năng đã bỏ — giữ lại để tham khảo
  pdf_backend.py
  pdf_tab.py
  review_backend.py
  review_tab.py
  review_vision_backend.py
```

### Sơ đồ phụ thuộc

```
streamlit_app.py
  ├── styles
  ├── auth
  └── word_tab
        ├── word_backend
        │     └── gemini
        ├── domain_glossary
        ├── gemini
        ├── ui_common
        └── config
```

---

## 📝 Tab Word — chi tiết tính năng

### Flow một nút

```
[Upload + direction]
   ▼
[🚀 Dịch]
   ▼
[Analyze + source/target + domain detect + seed glossary]
   ▼
[Translate chunks in parallel]
   ▼
[Validate media/format + Download + Rescan + OCR]
```

### Feature table

| Feature | Mô tả |
|---|---|
| **Direction picker** | UI chỉ còn 2 hướng `Anh → Việt` và `Việt → Anh`; prompt truyền rõ `source_lang` + `target_lang`. |
| **One-click flow** | Chỉ còn một nút `Dịch`; app tự phân tích, nạp glossary kỹ thuật, dịch và validate output. |
| **Seed domain glossary** | Auto-detect elevator/escalator → nạp thuật ngữ chuẩn ngành từ `data/`, sau đó merge với glossary AI. |
| **Media preservation** | Paragraph chỉ có ảnh/drawing được đánh role `media_only`, không gửi AI, không rebuild run; output validate media count với file gốc. |
| **Text-box / shapes** | Extract paragraphs inside `w:txbxContent` (text-box, shape) → dịch và ghi lại đúng vị trí. |
| **Table-aware translation** | Mỗi cell gắn prefix `(T# R# C#)` trong prompt → AI dịch consistent theo column, R1 = header ngắn gọn, giữ số/đơn vị. Auto-strip prefix ở output. |
| **Inline format** | Bold/italic/underline encode `<b><i><u>` tags → AI preserve qua dịch → rebuild runs bằng deep-copy `w:rPr` khi an toàn. |
| **Cross-chunk glossary** | 1 Gemini call build glossary → inject mọi chunk prompt → consistent terminology cross-chunk |
| **Parallel chunks** | `ThreadPoolExecutor` dịch 4 chunk cùng lúc → giảm thời gian ~4× |
| **Adaptive chunking** | Chia theo ký tự (~8k chars/chunk), min 8 / max 40 paragraph/chunk |
| **Per-chunk retry** | Mỗi chunk fail sẽ retry tối đa 3 lần với backoff 1s/2s/4s |
| **Thread-safe model fallback** | `gemini-3.5-flash` → `2.5-flash-lite` → `2.5-flash` → `2.0-flash` → ... (lock-protected) |
| **Cached DOCX rebuild** | Version counter → `apply_translations` chỉ chạy khi translations thay đổi |
| **H/F detection** | Detect cả H/F thật (docx structure) **và** text lặp lại ≥3 lần trong body (heuristic); H/F được dịch trong flow chính. |
| **Quét bỏ sót** | Phát hiện đoạn API fail (translation == original) → dịch lại các đoạn còn thiếu. |
| **Token + cost** | Cộng dồn input/output tokens → quy đổi USD/VND realtime từ pricing source-of-truth trong `config.py`. |
| **Role badges** | Header / Footer / 🔁 Lặp lại / 🔲 Text-box / Heading / Bullet / TOC / Cell / Note |

### Logic H/F detection

```
1. extract_docx_blocks() đọc tất cả paragraph (body + tables + headers + footers)
2. Block trong section.header / section.footer → role = "header"/"footer"
3. Block trong body có text lặp lại ≥ 3 lần (text dài ≥ 10 chars)
   → role = "body_repeated"  (heuristic bắt H/F-like trong body)
4. Lần dịch đầu: bỏ qua mọi role trong NO_TRANSLATE_ROLES
5. User bấm "Dịch Header/Footer" → translate đúng các role này
6. apply_translations() áp dụng MỌI translation trong dict
   (không phân biệt role) → file Word có đầy đủ
```

### Logic cross-chunk glossary

```
1. Sau extract, detect subdomain elevator/escalator từ keyword trong tài liệu
2. Nạp seed glossary theo hướng dịch từ:
   - data/glossary_elevator.json
   - data/glossary_escalator.json
3. Gọi Gemini 1 lần để gợi ý thêm thuật ngữ lặp lại trong document
4. Merge glossary theo precedence: user-imported > seed chuẩn > AI-extracted
5. Inject glossary vào mỗi chunk prompt ("USE THESE EXACT translations")
6. Rescan + H/F translation tái dùng glossary đã edit (stored in session_state)
```

### Logic domain thang máy / thang cuốn

```
1. domain_glossary.detect_subdomain(blocks) quét 200 block đầu
2. Nếu đủ keyword elevator/lift/hoistway/cabin/thang máy/giếng thang...
   → domain = elevator
3. Nếu đủ keyword escalator/handrail/comb plate/thang cuốn/tay vịn/tấm lược...
   → domain = escalator
4. seed_for_direction(subdomains, source_lang, target_lang) chọn en_vi hoặc vi_en
5. build_doc_context() inject style guide kỹ thuật:
   - giữ nguyên unit: mm, m, m/s, kg, kN, V, Hz, °C
   - giữ nguyên standard: EN 81-20, EN 81-50, ISO 22201, ASME A17.1, TCVN 6395
   - giữ part number, drawing number, revision code
   - văn phong kỹ thuật chính thức, tránh dịch văn nói
```

### Logic dịch một nút

```
1. User upload DOCX + chọn hướng Anh → Việt hoặc Việt → Anh
2. _run_analysis() extract block, detect domain, build seed/AI glossary, build doc_context
3. word_auto_translate trigger _run_full_translation() ngay sau khi phân tích xong
4. _run_full_translation() chunk blocks, gọi Gemini song song, validate media/format
5. UI hiển thị download chính, rescan nếu cần, và OCR tuỳ chọn
```

### Logic table-aware translation

```
1. extract: mỗi cell → block.table_cell = (T_idx, R_idx, C_idx) 1-based
2. prompt: prefix "(T1 R2 C3) cell text" → AI biết context
3. rule thêm vào prompt: "same C# = same terminology, R1 = header concise"
4. output: regex strip prefix (phòng AI giữ lại) → translation gọn
```

### Logic inline format (bold/italic/underline)

```
1. extract: runs_to_tagged_text() → "<b>Important</b>: read <i>carefully</i>"
2. prompt: thêm rule "PRESERVE <b><i><u> tags in translation"
3. AI giữ nguyên hoặc dịch text quanh tags
4. apply: replace_paragraph_with_tagged() parse tags → rebuild runs với format
5. Nếu paragraph/run có drawing/pict/object → fallback sang path giữ XML để không mất ảnh
6. validate_docx_output(out, original_bytes=orig) so media/drawing count với file gốc
```

---

## 🛠 Tuỳ chỉnh nhanh

| Bạn muốn… | Sửa file |
|---|---|
| Đổi Gemini model | `config.py` → `WORD_MODELS` (model đầu tiên = ưu tiên) |
| Đổi hướng dịch hiển thị | `config.py` → `TRANSLATION_DIRECTIONS` |
| Sửa glossary thang máy | `data/glossary_elevator.json` |
| Sửa glossary thang cuốn | `data/glossary_escalator.json` |
| Sửa keyword detect domain | `domain_glossary.py` |
| Đổi giá / tỉ giá | `config.py` → `PRICE_INPUT`, `PRICE_OUTPUT`, `USD_TO_VND` |
| Đổi chunk size | `config.py` → `TARGET_CHUNK_CHARS`, `MIN_CHUNK_BLOCKS`, `MAX_CHUNK_BLOCKS` |
| Đổi số luồng song song | `config.py` → `MAX_WORD_WORKERS` |
| Đổi số lần retry mỗi chunk | `config.py` → `CHUNK_RETRIES` |
| Đổi ngưỡng H/F repeating | `config.py` → `HF_REPEAT_THRESHOLD`, `HF_REPEAT_MIN_CHARS` |
| Đổi role bị bỏ qua | `config.py` → `NO_TRANSLATE_ROLES` |
| Đổi theme / màu | `styles.py` |
| Tinh chỉnh prompt | `word_backend.py` → `_build_chunk_prompt` |
| Thay đổi heuristic H/F | `word_backend.py` → `_mark_repeating_as_hf` |

---

## 🧪 Smoke tests

Không dùng pytest để giữ runtime dependencies gọn. Chạy toàn bộ smoke scripts:

```bash
python tests/run_smoke.py
```

Các nhóm hiện có:

- `smoke_config.py` — import config không cần secrets thật, pricing, directions, roles
- `smoke_direction.py` — direction resolver + batch ZIP filename sanitizer
- `smoke_docx.py` — extract/apply/validate DOCX, media preservation
- `smoke_domain.py` — detect domain, seed glossary, doc context domain rules
- `smoke_checkpoint.py` — checkpoint path portable + save/load/clear
- `smoke_ocr.py` — cost helpers (tile/token), prompt builder (source_lang + domain + glossary), occurrence extraction, caption selection + dedupe + edit override, overlay Pillow render, replace-by-occurrence với clone media khi shared

---

## 🖼 OCR ảnh trong DOCX

Sau khi dịch xong, expander **OCR & dịch text trong ảnh** cho 3 bước:

1. **Quét ảnh & ước tính chi phí** — liệt kê occurrence (mỗi `r:embed` trong `<w:p>` là 1 occurrence; cùng ảnh dùng nhiều nơi vẫn ra nhiều occurrence), hiển thị USD/VND ước tính theo kích thước ảnh × giá Gemini Vision.
2. **Xác nhận chi phí + chạy OCR** — Gemini Vision OCR + dịch song song, lưu actual cost per-image từ `usage_metadata` + aggregate batch.
3. **Review & xuất**:
   - Mỗi ảnh: checkbox **Đưa vào file xuất**, checkbox **Giữ ảnh gốc** (caption mode), text_area chỉnh bản dịch.
   - Nhóm "Không phát hiện chữ / lỗi" collapsed riêng, default không chọn.
   - Chọn tất cả / bỏ chọn tất cả.
   - Output mode (radio):
     - **Đưa text dưới ảnh** (default): chèn caption `[OCR] <text>` ngay dưới mỗi ảnh được chọn, italic, căn giữa. Dedupe khi xuất lại nhiều lần. Có thể chọn xoá ảnh gốc (chỉ giữ caption).
     - **Dịch trực tiếp trên ảnh**: che chữ gốc trong từng vùng OCR (Pillow `ImageDraw`, fill bằng avg color của vùng), vẽ bản dịch lên đúng bbox. Không làm song ngữ trên ảnh. Ảnh thiếu bbox → fallback caption per-image.
   - Replace ảnh trong DOCX theo **occurrence** — khi ảnh dùng nhiều chỗ (cùng `rId`), clone media file mới + tạo rId mới cho occurrence sau lần đầu, đảm bảo không thay nhầm ảnh khác.

File download riêng: `*_translated_<lang>_ocr_caption.docx` hoặc `*_translated_<lang>_ocr_overlay.docx`.

---

## 📦 Dependencies

- `streamlit` — UI framework
- `google-genai` — Gemini SDK
- `python-docx` — DOCX parsing + writing
- `pandas` — `data_editor` cho tab Word
- `lxml` — DOCX XML manipulation
- `Pillow` — image overlay rendering (OCR overlay mode) + đọc kích thước ảnh để estimate tokens

---

## 🔒 Lưu ý bảo mật

- **Không commit** `.streamlit/secrets.toml` (đã có trong `.gitignore`).
- App có password gate cơ bản — chỉ phù hợp dùng nội bộ. Không dùng cho production
  public mà chưa bổ sung HTTPS + rate limiting + audit log.

---

## 🗺 Roadmap

### Chất lượng

- Terminology QA so với glossary đã chốt trước khi xuất file.
- Post-edit pass riêng cho tài liệu kỹ thuật thang máy/thang cuốn.
- Consistency check cross-chunk cho thuật ngữ, số liệu, đơn vị, standard.
- Style-guide injection theo loại tài liệu: spec, manual, checklist, report.

### Tốc độ / chi phí

- Gemini Batch API cho tài liệu lớn để giảm chi phí.
- Prompt/context caching cho glossary + document context.
- Dashboard token/cost theo file và theo batch.
- Smarter chunk sizing theo độ phức tạp thay vì chỉ theo ký tự.

### Format / layout

- Bổ sung sample thật cho table image, textbox image, footnote/endnote, track changes.
- Kiểm tra comment anchor, nested table, drawing canvas, SmartArt.
- Mở rộng xử lý equation/OMML cho text label quanh công thức.
- Verify image alt-text pipeline trên tài liệu thực tế.

### UX / editor

- Side-by-side preview file gốc vs bản dịch.
- Diff bản dịch cũ vs bản mới khi rescan.
- Batch UI rõ hơn cho nhiều file lớn.

---

## 🗂 Archive

Các tính năng đã ngừng dùng được giữ lại trong `archive/`:

- `pdf_backend.py` + `pdf_tab.py` — dịch PDF (extract spans → translate → redact + rewrite)
- `review_backend.py` + `review_tab.py` + `review_vision_backend.py` — so sánh/đánh giá bản dịch

Để bật lại: di chuyển ra khỏi `archive/`, thêm lại import vào `streamlit_app.py`,
restore deps trong `requirements.txt` (`pymupdf`) và `packages.txt`
(`libreoffice-writer`, `libreoffice-core`, fonts).

---

## 📝 License

Internal use — Vi Nguyen.
