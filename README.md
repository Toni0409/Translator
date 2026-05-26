# ⬡ Translator — Word

App Streamlit dịch file **Microsoft Word (.docx)** sang nhiều ngôn ngữ bằng
**Gemini**, giữ nguyên layout / format gốc.

- 📝 **Word** — extract paragraph → **detect H/F + heuristic repeating** → dịch
  body song song (4 luồng) → preview + sửa inline → nút riêng dịch H/F
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

App mở tại http://localhost:8501. Nhập password → upload file → chọn ngôn ngữ → dịch.

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
config.py            # Hằng số: API key, model, prices, languages, thresholds
styles.py            # CSS dark theme
auth.py              # Password gate + logout
gemini.py            # Shared Gemini client, JSON parser, threaded call helper
ui_common.py         # timer_box / stat_box / log_adder / calc_cost
word_backend.py      # Word: extract_docx_blocks (+ H/F detection), translate_parallel
word_tab.py          # Word: Streamlit UI tab

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
        ├── gemini
        ├── ui_common
        └── config
```

---

## 📝 Tab Word — chi tiết tính năng

### Flow 2 phase

```
[Upload + lang]
   ▼
[🔬 Phân tích + Glossary]   ──or──  [⚡ Phân tích & dịch luôn]
   ▼                                       │
[Stats: body/H-F/textbox/cells]            │
[💾 TM preview]                            │
[📚 Glossary editor inline (optional)]     │
   ▼                                       │
[▶ Dịch X đoạn]  ◀────────────────────────┘
   ▼
[Download + Rescan + H/F + Inline editor]
```

### Feature table

| Feature | Mô tả |
|---|---|
| **2-phase flow** | Phân tích (extract+glossary) → review glossary → Dịch. Có nút `⚡ Phân tích & dịch luôn` để skip review. |
| **Glossary editor** | Hiển thị top-30 thuật ngữ AI suggest → user sửa/xóa/thêm trước khi dịch → bắt buộc terminology cụ thể. |
| **Translation Memory** | Hash text + target_lang → cache. Dịch lại doc cũ = 100% TM hit, $0 API. Persist xuyên session, có nút clear. |
| **Text-box / shapes** | Extract paragraphs inside `w:txbxContent` (text-box, shape) → dịch và ghi lại đúng vị trí. |
| **Table-aware translation** | Mỗi cell gắn prefix `(T# R# C#)` trong prompt → AI dịch consistent theo column, R1 = header ngắn gọn, giữ số/đơn vị. Auto-strip prefix ở output. |
| **Inline format** | Bold/italic/underline encode `<b><i><u>` tags → AI preserve qua dịch → rebuild runs giữ font/size/color |
| **Cross-chunk glossary** | 1 Gemini call build glossary → inject mọi chunk prompt → consistent terminology cross-chunk |
| **Parallel chunks** | `ThreadPoolExecutor` dịch 4 chunk cùng lúc → giảm thời gian ~4× |
| **Adaptive chunking** | Chia theo ký tự (~8k chars/chunk), min 8 / max 40 paragraph/chunk |
| **Per-chunk retry** | Mỗi chunk fail sẽ retry tối đa 3 lần với backoff 1s/2s/4s |
| **Thread-safe model fallback** | `gemini-3.5-flash` → `2.5-flash-lite` → `2.5-flash` → `2.0-flash` → ... (lock-protected) |
| **Cached DOCX rebuild** | Version counter → `apply_translations` chỉ chạy khi translations thay đổi |
| **H/F detection** | Detect cả H/F thật (docx structure) **và** text lặp lại ≥3 lần trong body (heuristic) |
| **Skip H/F by default** | Lần dịch đầu chỉ làm body. Nút riêng để dịch H/F (cũng dùng TM auto) |
| **Quét bỏ sót** | Phát hiện đoạn API fail (translation == original) → dịch lại (TM auto) |
| **Inline edit** | `data_editor` — sửa trực tiếp bản dịch, cả body + H/F. Vai trò hiện kèm coord bảng (T#R#C#) |
| **Filter editor** | "Chỉ hiển thị đoạn chưa dịch" + "Hiện cả Header/Footer" |
| **Token + cost** | Cộng dồn input/output tokens → quy đổi USD/VND realtime |
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
1. Sau extract, trích noun phrases + technical terms lặp lại ≥ 3 lần
2. Gọi Gemini 1 lần để dịch top-30 thuật ngữ → {en: vi}
3. User review/sửa ở data_editor inline (Phase 1) — optional
4. Inject glossary vào mỗi chunk prompt ("USE THESE EXACT translations")
5. Rescan + H/F translation tái dùng glossary đã edit (stored in session_state)
```

### Logic Translation Memory

```
1. Hash key = md5("{target_lang}|{text}")[:16]
   → khác lang = khác key (Việt vs Nhật cho cùng text → 2 entry riêng)
2. Trước translate: tm_lookup(blocks) → (cached_dict, remaining_blocks)
3. Chỉ chunk + gọi API cho `remaining` → tiết kiệm tokens
4. Sau translate: tm_store(remaining, new_translations) → grow TM
5. Persist trong session_state["word_tm"] — sống xuyên doc cho đến khi clear
6. Rescan + H/F cũng dùng TM (auto)
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
   Fallback về first-run-wins nếu paragraph có hyperlink/field
```

---

## 🛠 Tuỳ chỉnh nhanh

| Bạn muốn… | Sửa file |
|---|---|
| Đổi Gemini model | `config.py` → `WORD_MODELS` (model đầu tiên = ưu tiên) |
| Thêm ngôn ngữ | `config.py` → `LANGUAGES`, `LANG_EN` |
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

## 📦 Dependencies

- `streamlit` — UI framework
- `google-genai` — Gemini SDK
- `python-docx` — DOCX parsing + writing
- `pandas` — `data_editor` cho tab Word
- `lxml` — DOCX XML manipulation

---

## 🔒 Lưu ý bảo mật

- **Không commit** `.streamlit/secrets.toml` (đã có trong `.gitignore`).
- App có password gate cơ bản — chỉ phù hợp dùng nội bộ. Không dùng cho production
  public mà chưa bổ sung HTTPS + rate limiting + audit log.

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
