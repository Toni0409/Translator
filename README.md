# ⬡ Translator — PDF & Word

App Streamlit dịch file **PDF** và **Microsoft Word (.docx)** sang nhiều ngôn ngữ
bằng **Gemini**, giữ nguyên layout / format gốc.

- 📄 **PDF**: extract text spans → dịch → ghi lại đúng vị trí + font + màu + bold
- 📝 **Word**: extract paragraph → dịch chunk **song song (4 luồng)** với per-chunk retry
  → giữ format → **preview + sửa inline** + **rescan đoạn bỏ sót**
- 🔐 Password protection · ⏱ Live timer · 🔁 Exponential backoff · 💵 Cost tracking

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
streamlit_app.py     # Entry — set config, inject CSS, render 2 tab
config.py            # Hằng số: API key, model, prices, languages, fonts
styles.py            # CSS dark theme
auth.py              # Password gate + logout
gemini.py            # Shared Gemini client, JSON parser, threaded call helper
ui_common.py         # timer_box / stat_box / log_adder
pdf_backend.py       # PDF: extract_line_groups, translate_page, write_translated_pdf
pdf_tab.py           # PDF: Streamlit UI tab
word_backend.py      # Word: extract_docx_blocks, translate_chunk, apply_translations
word_tab.py          # Word: Streamlit UI tab
```

Mỗi file < 300 dòng, một trách nhiệm rõ ràng → AI / vibe code dễ đọc & sửa.

### Sơ đồ phụ thuộc

```
streamlit_app.py
  ├── styles
  ├── auth ──────────────┐
  ├── pdf_tab            │
  │     ├── pdf_backend  │
  │     │     └── gemini ┤
  │     ├── gemini       │
  │     ├── ui_common    │
  │     └── config ──────┤
  └── word_tab           │
        ├── word_backend │
        │     └── gemini │
        ├── gemini       │
        ├── ui_common    │
        └── config ──────┘
```

---

## 🛠 Tuỳ chỉnh nhanh

| Bạn muốn… | Sửa file |
|---|---|
| Đổi Gemini model | `config.py` → `PDF_MODEL`, `WORD_MODELS` |
| Thêm ngôn ngữ | `config.py` → `LANGUAGES`, `LANG_EN` |
| Đổi giá / tỉ giá | `config.py` → `PRICE_INPUT`, `PRICE_OUTPUT`, `USD_TO_VND` |
| Đổi chunk size Word | `config.py` → `TARGET_CHUNK_CHARS`, `MIN_CHUNK_BLOCKS`, `MAX_CHUNK_BLOCKS` |
| Đổi số luồng song song Word | `config.py` → `MAX_WORD_WORKERS` |
| Đổi số lần retry mỗi chunk | `config.py` → `CHUNK_RETRIES` |
| Đổi theme / màu | `styles.py` |
| Tinh chỉnh prompt PDF | `pdf_backend.py` → `translate_page` |
| Tinh chỉnh prompt Word | `word_backend.py` → `translate_chunk` |

---

## ⚡ Tab Word — chi tiết tính năng

| Feature | Mô tả |
|---|---|
| **Parallel chunks** | `ThreadPoolExecutor` dịch 4 chunk cùng lúc → giảm thời gian ~4× |
| **Adaptive chunking** | Chia theo ký tự (~8k chars/chunk) thay vì fixed 25 paragraphs |
| **Per-chunk retry** | Mỗi chunk fail sẽ retry tối đa 3 lần với backoff 1s/2s/4s |
| **Model fallback** | `gemini-2.5-flash-lite` → `flash` → `2.0-flash` → ... |
| **Token + cost** | Cộng dồn input/output tokens → quy đổi USD/VND realtime |
| **Inline edit** | Sau khi dịch, hiện bảng `data_editor` — sửa trực tiếp bản dịch |
| **Quét bỏ sót** | Phát hiện đoạn API fail (translation == original) → dịch lại |
| **Auto-detect role** | Heading / TOC / bullet / table_cell → prompt phù hợp từng loại |
| **Skip header/footer** | Giữ nguyên (cấu hình tại `NO_TRANSLATE_ROLES`) |

---

## 📦 Dependencies

- `streamlit` — UI framework
- `google-genai` — Gemini SDK
- `pymupdf` (fitz) — PDF parsing + writing
- `python-docx` — DOCX parsing + writing
- `pandas` — data_editor cho tab Word

---

## 🔒 Lưu ý bảo mật

- **Không commit** `.streamlit/secrets.toml` (đã có trong `.gitignore`).
- App có password gate cơ bản — chỉ phù hợp dùng nội bộ. Không dùng cho production
  public mà chưa bổ sung HTTPS + rate limiting + audit log.

---

## 📝 License

Internal use — Vi Nguyen.
