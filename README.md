# ⬡ Translator — PDF & Word

App Streamlit dịch file **PDF** và **Microsoft Word (.docx)** sang nhiều ngôn ngữ
bằng **Gemini**, giữ nguyên layout / format gốc.

- 📄 **PDF**: extract text spans → dịch → ghi lại đúng vị trí + font + màu + bold
- 📝 **Word**: extract paragraph → dịch theo chunk (model fallback) → giữ nguyên run format
- 🔐 Password protection · ⏱ Live timer · 🔁 Retry với exponential backoff

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
| Đổi chunk size Word | `config.py` → `CHUNK_SIZE` |
| Đổi theme / màu | `styles.py` |
| Tinh chỉnh prompt PDF | `pdf_backend.py` → `translate_page` |
| Tinh chỉnh prompt Word | `word_backend.py` → `translate_chunk` |

---

## 📦 Dependencies

- `streamlit` — UI framework
- `google-genai` — Gemini SDK
- `pymupdf` (fitz) — PDF parsing + writing
- `python-docx` — DOCX parsing + writing

---

## 🔒 Lưu ý bảo mật

- **Không commit** `.streamlit/secrets.toml` (đã có trong `.gitignore`).
- App có password gate cơ bản — chỉ phù hợp dùng nội bộ. Không dùng cho production
  public mà chưa bổ sung HTTPS + rate limiting + audit log.

---

## 📝 License

Internal use — Vi Nguyen.
