# ⬡ Translator — PDF & Word

App Streamlit dịch file **PDF** và **Microsoft Word (.docx)** sang nhiều ngôn ngữ
bằng **Gemini**, giữ nguyên layout / format gốc.

- 📄 **PDF** — extract text spans → **phát hiện bảng** → dịch với context (T# R# C#)
  → ghi lại đúng vị trí + font + màu + bold
- 📝 **Word** — extract paragraph → **detect H/F + heuristic repeating** → dịch
  body song song (4 luồng) → preview + sửa inline → nút riêng dịch H/F
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
config.py            # Hằng số: API key, model, prices, languages, fonts, thresholds
styles.py            # CSS dark theme
auth.py              # Password gate + logout
gemini.py            # Shared Gemini client, JSON parser, threaded call helper
ui_common.py         # timer_box / stat_box / log_adder / calc_cost
pdf_backend.py       # PDF: extract_line_groups (+ table detection), translate_page
pdf_tab.py           # PDF: Streamlit UI tab
word_backend.py      # Word: extract_docx_blocks (+ H/F detection), translate_parallel
word_tab.py          # Word: Streamlit UI tab
```

Mỗi file < 500 dòng, một trách nhiệm rõ ràng → AI / vibe code dễ đọc & sửa.

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

## 📄 Tab PDF — chi tiết tính năng

| Feature | Mô tả |
|---|---|
| **Layout-preserving** | Redact text gốc → in lại bản dịch vào đúng bbox cũ |
| **Font + style** | Auto-detect font Unicode, bold flag, màu chữ giữ nguyên |
| **Page range** | Nhập "1-5,8,10-12" để dịch page cụ thể |
| **Table detection** | `page.find_tables()` → kèm `(T# R# C#)` vào prompt để AI dịch cell consistent |
| **Smart prompt** | Khi có bảng → thêm rules: header row ngắn gọn, cùng cột dùng cùng thuật ngữ, giữ nguyên số/đơn vị |
| **Live timer** | Đếm giây realtime trong lúc chờ Gemini API |
| **Rate-limit retry** | Exponential backoff 5/10/20/40/80s khi gặp 429 |
| **Cost tracking** | Token in/out → USD + VND realtime |

### Logic table detection

```
1. page.find_tables() → list of tables (mỗi table có bbox + cells)
2. Với mỗi line bbox, tìm cell chứa tâm dòng
3. Group có cell info → format prompt: "[5] (T1 R2 C3) cell text"
4. AI thấy context bảng → dịch consistent + giữ format số
5. Bản dịch trả về vẫn theo index [0]...[N-1] — không có prefix (T# R# C#)
```

---

## 📝 Tab Word — chi tiết tính năng

| Feature | Mô tả |
|---|---|
| **Parallel chunks** | `ThreadPoolExecutor` dịch 4 chunk cùng lúc → giảm thời gian ~4× |
| **Adaptive chunking** | Chia theo ký tự (~8k chars/chunk), min 8 / max 40 paragraph/chunk |
| **Per-chunk retry** | Mỗi chunk fail sẽ retry tối đa 3 lần với backoff 1s/2s/4s |
| **Model fallback** | `gemini-2.5-flash-lite` → `flash` → `2.0-flash` → ... |
| **Token + cost** | Cộng dồn input/output tokens → quy đổi USD/VND realtime |
| **H/F detection** | Detect cả H/F thật (docx structure) **và** text lặp lại ≥3 lần trong body (heuristic) |
| **Skip H/F by default** | Lần dịch đầu chỉ làm body. Bấm nút riêng để dịch H/F |
| **Nút Dịch H/F** | Hiện ra với count `(missed/total)` — dịch song song như body |
| **Inline edit** | `data_editor` — sửa trực tiếp bản dịch, cả body + H/F |
| **Filter editor** | "Chỉ hiển thị đoạn chưa dịch" + "Hiện cả Header/Footer" |
| **Quét bỏ sót** | Phát hiện đoạn API fail (translation == original) → dịch lại |
| **Role badges** | Header / Footer / 🔁 Lặp lại / Heading / Bullet / TOC / Cell / Note |
| **Auto-detect role** | Heading / TOC / bullet / table_cell → prompt phù hợp từng loại |

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
| Đổi ngưỡng H/F repeating | `config.py` → `HF_REPEAT_THRESHOLD`, `HF_REPEAT_MIN_CHARS` |
| Đổi role bị bỏ qua | `config.py` → `NO_TRANSLATE_ROLES` |
| Đổi theme / màu | `styles.py` |
| Tinh chỉnh prompt PDF | `pdf_backend.py` → `translate_page` |
| Tinh chỉnh prompt Word | `word_backend.py` → `_build_chunk_prompt` |
| Thay đổi heuristic H/F | `word_backend.py` → `_mark_repeating_as_hf` |
| Thay đổi table detection | `pdf_backend.py` → `detect_tables_on_page` |

---

## 📦 Dependencies

- `streamlit` — UI framework
- `google-genai` — Gemini SDK
- `pymupdf` (fitz) — PDF parsing + writing + table detection
- `python-docx` — DOCX parsing + writing
- `pandas` — `data_editor` cho tab Word

---

## 🔒 Lưu ý bảo mật

- **Không commit** `.streamlit/secrets.toml` (đã có trong `.gitignore`).
- App có password gate cơ bản — chỉ phù hợp dùng nội bộ. Không dùng cho production
  public mà chưa bổ sung HTTPS + rate limiting + audit log.

---

## 📝 License

Internal use — Vi Nguyen.
