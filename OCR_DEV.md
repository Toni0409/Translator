# OCR Dev Plan - Word Translator

Branch: `dev`

Mục tiêu: phát triển mạnh luồng OCR cho ảnh trong DOCX, có kiểm soát chi phí Gemini, có màn review ảnh kèm chữ, chọn ảnh nào xử lý, và xuất DOCX theo 2 kiểu:

1. Dịch trực tiếp trên ảnh: chèn chữ dịch đè lên ảnh đúng vùng.
2. Dịch dưới ảnh: đưa text dịch xuống dưới ảnh như caption/code hiện tại.

File này là nguồn điều phối OCR trên branch `dev`. Khi AI/code xong một mục, phải tick mục đó, ghi ngày, commit, test đã chạy và note ngắn vào `Completion log` để tránh duplicate việc.

## Cần Làm Trước

- [ ] `P0.1` Chốt data model OCR theo từng lần ảnh xuất hiện trong DOCX, không chỉ theo file `word/media/*`.
  - Lý do: một ảnh media có thể được dùng nhiều nơi; user cần tick chọn từng ảnh theo vị trí hiển thị.
  - Output mong muốn: `ImageOccurrence` có `id`, `filename`, `content_type`, `data`, `doc_part`, `rels_path`, `rId`, `paragraph_index`, `width`, `height`, `occurrence_index`.
  - Done note:

- [ ] `P0.2` Tách state OCR trong `st.session_state` thành object rõ ràng.
  - Hiện tại: `word_image_ocr = (imgs, results)`.
  - Mục tiêu: `word_ocr_state = {"images": ..., "results": ..., "selection": ..., "cost": ..., "mode": ...}`.
  - Done note:

- [ ] `P1.1` Hiển thị ước tính chi phí trước khi bấm OCR.
  - Dùng số ảnh, kích thước ảnh, model hiện tại và giá trong `config.py`.
  - Show cả USD và VND.
  - Ghi rõ đây là ước tính; actual cost sau khi chạy lấy từ `response.usage_metadata`.
  - Done note:

- [ ] `P1.2` Lưu actual cost sau OCR theo từng ảnh và tổng batch.
  - `_ocr_single_image()` phải trả thêm `tok_in`, `tok_out`, `total_tokens`, `usd`, `vnd`, `model`.
  - `ocr_and_translate_images()` aggregate tổng chi phí.
  - UI show: tổng input/output tokens, USD, VND, model.
  - Done note:

- [ ] `P2.1` Sau khi bấm "Quét và dịch ảnh", render màn review ảnh bên trái, OCR/chữ dịch bên phải.
  - Mỗi ảnh có checkbox `Chọn đưa vào DOCX`.
  - Có nút chọn tất cả / bỏ chọn tất cả.
  - Có `text_area` cho user sửa bản dịch OCR trước khi xuất.
  - Ảnh không có text hoặc lỗi OCR phải nằm nhóm riêng, mặc định không chọn.
  - Done note:

- [ ] `P2.2` Thêm checkbox giữ/bỏ ảnh gốc cho từng ảnh.
  - Default: giữ ảnh gốc.
  - Nếu user bỏ ảnh gốc ở mode caption: xóa/ẩn ảnh và chỉ giữ text dịch dưới vị trí đó.
  - Nếu user bỏ ảnh gốc ở mode overlay: không hợp lệ hoặc phải cảnh báo vì overlay cần nền ảnh; ưu tiên disable checkbox trong mode overlay.
  - Done note:

- [ ] `P3.1` Hoàn thiện option 2 - đưa text dưới ảnh.
  - Refactor `insert_ocr_captions_into_docx()` để nhận `selected_ids` và `edited_translations`.
  - Chỉ chèn caption cho ảnh được tick.
  - Không chèn duplicate nếu user bấm xuất lại nhiều lần trong cùng session.
  - Preserve behavior hiện tại: caption căn giữa, font nhỏ, prefix `[OCR]`.
  - Done note:

- [ ] `P4.1` Thiết kế option 1 - dịch trực tiếp trên ảnh.
  - OCR prompt phải trả thêm vùng chữ: `regions[]` với `bbox`, `ocr`, `translation`, `confidence`.
  - Chuẩn hóa `bbox` về `[x, y, w, h]` theo tỷ lệ 0..1 để độc lập kích thước ảnh.
  - Nếu Gemini không trả bbox đáng tin, fallback rõ sang caption mode.
  - Done note:

- [ ] `P4.2` Implement render overlay lên ảnh.
  - Cần dependency xử lý ảnh, ưu tiên `Pillow`.
  - Vẽ nền mờ/opaque phía sau text dịch để đọc được nhưng vẫn giữ context ảnh.
  - Auto-fit font theo bbox, wrap text, tránh tràn vùng.
  - Output ảnh mới dạng PNG/JPEG, giữ đúng aspect ratio.
  - Done note:

- [ ] `P4.3` Replace ảnh trong DOCX đúng occurrence được chọn.
  - Không được thay nhầm mọi ảnh dùng chung cùng media part.
  - Nếu technical constraint của DOCX bắt buộc replace theo media part, phải clone media part riêng cho occurrence đó.
  - Preserve relationship, kích thước hiển thị, layout paragraph/table.
  - Done note:

## Hiện Trạng Code OCR

- `word_tab.py` đã có expander `OCR & dịch text trong ảnh`.
- `word_backend.extract_images_from_docx()` đang quét `word/media/*`, trả list ảnh theo file media.
- `word_backend._ocr_single_image()` gọi Gemini Vision, trả JSON `{has_text, ocr, translation}`.
- `word_backend.ocr_and_translate_images()` chạy song song, skip ảnh dưới 5 KB như icon/decor.
- `word_backend.insert_ocr_captions_into_docx()` đã có option gần giống "đưa text dưới ảnh".
- Thiếu lớn nhất: chi phí OCR, chọn từng ảnh, sửa translation trước xuất, giữ/bỏ ảnh, mode overlay trực tiếp trên ảnh, test cho per-image behavior.

## Quyết Định Sản Phẩm

| Vấn đề | Quyết định |
|---|---|
| Màn sau OCR | Hiển thị từng ảnh bên trái, text OCR + text dịch bên phải |
| Chọn ảnh | Checkbox từng ảnh, thêm chọn tất cả / bỏ chọn tất cả |
| Ảnh không có text | Mặc định không chọn, show ở nhóm "Không phát hiện chữ" |
| Text dịch | Cho phép user sửa trước khi xuất DOCX |
| Chi phí | Show ước tính trước OCR và actual cost sau OCR |
| Output mode | Radio: `Dịch trực tiếp trên ảnh` / `Đưa text dưới ảnh` |
| Default output | `Đưa text dưới ảnh`, vì đã có code nền và ít rủi ro layout hơn |
| Overlay | Thực hiện theo phase riêng, có fallback caption nếu thiếu bbox |
| Domain style | OCR translation dùng cùng `source_lang`, `target_lang`, glossary và văn phong kỹ thuật thang máy/thang cuốn như dịch Word |

## Chi Phí OCR

Nguồn giá cần theo Google AI pricing chính thức, kiểm tra lại trước khi code nếu Google đổi model/giá.

Snapshot 2026-05-26 cho model đang ưu tiên trong app:

| Model | Mode | Input | Output |
|---|---|---:|---:|
| `gemini-3.5-flash` | Standard Paid | `$1.50 / 1M tokens` | `$9.00 / 1M tokens` |
| `gemini-3.5-flash` | Batch | `$0.75 / 1M tokens` | `$4.50 / 1M tokens` |

Ghi chú token ảnh:

- Gemini tính token cho cả text, image và modality khác.
- Với image input: ảnh <= 384 px mỗi chiều tính 258 tokens; ảnh lớn hơn được chia/tile, mỗi tile 768x768 tính 258 tokens.
- Actual cost sau call phải đọc từ `response.usage_metadata.prompt_token_count` và `response.usage_metadata.candidates_token_count`.
- Công thức trong app phải dùng `ui_common.calc_cost(tok_in, tok_out)` để thống nhất với `config.py`.

UI cần có 2 lớp giá:

1. Ước tính trước OCR:
   - Tổng số ảnh sẽ quét.
   - Số ảnh bị skip vì quá nhỏ.
   - Ước tính input image tokens theo kích thước ảnh.
   - Ước tính output tokens theo ngưỡng mặc định, ví dụ 200-400 tokens/ảnh có text.
   - USD/VND ước tính.

2. Chi phí thực tế sau OCR:
   - Per-image: model, input tokens, output tokens, USD/VND.
   - Tổng batch: input tokens, output tokens, USD/VND.
   - Nếu API không trả usage metadata, show warning "Không lấy được usage metadata; chỉ có ước tính".

## Hướng Kỹ Thuật

### Backend OCR

- [ ] `B1` Tạo data class/helper cho OCR image occurrence.
  - File đề xuất: `word_backend.py` trước block OCR hoặc file mới `ocr_backend.py` nếu muốn tách.
  - Không làm refactor lớn ngoài OCR trong lần đầu.
  - Done note:

- [ ] `B2` Nâng `extract_images_from_docx()`.
  - Vẫn đọc bytes từ `word/media/*`.
  - Thêm parse `document.xml`, tables, headers/footers nếu đang support ảnh ở các part đó.
  - Map `r:embed` -> media target qua `.rels`.
  - Trả occurrence theo vị trí xuất hiện, không chỉ media filename.
  - Done note:

- [ ] `B3` Nâng prompt OCR.
  - Pass `source_lang` + `target_lang`.
  - Inject domain technical style nếu detect elevator/escalator.
  - Preserve numbers, units, standard codes, part numbers.
  - Return JSON strict:
    - `has_text`
    - `ocr`
    - `translation`
    - `regions`: list bbox + text pairs
    - `confidence`
  - Done note:

- [ ] `B4` Token/cost tracking.
  - Tạo helper `_usage_tokens(response) -> (input, output, total)`.
  - Reuse hoặc align với token tracking đang dùng cho Word chunks.
  - Mọi retry chỉ tính lần thành công; nếu muốn audit lỗi, ghi thêm `attempts`.
  - Done note:

- [ ] `B5` Caption insertion selectable.
  - Function signature đề xuất:
    `insert_ocr_captions_into_docx(docx_bytes, images, ocr_results, selected_ids, edited_translations=None, remove_original_ids=None)`.
  - Nếu `remove_original_ids` có ảnh, xóa drawing paragraph cẩn thận hoặc thay bằng caption-only paragraph.
  - Done note:

- [ ] `B6` Overlay renderer.
  - Function đề xuất:
    `render_translated_overlay(image_bytes, content_type, regions, edited_translation, options) -> (bytes, content_type)`.
  - Options: font size min/max, opacity, background fill, text color.
  - Cần test với ảnh ngang, ảnh dọc, ảnh nhỏ, text dài.
  - Done note:

- [ ] `B7` DOCX image replacement by occurrence.
  - Function đề xuất:
    `replace_docx_image_occurrences(docx_bytes, replacements_by_occurrence_id)`.
  - Nếu một media part được dùng nhiều occurrence, clone image part trước khi replace occurrence được chọn.
  - Done note:

### UI OCR

- [ ] `U1` Preflight panel trước OCR.
  - Show số ảnh, ảnh nhỏ sẽ skip, estimate chi phí.
  - Có checkbox xác nhận "Tôi đồng ý chạy OCR và phát sinh chi phí Gemini".
  - Disable button OCR nếu chưa xác nhận khi estimate > 0.
  - Done note:

- [ ] `U2` Review layout sau OCR.
  - Layout mỗi item: image preview trái, OCR/text dịch phải.
  - Checkbox `Đưa ảnh này vào file xuất`.
  - Checkbox `Giữ ảnh gốc` nếu mode cho phép.
  - `text_area` edit translation.
  - Show cost nhỏ theo ảnh.
  - Done note:

- [ ] `U3` Output mode selector.
  - `radio("Cách đưa OCR vào DOCX", ["Đưa text dưới ảnh", "Dịch trực tiếp trên ảnh"])`.
  - Mode caption dùng code hiện có sau refactor.
  - Mode overlay show warning nếu thiếu bbox hoặc thiếu Pillow.
  - Done note:

- [ ] `U4` Download output riêng cho OCR.
  - Tên file đề xuất:
    - `_translated_ocr_caption.docx`
    - `_translated_ocr_overlay.docx`
  - Ghi summary trước download: số ảnh chọn, số ảnh giữ, số ảnh bỏ, cost actual.
  - Done note:

### Test Và Verify

- [ ] `T1` Unit/smoke test cost helper OCR.
  - Fake response có `usage_metadata`.
  - Verify input/output tokens -> USD/VND đúng theo `calc_cost`.
  - Done note:

- [ ] `T2` Test caption mode chọn từng ảnh.
  - DOCX fixture có 2 ảnh có chữ.
  - Tick 1 ảnh -> chỉ có 1 caption.
  - Tick 0 ảnh -> DOCX unchanged.
  - Done note:

- [ ] `T3` Test same-media multiple occurrence.
  - Cùng một image file xuất hiện 2 lần.
  - Chọn occurrence thứ 2 -> occurrence thứ 1 không bị đổi.
  - Đây là test bắt buộc trước khi overlay merge.
  - Done note:

- [ ] `T4` Test overlay renderer.
  - Nếu dùng Pillow: render text lên bbox, output image mở được.
  - Text dài phải wrap, không crash.
  - Done note:

- [ ] `T5` Manual QA với file thật.
  - Ảnh technical diagram thang máy.
  - Ảnh bảng thông số.
  - Ảnh thang cuốn có label.
  - Ảnh logo/icon nhỏ.
  - DOCX có ảnh trong table.
  - Done note:

## Thứ Tự Commit Đề Xuất

1. `docs: add OCR development plan`
2. `refactor: model OCR image occurrences`
3. `feat: show OCR cost estimate and usage`
4. `feat: add OCR review selections`
5. `feat: make OCR caption insertion selectable`
6. `feat: add OCR overlay renderer`
7. `feat: replace selected DOCX image occurrences`
8. `test: cover OCR cost and image insertion modes`
9. `docs: update OCR workflow in README`

## Rủi Ro

- Overlay "đúng vị trí" phụ thuộc bbox. Gemini Vision có thể đọc chữ tốt nhưng bbox không phải lúc nào ổn định như OCR layout engine chuyên dụng.
- DOCX media part có thể được reuse nhiều nơi; replace theo filename sẽ làm đổi nhầm nhiều ảnh.
- Xóa ảnh gốc trong DOCX dễ làm vỡ layout nếu ảnh nằm trong paragraph/table complex.
- Thêm Pillow cần cập nhật `requirements.txt`; môi trường deploy phải install được.
- Ảnh scan mờ/nhỏ sẽ cần `media_resolution` cao hơn, kéo chi phí và latency tăng.
- Actual cost phải đọc sau call; preflight chỉ là ước tính.

## Definition Of Done

- User upload DOCX, bấm OCR, thấy từng ảnh cạnh text OCR và bản dịch.
- User tick được ảnh muốn đưa vào output, sửa text dịch trước khi xuất.
- App hiển thị ước tính chi phí trước OCR và chi phí thực tế sau OCR bằng USD/VND.
- Mode "Đưa text dưới ảnh" chỉ chèn caption cho ảnh được chọn.
- Mode "Dịch trực tiếp trên ảnh" thay đúng ảnh được chọn, không thay nhầm ảnh khác.
- Có test/smoke cho cost, selection, caption insertion, same-media occurrence.
- README hoặc plan được update sau khi code xong.

## Completion Log

- 2026-05-26: Created `OCR_DEV.md` trên branch `dev`. Đây mới là plan, chưa code OCR mới.
