# OCR Dev Plan - Word Translator

Branch: `dev`

Mục tiêu: phát triển mạnh luồng OCR cho ảnh trong DOCX, có kiểm soát chi phí Gemini, có màn review ảnh kèm chữ, chọn ảnh nào xử lý, và xuất DOCX theo 2 kiểu:

1. Dịch trực tiếp trên ảnh: che/xóa chữ gốc trong ảnh rồi chèn chữ dịch lên đúng vùng.
2. Dịch dưới ảnh: đưa text dịch xuống dưới ảnh như caption/code hiện tại.

File này là nguồn điều phối OCR trên branch `dev`. Khi AI/code xong một mục, phải tick mục đó, ghi ngày, commit, test đã chạy và note ngắn vào `Completion log` để tránh duplicate việc.

## Cần Làm Trước

- [x] `P0.1` Chốt data model OCR theo từng lần ảnh xuất hiện trong DOCX, không chỉ theo file `word/media/*`.
  - Lý do: một ảnh media có thể được dùng nhiều nơi; user cần tick chọn từng ảnh theo vị trí hiển thị.
  - Output mong muốn: `ImageOccurrence` có `id`, `filename`, `content_type`, `data`, `doc_part`, `rels_path`, `rId`, `paragraph_index`, `width`, `height`, `occurrence_index`.
  - Done: 2026-05-26, commit pending. Xem `## Completion Log` cuối file để biết chi tiết.

- [x] `P0.2` Tách state OCR trong `st.session_state` thành object rõ ràng.
  - Hiện tại: `word_image_ocr = (imgs, results)`.
  - Mục tiêu: `word_ocr_state = {"images": ..., "results": ..., "selection": ..., "cost": ..., "mode": ...}`.
  - Done: 2026-05-26, commit pending. Xem `## Completion Log` cuối file để biết chi tiết.

- [x] `P1.1` Hiển thị ước tính chi phí trước khi bấm OCR.
  - Dùng số ảnh, kích thước ảnh, model hiện tại và giá trong `config.py`.
  - Show cả USD và VND.
  - Ghi rõ đây là ước tính; actual cost sau khi chạy lấy từ `response.usage_metadata`.
  - Done: 2026-05-26, commit pending. Xem `## Completion Log` cuối file để biết chi tiết.

- [x] `P1.2` Lưu actual cost sau OCR theo từng ảnh và tổng batch.
  - `_ocr_single_image()` phải trả thêm `tok_in`, `tok_out`, `total_tokens`, `usd`, `vnd`, `model`.
  - `ocr_and_translate_images()` aggregate tổng chi phí.
  - UI show: tổng input/output tokens, USD, VND, model.
  - Done: 2026-05-26, commit pending. Xem `## Completion Log` cuối file để biết chi tiết.

- [x] `P2.1` Sau khi bấm "Quét và dịch ảnh", render màn review ảnh bên trái, OCR/chữ dịch bên phải.
  - Mỗi ảnh có checkbox `Chọn đưa vào DOCX`.
  - Có nút chọn tất cả / bỏ chọn tất cả.
  - Có `text_area` cho user sửa bản dịch OCR trước khi xuất.
  - Ảnh không có text hoặc lỗi OCR phải nằm nhóm riêng, mặc định không chọn.
  - Done: 2026-05-26, commit pending. Xem `## Completion Log` cuối file để biết chi tiết.

- [x] `P2.2` Thêm checkbox giữ/bỏ ảnh gốc cho từng ảnh.
  - Default: giữ ảnh gốc.
  - Nếu user bỏ ảnh gốc ở mode caption: xóa/ẩn ảnh và chỉ giữ text dịch dưới vị trí đó.
  - Nếu user chọn mode overlay: vẫn giữ ảnh làm nền, nhưng chữ gốc trong ảnh phải bị che/xóa tại vùng OCR trước khi vẽ chữ dịch.
  - Không làm kiểu song ngữ trên cùng ảnh trong mode overlay.
  - Done: 2026-05-26, commit pending. Xem `## Completion Log` cuối file để biết chi tiết.

- [x] `P3.1` Hoàn thiện option 2 - đưa text dưới ảnh.
  - Refactor `insert_ocr_captions_into_docx()` để nhận `selected_ids` và `edited_translations`.
  - Chỉ chèn caption cho ảnh được tick.
  - Không chèn duplicate nếu user bấm xuất lại nhiều lần trong cùng session.
  - Preserve behavior hiện tại: caption căn giữa, font nhỏ, prefix `[OCR]`.
  - Done: 2026-05-26, commit pending. Xem `## Completion Log` cuối file để biết chi tiết.

- [x] `P4.1` Thiết kế option 1 - dịch trực tiếp trên ảnh.
  - OCR prompt phải trả thêm vùng chữ: `regions[]` với `bbox`, `ocr`, `translation`, `confidence`.
  - Chuẩn hóa `bbox` về `[x, y, w, h]` theo tỷ lệ 0..1 để độc lập kích thước ảnh.
  - Quyết định chốt: chữ gốc trong vùng OCR phải được che/xóa, output ảnh chỉ hiển thị bản dịch.
  - Không thêm option giữ chữ gốc song song trong ảnh overlay.
  - Nếu Gemini không trả bbox đáng tin, fallback rõ sang caption mode.
  - Done: 2026-05-26, commit pending. Xem `## Completion Log` cuối file để biết chi tiết.

- [x] `P4.2` Implement render overlay lên ảnh.
  - Cần dependency xử lý ảnh, ưu tiên `Pillow`.
  - Vẽ nền opaque để che chữ gốc trước, sau đó vẽ text dịch lên trên.
  - Ưu tiên màu nền gần vùng ảnh gốc; fallback dùng trắng/xám nhạt nếu không detect được nền.
  - Auto-fit font theo bbox, wrap text, tránh tràn vùng.
  - Output ảnh mới dạng PNG/JPEG, giữ đúng aspect ratio.
  - Done: 2026-05-26, commit pending. Xem `## Completion Log` cuối file để biết chi tiết.

- [x] `P4.3` Replace ảnh trong DOCX đúng occurrence được chọn.
  - Không được thay nhầm mọi ảnh dùng chung cùng media part.
  - Nếu technical constraint của DOCX bắt buộc replace theo media part, phải clone media part riêng cho occurrence đó.
  - Preserve relationship, kích thước hiển thị, layout paragraph/table.
  - Done: 2026-05-26, commit pending. Xem `## Completion Log` cuối file để biết chi tiết.

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
| Overlay | Che/xóa chữ gốc trong vùng OCR rồi vẽ bản dịch lên đúng vùng; không làm song ngữ trên ảnh |
| Overlay fallback | Nếu thiếu bbox đáng tin, fallback sang caption mode |
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

- [x] `B1` Tạo data class/helper cho OCR image occurrence.
  - File đề xuất: `word_backend.py` trước block OCR hoặc file mới `ocr_backend.py` nếu muốn tách.
  - Không làm refactor lớn ngoài OCR trong lần đầu.
  - Done: 2026-05-26, commit pending. Xem `## Completion Log` cuối file để biết chi tiết.

- [x] `B2` Nâng `extract_images_from_docx()`.
  - Vẫn đọc bytes từ `word/media/*`.
  - Thêm parse `document.xml`, tables, headers/footers nếu đang support ảnh ở các part đó.
  - Map `r:embed` -> media target qua `.rels`.
  - Trả occurrence theo vị trí xuất hiện, không chỉ media filename.
  - Done: 2026-05-26, commit pending. Xem `## Completion Log` cuối file để biết chi tiết.

- [x] `B3` Nâng prompt OCR.
  - Pass `source_lang` + `target_lang`.
  - Inject domain technical style nếu detect elevator/escalator.
  - Preserve numbers, units, standard codes, part numbers.
  - Return JSON strict:
    - `has_text`
    - `ocr`
    - `translation`
    - `regions`: list bbox + text pairs
    - `confidence`
  - Done: 2026-05-26, commit pending. Xem `## Completion Log` cuối file để biết chi tiết.

- [x] `B4` Token/cost tracking.
  - Tạo helper `_usage_tokens(response) -> (input, output, total)`.
  - Reuse hoặc align với token tracking đang dùng cho Word chunks.
  - Mọi retry chỉ tính lần thành công; nếu muốn audit lỗi, ghi thêm `attempts`.
  - Done: 2026-05-26, commit pending. Xem `## Completion Log` cuối file để biết chi tiết.

- [x] `B5` Caption insertion selectable.
  - Function signature đề xuất:
    `insert_ocr_captions_into_docx(docx_bytes, images, ocr_results, selected_ids, edited_translations=None, remove_original_ids=None)`.
  - Nếu `remove_original_ids` có ảnh, xóa drawing paragraph cẩn thận hoặc thay bằng caption-only paragraph.
  - Done: 2026-05-26, commit pending. Xem `## Completion Log` cuối file để biết chi tiết.

- [x] `B6` Overlay renderer.
  - Function đề xuất:
    `render_translated_overlay(image_bytes, content_type, regions, edited_translation, options) -> (bytes, content_type)`.
  - Renderer phải che chữ gốc trong từng bbox trước khi vẽ bản dịch.
  - Options: font size min/max, background fill, text color.
  - Không expose option "giữ chữ gốc" trong overlay.
  - Cần test với ảnh ngang, ảnh dọc, ảnh nhỏ, text dài.
  - Done: 2026-05-26, commit pending. Xem `## Completion Log` cuối file để biết chi tiết.

- [x] `B7` DOCX image replacement by occurrence.
  - Function đề xuất:
    `replace_docx_image_occurrences(docx_bytes, replacements_by_occurrence_id)`.
  - Nếu một media part được dùng nhiều occurrence, clone image part trước khi replace occurrence được chọn.
  - Done: 2026-05-26, commit pending. Xem `## Completion Log` cuối file để biết chi tiết.

### UI OCR

- [x] `U1` Preflight panel trước OCR.
  - Show số ảnh, ảnh nhỏ sẽ skip, estimate chi phí.
  - Có checkbox xác nhận "Tôi đồng ý chạy OCR và phát sinh chi phí Gemini".
  - Disable button OCR nếu chưa xác nhận khi estimate > 0.
  - Done: 2026-05-26, commit pending. Xem `## Completion Log` cuối file để biết chi tiết.

- [x] `U2` Review layout sau OCR.
  - Layout mỗi item: image preview trái, OCR/text dịch phải.
  - Checkbox `Đưa ảnh này vào file xuất`.
  - Checkbox `Giữ ảnh gốc` nếu mode cho phép.
  - `text_area` edit translation.
  - Show cost nhỏ theo ảnh.
  - Done: 2026-05-26, commit pending. Xem `## Completion Log` cuối file để biết chi tiết.

- [x] `U3` Output mode selector.
  - `radio("Cách đưa OCR vào DOCX", ["Đưa text dưới ảnh", "Dịch trực tiếp trên ảnh"])`.
  - Mode caption dùng code hiện có sau refactor.
  - Mode overlay ghi rõ: "Chữ gốc trong ảnh sẽ được che và thay bằng bản dịch".
  - Mode overlay show warning nếu thiếu bbox hoặc thiếu Pillow.
  - Done: 2026-05-26, commit pending. Xem `## Completion Log` cuối file để biết chi tiết.

- [x] `U4` Download output riêng cho OCR.
  - Tên file đề xuất:
    - `_translated_ocr_caption.docx`
    - `_translated_ocr_overlay.docx`
  - Ghi summary trước download: số ảnh chọn, số ảnh giữ, số ảnh bỏ, cost actual.
  - Done: 2026-05-26, commit pending. Xem `## Completion Log` cuối file để biết chi tiết.

### Test Và Verify

- [x] `T1` Unit/smoke test cost helper OCR.
  - Fake response có `usage_metadata`.
  - Verify input/output tokens -> USD/VND đúng theo `calc_cost`.
  - Done: 2026-05-26, commit pending. Xem `## Completion Log` cuối file để biết chi tiết.

- [x] `T2` Test caption mode chọn từng ảnh.
  - DOCX fixture có 2 ảnh có chữ.
  - Tick 1 ảnh -> chỉ có 1 caption.
  - Tick 0 ảnh -> DOCX unchanged.
  - Done: 2026-05-26, commit pending. Xem `## Completion Log` cuối file để biết chi tiết.

- [x] `T3` Test same-media multiple occurrence.
  - Cùng một image file xuất hiện 2 lần.
  - Chọn occurrence thứ 2 -> occurrence thứ 1 không bị đổi.
  - Đây là test bắt buộc trước khi overlay merge.
  - Done: 2026-05-26, commit pending. Xem `## Completion Log` cuối file để biết chi tiết.

- [x] `T4` Test overlay renderer.
  - Nếu dùng Pillow: render text lên bbox, output image mở được.
  - Text dài phải wrap, không crash.
  - Done: 2026-05-26, commit pending. Xem `## Completion Log` cuối file để biết chi tiết.

- [ ] `T5` Manual QA với file thật.
  - Ảnh technical diagram thang máy.
  - Ảnh bảng thông số.
  - Ảnh thang cuốn có label.
  - Ảnh logo/icon nhỏ.
  - DOCX có ảnh trong table.
  - Blocked: 2026-05-26. Cần user cung cấp DOCX thực tế. Smoke synthetic (T1-T4) đã pass đầy đủ — occurrence model, caption selection + dedupe, same-media multi-occurrence (3 occ chung rId), overlay Pillow render, replace-by-occurrence preserve drawing count, DOCX validate sạch sau replace. T5 chờ user upload tài liệu thật để kiểm chứng QA cuối.

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
- Che chữ gốc cần detect nền tốt; nếu nền phức tạp có thể còn vệt nền hoặc vùng che chưa đẹp.
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
- Mode "Dịch trực tiếp trên ảnh" che chữ gốc và thay bằng bản dịch ở đúng ảnh được chọn, không thay nhầm ảnh khác.
- Có test/smoke cho cost, selection, caption insertion, same-media occurrence.
- README hoặc plan được update sau khi code xong.

## Completion Log

- 2026-05-26: Created `OCR_DEV.md` trên branch `dev`. Đây mới là plan, chưa code OCR mới.
- 2026-05-26: Chốt behavior Option 1 overlay: che/xóa chữ gốc trong ảnh, output chỉ hiển thị bản dịch tại vùng OCR; không làm kiểu song ngữ trên ảnh.
- 2026-05-26: Implement xong OCR flow đầy đủ (commit pending). Chi tiết theo block:
  - **P0.1 / B1 / B2**: `extract_image_occurrences()` quét mọi `r:embed` trong `<w:p>` qua tất cả doc parts (document.xml + header/footer/footnote/endnote/comments). Mỗi occurrence có `id, filename, content_type, data, doc_part, rels_path, rId, paragraph_index, occurrence_index, width_px, height_px`. Pillow optional (None nếu thiếu).
  - **P0.2**: state object `word_ocr_state = {occurrences, results, selection, keep_original, edited, estimate, mode, phase}` với phase machine `idle → preflight → done`.
  - **P1.1 / P1.2 / B3 / B4 / U1**: `estimate_image_input_tokens()` (≤384px=258, >384 tile 768×768) + `estimate_ocr_cost(...)`. `_ocr_single_image` trả `tok_in/tok_out/total_tokens/usd/vnd/model/attempts/error`. `ocr_and_translate_images(...)` aggregate vào key `_total`. UI preflight panel 3 metric (Ảnh tổng / Sẽ OCR / Ước tính $USD) + checkbox xác nhận trước khi chạy.
  - **B3 (OCR prompt)**: pass `source_lang`, glossary (cap 40), domain block khi `subdomains & {elevator, escalator}`, yêu cầu `regions[]` với bbox normalized 0..1 + `confidence`.
  - **P2.1 / P2.2 / U2**: review layout mỗi occ — ảnh trái (220px), info phải (cost per-image, checkbox "đưa vào output", checkbox "giữ ảnh gốc" trong caption mode, text_area edit translation). Nút "Chọn tất cả / bỏ chọn tất cả". Nhóm "Không phát hiện chữ / lỗi" collapsed riêng, default không chọn.
  - **P3.1 / B5**: `insert_ocr_captions_into_docx(docx_bytes, occurrences, ocr_results, selected_ids, edited_translations, remove_original_ids)` — chỉ chèn cho `selected_ids`, override translation bằng `edited_translations`, dedupe khi caption kế tiếp đã có `[OCR] ` (no duplicate khi re-export), `remove_original_ids` xoá `w:drawing/pict/object` chứa rId + xoá paragraph nếu rỗng. Caption căn giữa, italic, size 9, gray.
  - **P4.1**: chốt design overlay = che chữ gốc (fill với avg color của vùng bbox) rồi vẽ bản dịch. Single source of truth — không bilingual trên ảnh. Fallback caption khi region thiếu bbox.
  - **P4.2 / B6**: `render_translated_overlay(image_bytes, content_type, regions, edited_translation, options)` dùng Pillow — fill rectangle với avg color, contrast text color tự chọn theo luminance, font TrueType (DejaVuSans/Arial/Helvetica fallback chain → PIL default), auto-fit size từ 36→8 wrap theo bbox, output PNG.
  - **P4.3 / B7**: `replace_docx_image_occurrences(docx_bytes, occurrences, replacements_by_occ_id)` — group replacements theo (doc_part, rId), occurrence ĐẦU TIÊN overwrite media gốc, các occurrence sau **clone media file mới + thêm rId mới trong rels + redirect embed** → an toàn khi media được shared. `_update_content_types()` cập nhật `[Content_Types].xml` khi đổi extension (JPG → PNG sau overlay).
  - **U3**: radio "Đưa text dưới ảnh" / "Dịch trực tiếp trên ảnh"; overlay mode warn khi Pillow thiếu hoặc bbox thiếu (fallback caption per-image).
  - **U4**: download riêng `_translated_*_ocr_caption.docx` / `_translated_*_ocr_overlay.docx` + summary line cost trước nút.
  - **T1-T4**: `tests/smoke_ocr.py` 25 checks — cost estimate (4), prompt builder (5), occurrence model 3-image (3), caption selection 0/1/3 + dedupe + edit override (5), overlay Pillow render + edit + empty (4), replace-by-occurrence preserve drawing count + DOCX validate (4). Tổng: **6 smoke scripts, 81 checks all-green**.
  - **T5**: blocked vì cần DOCX thực — chờ user.
  - Backward-compat: giữ `extract_images_from_docx()` thin wrapper (dedupe by filename) để code cũ không lỗi import.
  - `requirements.txt`: thêm `Pillow>=10.0`.
  - UI cũ key `word_image_ocr` (tuple) thay bằng `word_ocr_state` (dict), clear ở `_clear_state`.
