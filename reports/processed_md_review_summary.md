# Tổng hợp rà soát processed_md

Thời điểm rà soát: 2026-06-21

## Phạm vi

- Đã đối chiếu 12 file trong `files/` với 12 file Markdown tương ứng trong `processed_md/`.
- Đã chạy render PDF bằng `pypdfium2` qua `src/capture_pdf_pages.py`.
- Ảnh PDF đã lưu tại `files_images/`: 3 PDF, tổng 12 trang PNG.
- Không có file `.csv` trong `files/`, nên không có mẫu CSV để kiểm tra.
- Báo cáo tự động chi tiết: `reports/processed_md_audit.md` và `reports/processed_md_audit.json`.

## Kết luận nhanh

- Tất cả file gốc đều có file Markdown tương ứng.
- Không phát hiện `NaN` hoặc `Unnamed: X` trong `processed_md/`.
- Các file PDF đã được render đủ số trang; kiểm tra thủ công các bảng chính cho thấy Markdown giữ được hàng/cột và giá trị chính.
- Các file Excel lớn nhìn chung giữ bảng và chuyển xuống dòng trong ô sang `<br>`.
- Các lỗi fidelity/format đã xác nhận ở lượt rà soát trước hiện đã được khắc phục trong `processed_md/`.

## Lỗi đã sửa (Đã khắc phục ngày 2026-06-21)

1. `processed_md/Kế hoạch thực hiện hoạt động NCKH SV năm 2026_K. KH-KTMT.md`

- [x] Đã đổi numbering tự động của Markdown ở các dòng 20-22, 56-59, 67 thành số cứng in đậm (ví dụ: `**1.**`, `**2.**`, ...).

2. `processed_md/Bảng tổng hợp đề xuất nckh sv_KHUD_2026_V3 (1).md`

- [x] Đã loại bỏ dấu `'` thừa ở đầu ô ở dòng 80, chỉ giữ lại `- Xác định...`.

3. `processed_md/BM Ke hoach NCKH SV 2026_KHUD_V3 (1).md`

- [x] Đã sửa chữ `đều` thành `đề` tại dòng 41 để khớp chính xác với bản DOCX gốc.

4. `processed_md/KTGT_KE HOACH TRIEN KHAI DE TAI NCKH SINH VIEN NAM 2026 LTHien revised.md`

- [x] Đã khôi phục header cột tại dòng 15 từ `Chương trình` về `Sinh viên thuộc chương trình` theo đúng DOCX gốc.

## Ghi chú kiểm tra PDF

- `20260615 - Kế hoạch triển khai đề tài NCKH SV.pdf`: ảnh trang 1 có bảng 3 cột; Markdown giữ đủ các hàng thời gian và bảng kinh phí.
- `To trinh_NCKHSV_Khoa KTHH_2026.signed.pdf`: ảnh trang 4 có bảng biểu mẫu và bảng kinh phí; Markdown giữ đủ mã biểu mẫu, hạng mục và giá trị chính. Dòng `~ 61% tổng KP` trong Markdown khớp ảnh nguồn dù giá trị này có vẻ bất thường về mặt tính toán.
- `2026.06.05_To trinh_NCKHSV_KhoaCoKhi_2026_v1.signed.pdf`: bảng biểu mẫu và bảng kinh phí trong Markdown giữ cấu trúc, giá trị chính, và phần ghi chú.

## Ghi chú về báo cáo tự động

`reports/processed_md_audit.md` báo nhiều `major` do kiểm tra exact-match trên cả ô/page text dài. Khi rà lại thủ công, nhiều cảnh báo là false positive vì Markdown đã tách ô nhiều dòng bằng `<br>` hoặc chuyển bảng sang Markdown hợp lệ. Không nên coi toàn bộ `major` trong báo cáo tự động là lỗi cần sửa ngay; các lỗi xác nhận đã được ghi ở mục "Lỗi đã sửa".
