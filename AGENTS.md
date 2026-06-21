# AGENTS Guidelines for This Repository

Ngôn ngữ giao tiếp: Tiếng Việt (dùng cho trả lời yêu cầu)
Ngôn ngữ tài liệu: Tiếng Việt
Tools language: Python, Markdown.

Môi trường: Python Virtual Environment, WSL.
> Load môi trường (markitdown) trước khi chạy:
```bash
source .venv/bin/activate
# Cài đặt các thư viện cần thiết nếu chưa có: pip install -r requirements.txt
```

**Important Rules**:
1. **Version control**: Luôn fetch và pull trước khi thực hiện modify code.
2. **Môi trường hoạt động**: Luôn làm việc và thực thi script trong môi trường venv `markitdown`.
3. **Cấu trúc mã nguồn**: Toàn bộ mã nguồn Python, script xử lý hoặc helper phải được tổ chức trong thư mục `src/` (hoặc các thư mục chức năng rõ ràng), không viết script chạy ad-hoc trực tiếp tại thư mục gốc của dự án.
4. **Cơ chế xử lý lỗi và Logging**: Do các tài liệu đầu vào (`.pdf`, `.xlsx`, `.docx`) của các khoa có thể khác nhau về cấu trúc hoặc bị lỗi định dạng, mã nguồn phải được cài đặt cơ chế try-catch chặt chẽ cho từng file đơn lẻ. Ghi nhận log cụ thể (lỗi ở file nào, dòng nào) thay vì để chương trình dừng đột ngột (crash).
5. **Quy tắc làm sạch tài liệu (Cleaned Markdown)**:
   * **Đối với tài liệu `.docx`**: Tuyệt đối không dùng list numbering tự động của markdown (như các dòng bắt đầu bằng `1. ` liên tục) vì parser sẽ tự reset số thứ tự khi gặp đoạn văn đứt quãng. Bắt buộc viết số thứ tự cứng cố định dạng in đậm (ví dụ: `**1. Mục tiêu**`, `**2. Tiến độ**`, `**5.1. Bối cảnh**`).
   * **Đối với tài liệu `.xlsx`/`.xls`**: Loại bỏ triệt để các ô rác chứa `NaN` và `Unnamed: X`. Thay thế các ký tự xuống dòng `\n` trong ô bảng thành thẻ `<br>` để tránh vỡ cấu trúc cột của bảng biểu Markdown.
   * **Đối với tài liệu `.pdf`**: Luôn sử dụng thư viện `pypdfium2` để render các trang PDF thành ảnh PNG lưu trong `files_images/`, sau đó dùng ảnh này để đối chiếu trực quan và khôi phục các cấu trúc bảng biểu, cột bị vỡ khi chuyển đổi.