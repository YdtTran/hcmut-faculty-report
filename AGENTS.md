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