# Architecture

Dự án thực hiện tổng hợp báo cáo từ các file tài liệu đơn lẻ thành một tài liệu tổng hợp duy nhất thông qua pipeline 4 bước dưới đây:

## Pipeline xử lý dữ liệu

1. **Convert sang Markdown (Raw Markdown)**
   * **Nội dung**: Chuyển đổi các định dạng tài liệu gốc từ các khoa như `.pdf`, `.xlsx`, `.docx` (nằm trong thư mục `files/`) sang dạng raw markdown để dễ dàng trích xuất thông tin.
   * **Công cụ**: Sử dụng thư viện `markitdown`.
   * **Đường dẫn lưu trữ**: `./raw_md/` (Tên file markdown trùng với tên file gốc ban đầu).

2. **Clean Markdown (Processed Markdown)**
   * **Nội dung**: Lọc bỏ các thông tin thừa, sửa lỗi định dạng, chuẩn hóa văn bản từ raw markdown để tạo ra bản markdown sạch.
   * **Công cụ**: Sử dụng mã nguồn Python (xử lý regex hoặc LLM API thích hợp).
   * **Đường dẫn lưu trữ**: `./processed_md/`

3. **Align Layout (Final Markdown)**
   * **Nội dung**: Sắp xếp, điều chỉnh lại cấu trúc cleaned markdown để đảm bảo khớp với bố cục ban đầu của các file gốc (bao gồm bảng biểu, các mục phân cấp lớn nhỏ).
   * **Đường dẫn lưu trữ**: `./final_md/`

4. **Đóng gói & Xuất bản (Report Export)**
   * **Nội dung**: Tổng hợp các file từ `./final_md/` thành 1 bộ tài liệu tổng hợp duy nhất bao gồm 4 định dạng bắt buộc dưới đây:
     * **`.docx`**: File báo cáo chính, được định dạng đẹp và chuyên nghiệp để in ấn, trình ký.
     * **`.md`**: File markdown tổng hợp sạch sẽ, dùng để feed vào các mô hình AI khác hỗ trợ hỏi đáp (RAG, Chatbot).
     * **`.html`**: File dùng để hiển thị và xem nhanh trực quan trên trình duyệt.
     * **`.tex`**: File mã nguồn LaTeX giúp visualize bằng code trực quan, hỗ trợ tùy chỉnh định dạng nâng cao và biên dịch chất lượng cao.
   * **Công cụ đóng gói**: Sử dụng Python script (`python-docx`, `pypandoc`) hoặc tích hợp công cụ `pandoc` trên hệ thống.
   * **Đường dẫn lưu trữ**: `report/` (Ví dụ: `report/tai_lieu_tong_hop_11_khoa.docx`, `report/tai_lieu_tong_hop_11_khoa.md`, v.v.).
