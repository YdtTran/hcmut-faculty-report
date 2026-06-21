#!/usr/bin/env python3
import os
import sys
import logging
from pathlib import Path
from dotenv import load_dotenv

# Thử import google.generativeai, nếu chưa có thì hướng dẫn cài đặt
try:
    import google.generativeai as genai
    from PIL import Image
except ImportError:
    print("Vui lòng cài đặt google-generativeai và Pillow bằng lệnh:")
    print("  .venv/bin/pip install google-generativeai pillow python-dotenv")
    sys.exit(1)

# Load file .env nếu có
load_dotenv()

# Thiết lập log
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler("clean_markdown.log", encoding="utf-8")
    ]
)
logger = logging.getLogger(__name__)

def main():
    api_key = os.getenv("GEMINI_API_KEY")
    if not api_key:
        logger.error("Không tìm thấy GEMINI_API_KEY trong biến môi trường hoặc tệp .env.")
        logger.error("Vui lòng tạo tệp .env ở gốc dự án với nội dung: GEMINI_API_KEY=your_key_here")
        sys.exit(1)
        
    genai.configure(api_key=api_key)
    
    # Sử dụng gemini-1.5-flash hỗ trợ multimodal rất tốt
    model_name = 'gemini-1.5-flash'
    
    raw_dir = Path("raw_md")
    processed_dir = Path("processed_md")
    images_dir = Path("files_images")
    files_dir = Path("files")
    
    if not raw_dir.exists():
        logger.error(f"Thư mục '{raw_dir}' không tồn tại. Hãy chạy convert_to_raw_md.py trước.")
        sys.exit(1)
        
    processed_dir.mkdir(parents=True, exist_ok=True)
    
    raw_files = sorted([f for f in raw_dir.iterdir() if f.is_file() and f.suffix == '.md'])
    
    if not raw_files:
        logger.warning("Không tìm thấy file markdown nào trong raw_md/.")
        return
        
    logger.info(f"Bắt đầu làm sạch {len(raw_files)} file markdown.")
    
    for raw_path in raw_files:
        logger.info(f"Đang xử lý: {raw_path.name}")
        processed_path = processed_dir / raw_path.name
        
        # Đọc nội dung raw markdown
        with open(raw_path, "r", encoding="utf-8") as f:
            raw_content = f.read()
            
        # Xác định xem file gốc có phải PDF không
        pdf_file = files_dir / f"{raw_path.stem}.pdf"
        pdf_images_path = images_dir / raw_path.stem
        
        system_instruction = (
            "Bạn là một chuyên gia chuẩn hóa tài liệu văn bản tiếng Việt. Nhiệm vụ của bạn là làm sạch "
            "nội dung markdown thô (raw markdown) của tài liệu và định dạng lại cấu trúc sao cho "
            "bố cục và định dạng giống nhất với tài liệu gốc.\n"
            "Các quy tắc bắt buộc:\n"
            "1. GIỮ NGUYÊN HOÀN TOÀN nội dung chữ tiếng Việt gốc, không tự ý tóm tắt, viết lại, hoặc sửa từ ngữ.\n"
            "2. Chuẩn hóa cấu trúc markdown: các tiêu đề (sử dụng #, ##, ###), danh sách dấu đầu dòng, danh sách số.\n"
            "3. Chuyển đổi các bảng biểu thô thành định dạng bảng Markdown chuẩn (dùng | và ---). Đảm bảo các hàng, cột thẳng hàng, đẹp đẽ.\n"
            "4. Loại bỏ các ngắt dòng lỗi ở giữa câu do chuyển đổi định dạng gây ra, đảm bảo các đoạn văn liền mạch.\n"
            "5. Không thêm bất kỳ nội dung giải thích nào của bạn (như 'Dưới đây là...', 'Đây là file...'). Chỉ trả về duy nhất nội dung markdown sạch."
        )
        
        try:
            model = genai.GenerativeModel(
                model_name=model_name,
                system_instruction=system_instruction
            )
            
            # Nếu có ảnh chụp trang PDF, dùng chế độ Multimodal
            if pdf_file.exists() and pdf_images_path.exists():
                logger.info(f"  -> Phát hiện PDF gốc. Sử dụng ảnh chụp trang để đối chiếu layout.")
                image_files = sorted(
                    [img for img in pdf_images_path.iterdir() if img.is_file() and img.suffix.lower() == '.png'],
                    key=lambda x: int(x.stem.split('_')[-1]) if '_' in x.stem else 0
                )
                
                pil_images = []
                for img_path in image_files:
                    pil_images.append(Image.open(img_path))
                    
                logger.info(f"  -> Đã load {len(pil_images)} ảnh trang để gửi kèm.")
                
                prompt = [
                    f"Dưới đây là {len(pil_images)} ảnh chụp tương ứng với các trang của tài liệu gốc theo thứ tự. "
                    "Và đây là nội dung raw markdown chuyển đổi thô từ tài liệu này:\n\n"
                    "--- RAW MARKDOWN START ---\n"
                    f"{raw_content}\n"
                    "--- RAW MARKDOWN END ---\n\n"
                    "Hãy đối chiếu với ảnh chụp thực tế và sửa đổi, làm sạch raw markdown trên để tạo ra cleaned markdown có layout chuẩn xác nhất. "
                    "Hãy trả về kết quả markdown sạch."
                ]
                prompt.extend(pil_images)
                
                response = model.generate_content(prompt)
                
            else:
                # Đối với DOCX/XLSX không có ảnh chụp, chỉ làm sạch văn bản thuần
                logger.info(f"  -> Không dùng ảnh đối chiếu. Tiến hành làm sạch văn bản thuần.")
                prompt = (
                    "Dưới đây là nội dung raw markdown của tài liệu. Hãy làm sạch, chuẩn hóa ngắt dòng "
                    "và cấu trúc bảng biểu, giữ nguyên chính xác nội dung chữ:\n\n"
                    "--- RAW MARKDOWN START ---\n"
                    f"{raw_content}\n"
                    "--- RAW MARKDOWN END ---"
                )
                response = model.generate_content(prompt)
                
            # Lưu kết quả
            cleaned_text = response.text.strip()
            if cleaned_text.startswith("```markdown"):
                cleaned_text = cleaned_text[len("```markdown"):].strip()
            if cleaned_text.endswith("```"):
                cleaned_text = cleaned_text[:-3].strip()
                
            with open(processed_path, "w", encoding="utf-8") as f:
                f.write(cleaned_text)
                
            logger.info(f"  -> Đã lưu cleaned markdown vào: {processed_path.name}")
            
        except Exception as e:
            logger.error(f"Lỗi khi làm sạch file {raw_path.name}: {e}")
            
    logger.info("Hoàn tất quá trình làm sạch markdown.")

if __name__ == "__main__":
    main()
