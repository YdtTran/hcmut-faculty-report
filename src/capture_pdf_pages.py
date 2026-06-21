#!/usr/bin/env python3
import os
import sys
import logging
from pathlib import Path
import pypdfium2 as pdfium

# Thiết lập log
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler("capture_pdf.log", encoding="utf-8")
    ]
)
logger = logging.getLogger(__name__)

def main():
    input_dir = Path("files")
    output_base_dir = Path("files_images")
    
    if not input_dir.exists():
        logger.error(f"Thư mục '{input_dir}' không tồn tại.")
        sys.exit(1)
        
    output_base_dir.mkdir(parents=True, exist_ok=True)
    
    # Tìm các file PDF
    pdf_files = sorted([f for f in input_dir.iterdir() if f.is_file() and f.suffix.lower() == '.pdf'])
    
    if not pdf_files:
        logger.warning("Không tìm thấy file PDF nào trong thư mục files/.")
        return
        
    logger.info(f"Tìm thấy {len(pdf_files)} file PDF cần chụp ảnh các trang.")
    
    for pdf_path in pdf_files:
        logger.info(f"Đang xử lý file PDF: {pdf_path.name}")
        # Tạo thư mục riêng cho từng file PDF để chứa ảnh các trang
        pdf_image_dir = output_base_dir / pdf_path.stem
        pdf_image_dir.mkdir(parents=True, exist_ok=True)
        
        try:
            doc = pdfium.PdfDocument(str(pdf_path))
            num_pages = len(doc)
            logger.info(f"File '{pdf_path.name}' có {num_pages} trang.")
            
            for i, page in enumerate(doc):
                # Render trang ở độ phân giải tốt (scale=2.0)
                image = page.render(scale=2.0).to_pil()
                image_path = pdf_image_dir / f"page_{i + 1}.png"
                image.save(image_path, "PNG")
                logger.info(f"  -> Lưu trang {i + 1}/{num_pages} vào {image_path.relative_to(output_base_dir)}")
                
            logger.info(f"Hoàn thành chụp ảnh file PDF: {pdf_path.name}")
            
        except Exception as e:
            logger.error(f"Lỗi khi xử lý file PDF {pdf_path.name}: {e}")
            
    logger.info("Hoàn tất quá trình chụp ảnh các trang PDF.")

if __name__ == "__main__":
    main()
