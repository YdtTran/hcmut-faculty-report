#!/usr/bin/env python3
import os
import sys
import logging
from pathlib import Path
from markitdown import MarkItDown

# Thiết lập log
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler("convert_raw_md.log", encoding="utf-8")
    ]
)
logger = logging.getLogger(__name__)

def main():
    input_dir = Path("files")
    output_dir = Path("raw_md")
    
    if not input_dir.exists():
        logger.error(f"Thư mục đầu vào '{input_dir}' không tồn tại.")
        sys.exit(1)
        
    output_dir.mkdir(parents=True, exist_ok=True)
    
    # Danh sách các định dạng được hỗ trợ
    supported_extensions = {'.pdf', '.docx', '.xlsx', '.xls'}
    
    # Lấy danh sách file và sắp xếp để xử lý tuần tự
    files_to_convert = sorted(
        [f for f in input_dir.iterdir() if f.is_file() and f.suffix.lower() in supported_extensions]
    )
    
    if not files_to_convert:
        logger.warning(f"Không tìm thấy file tài liệu nào trong thư mục '{input_dir}'.")
        return
        
    logger.info(f"Tìm thấy {len(files_to_convert)} file tài liệu cần chuyển đổi.")
    
    # Khởi tạo MarkItDown
    try:
        md = MarkItDown()
    except Exception as e:
        logger.error(f"Không thể khởi tạo MarkItDown: {e}")
        sys.exit(1)
        
    success_count = 0
    fail_count = 0
    
    for file_path in files_to_convert:
        logger.info(f"Đang xử lý file: {file_path.name}")
        out_file_path = output_dir / f"{file_path.stem}.md"
        
        try:
            # Chuyển đổi file
            result = md.convert(str(file_path))
            
            # Ghi nội dung ra file markdown
            with open(out_file_path, "w", encoding="utf-8") as f:
                f.write(result.text_content)
                
            logger.info(f"Chuyển đổi thành công: {file_path.name} -> {out_file_path.name}")
            success_count += 1
            
        except Exception as e:
            logger.error(f"Lỗi khi chuyển đổi file {file_path.name}: {e}")
            fail_count += 1
            
    logger.info("=" * 50)
    logger.info("KẾT QUẢ CHUYỂN ĐỔI:")
    logger.info(f"Tổng số file: {len(files_to_convert)}")
    logger.info(f"Thành công: {success_count}")
    logger.info(f"Thất bại: {fail_count}")
    logger.info("=" * 50)

if __name__ == "__main__":
    main()
