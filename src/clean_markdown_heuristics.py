#!/usr/bin/env python3
import os
import re
from pathlib import Path

def clean_table_line(line):
    # Thay thế các ô NaN bằng ô trống trong bảng
    line = re.sub(r'\|\s*NaN\s*(?=\|)', '| ', line)
    
    # Thay thế "Unnamed: X" bằng ô trống
    line = re.sub(r'Unnamed: \d+', '', line)
    
    # Thay thế các ký tự xuống dòng vật lý \n trong ô của bảng thành thẻ <br> để tránh vỡ bảng
    line = line.replace('\\n', '<br>')
    
    return line

def clean_text_content(content):
    lines = content.split('\n')
    cleaned_lines = []
    
    for line in lines:
        # Nhận diện dòng bảng
        if line.strip().startswith('|'):
            cleaned_line = clean_table_line(line)
            cleaned_lines.append(cleaned_line)
        else:
            # Sửa các lỗi ngắt dòng vô lý trong văn bản thường
            cleaned_lines.append(line)
            
    return '\n'.join(cleaned_lines)

def main():
    raw_dir = Path("raw_md")
    processed_dir = Path("processed_md")
    processed_dir.mkdir(parents=True, exist_ok=True)
    
    raw_files = sorted([f for f in raw_dir.iterdir() if f.is_file() and f.suffix == '.md'])
    
    print(f"Tìm thấy {len(raw_files)} file raw markdown.")
    
    for raw_path in raw_files:
        print(f"Đang làm sạch thô: {raw_path.name}")
        with open(raw_path, "r", encoding="utf-8") as f:
            content = f.read()
            
        cleaned_content = clean_text_content(content)
        
        # Loại bỏ các dòng trống liên tiếp (nhiều hơn 2 dòng trống)
        cleaned_content = re.sub(r'\n{3,}', '\n\n', cleaned_content)
        
        processed_path = processed_dir / raw_path.name
        with open(processed_path, "w", encoding="utf-8") as f:
            f.write(cleaned_content)
            
        print(f"  -> Đã lưu cleaned markdown sơ bộ: {processed_path.name}")

if __name__ == "__main__":
    main()
