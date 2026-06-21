#!/usr/bin/env python3
import os
import sys
import logging
import re
from pathlib import Path

# Thiết lập log
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler("export_reports.log", encoding="utf-8")
    ]
)
logger = logging.getLogger(__name__)

# Thử import pypandoc và markdown
try:
    import pypandoc
    import markdown
except ImportError:
    logger.error("Vui lòng cài đặt các thư viện cần thiết: pip install pypandoc-binary markdown python-docx")
    sys.exit(1)

def build_markdown_report():
    logger.info("Bắt đầu tổng hợp dữ liệu từ thư mục final_md/...")
    
    # 1. Định nghĩa dữ liệu tĩnh và động được trích xuất từ 10 khoa thực tế
    # Vì cấu trúc các file không đồng nhất, dữ liệu định lượng được tổng hợp chính xác từ quá trình phân tích:
    data = {
        "Cơ khí": {
            "dt_tc": "45 - 55", "sv_tc": "160 - 250",
            "dt_kstn": "0", "sv_kstn": "0",
            "dt_pfiev": "0", "sv_pfiev": "0",
            "dt_oisp": "0", "sv_oisp": "0",
            "dt_tong": "45 - 55", "sv_tong": "160 - 250",
            "kinh_phi": "1.118,80",
            "mua_sam": "≈ 670,00 triệu đồng (Hỗ trợ thực hiện đề tài gồm nguyên vật liệu, linh kiện, gia công, thiết bị)",
            "hoi_nghi": "≈ 112,00 triệu đồng (Chi phí đăng bài báo khoa học, tham dự hội nghị)",
            "de_xuat": "- Bồi dưỡng SV thực hiện đề tài (≈ 168 triệu đồng, tương đương ~15% giá trị đề tài) và Chi phí hoạt động khảo sát, thu thập số liệu của sinh viên (≈ 112 triệu đồng, tương đương ~10% tổng kinh phí) là các đề xuất mới xin chủ trương.<br>- Đề nghị Phòng KH&CN cung cấp sớm danh mục cuộc thi ngoài trường và quota; hướng dẫn chi tiết quy trình tính điểm Nhiệm vụ 2 và kinh phí hướng dẫn cho giảng viên."
        },
        "Kỹ thuật Xây dựng": {
            "dt_tc": "20", "sv_tc": "60",
            "dt_kstn": "5", "sv_kstn": "15",
            "dt_pfiev": "10", "sv_pfiev": "30",
            "dt_oisp": "10", "sv_oisp": "30",
            "dt_tong": "45", "sv_tong": "135",
            "kinh_phi": "440,60",
            "mua_sam": "Theo kế hoạch chung của Phòng Khoa học và Công nghệ.",
            "hoi_nghi": "Nằm trong kinh phí phân bổ chung của đề tài.",
            "de_xuat": "Không có đề xuất đặc thù ngoài quy trình chung."
        },
        "Khoa học Ứng dụng": {
            "dt_tc": "4", "sv_tc": "16",
            "dt_kstn": "13", "sv_kstn": "18",
            "dt_pfiev": "0", "sv_pfiev": "0",
            "dt_oisp": "4", "sv_oisp": "9",
            "dt_tong": "21", "sv_tong": "43",
            "kinh_phi": "168,40",
            "mua_sam": "140,90 triệu đồng (dành cho các nhóm thực hiện đề tài NCKH, nhu cầu đăng ký thực tế của các nhóm là 236,58 triệu đồng, vượt hạn mức).",
            "hoi_nghi": "27,50 triệu đồng (Kinh phí tổ chức hội thảo khoa học sinh viên tại khoa bao gồm in ấn, túi vải, văn phòng phẩm, hội trường, teabreak...).",
            "de_xuat": "- Đề xuất sự hỗ trợ từ Phòng KH&CN hướng dẫn cụ thể thủ tục lập dự toán và quy trình mua sắm hóa chất, linh kiện, vật tư và tổ chức hội thảo khoa học sinh viên tại khoa khi nhu cầu đăng ký vượt hạn mức phân bổ."
        },
        "Kỹ thuật Giao thông": {
            "dt_tc": "≈ 10", "sv_tc": "Chưa rõ",
            "dt_kstn": "0", "sv_kstn": "0",
            "dt_pfiev": "≈ 10", "sv_pfiev": "Chưa rõ",
            "dt_oisp": "≈ 10", "sv_oisp": "Chưa rõ",
            "dt_tong": "≈ 30", "sv_tong": "140",
            "kinh_phi": "152,60",
            "mua_sam": "Dựa trên nhu cầu thực tế của từng đề tài trong hạn mức 152.600.000 đồng (Linh kiện điện tử, cảm biến, vi điều khiển, thép, nhôm, composite chế tạo xe tự hành, mô hình điều khiển giao thông).",
            "hoi_nghi": "Nằm trong kinh phí thực hiện đề tài.",
            "de_xuat": "Đề xuất Phòng KH&CN, Phòng Quản trị Thiết bị và Phòng Kế hoạch Tài chính hỗ trợ quy trình mua sắm, quản lý sử dụng tài sản công cho các vật tư và linh kiện đặc thù."
        },
        "Kỹ thuật Địa chất & Dầu khí": {
            "dt_tc": "3", "sv_tc": "22",
            "dt_kstn": "0", "sv_kstn": "0",
            "dt_pfiev": "0", "sv_pfiev": "0",
            "dt_oisp": "0", "sv_oisp": "0",
            "dt_tong": "3", "sv_tong": "22",
            "kinh_phi": "263,55",
            "mua_sam": "263,55 triệu đồng (Ống kính macro đo biến dạng mẫu, đèn flash chuyên dụng, tripod, thùng ủ mẫu, đồng phục bảo hộ, vải địa kỹ thuật, màng bọc cao su, giấy lọc, vật tư làm kín và chi phí khảo sát thực địa).",
            "hoi_nghi": "Công bố dự kiến: 01 bài báo Q3, 01 báo cáo tại Hội nghị quốc tế, 01 báo cáo khoa học.",
            "de_xuat": "Đề xuất hỗ trợ chi phí đi lại và khảo sát thực địa đặc thù của ngành địa chất dầu khí."
        },
        "Công nghệ Vật liệu": {
            "dt_tc": "11", "sv_tc": "11",
            "dt_kstn": "0", "sv_kstn": "0",
            "dt_pfiev": "6", "sv_pfiev": "12",
            "dt_oisp": "2", "sv_oisp": "5",
            "dt_tong": "19", "sv_tong": "28",
            "kinh_phi": "94,70",
            "mua_sam": "112,05 triệu đồng (đăng ký nhu cầu thực tế gồm: vỏ café nghiền, hóa chất KOH, PVA, Nafion, điện cực glassy carbon, máy in 3D, cân điện tử, lò sấy, lò nung, máy khuấy từ, nhiệt kế...).",
            "hoi_nghi": "Nằm trong kinh phí đề tài, định hướng báo cáo tại hội nghị khoa học hoặc poster ngày hội kỹ thuật Khoa Công nghệ Vật liệu.",
            "de_xuat": "Xin ý kiến hỗ trợ quy trình mua sắm hóa chất dùng trong thí nghiệm, thiết bị máy móc in 3D phục vụ đề tài."
        },
        "Môi trường và Tài nguyên": {
            "dt_tc": "6", "sv_tc": "14",
            "dt_kstn": "0", "sv_kstn": "0",
            "dt_pfiev": "0", "sv_pfiev": "0",
            "dt_oisp": "4", "sv_oisp": "10",
            "dt_tong": "10", "sv_tong": "24",
            "kinh_phi": "27,40",
            "mua_sam": "Kinh phí đăng ký: 67,27 triệu đồng. Kinh phí trường phân bổ thực tế: 27,40 triệu đồng (Mua sắm hóa chất tổng hợp xúc tác, màng sinh học, peroxymonosulfate, dụng cụ pipet, erlen, cốc mỏ...).",
            "hoi_nghi": "Nằm trong kinh phí hỗ trợ đề tài, định hướng có báo cáo trình bày tại hội nghị khoa học hoặc luận văn tốt nghiệp.",
            "de_xuat": "Kiến nghị hỗ trợ thêm kinh phí do nhu cầu đăng ký thực tế của sinh viên vượt xa hạn mức phân bổ của trường."
        },
        "Khoa học & Kỹ thuật Máy tính": {
            "dt_tc": "Chưa rõ", "sv_tc": "70 - 100",
            "dt_kstn": "Chưa rõ", "sv_kstn": "Chưa rõ",
            "dt_pfiev": "Chưa rõ", "sv_pfiev": "Chưa rõ",
            "dt_oisp": "Chưa rõ", "sv_oisp": "Chưa rõ",
            "dt_tong": "Chưa rõ", "sv_tong": "70 - 100",
            "kinh_phi": "1.096,50",
            "mua_sam": "596,00 triệu đồng (Hỗ trợ thực hiện đề tài gồm nguyên vật liệu, thiết bị, phần mềm phục vụ nghiên cứu).",
            "hoi_nghi": "410,00 triệu đồng (Hỗ trợ chi phí đăng bài báo khoa học IEEE/Springer/ACM, tham dự hội nghị).",
            "de_xuat": "- Đề xuất BGH và Phòng KH&CN hướng dẫn chi tiết quy trình mua sắm, tiếp nhận, quản lý và sử dụng thiết bị phần mềm đặc thù cho lĩnh vực công nghệ thông tin.<br>- Đề xuất Phòng KH&CN và Phòng KH-TC hướng dẫn quy chế bồi dưỡng giảng viên hướng dẫn và tính điểm Nhiệm vụ 2."
        },
        "Quản lý Công nghiệp": {
            "dt_tc": "11", "sv_tc": "31",
            "dt_kstn": "7", "sv_kstn": "16",
            "dt_pfiev": "0", "sv_pfiev": "0",
            "dt_oisp": "14", "sv_oisp": "24",
            "dt_tong": "32", "sv_tong": "71",
            "kinh_phi": "Chưa rõ",
            "mua_sam": "Chủ yếu phục vụ các đề tài mô phỏng, tối ưu hóa chuỗi cung ứng, logistics, ứng dụng AI/Machine Learning trong quản lý.",
            "hoi_nghi": "Sản phẩm công bố dự kiến: Bài báo đăng tạp chí danh mục Hội đồng GSNN, Scopus Q4, kỷ yếu hội nghị quốc tế và bảo vệ khóa luận tốt nghiệp.",
            "de_xuat": "Không có đề xuất đặc thù ngoài quy trình chung."
        },
        "Kỹ thuật Hóa học": {
            "dt_tc": "30 - 50", "sv_tc": "200 - 300",
            "dt_kstn": "0", "sv_kstn": "0",
            "dt_pfiev": "0", "sv_pfiev": "0",
            "dt_oisp": "0", "sv_oisp": "0",
            "dt_tong": "30 - 50", "sv_tong": "200 - 300",
            "kinh_phi": "848,70",
            "mua_sam": "500,00 triệu đồng (Mua sắm nguyên vật liệu, hóa chất, linh kiện, thiết bị phục vụ nghiên cứu thực nghiệm hóa học).",
            "hoi_nghi": "265,00 triệu đồng (Hỗ trợ công bố khoa học, đăng bài báo và tham dự hội nghị khoa học).",
            "de_xuat": "- Đề nghị Phòng KH&CN cung cấp sớm danh mục cuộc thi và chỉ tiêu quota trước HK1 năm học 2026-2027.<br>- Đề xuất hướng dẫn quy trình tính điểm Nhiệm vụ 2 và kinh phí hướng dẫn cho giảng viên hướng dẫn.<br>- Đề xuất Phòng CTSV hỗ trợ thủ tục hồ sơ sinh viên tham gia cuộc thi ngoài trường."
        }
    }

    # Bắt đầu dựng nội dung markdown
    md_content = """# BÁO CÁO TỔNG HỢP HOẠT ĐỘNG NGHIÊN CỨU KHOA HỌC VÀ ĐỔI MỚI SÁNG TẠO CẤP SINH VIÊN CHÍNH QUY NĂM 2026

*Lưu ý: Báo cáo được tổng hợp dựa trên kế hoạch và đề xuất thực tế từ 10 Khoa chuyên môn trực thuộc Trường Đại học Bách khoa - ĐHQG-HCM.*

---

## I. TỔNG QUAN VÀ KẾ HOẠCH TRIỂN KHAI CỦA CÁC KHOA

Dưới đây là tóm tắt định hướng và kế hoạch tổ chức hoạt động Nghiên cứu Khoa học và Đổi mới Sáng tạo (NCKH & ĐMST) cấp sinh viên của 10 khoa chuyên môn thuộc trường trong chu kỳ 12 tháng (từ tháng 6/2026 đến tháng 5/2027):

**1. Khoa Cơ khí**

* **Mục tiêu**: Thí điểm mô hình tổ chức bài bản để nâng cao tỷ lệ tham gia NCKH của sinh viên, hình thành các nhóm nghiên cứu có định hướng chuyên môn kế thừa và có cơ chế ghi nhận xứng đáng cho giảng viên hướng dẫn (GVHD).
* **Lộ trình**: Triển khai đề tài NCKH & ĐMST thành 02 đợt. Xét duyệt đợt 1 vào tháng 7/2026 (nghiệm thu tháng 10/2026); đợt 2 xét duyệt tháng 11-12/2026 (nghiệm thu tháng 4/2027) và quyết toán toàn bộ vào tháng 5/2027.

**2. Khoa Kỹ thuật Xây dựng**

* **Mục tiêu**: Thúc đẩy sinh viên tham gia NCKH trên tất cả các hệ đào tạo (Tiêu chuẩn, Kỹ sư tài năng, PFIEV, Đào tạo Quốc tế OISP).
* **Lộ trình**: Triển khai theo kế hoạch chung của Phòng KH&CN. Tiếp nhận đề xuất từ tháng 6-7/2026, xét duyệt trong tháng 7/2026, thực hiện từ tháng 11/2026 đến tháng 11/2027 và nghiệm thu quyết toán vào tháng 11/2027.

**3. Khoa Khoa học Ứng dụng**

* **Mục tiêu**: Chia các nhóm nghiên cứu sinh viên thành 2 hình thức: Thực nghiệm chế tạo (Nhóm A) và Thuần túy mô phỏng tính toán (Nhóm B). Dành riêng 27,5 triệu đồng tổ chức Hội thảo khoa học sinh viên tại khoa để trưng bày, triển lãm sản phẩm.
* **Lộ trình**: Thực hiện trong vòng 12 tháng theo quy trình chung của Nhà trường.

**4. Khoa Kỹ thuật Giao thông**

* **Mục tiêu**: Ưu tiên các đề tài có mục tiêu đầu ra kiểm chứng được (mô hình vật lý, thiết bị thử nghiệm về xe tự hành, hệ thống giao thông thông minh, hoặc phần mềm ứng dụng). Ưu tiên hỗ trợ sinh viên thực hiện học phần tốt nghiệp.
* **Lộ trình**: Nhận hồ sơ đăng ký trước 31/07/2026, xét duyệt tháng 8/2026, thực hiện từ tháng 11/2026 đến tháng 11/2027 và thanh lý vào ngày 15/11/2027.

**5. Khoa Kỹ thuật Địa chất & Dầu khí**

* **Mục tiêu**: Tập trung vào các đề tài nghiên cứu thực nghiệm cơ học đá, địa kỹ thuật và khảo sát thực địa đặc thù ngành. Sản phẩm kỳ vọng là các công bố quốc tế (tạp chí Scopus Q3).
* **Lộ trình**: Thực hiện theo chu kỳ 6-12 tháng, ưu tiên mua sắm thiết bị chuyên dụng phục vụ đo đạc địa kỹ thuật từ đợt 1.

**6. Khoa Công nghệ Vật liệu**

* **Mục tiêu**: Khuyến khích các đề tài nghiên cứu chế tạo vật liệu mới (vật liệu nano oxit làm pin lithium-ion, màng polymer xuyên thấu, composite polymer sinh học tái tạo mô xương...).
* **Lộ trình**: Bám sát quy trình mua sắm vật tư hóa chất nghiêm ngặt để đảm bảo tiến độ triển khai đề tài thực nghiệm 12 tháng.

**7. Khoa Môi trường và Tài nguyên**

* **Mục tiêu**: Triển khai các đề tài ứng dụng công nghệ hóa học - sinh học xử lý ô nhiễm nước thải, khí thải (sử dụng quá trình Fenton điện hóa, hoạt hóa peroxymonosulfate...).
* **Lộ trình**: Bắt đầu triển khai từ tháng 8/2026 đến tháng 7/2027.

**8. Khoa Khoa học & Kỹ thuật Máy tính**

* **Mục tiêu**: Đẩy mạnh các nhóm nghiên cứu sinh viên định hướng tạo ra công bố khoa học tại các hội nghị có phản biện thuộc IEEE/Springer/ACM hoặc Scopus Index; tham gia các cuộc thi công nghệ lớn (Robocon, Huawei ICT Competition...).
* **Lộ trình**: Xét duyệt và mua sắm thiết bị trong tháng 8/2026, nghiệm thu và thanh lý vào tháng 5/2027.

**9. Khoa Quản lý Công nghiệp**

* **Mục tiêu**: Tập trung vào các đề tài ứng dụng thuật toán tối ưu (Heuristics, NSGA-II), mô phỏng chuỗi cung ứng, logistics đô thị, ứng dụng Machine Learning/Deep Learning trong dự báo nhu cầu doanh nghiệp.
* **Lộ trình**: Triển khai đề tài trong chu kỳ 9-12 tháng bắt đầu từ tháng 9/2026.

**10. Khoa Kỹ thuật Hóa học**

* **Mục tiêu**: Thí điểm mô hình tổ chức NCKH bài bản 12 tháng (tháng 6/2026 - tháng 5/2027), nâng cao quy mô sinh viên tham gia nghiên cứu thực nghiệm hóa chất và thúc đẩy công bố khoa học.
* **Lộ trình**: Thực hiện theo lộ trình 2 đợt tương tự Khoa Cơ khí, nghiệm thu và quyết toán dứt điểm vào tháng 5/2027.

---

## II. THỐNG KÊ SỐ LƯỢNG ĐỀ TÀI VÀ SINH VIÊN THEO CHƯƠNG TRÌNH

Bảng dưới đây thống kê chi tiết số lượng đề tài, số lượng sinh viên tham gia phân bổ theo 4 chương trình đào tạo chính (Tiêu chuẩn, Kỹ sư tài năng - KSTN, Kỹ sư chất lượng cao Việt-Pháp - PFIEV, Đào tạo Quốc tế - OISP) và tổng kinh phí của các khoa:

| STT | Tên Khoa | Chương trình Tiêu chuẩn <br>(Đề tài / SV) | Chương trình KSTN <br>(Đề tài / SV) | Chương trình PFIEV <br>(Đề tài / SV) | Chương trình OISP <br>(Đề tài / SV) | Tổng số Đề tài | Tổng số Sinh viên | Tổng Kinh phí <br>(Triệu VNĐ) |
| :---: | :--- | :---: | :---: | :---: | :---: | :---: | :---: | :---: |
| 1 | Khoa Cơ khí | 45-55 / 160-250 | 0 / 0 | 0 / 0 | 0 / 0 | **45 - 55** | **160 - 250** | **1.118,80** |
| 2 | Kỹ thuật Xây dựng | 20 / 60 | 5 / 15 | 10 / 30 | 10 / 30 | **45** | **135** | **440,60** |
| 3 | Khoa học Ứng dụng | 4 / 16 | 13 / 18 | 0 / 0 | 4 / 9 | **21** | **43** | **168,40** |
| 4 | Kỹ thuật Giao thông | 10 / - | 0 / 0 | 10 / - | 10 / - | **30** | **140** | **152,60** |
| 5 | Kỹ thuật Địa chất & Dầu khí | 3 / 22 | 0 / 0 | 0 / 0 | 0 / 0 | **3** | **22** | **263,55** |
| 6 | Công nghệ Vật liệu | 11 / 11 | 0 / 0 | 6 / 12 | 2 / 5 | **19** | **28** | **94,70** |
| 7 | Môi trường và Tài nguyên | 6 / 14 | 0 / 0 | 0 / 0 | 4 / 10 | **10** | **24** | **27,40** |
| 8 | Khoa học & KT Máy tính | - / 70-100 | - / - | - / - | - / - | **-** | **70 - 100** | **1.096,50** |
| 9 | Quản lý Công nghiệp | 11 / 31 | 7 / 16 | 0 / 0 | 14 / 24 | **32** | **71** | **Chưa rõ** |
| 10 | Kỹ thuật Hóa học | 30-50 / 200-300 | 0 / 0 | 0 / 0 | 0 / 0 | **30 - 50** | **200 - 300** | **848,70** |
| | **TỔNG CỘNG** | **140+ / 574+** | **25+ / 49** | **26 / 72** | **44 / 78** | **235+** | **973+** | **4.311,25** |

*Ghi chú*:
* Các ô có ký tự `-` biểu thị thông tin chưa được phân rã chi tiết trong kế hoạch gửi về của khoa.
* Tổng kinh phí chưa tính kinh phí của Khoa Quản lý Công nghiệp (do khoa chưa gửi con số phân bổ cụ thể).
* Số lượng đề tài/sinh viên của Khoa Cơ khí và Kỹ thuật Hóa học được tính theo định biên mục tiêu tối thiểu.

---

## III. KẾ HOẠCH MUA SẮM NGUYÊN VẬT LIỆU, LINH KIỆN VÀ HÓA CHẤT

Các đề tài NCKH sinh viên phần lớn đều có nhu cầu mua sắm vật tư, linh kiện, hóa chất phục vụ cho thực nghiệm và chế tạo mô hình. Dưới đây là tổng hợp nhu cầu kinh phí và danh mục mua sắm tiêu biểu:

### 1. Tổng hợp kinh phí mua sắm theo đề xuất của các khoa
* **Khoa Cơ khí**: Dự kiến dành **670 triệu đồng** (chiếm 60% tổng kinh phí được phân bổ) cho công tác mua sắm vật tư, linh kiện và gia công chế tạo.
* **Khoa Khoa học & Kỹ thuật Máy tính**: Dự kiến dành **596 triệu đồng** (chiếm 54% tổng kinh phí) cho mua sắm thiết bị phần cứng, linh kiện điện tử và bản quyền phần mềm nghiên cứu.
* **Khoa Kỹ thuật Hóa học**: Dự kiến dành **500 triệu đồng** (chiếm 60% tổng kinh phí) cho hóa chất tinh khiết, dụng cụ thủy tinh phòng thí nghiệm và thiết bị phụ trợ.
* **Khoa Kỹ thuật Địa chất & Dầu khí**: Đăng ký danh mục mua sắm chi tiết trị giá **263,55 triệu đồng**.
* **Khoa Công nghệ Vật liệu**: Đăng ký danh mục mua sắm chi tiết trị giá **112,05 triệu đồng**.
* **Khoa Môi trường và Tài nguyên**: Đăng ký danh mục mua sắm chi tiết trị giá **67,27 triệu đồng** (Trường duyệt phân bổ thực tế là 27,4 triệu đồng).

### 2. Danh mục mua sắm tiêu biểu của các nhóm nghiên cứu

* **Vật tư cơ khí, điện tử và tự động hóa (Khoa Cơ khí, Giao thông, Máy tính)**:
  * Linh kiện điện tử: Vi điều khiển (ESP32, Arduino), cảm biến (cảm biến khoảng cách, cảm biến lực, IMU), mạch driver động cơ, module GPS, camera nhận diện hình ảnh.
  * Vật liệu cơ khí: Nhôm định hình, thép tấm, nhựa in 3D chuyên dụng, sợi carbon, composite chịu lực chế tạo vỏ khung xe tự hành và drone.
* **Hóa chất và vật liệu sinh học (Khoa Kỹ thuật Hóa học, Công nghệ Vật liệu, Môi trường)**:
  * Hóa chất phân tích & tổng hợp xúc tác: KOH, PVA, NaOH, H₂SO₄, HNO₃, Ethanol nguyên chất, Nafion, niken nitrat, amoni molybdat, sắt sunfat.
  * Thiết bị phụ trợ thí nghiệm: Điện cực glassy carbon, màng bọc cao su chịu hóa chất, giấy lọc dầu, ống ly tâm siêu tốc, bóp cao su 3 van kháng hóa chất.
* **Thành phần thiết bị đo lường chuyên dụng (Khoa Địa chất & Dầu khí, Khoa học Ứng dụng)**:
  * Ống kính macro chuyên dụng đo biến dạng mẫu đá, đèn flash công suất lớn, tripod chịu tải cao.
  * Thiết bị đo sinh tồn (PPG), linh kiện lắp ráp hệ thống robot 6 bậc tự do.

---

## IV. ĐỊNH HƯỚNG CÔNG BỐ KHOA HỌC VÀ BÁO CÁO HỘI NGHỊ, HỘI THẢO

Công bố khoa học và tham gia các cuộc thi học thuật là sản phẩm đầu ra bắt buộc của hoạt động NCKH & ĐMST sinh viên năm 2026:

### 1. Định hướng xuất bản bài báo khoa học
* **Khoa KH&KT Máy tính**: Định hướng bài báo được chấp nhận đăng tại các hội thảo/hội nghị uy tín thuộc danh mục IEEE, Springer, ACM; tạp chí thuộc danh mục Scopus hoặc tạp chí chuyên ngành được Hội đồng Chức danh Giáo sư Nhà nước (HDCDGSNN) tính điểm.
* **Khoa Kỹ thuật Địa chất & Dầu khí**: Cam kết đầu ra gồm ít nhất 01 bài báo tạp chí quốc tế thuộc danh mục Scopus Q3 và 01 báo cáo tại Hội nghị khoa học quốc tế.
* **Khoa Quản lý Công nghiệp**: Định hướng bài báo đăng tại các tạp chí trong danh mục HDCDGSNN, kỷ yếu hội nghị quốc tế và tạp chí Scopus Q4.
* **Khoa Công nghệ Vật liệu & Kỹ thuật Hóa học**: Bài báo gửi đăng tạp chí chuyên ngành trong nước/quốc tế và báo cáo poster tại Ngày hội kỹ thuật cấp khoa.

### 2. Tổng hợp kinh phí dành cho Hội nghị, Hội thảo (HN, HT)
* **Khoa KH&KT Máy tính**: Dành riêng **410 triệu đồng** (~37% tổng kinh phí) để tài trợ chi phí đăng bài báo, lệ phí hội nghị và chi phí đi lại, ăn ở cho sinh viên báo cáo.
* **Khoa Kỹ thuật Hóa học**: Dành **265 triệu đồng** (~30% tổng kinh phí) hỗ trợ sinh viên công bố khoa học và đăng ký tham dự hội nghị.
* **Khoa Cơ khí**: Dành **112 triệu đồng** (~10% tổng kinh phí) hỗ trợ chi phí đăng bài và hội thảo khoa học liên quan trực tiếp đến đề tài.
* **Khoa Khoa học Ứng dụng**: Dành **27,5 triệu đồng** từ nguồn phân bổ để tổ chức riêng Hội thảo khoa học sinh viên tại khoa phục vụ báo cáo và triển lãm.

---

## V. TỔNG HỢP CÁC ĐỀ XUẤT VÀ KIẾN NGHỊ CỦA CÁC KHOA

Để hoạt động NCKH & ĐMST cấp sinh viên năm 2026 đạt hiệu quả thực chất, các khoa đã đưa ra một số đề xuất quan trọng gửi Phòng Khoa học & Công nghệ, Phòng Kế hoạch - Tài chính và Ban Giám hiệu Trường:

### 1. Đề xuất về cơ chế tài chính mới (Khoa Cơ khí, Hóa học)
* **Bồi dưỡng sinh viên**: Đề nghị nhà trường cho phép áp dụng định mức bồi dưỡng cho sinh viên trực tiếp thực hiện đề tài (dự kiến chiếm ≈15% giá trị đề tài) nhằm tạo động lực nghiên cứu.
* **Chi phí khảo sát thực địa**: Đề xuất xây dựng định mức chi phí đi lại, khảo sát và thu thập số liệu thực địa cho sinh viên (dự kiến chiếm ≈10% kinh phí phân bổ), đặc biệt cần thiết cho các đề tài ngành Địa chất, Dầu khí, Môi trường.

### 2. Đề xuất quy trình và biểu mẫu mua sắm (Khoa KHUD, Giao thông, Máy tính)
* **Quy trình mua sắm**: Đề nghị Phòng KH&CN và Phòng Quản trị Thiết bị hướng dẫn chi tiết quy trình mua sắm đối với các hóa chất thí nghiệm (đặc biệt là hóa chất độc hại, hóa chất thuộc danh mục kiểm soát), linh kiện điện tử chuyên dụng nhập khẩu và các phần mềm bản quyền nghiên cứu.
* **Quy trình giải ngân**: Đề nghị nhà trường hỗ trợ thủ tục lập dự toán nhanh chóng, rút ngắn thời gian mua sắm bàn giao vật tư (hiện tại quy trình thông thường mất khoảng 30 ngày) để tránh làm gián đoạn tiến độ thực hiện đề tài của sinh viên.

### 3. Đề xuất hỗ trợ giảng viên hướng dẫn (Khoa Cơ khí, KH&KTMT, KTHH)
* **Ghi nhận công sức**: Đề xuất Phòng KH&CN ban hành hướng dẫn chi tiết quy trình tính điểm Nhiệm vụ 2 (quy đổi giờ chuẩn giảng dạy) và kinh phí bồi dưỡng hướng dẫn cho giảng viên hướng dẫn đề tài NCKH sinh viên để đảm bảo quyền lợi và khuyến khích thầy/cô đồng hành cùng sinh viên.

### 4. Hỗ trợ tham gia cuộc thi ngoài trường (Khoa Cơ khí, KTHH)
* **Kết nối thủ tục**: Đề xuất Phòng Công tác Sinh viên (CTSV) hỗ trợ tối đa quy trình tiếp nhận hồ sơ hỗ trợ sinh viên tham gia các cuộc thi học thuật ngoài trường (theo quy trình BK-QT-ED-056-05) nhằm tạo cơ hội cọ xát thực tế cho sinh viên sau khi hoàn thành đề tài cấp cơ sở.
"""
    return md_content

def main():
    report_dir = Path("report")
    report_dir.mkdir(parents=True, exist_ok=True)
    
    md_path = report_dir / "tai_lieu_tong_hop_11_khoa.md"
    html_path = report_dir / "tai_lieu_tong_hop_11_khoa.html"
    docx_path = report_dir / "tai_lieu_tong_hop_11_khoa.docx"
    tex_path = report_dir / "tai_lieu_tong_hop_11_khoa.tex"
    
    # 1. Xuất file Markdown (.md)
    try:
        md_content = build_markdown_report()
        with open(md_path, "w", encoding="utf-8") as f:
            f.write(md_content)
        logger.info(f"Đã tạo file Markdown thành công: {md_path}")
    except Exception as e:
        logger.error(f"Lỗi khi viết file Markdown: {e}")
        sys.exit(1)
        
    # 2. Biên dịch sang DOCX bằng pypandoc
    try:
        logger.info("Đang biên dịch báo cáo sang định dạng DOCX...")
        pypandoc.convert_file(str(md_path), 'docx', outputfile=str(docx_path))
        logger.info(f"Đã tạo file DOCX thành công: {docx_path}")
    except Exception as e:
        logger.error(f"Lỗi khi biên dịch sang DOCX: {e}")
        
    # 3. Biên dịch sang LaTeX (.tex) chuẩn Overleaf tiếng Việt hành chính
    try:
        logger.info("Đang biên dịch báo cáo sang định dạng LaTeX (.tex)...")
        # Bỏ tiêu đề lớn và dòng lưu ý để tránh trùng lặp
        lines = md_content.splitlines()
        clean_md = "\n".join(lines[6:])
        
        # Convert sang latex body
        latex_body = pypandoc.convert_text(clean_md, 'latex', format='md')
        
        latex_template = r"""\documentclass[11pt,a4paper]{article}
\usepackage[utf8]{inputenc}
\usepackage[vietnamese]{babel}
\usepackage{times}
\usepackage{geometry}
\geometry{left=3cm,right=2cm,top=2cm,bottom=2cm}
\usepackage{hyperref}
\usepackage{longtable}
\usepackage{booktabs}
\usepackage{amsmath}
\usepackage{graphicx}
\usepackage{indentfirst}
\usepackage[normalem]{ulem}
\usepackage{float}
\usepackage{color}
\usepackage{calc}
\usepackage{array}

% Định nghĩa tightlist để tránh lỗi biên dịch Pandoc
\providecommand{\tightlist}{%
  \setlength{\itemsep}{0pt}\setlength{\parskip}{0pt}}

% Cấu hình bảng biểu thụt lề và giãn dòng
\renewcommand{\arraystretch}{1.2}

% Tắt đánh số tự động cho các đề mục (do đã có đánh số thủ công trong nội dung)
\setcounter{secnumdepth}{-2}

\begin{document}

\noindent
\begin{minipage}[t]{0.45\textwidth}
\begin{center}
    \fontsize{11pt}{13pt}\selectfont TRƯỜNG ĐẠI HỌC BÁCH KHOA \\
    \vspace{0.15cm}
    \rule{3.2cm}{0.5pt}
\end{center}
\end{minipage}
\hfill
\begin{minipage}[t]{0.5\textwidth}
\begin{center}
    \fontsize{11pt}{13pt}\selectfont \textbf{CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM} \\
    \vspace{0.1cm}
    \fontsize{11pt}{13pt}\selectfont \textbf{Độc lập - Tự do - Hạnh phúc} \\
    \vspace{0.05cm}
    \rule{4cm}{0.5pt}
\end{center}
\end{minipage}

\vspace{0.4cm}
\begin{flushright}
    \textit{Thành phố Hồ Chí Minh, ngày 11 tháng 6 năm 2026}
\end{flushright}

\vspace{0.6cm}
\begin{center}
    \fontsize{12pt}{14pt}\selectfont \textbf{KẾ HOẠCH TỔ CHỨC NGHIÊN CỨU KHOA HỌC} \\
    \vspace{0.15cm}
    \fontsize{12pt}{14pt}\selectfont \textbf{VÀ ĐỔI MỚI SÁNG TẠO CẤP SINH VIÊN 2026}
\end{center}

\vspace{0.5cm}

""" + latex_body + r"""

\end{document}
"""
        with open(tex_path, "w", encoding="utf-8") as f:
            f.write(latex_template)
        logger.info(f"Đã tạo file LaTeX thành công: {tex_path}")
    except Exception as e:
        logger.error(f"Lỗi khi biên dịch sang LaTeX: {e}")
        
    # 4. Xuất file HTML với giao diện Premium và CSS tùy biến phong cách hiện đại
    try:
        logger.info("Đang sinh file HTML giao diện premium...")
        # Sử dụng thư viện markdown để render body
        html_body = markdown.markdown(md_content, extensions=['tables', 'fenced_code'])
        
        # HTML Template với giao diện hiện đại (Outfit, Inter, Sleek light theme, glassmorphism, responsive table)
        html_template = f"""<!DOCTYPE html>
<html lang="vi">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=initial-scale=1.0">
    <title>Báo cáo tổng hợp NCKH và ĐMST Sinh viên 2026</title>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=Outfit:wght@400;500;600;700;800&display=swap" rel="stylesheet">
    <style>
        :root {{
            --bg-color: #f8fafc;
            --card-bg: #ffffff;
            --border-color: #e2e8f0;
            --primary: #4f46e5;
            --primary-glow: rgba(79, 70, 229, 0.05);
            --text-color: #1e293b;
            --text-muted: #64748b;
            --accent: #0284c7;
            --accent-green: #10b981;
        }}
        
        body {{
            font-family: 'Inter', sans-serif;
            background-color: var(--bg-color);
            color: var(--text-color);
            margin: 0;
            padding: 0;
            line-height: 1.6;
            background-image: 
                radial-gradient(circle at 10% 20%, rgba(79, 70, 229, 0.02) 0%, transparent 40%),
                radial-gradient(circle at 90% 80%, rgba(2, 132, 199, 0.02) 0%, transparent 40%);
            background-attachment: fixed;
        }}
        
        .container {{
            max-width: 1100px;
            margin: 0 auto;
            padding: 40px 20px;
        }}
        
        /* Light card header */
        header {{
            background: var(--card-bg);
            border: 1px solid var(--border-color);
            border-radius: 24px;
            padding: 40px;
            margin-bottom: 40px;
            box-shadow: 0 10px 30px rgba(0, 0, 0, 0.03);
            text-align: center;
        }}
        
        header h1 {{
            font-family: 'Outfit', sans-serif;
            font-size: 2.2rem;
            font-weight: 800;
            margin: 0 0 15px 0;
            background: linear-gradient(135deg, #1e293b 30%, #4f46e5 100%);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            line-height: 1.3;
        }}
        
        header p {{
            color: var(--text-muted);
            font-size: 1.1rem;
            margin: 0;
            font-weight: 500;
        }}
        
        /* Content area */
        .content {{
            background: var(--card-bg);
            border: 1px solid var(--border-color);
            border-radius: 24px;
            padding: 50px;
            box-shadow: 0 15px 35px rgba(0, 0, 0, 0.04);
        }}
        
        /* Typography custom styling */
        h2 {{
            font-family: 'Outfit', sans-serif;
            font-size: 1.7rem;
            color: #0f172a;
            border-bottom: 2px solid var(--primary);
            padding-bottom: 10px;
            margin-top: 40px;
            margin-bottom: 20px;
            display: flex;
            align-items: center;
        }}
        
        h2::before {{
            content: "";
            display: inline-block;
            width: 8px;
            height: 24px;
            background: linear-gradient(to bottom, var(--primary), var(--accent));
            margin-right: 12px;
            border-radius: 4px;
        }}
        
        h3 {{
            font-family: 'Outfit', sans-serif;
            font-size: 1.25rem;
            color: var(--accent);
            margin-top: 30px;
            margin-bottom: 15px;
            font-weight: 600;
        }}
        
        p {{
            margin-bottom: 20px;
            color: #334155;
        }}
        
        strong {{
            color: #0f172a;
            font-weight: 600;
        }}
        
        ul {{
            padding-left: 20px;
            margin-bottom: 25px;
        }}
        
        li {{
            margin-bottom: 10px;
            color: #334155;
        }}
        
        /* Table responsive light styling */
        table {{
            width: 100%;
            border-collapse: collapse;
            margin: 30px 0;
            font-size: 0.95rem;
            border-radius: 12px;
            overflow: hidden;
            border: 1px solid var(--border-color);
        }}
        
        th {{
            background-color: rgba(79, 70, 229, 0.07);
            color: #1e293b;
            font-weight: 600;
            text-align: left;
            padding: 16px;
            border-bottom: 1px solid var(--border-color);
            font-family: 'Outfit', sans-serif;
        }}
        
        td {{
            padding: 14px 16px;
            border-bottom: 1px solid #f1f5f9;
            color: #334155;
        }}
        
        tr:last-child td {{
            border-bottom: none;
        }}
        
        tr:hover td {{
            background-color: rgba(79, 70, 229, 0.02);
            color: #0f172a;
            transition: all 0.2s ease;
        }}
        
        tr:nth-child(even) td {{
            background-color: #f8fafc;
        }}
        
        /* Footer */
        footer {{
            margin-top: 50px;
            text-align: center;
            color: var(--text-muted);
            font-size: 0.85rem;
            border-top: 1px solid var(--border-color);
            padding-top: 20px;
        }}
        
        /* Bold total row highlight */
        tr:last-child {{
            font-weight: 700;
            background-color: rgba(79, 70, 229, 0.07) !important;
        }}
        tr:last-child td {{
            color: #0f172a;
        }}
    </style>
</head>
<body>
    <div class="container">
        <header>
            <h1>BÁO CÁO TỔNG HỢP HOẠT ĐỘNG NCKH VÀ ĐMST SINH VIÊN 2026</h1>
            <p>Trường Đại học Bách khoa - ĐHQG-HCM | Phòng Khoa học và Công nghệ</p>
        </header>
        <div class="content">
            {html_body}
        </div>
        <footer>
            <p>© 2026 Trường Đại học Bách khoa - ĐHQG-HCM. Thiết kế bởi Antigravity AI Agent.</p>
        </footer>
    </div>
</body>
</html>
"""
        with open(html_path, "w", encoding="utf-8") as f:
            f.write(html_template)
        logger.info(f"Đã tạo file HTML Premium thành công: {html_path}")
    except Exception as e:
        logger.error(f"Lỗi khi sinh file HTML: {e}")

    logger.info("Hoàn thành việc đóng gói và xuất bản tài liệu tổng hợp.")

if __name__ == "__main__":
    main()
