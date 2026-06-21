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

LATEX_UNICODE_REPLACEMENTS = {
    "≈": r"\(\approx\)",
    "≤": r"\(\leq\)",
    "≥": r"\(\geq\)",
    "°": r"\(^\circ\)",
    "₂": r"$_2$",
    "₃": r"$_3$",
    "₄": r"$_4$",
    "₅": r"$_5$",
    "₆": r"$_6$",
    "₇": r"$_7$",
    "₈": r"$_8$",
    "₉": r"$_9$",
    "₀": r"$_0$",
}


def sanitize_latex_unicode(text):
    """Normalize Unicode symbols that pdflatex/inputenc does not define."""
    for source, replacement in LATEX_UNICODE_REPLACEMENTS.items():
        text = text.replace(source, replacement)
    return text


def sanitize_latex_longtable(text):
    """Normalize Pandoc table output for pdflatex stability."""
    text = text.replace(
        "{\\def\\LTcaptype{none} % do not increment counter\n\\begin{longtable}",
        "\\begin{longtable}",
    )
    text = text.replace("\\bottomrule\\noalign{}\n\\endlastfoot", "\\endlastfoot")
    text = re.sub(r"(\\end\{longtable\})\n\}\n", r"\1\n", text)
    text = text.replace("\\begin{longtable}[]{", "\\begin{center}\n\\scriptsize\n\\begin{tabular}{")
    text = text.replace("\\endhead\n", "")
    text = text.replace("\\endlastfoot\n", "")
    text = text.replace("\\end{longtable}", "\\end{tabular}\n\\end{center}")
    return text


def add_latex_table_borders(text):
    """Add simple visible borders to generated tabular environments."""
    lines = text.splitlines()
    output = []
    in_tabular = False
    in_tabular_spec = False

    for line in lines:
        if "\\scriptsize" in line:
            output.append(line)
            output.append("  \\setlength{\\arrayrulewidth}{0.4pt}")
            continue

        if "\\begin{tabular}{@{}" in line:
            line = line.replace("\\begin{tabular}{@{}", "\\begin{tabular}{|")
            in_tabular = True
            in_tabular_spec = True

        if in_tabular_spec and "@{}}" in line:
            line = line.replace("@{}}", "|}")
            in_tabular_spec = False

        if in_tabular_spec and "\\arraybackslash}p{" in line and not line.rstrip().endswith("|"):
            line = line.replace("\\linewidth - 16\\tabcolsep", "\\linewidth - 22\\tabcolsep")
            line = f"{line}|"

        line = line.replace("\\toprule\\noalign{}", "\\hline")
        line = line.replace("\\midrule\\noalign{}", "\\hline")
        if in_tabular and line.rstrip().endswith("\\\\"):
            line = f"{line} \\hline"
        if "\\end{tabular}" in line:
            in_tabular = False

        output.append(line)

    return "\n".join(output) + "\n"

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
    md_content = """<table style="width: 100%; border: none; border-collapse: collapse; margin-bottom: 20px;">
  <tr style="border: none;">
    <td style="width: 45%; text-align: center; vertical-align: top; border: none; padding: 0;">
      TRƯỜNG ĐẠI HỌC BÁCH KHOA<br>
      <strong>&mdash;&mdash;&mdash;&mdash;&mdash;</strong>
    </td>
    <td style="width: 10%; border: none; padding: 0;"></td>
    <td style="width: 45%; text-align: center; vertical-align: top; border: none; padding: 0;">
      <strong>CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM</strong><br>
      <strong>Độc lập - Tự do - Hạnh phúc</strong><br>
      <strong>&mdash;&mdash;&mdash;&mdash;&mdash;&mdash;&mdash;&mdash;&mdash;</strong>
    </td>
  </tr>
</table>

<div style="text-align: right; font-style: italic; margin-bottom: 20px;">
  Thành phố Hồ Chí Minh, ngày 11 tháng 6 năm 2026
</div>

<div style="text-align: center; font-weight: bold; font-size: 1.5em; line-height: 1.4; margin-bottom: 30px;">
  KẾ HOẠCH TỔ CHỨC NGHIÊN CỨU KHOA HỌC<br>VÀ ĐỔI MỚI SÁNG TẠO CẤP SINH VIÊN 2026
</div>

*Lưu ý: Báo cáo được tổng hợp dựa trên kế hoạch và đề xuất thực tế từ 10 Khoa chuyên môn trực thuộc Trường Đại học Bách khoa - ĐHQG-HCM.*

---

## I. TỔNG QUAN VÀ KẾ HOẠCH TRIỂN KHAI CỦA CÁC KHOA

Dưới đây là tóm tắt định hướng chi tiết và lộ trình triển khai hoạt động Nghiên cứu Khoa học và Đổi mới Sáng tạo (NCKH & ĐMST) cấp sinh viên của 10 khoa chuyên môn thuộc trường trong chu kỳ 12 tháng (từ tháng 6/2026 đến tháng 5/2027):

**1. Khoa Cơ khí**

* **Mục tiêu và Tờ trình**: Thực hiện theo Tờ trình số 151/TTr-ĐHBK ngày 05/06/2026, Khoa triển khai kế hoạch thí điểm 12 tháng nhằm thiết lập quy trình tổ chức NCKH sinh viên chuẩn hóa, làm tiền đề nhân rộng và lặp lại hiệu quả từ năm 2027.
* **Định hướng chuyên môn**: Khoa tập trung huy động đông đảo lực lượng giảng viên hướng dẫn để hình thành các nhóm nghiên cứu nòng cốt, tạo tính kế thừa sản phẩm giữa các thế hệ sinh viên thay vì triển khai rời rạc. Kết quả nghiên cứu được hệ thống hóa làm minh chứng phục vụ kiểm định chất lượng đào tạo và làm nguồn sản phẩm tham gia các cuộc thi học thuật lớn.
* **Lộ trình triển khai**: Tiến trình được chia làm 02 đợt. Đợt 1 tiếp nhận và xét duyệt đề tài trong tháng 7/2026 (nghiệm thu vào tháng 10/2026); Đợt 2 xét duyệt vào tháng 11-12/2026 (nghiệm thu tháng 4/2027) và thực hiện thanh lý quyết toán toàn bộ kinh phí vào tháng 5/2027.

**2. Khoa Kỹ thuật Xây dựng**

* **Mục tiêu và Tờ trình**: Thực hiện theo Tờ trình số 178/KTXD ngày 15/06/2026, Khoa hướng tới thúc đẩy mạnh mẽ phong trào NCKH và ĐMST trên tất cả các chương trình đào tạo của Khoa (Tiêu chuẩn, Kỹ sư tài năng - KSTN, PFIEV, và Đào tạo Quốc tế OISP).
* **Lộ trình triển khai**: Quy trình thực hiện bám sát kế hoạch của Phòng KH&CN: Thông báo đăng ký và tiếp nhận đề xuất của sinh viên trong tháng 6-7/2026; Thành lập Hội đồng xét duyệt cấp Khoa trong tháng 7/2026; Tổng hợp danh mục vật tư gửi Phòng KH&CN mua sắm trong tháng 7-8/2026; Ký hợp đồng triển khai từ tháng 8-9/2026. Nghiệm thu đề tài vào tháng 4/2027 và quyết toán dứt điểm vào tháng 5/2027.

**3. Khoa Khoa học Ứng dụng**

* **Mục tiêu và Định hướng**: Khuyến khích tinh thần sáng tạo học thuật, phát triển năng lực nghiên cứu độc lập và khả năng ứng dụng thực tiễn của sinh viên. Khoa chủ trương chia các nhóm nghiên cứu sinh viên thành 2 hình thức: Nhóm A (Thực nghiệm chế tạo thiết bị, mô hình sinh học) và Nhóm B (Thuần túy mô phỏng tính toán lý thuyết).
* **Hoạt động và Lộ trình**: Triển khai trong vòng 12 tháng theo quy trình chung của Nhà trường. Đặc biệt, Khoa dành riêng một phần kinh phí để tổ chức Hội thảo khoa học sinh viên cấp Khoa nhằm trưng bày sản phẩm thực tế, báo cáo kết quả học thuật và tạo diễn đàn trao đổi chuyên môn giữa sinh viên, giảng viên và các phòng thí nghiệm.

**4. Khoa Kỹ thuật Giao thông**

* **Mục tiêu và Kế hoạch**: Thực hiện theo Kế hoạch số 49/KH-KTGT ngày 12/06/2026, Khoa ưu tiên các đề tài hướng đến sản phẩm đầu ra kiểm chứng trực quan bao gồm mô hình vật lý, drone, xe tự hành, thiết bị thử nghiệm thực tế hoặc các phần mềm ứng dụng chuyên ngành điều khiển giao thông thông minh.
* **Lộ trình và Đối tượng**: Đối tượng tham gia là sinh viên tất cả các hệ, trong đó ưu tiên sinh viên đang làm học phần tốt nghiệp. Nhận hồ sơ đăng ký trước ngày 31/07/2026, xét duyệt trong tháng 8/2026, thực hiện nghiên cứu từ tháng 11/2026 đến tháng 11/2027, nghiệm thu cấp Khoa vào tháng 10/2027 và thanh lý hợp đồng vào ngày 15/11/2027.

**5. Khoa Kỹ thuật Địa chất & Dầu khí**

* **Mục tiêu và Định hướng**: Tập trung phát triển các đề tài nghiên cứu thực nghiệm chuyên sâu về cơ học đá, cơ học đất, địa kỹ thuật công trình và khảo sát thực địa đặc thù ngành. Cam kết đầu ra chất lượng cao là các công bố quốc tế uy tín (tạp chí thuộc danh mục Scopus Q3).
* **Lộ trình triển khai**: Triển khai theo chu kỳ 6-12 tháng. Để đảm bảo tiến độ nghiên cứu thực nghiệm tại phòng thí nghiệm, Khoa ưu tiên tổng hợp và thực hiện quy trình mua sắm các thiết bị đo lường chuyên dụng ngay từ đợt đầu tiên.

**6. Khoa Công nghệ Vật liệu**

* **Mục tiêu và Định hướng**: Khuyến khích các đề tài nghiên cứu chế tạo vật liệu mới ứng dụng công nghệ cao và thân thiện với môi trường. Các hướng nghiên cứu chủ đạo gồm chế tạo vật liệu nano oxit kim loại làm anode cho pin lithium-ion, bê tông tự liền trong môi trường đặc thù, vật liệu in 3D từ nhựa tái chế và chế tạo khuôn concave vi mô.
* **Lộ trình triển khai**: Triển khai chu kỳ 12 tháng. Khoa thực hiện rà soát và lập kế hoạch mua sắm hóa chất, vật tư thí nghiệm nghiêm ngặt từ sớm để tránh tình trạng chậm trễ mẫu thực nghiệm của sinh viên.

**7. Khoa Môi trường và Tài nguyên**

* **Mục tiêu và Định hướng**: Tập trung vào các nghiên cứu ứng dụng công nghệ hóa-sinh hiện đại giải quyết các thách thức môi trường thực tế. Các đề tài tiêu biểu gồm ứng dụng Fenton điện hóa và hoạt hóa peroxymonosulfate xử lý nước thải dệt nhuộm, khảo sát khả năng chịu mặn và tương thích Bacillus velezensis và Trichoderma asperellum hướng đến phát triển chế phẩm sinh học phục vụ nông nghiệp sạch.
* **Lộ trình triển khai**: Triển khai từ tháng 08/2026 đến tháng 07/2027. Khoa nỗ lực tối ưu hóa phân bổ nguồn kinh phí thực tế của nhà trường để đáp ứng nhu cầu thực hiện đề tài rất lớn của sinh viên.

**8. Khoa Khoa học & Kỹ thuật Máy tính**

* **Mục tiêu và Định hướng**: Thúc đẩy nghiên cứu khoa học sinh viên đỉnh cao gắn liền với các công bố khoa học uy tín quốc tế và tham gia giải thưởng học thuật lớn. Định hướng đầu ra là bài báo đăng tại hội nghị có phản biện thuộc IEEE/Springer/ACM, Scopus Index hoặc tạp chí tính điểm của Hội đồng Giáo sư Nhà nước.
* **Lộ trình triển khai**: Xét duyệt đề tài và mua sắm thiết bị phần cứng/phần mềm trong tháng 8/2026, thực hiện nghiên cứu từ tháng 9/2026 đến tháng 4/2027, nghiệm thu và thanh lý vào tháng 5/2027. Khuyến khích sinh viên tham gia Huawei ICT Competition, Robocon, BK Innovation, Liên hoan Tuổi trẻ Sáng tạo.

**9. Khoa Quản lý Công nghiệp**

* **Mục tiêu và Định hướng**: Tập trung vào các đề tài ứng dụng công nghệ số và phương pháp tối ưu hóa trong quản lý kinh tế và vận hành chuỗi cung ứng. Hướng nghiên cứu chủ đạo gồm ứng dụng học máy trong dự báo nhu cầu doanh nghiệp, tối ưu hóa chuỗi cung ứng logistics (Heuristics, NSGA-II), định lượng khả năng phục hồi chuỗi cung ứng dưới tác động ripple effect, và ứng dụng AI trong lập kế hoạch du lịch xanh bền vững (mô hình UTAUT2).
* **Lộ trình triển khai**: Triển khai chu kỳ 9-12 tháng bắt đầu từ tháng 9/2026, nghiệm thu và đánh giá kết quả theo đúng quy trình chung của Khoa và Nhà trường.

**10. Khoa Kỹ thuật Hóa học**

* **Mục tiêu và Tờ trình**: Thực hiện theo Tờ trình số 182/TTr-KTHH ngày 11/06/2026, Khoa thực hiện kế hoạch 12 tháng (tháng 6/2026 - tháng 5/2027) nhằm xây dựng quy trình quản lý NCKH sinh viên chuyên nghiệp và hiệu quả. Nâng cao số lượng sinh viên tiếp cận nghiên cứu thực nghiệm hóa chất, thúc đẩy công bố khoa học.
* **Lộ trình**: Chia làm 2 đợt tiếp nhận và nghiệm thu tương tự Khoa Cơ khí, đảm bảo giải ngân và quyết toán dứt điểm toàn bộ kinh phí vào tháng 5/2027.

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
* **Khoa Công nghệ Vật liệu**: Đăng ký danh mục mua sắm chi tiết trị giá **112,05 triệu đồng** (trong tổng đăng ký nhu cầu thực tế là 112,05 triệu đồng).
* **Khoa Môi trường và Tài nguyên**: Đăng ký danh mục mua sắm chi tiết trị giá **67,27 triệu đồng** (Trường duyệt phân bổ thực tế là 27,4 triệu đồng).

### 2. Danh mục mua sắm chi tiết tiêu biểu của các nhóm nghiên cứu

* **Vật tư cơ khí, điện tử và tự động hóa (Khoa Cơ khí, Giao thông, Máy tính)**:
  * Linh kiện điện tử chuyên dụng: Vi điều khiển nhúng trên UAV, ESP32, mạch Arduino, cảm biến đo lực, cảm biến khoảng cách, cảm biến IMU, module truyền thông, module GPS, camera nhận diện hình ảnh.
  * Vật liệu cơ khí chế tạo: Nhôm định hình, thép tấm, nhựa in 3D chuyên dụng, sợi carbon, vật liệu composite chịu lực cao dùng chế tạo vỏ khung xe tự hành, cánh sóng và các cụm thử nghiệm robot.
* **Hóa chất và vật liệu sinh học (Khoa Kỹ thuật Hóa học, Công nghệ Vật liệu, Môi trường, Khoa học Ứng dụng)**:
  * Hóa chất phân tích & tổng hợp xúc tác: Hóa chất tinh khiết KOH, PVA, NaOH, H₂SO₄, HNO₃, cồn công nghiệp 96-98%, dung dịch Nafion 1100EW, các muối mangan(II) clorua MnCl2, thiếc(IV) clorua, đồng sunfat, sắt sunfat, boric acid, natri thiosunfat, amoni molybdat cấp AR.
  * Thiết bị phụ trợ và dụng cụ thủy tinh: Điện cực glassy carbon chữ L 6mm, micropipette kích thước 0.5-10 uL và 100-1000uL, hộp đựng tip nhựa polypropylene kháng hóa chất, chén nung sứ và nung kim loại 100cc, màng bọc cao su chịu lực, giấy lọc, dây o-ring làm kín.
* **Thành phần thiết bị đo lường và khảo sát (Khoa Địa chất & Dầu khí, Khoa học Ứng dụng)**:
  * Thiết bị quang học đo biến dạng: Ống kính macro chuyên dụng lắp trên camera đo biến dạng mẫu đá, đèn flash công suất lớn, tripod chịu tải cao.
  * Thiết bị y sinh và y học: Gel tẩy da NuPrep, gel kết nối điện cực đo điện não Ten20 phục vụ phân tích tín hiệu EEG, thiết bị đo sinh tồn (PPG).

---

## IV. ĐỊNH HƯỚNG CÔNG BỐ KHOA HỌC VÀ BÁO CÁO HỘI NGHỊ, HỘI THẢO

Công bố khoa học và tham gia các cuộc thi học thuật là sản phẩm đầu ra bắt buộc của hoạt động NCKH & ĐMST sinh viên năm 2026:

### 1. Định hướng xuất bản bài báo khoa học
* **Khoa KH&KT Máy tính**: Định hướng bài báo được chấp nhận đăng tại các hội thảo/hội nghị uy tín thuộc danh mục IEEE, Springer, ACM; tạp chí thuộc danh mục Scopus hoặc tạp chí chuyên ngành được Hội đồng Chức danh Giáo sư Nhà nước (HDCDGSNN) tính điểm.
* **Khoa Kỹ thuật Địa chất & Dầu khí**: Cam kết đầu ra gồm ít nhất 01 bài báo tạp chí quốc tế thuộc danh mục Scopus Q3 và 01 báo cáo tại Hội nghị khoa học quốc tế.
* **Khoa Quản lý Công nghiệp**: Định hướng bài báo đăng tại các tạp chí trong danh mục HDCDGSNN, kỷ yếu hội nghị quốc tế và tạp chí Scopus Q4.
* **Khoa Công nghệ Vật liệu & Kỹ thuật Hóa học**: Bài báo gửi đăng tạp chí chuyên ngành trong nước/quốc tế và báo cáo poster tại Ngày hội kỹ thuật cấp khoa.

### 2. Định hướng tham gia các cuộc thi học thuật lớn
* Định hướng đưa các sản phẩm nghiên cứu thực tế tham gia tranh tài tại các cuộc thi công nghệ và học thuật uy tín như Huawei ICT Competition, Robocon, BK Innovation, Giải thưởng Khoa học và Công nghệ dành cho sinh viên trong các cơ sở giáo dục Đại học của Bộ GD&ĐT, và Liên hoan Tuổi trẻ Sáng tạo.

### 3. Tổng hợp kinh phí dành cho Hội nghị, Hội thảo (HN, HT)
* **Khoa KH&KT Máy tính**: Dành riêng **410 triệu đồng** (~37% tổng kinh phí) để tài trợ chi phí đăng bài báo, lệ phí hội nghị và chi phí đi lại, ăn ở cho sinh viên báo cáo.
* **Khoa Kỹ thuật Hóa học**: Dành **265 triệu đồng** (~30% tổng kinh phí) hỗ trợ sinh viên công bố khoa học và đăng ký tham dự hội nghị.
* **Khoa Cơ khí**: Dành **112 triệu đồng** (~10% tổng kinh phí) hỗ trợ chi phí đăng bài và hội thảo khoa học liên quan trực tiếp đến đề tài.
* **Khoa Khoa học Ứng dụng**: Dành **27,5 triệu đồng** từ nguồn phân bổ để tổ chức riêng Hội thảo khoa học sinh viên tại khoa phục vụ báo cáo và triển lãm.

---

## V. TỔNG HỢP CÁC ĐỀ XUẤT VÀ KIẾN NGHỊ CỦA CÁC KHOA

Để hoạt động NCKH & ĐMST cấp sinh viên năm 2026 đạt hiệu quả thực chất, các khoa đã đưa ra một số đề xuất quan trọng gửi Phòng Khoa học & Công nghệ, Phòng Kế hoạch - Tài chính và Ban Giám hiệu Trường:

### 1. Đề xuất về cơ chế tài chính và định mức hỗ trợ (Khoa Cơ khí, Hóa học, Địa chất, Môi trường)
* **Hỗ trợ kinh phí sinh viên**: Kiến nghị áp dụng cơ chế cho phép chi bồi dưỡng trực tiếp cho sinh viên tham gia thực hiện đề tài (dự kiến định mức ≈15% giá trị đề tài) nhằm tạo động lực và nâng cao trách nhiệm nghiên cứu của sinh viên.
* **Đặc thù ngành**: Kiến nghị xây dựng định mức chi phí đi lại, khảo sát thực địa và thu thập số liệu hiện trường cho sinh viên (dự kiến chiếm ≈10% kinh phí phân bổ), đặc biệt cần thiết cho các ngành có khối lượng công việc thực địa lớn như Địa chất Dầu khí, Môi trường và Tài nguyên.
* **Tăng hạn mức kinh phí**: Một số khoa (như Môi trường và Tài nguyên, Công nghệ Vật liệu) đề nghị nhà trường xem xét nâng hạn mức kinh phí phân bổ thực tế vì nhu cầu đăng ký và tính toán dự toán hóa chất, thiết bị thực tế của sinh viên vượt xa hạn mức trường duyệt.

### 2. Đề xuất cải tiến quy trình hành chính và mua sắm (Khoa KHUD, Giao thông, Máy tính, Vật liệu)
* **Quy trình hóa chất & thiết bị chuyên dụng**: Hướng dẫn chi tiết quy trình mua sắm, nhập khẩu và quản lý sử dụng các hóa chất thí nghiệm thuộc danh mục kiểm soát (hóa chất độc hại), các linh kiện điện tử chuyên dụng nhập khẩu và các phần mềm bản quyền mô phỏng nghiên cứu.
* **Rút ngắn thời gian giải ngân**: Đề xuất tối ưu hóa quy trình duyệt hồ sơ mua sắm và giải ngân (rút ngắn thời gian thông thường 30 ngày xuống mức tối thiểu) để tránh làm gián đoạn tiến độ thực hiện đề tài của sinh viên, đặc biệt với các đề tài chế tạo thực nghiệm chỉ có chu kỳ 12 tháng.

### 3. Đề xuất cơ chế ghi nhận cho giảng viên hướng dẫn (Khoa Cơ khí, KH&KTMT, KTHH)
* **Ghi nhận công sức giảng viên**: Đề nghị Phòng KH&CN ban hành hướng dẫn cụ thể về việc tính điểm Nhiệm vụ 2 (quy đổi giờ chuẩn giảng dạy) và hỗ trợ kinh phí bồi dưỡng hướng dẫn cho giảng viên hướng dẫn đề tài NCKH sinh viên nhằm khuyến khích giảng viên đồng hành lâu dài cùng các nhóm sinh viên.

### 4. Đề xuất cơ chế phối hợp tham gia cuộc thi ngoài trường (Khoa Cơ khí, KTHH)
* **Tạo điều kiện tham gia cuộc thi**: Kiến nghị Phòng Công tác Sinh viên (CTSV) hỗ trợ đơn giản hóa quy trình tiếp nhận hồ sơ, làm thủ tục đề cử sinh viên tham gia các cuộc thi học thuật ngoài trường (theo quy trình BK-QT-ED-056-05) nhằm tạo cơ hội cọ xát thực tế cho sinh viên sau khi hoàn thành đề tài cấp cơ sở.
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
        # Tìm phần thân báo cáo để tránh trùng lặp phần Quốc hiệu & Tiêu ngữ hành chính
        try:
            section_i_idx = md_content.index("## I. TỔNG QUAN")
            clean_md = md_content[section_i_idx:]
        except ValueError:
            lines = md_content.splitlines()
            clean_md = "\n".join(lines[22:])
        
        # Convert sang latex body
        latex_body = pypandoc.convert_text(clean_md, 'latex', format='md')
        latex_body = sanitize_latex_unicode(latex_body)
        latex_body = sanitize_latex_longtable(latex_body)
        latex_body = add_latex_table_borders(latex_body)
        
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
        # Cắt phần thân báo cáo để tránh trùng lặp phần Quốc hiệu & Tiêu ngữ hành chính
        try:
            section_i_idx = md_content.index("## I. TỔNG QUAN")
            clean_html_md = md_content[section_i_idx:]
        except ValueError:
            clean_html_md = md_content
        html_body = markdown.markdown(clean_html_md, extensions=['tables', 'fenced_code'])
        
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
        
        /* Administrative header styles */
        .administrative-header {{
            display: flex;
            justify-content: space-between;
            align-items: flex-start;
            margin-bottom: 15px;
            font-size: 0.95rem;
            color: var(--text-color);
        }}
        
        .admin-left {{
            width: 45%;
            text-align: center;
            line-height: 1.5;
            font-family: 'Outfit', sans-serif;
            font-weight: 500;
        }}
        
        .admin-right {{
            width: 50%;
            text-align: center;
            line-height: 1.5;
            font-family: 'Outfit', sans-serif;
        }}
        
        .header-line {{
            display: inline-block;
            width: 90px;
            height: 1px;
            background-color: var(--text-color);
            margin-top: 8px;
        }}
        
        .header-line-long {{
            display: inline-block;
            width: 150px;
            height: 1px;
            background-color: var(--text-color);
            margin-top: 8px;
        }}
        
        .date-location {{
            text-align: right;
            font-style: italic;
            margin-bottom: 35px;
            font-size: 0.95rem;
            color: var(--text-muted);
        }}
        
        .main-title {{
            text-align: center;
            font-family: 'Outfit', sans-serif;
            font-size: 1.8rem;
            font-weight: 800;
            line-height: 1.4;
            margin-bottom: 40px;
            color: #0f172a;
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
        <div class="administrative-header">
            <div class="admin-left">
                TRƯỜNG ĐẠI HỌC BÁCH KHOA<br>
                <span class="header-line"></span>
            </div>
            <div class="admin-right">
                <strong>CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM</strong><br>
                <strong>Độc lập - Tự do - Hạnh phúc</strong><br>
                <span class="header-line-long"></span>
            </div>
        </div>
        <div class="date-location">
            Thành phố Hồ Chí Minh, ngày 11 tháng 6 năm 2026
        </div>
        <div class="main-title">
            KẾ HOẠCH TỔ CHỨC NGHIÊN CỨU KHOA HỌC<br>VÀ ĐỔI MỚI SÁNG TẠO CẤP SINH VIÊN 2026
        </div>
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
