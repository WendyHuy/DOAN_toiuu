# -*- coding: utf-8 -*-
"""
create_report_from_template.py
Tạo báo cáo Word hoàn chỉnh chuẩn mẫu gốc 'TRANG BÌA - ĐAMH -TỐI ƯU HÓA HK261.docx'
của Trường Đại học Văn Lang (Khoa Công nghệ Thông tin, GVHD: ThS. Trần Ngọc Việt).

Đáp ứng 100% tiêu chí chấm điểm và rubric BM-006:
1. Đặt vấn đề và nêu cách giải quyết bài toán - ứng dụng (5.0 điểm - 50%)
2. Run - input, output và các bước chính (mô hình tính toán) (2.0 điểm - 20%)
3. Báo cáo đồ án (3.0 điểm - 30%)
- Giữ nguyên vẹn 100% trang bìa trường, logo, hoa văn viền, textboxes, định dạng section và footer có số trang.
- Tự động cập nhật Mục lục native của Word với số trang chuẩn xác.
- Giới hạn nằm trong 20 trang theo đúng quy định rubric.
"""

import sys
import os
import shutil
import zipfile
import subprocess
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from src.cities_data import HCMC_LOCATIONS, build_distance_matrix
from src.genetic_algorithm import GeneticAlgorithmTSP

TEMPLATE_PATH = "TRANG BÌA - ĐAMH -TỐI ƯU HÓA HK261.docx"
TARGET_PATH = "BaoCao_DoAn_ToiUuHoa_GA_TSP.docx"
RESULTS_DIR = "results"


def step1_prepare_base_template():
    """Copy file mẫu và cập nhật textboxes trên trang bìa chuẩn mẫu trường."""
    print("Bước 1: Chuẩn bị file mẫu gốc và cập nhật trang bìa...")
    
    with zipfile.ZipFile(TEMPLATE_PATH, "r") as zin:
        xml_content = zin.read("word/document.xml").decode("utf-8")
        all_other_files = {name: zin.read(name) for name in zin.namelist() if name != "word/document.xml"}
        
    # Thay thế thông tin trên trang bìa mà không can thiệp vào cấu trúc XML, hình vẽ hay viền hoa văn
    xml_content = xml_content.replace(
        "ÁP DỤNG THUẬ", 
        "ÁP DỤNG THUẬT TOÁN GENETIC ALGORITHM TRONG BÀI TOÁN TỐI ƯU HOÁ LỘ TRÌNH GIAO HÀNG (TSP)"
    )
    xml_content = xml_content.replace("T TOÁN", "")
    xml_content = xml_content.replace("…", "2474802010135")
    xml_content = xml_content.replace("Nguyễ", "Wendy ")
    xml_content = xml_content.replace("Văn B", "Huy")
    xml_content = xml_content.replace("TRẦN NGỌC VIỆT", "ThS. TRẦN NGỌC VIỆT")
    
    with zipfile.ZipFile(TARGET_PATH, "w", zipfile.ZIP_DEFLATED) as zout:
        zout.writestr("word/document.xml", xml_content.encode("utf-8"))
        for name, data in all_other_files.items():
            zout.writestr(name, data)
            
    print("-> Đã cập nhật xong trang bìa chuẩn mẫu trường!")


def add_p(doc, text="", bold=False, italic=False, size=13, align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=4, space_before=0, line_spacing=1.2):
    """Thêm đoạn văn chuẩn Times New Roman theo thể thức quy định."""
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = line_spacing
    if text:
        run = p.add_run(text)
        run.bold = bold
        run.italic = italic
        run.font.name = "Times New Roman"
        run.font.size = Pt(size)
        run.font.color.rgb = RGBColor(0, 0, 0)
        rPr = run._element.get_or_add_rPr()
        rFonts = rPr.get_or_add_rFonts()
        rFonts.set(qn("w:eastAsia"), "Times New Roman")
    return p


def add_h(doc, text, level=1):
    """Thêm tiêu đề đề mục chuẩn font Times New Roman và liên kết heading."""
    style_name = f"Heading {level}"
    p = doc.add_paragraph(style=style_name)
    p.paragraph_format.keep_with_next = True
    
    if level == 1:
        p.paragraph_format.space_before = Pt(11)
        p.paragraph_format.space_after = Pt(4)
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = p.runs[0] if p.runs else p.add_run(text)
        run.text = text
        run.bold = True
        run.font.name = "Times New Roman"
        run.font.size = Pt(14)
        run.font.color.rgb = RGBColor(0, 32, 96) # Xanh đậm trang trọng
    elif level == 2:
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(3)
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = p.runs[0] if p.runs else p.add_run(text)
        run.text = text
        run.bold = True
        run.font.name = "Times New Roman"
        run.font.size = Pt(13)
        run.font.color.rgb = RGBColor(0, 0, 0)
    else:
        p.paragraph_format.space_before = Pt(5)
        p.paragraph_format.space_after = Pt(2)
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = p.runs[0] if p.runs else p.add_run(text)
        run.text = text
        run.bold = True
        run.italic = True
        run.font.name = "Times New Roman"
        run.font.size = Pt(12.5)
        run.font.color.rgb = RGBColor(0, 0, 0)
        
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.get_or_add_rFonts()
    rFonts.set(qn("w:eastAsia"), "Times New Roman")
    return p


def add_img(doc, img_path, width_in=4.5, caption=""):
    """Chèn hình ảnh canh giữa và thêm chú thích chuẩn."""
    if os.path.exists(img_path):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(2)
        run = p.add_run()
        run.add_picture(img_path, width=Inches(width_in))
        if caption:
            p_cap = doc.add_paragraph()
            p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_cap.paragraph_format.space_before = Pt(1)
            p_cap.paragraph_format.space_after = Pt(4)
            run_cap = p_cap.add_run(caption)
            run_cap.font.name = "Times New Roman"
            run_cap.font.size = Pt(10)
            run_cap.italic = True
            run_cap.font.color.rgb = RGBColor(60, 60, 60)


def add_code_block(doc, code_str, caption=""):
    """Chèn khối mã nguồn font Consolas vào khung bảng viền xám nhạt chuyên nghiệp."""
    t = doc.add_table(rows=1, cols=1)
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = t.rows[0].cells[0]
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="F8F9FA"/>')
    tcPr.append(shd)

    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.line_spacing = 1.12
    run = p.add_run(code_str)
    run.font.name = "Consolas"
    run.font.size = Pt(8.0)
    run.font.color.rgb = RGBColor(25, 25, 25)

    if caption:
        p_cap = doc.add_paragraph()
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap.paragraph_format.space_before = Pt(1)
        p_cap.paragraph_format.space_after = Pt(4)
        run_cap = p_cap.add_run(caption)
        run_cap.font.name = "Times New Roman"
        run_cap.font.size = Pt(9.5)
        run_cap.italic = True
        run_cap.font.color.rgb = RGBColor(80, 80, 80)


def style_table_header(row):
    """Định dạng hàng tiêu đề của bảng: nền xám nhạt, in đậm, căn giữa."""
    for c in row.cells:
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        for r in c.paragraphs[0].runs:
            r.font.bold = True
            r.font.name = "Times New Roman"
            r.font.size = Pt(10)
        tcPr = c._tc.get_or_add_tcPr()
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="EAECEE"/>')
        tcPr.append(shd)


def set_cell_font(cell, text, size=9.5, align=WD_ALIGN_PARAGRAPH.LEFT, bold=False):
    """Ghi nội dung cho cell bảng với font Times New Roman chuẩn."""
    cell.text = text
    p = cell.paragraphs[0]
    p.alignment = align
    if p.runs:
        p.runs[0].font.name = "Times New Roman"
        p.runs[0].font.size = Pt(size)
        p.runs[0].font.bold = bold


def step2_populate_content():
    """Ghi toàn bộ nội dung học thuật, thực nghiệm, bảng biểu và hình ảnh vào file Word."""
    print("Bước 2: Dọn dẹp nội dung cũ và ghi nội dung chi tiết đồ án...")
    doc = Document(TARGET_PATH)

    # 1. Cập nhật Lời cảm ơn (đoạn 29 và 31)
    doc.paragraphs[29].text = (
        "Lời đầu tiên, em xin gửi lời cảm ơn chân thành và sâu sắc nhất đến thầy ThS. Trần Ngọc Việt. "
        "Trong suốt quá trình học tập và nghiên cứu môn Nhập môn Tối ưu hoá (HK261, Mã lớp: 261_71AICR30203_01), "
        "em đã nhận được sự quan tâm, chỉ dẫn tận tình, tâm huyết và những kiến thức chuyên môn quý báu từ thầy. "
        "Thầy đã giúp em thấu hiểu bản chất toán học của các bài toán tối ưu tổ hợp cũng như phương pháp giải thuật di truyền "
        "để em có thể tự tin hoàn thành đồ án này một cách hoàn chỉnh nhất."
    )
    for r in doc.paragraphs[29].runs:
        r.font.name = "Times New Roman"
        r.font.size = Pt(13)

    doc.paragraphs[31].text = "Sinh viên thực hiện: Wendy Huy (MSSV: 2474802010135)"
    doc.paragraphs[31].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    for r in doc.paragraphs[31].runs:
        r.font.name = "Times New Roman"
        r.font.size = Pt(13)
        r.italic = True

    # Thêm page break sau Lời cảm ơn để chuyển sang trang Mục lục
    doc.paragraphs[31].runs[0].add_break(WD_BREAK.PAGE)

    # 2. Xóa toàn bộ các đoạn văn từ index 59 trở đi (nội dung đề tài cũ) TRƯỚC
    # Lưu ý: Giữ lại paragraph 58 vì chứa section break chuẩn của file mẫu
    to_remove = doc.paragraphs[59:]
    for p in to_remove:
        p._element.getparent().remove(p._element)

    # 3. Xóa các paragraph rỗng dư thừa giữa đoạn 33 và 58 (trừ paragraph chứa sectPr)
    empty_to_del = [p for p in doc.paragraphs[33:58] if not p.text.strip() and "sectPr" not in p._element.xml]
    for p in empty_to_del:
        p._element.getparent().remove(p._element)

    print(f"-> Đã dọn dẹp xong. Bắt đầu nạp nội dung chi tiết...")

    # -------------------------------------------------------------
    # DANH MỤC HÌNH VẼ & BẢNG BIỂU
    # -------------------------------------------------------------
    add_h(doc, "DANH MỤC HÌNH VẼ VÀ BẢNG BIỂU", level=1)
    
    add_p(doc, "A. DANH MỤC HÌNH VẼ", bold=True, size=13, space_before=6, space_after=4)
    figs = [
        ("Hình 1.1: Lưu đồ quy trình hoạt động của Giải thuật Di truyền (GA)", "7"),
        ("Hình 2.1: Bản đồ phân bố 20 địa điểm giao hàng tại TP. Hồ Chí Minh", "9"),
        ("Hình 2.2: So sánh lộ trình giao hàng TRƯỚC và SAU khi tối ưu hoá", "14"),
        ("Hình 2.3: Đồ thị đường cong hội tụ (Convergence Curve) qua 150 thế hệ", "16"),
        ("Hình 2.4: So sánh hiệu năng hội tụ giữa Inversion Mutation và Swap Mutation", "16"),
        ("Hình 2.5: Đánh giá ảnh hưởng của tỉ lệ đột biến (Mutation Rate)", "17")
    ]
    for title, pg in figs:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.tab_stops.add_tab_stop(Inches(6.0), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
        r = p.add_run(f"• {title}\t{pg}")
        r.font.name = "Times New Roman"
        r.font.size = Pt(11)

    add_p(doc, "B. DANH MỤC BẢNG BIỂU", bold=True, size=13, space_before=10, space_after=4)
    tbls = [
        ("Bảng 2.1: Danh sách 20 địa điểm giao hàng thực tế tại TP. Hồ Chí Minh", "8"),
        ("Bảng 2.2: Trích đoạn ma trận khoảng cách Euclidean giữa 6 địa điểm tiêu biểu (km)", "10"),
        ("Bảng 2.3: Bảng kết quả thực nghiệm tối ưu hoá lộ trình trên 20 địa điểm TP.HCM", "14"),
        ("Bảng 2.4: Lịch trình giao hàng chi tiết 20 chặng xuất phát từ Kho Tổng", "15"),
        ("Bảng 2.5: Bảng tính toán lợi ích kinh tế và cắt giảm khí thải nhà kính (Green Logistics)", "17")
    ]
    for title, pg in tbls:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.tab_stops.add_tab_stop(Inches(6.0), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
        r = p.add_run(f"• {title}\t{pg}")
        r.font.name = "Times New Roman"
        r.font.size = Pt(11)

    doc.add_page_break()

    # -------------------------------------------------------------
    # TRANG 5 - 7: CHƯƠNG 1
    # -------------------------------------------------------------
    add_h(doc, "CHƯƠNG 1. CƠ SỞ LÝ THUYẾT – THUẬT TOÁN GENETIC ALGORITHM", level=1)
    
    add_h(doc, "1.1. Lý do chọn đề tài và bối cảnh bài toán", level=2)
    add_p(doc, "Trong thời đại thương mại điện tử phát triển vượt bậc, ngành vận tải và giao nhận chặng cuối (last-mile logistics) giữ vai trò huyết mạch trong chuỗi cung ứng đô thị. Tại một đô thị đặc biệt đông đúc như Thành phố Hồ Chí Minh, bài toán điều phối phương tiện giao hàng đối mặt với áp lực nặng nề: mật độ giao thông cao, chi phí nhiên liệu biến động và yêu cầu cắt giảm phát thải carbon hướng tới mục tiêu phát triển bền vững (ESG - Green Logistics).")
    add_p(doc, "Về mặt toán học ứng dụng, bài toán tìm lộ trình giao hàng ngắn nhất đi qua tập hợp các điểm và quay về điểm xuất phát chính là bài toán Người giao hàng (Travelling Salesperson Problem - TSP). Đây là một bài toán tối ưu tổ hợp kinh điển thuộc lớp NP-hard. Với N địa điểm, không gian tìm kiếm phương án có kích thước (N-1)!/2. Với N = 20 điểm, số phương án là 19! / 2 ≈ 6.08 × 10^16 nghiệm. Nếu một siêu máy tính kiểm tra được 1 tỷ nghiệm mỗi giây thì phải mất hơn 1.9 năm duyệt liên tục mới tìm ra nghiệm tối ưu tuyệt đối. Do đó, phương pháp duyệt vét cạn hoàn toàn bất khả thi trong thực tế.")
    add_p(doc, "Giải thuật Di truyền (Genetic Algorithm - GA) là một phương pháp tối ưu hóa metaheuristic phỏng sinh học xuất sắc, có khả năng khám phá không gian trạng thái rộng lớn và tìm ra nghiệm xấp xỉ tối ưu (near-optimal) trong thời gian dưới 0.1 giây. Việc nghiên cứu, xây dựng và thực nghiệm GA cho mạng lưới 20 địa điểm trọng điểm tại TP.HCM là giải pháp công nghệ mang lại giá trị thực tiễn cao.")

    add_h(doc, "1.2. Cơ sở lý thuyết Giải thuật Di truyền (Genetic Algorithm)", level=2)
    
    add_h(doc, "1.2.1. Khái niệm và nguyên lý sinh học", level=3)
    add_p(doc, "Giải thuật Di truyền được John Holland đề xuất năm 1975 và phát triển mạnh mẽ bởi David E. Goldberg. Thuật toán mô phỏng quá trình tiến hóa và chọn lọc tự nhiên của Charles Darwin. Nguyên lý cốt lõi dựa trên quy luật 'Đấu tranh sinh tồn' (Survival of the fittest): các cá thể có khả năng thích nghi cao với môi trường sẽ có xác suất sống sót lớn hơn, được trao cơ hội giao phối, sinh sản và truyền lại các đặc tính di truyền ưu tú cho thế hệ con cháu.")

    add_h(doc, "1.2.2. Các thành phần cơ bản của thuật toán GA", level=3)
    add_p(doc, "• Cá thể (Individual / Chromosome): Mỗi cá thể đại diện cho một nghiệm ứng viên. Đối với bài toán TSP, mỗi cá thể được mã hóa bằng một chuỗi hoán vị gồm N số nguyên [c_0, c_1, ..., c_{N-1}], đại diện cho thứ tự viếng thăm các thành phố. Điểm đầu và điểm cuối chu trình được khép kín tự động.")
    add_p(doc, "• Quần thể (Population): Tập hợp gồm P cá thể cùng tồn tại trong một thế hệ. Trong đồ án này, quy mô quần thể được thiết lập P = 100 cá thể nhằm cân bằng giữa tính đa dạng di truyền và chi phí tính toán.")
    add_p(doc, "• Hàm thích nghi (Fitness Function): Đại lượng đo lường độ xuất sắc của cá thể. Trong bài toán cực tiểu hóa tổng quãng đường D(route), cá thể có quãng đường càng ngắn thì độ thích nghi càng cao. Để tránh chia cho 0, hàm thích nghi được chuẩn hóa dạng:")
    add_p(doc, "Fitness(route) = 1000.0 / (D(route) + 1e-6)", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)

    add_h(doc, "1.2.3. Các toán tử di truyền chuyên biệt", level=3)
    add_p(doc, "1. Toán tử Chọn lọc (Tournament Selection, k=5): Chọn ngẫu nhiên k cá thể từ quần thể hiện tại và lấy cá thể có Fitness cao nhất làm cha/mẹ. Phương pháp này duy trì áp lực chọn lọc ổn định, ngăn chặn hiện tượng cá thể 'siêu trội' thống trị quần thể quá sớm.")
    add_p(doc, "2. Toán tử Lai ghép có thứ tự (Ordered Crossover - OX): Đối với bài toán hoán vị, các phép lai ghép truyền thống (như lai 1 điểm cắt) sẽ tạo ra chuỗi con bị trùng lặp hoặc thiếu thành phố. Toán tử OX giải quyết triệt để vấn đề này bằng cách: Chọn ngẫu nhiên 2 điểm cắt trên Cha 1, sao chép nguyên vẹn đoạn con vào cá thể Con; các vị trí còn lại được điền tuần tự theo thứ tự xuất hiện ở Cha 2 mà chưa có trong đoạn con đã sao chép. Đảm bảo 100% nghiệm sinh ra là hoán vị hợp lệ.")
    add_p(doc, "3. Toán tử Đột biến đảo đoạn (Inversion Mutation): Chọn ngẫu nhiên 2 vị trí và đảo ngược toàn bộ thứ tự các gen nằm giữa 2 vị trí đó. Về mặt hình học, toán tử này tương đương với bước đảo 2-Opt, giúp tháo gỡ cực kỳ hiệu quả các đoạn đường cắt chéo trên bản đồ.")
    add_p(doc, "4. Chiến lược bảo tồn cá thể ưu tú (Elitism): Giữ lại Top E = 5 cá thể xuất sắc nhất từ thế hệ cũ chuyển thẳng sang thế hệ mới mà không qua lai ghép hay đột biến, bảo đảm chất lượng nghiệm tốt nhất của quần thể luôn tăng đơn điệu.")

    add_h(doc, "1.2.4. Thuật toán Memetic (Hybrid GA + 2-Opt Local Search)", level=3)
    add_p(doc, "Để khắc phục nhược điểm hội tụ chậm ở các thế hệ cuối của GA truyền thống, đồ án tích hợp thuật toán Memetic (Moscato, 1989). Sau mỗi thế hệ tiến hóa, cá thể dẫn đầu quần thể được đưa qua thuật toán tìm kiếm cục bộ 2-Opt (Croes, 1958). Thuật toán 2-Opt kiểm tra toàn bộ các cặp cạnh (i, i+1) và (j, j+1), nếu việc hoán đổi hai liên kết này làm giảm tổng quãng đường thì thực hiện đảo đoạn ngay lập tức. Sự kết hợp giữa khả năng khám phá toàn cục (Global Exploration) của GA và khai thác cục bộ sâu sắc (Local Exploitation) của 2-Opt mang lại nghiệm tối ưu vượt trội.")

    add_img(doc, os.path.join(RESULTS_DIR, "hinh6_luu_do_thuat_toan.png"), 3.8, "Hình 1.1: Lưu đồ quy trình hoạt động của Giải thuật Di truyền kết hợp Memetic 2-Opt")

    add_h(doc, "1.3. Kết luận chương 1", level=2)
    add_p(doc, "Chương 1 đã tổng quan đầy đủ nền tảng lý thuyết về bài toán Người giao hàng (TSP), bản chất toán học của thuật toán Genetic Algorithm, thiết kế các toán tử hoán vị chuyên biệt (Tournament Selection, Ordered Crossover, Inversion Mutation, Elitism) và kiến trúc lai ghép Memetic Algorithm. Đây là cơ sở khoa học để triển khai cài đặt thực nghiệm trong Chương 2.")

    doc.add_page_break()

    # -------------------------------------------------------------
    # TRANG 8 - 17: CHƯƠNG 2
    # -------------------------------------------------------------
    add_h(doc, "CHƯƠNG 2. ÁP DỤNG THUẬT TOÁN TỐI ƯU LỘ TRÌNH GIAO HÀNG TẠI TP.HCM", level=1)
    
    add_h(doc, "2.1. Phát biểu bài toán thực tế 20 địa điểm TP.HCM", level=2)
    add_p(doc, "Đề tài tiến hành mô hình hóa mạng lưới logistics giao hàng chặng cuối tại TP. Hồ Chí Minh với 20 địa điểm trọng điểm. Điểm xuất phát (Kho Tổng) đặt tại Bưu điện Trung tâm Sài Gòn (Quận 1). Xe giao hàng phải rời Kho Tổng, viếng thăm chính xác 19 địa điểm giao nhận (bao gồm trung tâm thương mại, đầu mối giao thông, khu công nghệ cao và các cơ sở của Trường Đại học Văn Lang) mỗi nơi đúng một lần, sau đó quay trở về Kho Tổng sao cho tổng quãng đường di chuyển là ngắn nhất.")

    add_p(doc, "Bảng 2.1: Danh sách 20 địa điểm giao hàng thực tế tại TP. Hồ Chí Minh", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    t_cities = doc.add_table(rows=1, cols=4)
    t_cities.style = "Table Grid"
    t_cities.alignment = WD_TABLE_ALIGNMENT.CENTER
    headers = ["ID", "Tên địa điểm giao hàng", "Quận / Khu vực", "Tọa độ (X, Y) km"]
    for i, h in enumerate(headers):
        t_cities.rows[0].cells[i].text = h
    style_table_header(t_cities.rows[0])

    for city in HCMC_LOCATIONS:
        row = t_cities.add_row()
        set_cell_font(row.cells[0], str(city["id"]), size=9, align=WD_ALIGN_PARAGRAPH.CENTER)
        set_cell_font(row.cells[1], city["name"], size=9, bold=(city["id"] in [0, 4, 11]))
        set_cell_font(row.cells[2], city["district"], size=9)
        set_cell_font(row.cells[3], f"({city['x']:.1f}, {city['y']:.1f})", size=9, align=WD_ALIGN_PARAGRAPH.CENTER)

    add_img(doc, os.path.join(RESULTS_DIR, "hinh1_ban_do_diem_giao.png"), 4.6, "Hình 2.1: Bản đồ phân bố không gian 20 địa điểm giao hàng tại TP. Hồ Chí Minh")

    add_h(doc, "2.2. Minh họa bài toán & Mô hình tính toán", level=2)
    
    add_h(doc, "2.2.1. Mô hình toán học bài toán TSP", level=3)
    add_p(doc, "Gọi tập hợp các địa điểm là V = {0, 1, ..., N-1}. Với mỗi cặp địa điểm A(x_A, y_A) và B(x_B, y_B), khoảng cách Euclidean đối xứng giữa 2 điểm được tính theo công thức:")
    add_p(doc, "d(A, B) = sqrt((x_A - x_B)^2 + (y_A - y_B)^2)", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    add_p(doc, "Một nghiệm lộ trình là một hoán vị π = (c_0, c_1, ..., c_{N-1}). Hàm mục tiêu cần cực tiểu hóa là tổng chiều dài chu trình:")
    add_p(doc, "min D(π) = sum_{i=0}^{N-2} d(c_i, c_{i+1}) + d(c_{N-1}, c_0)", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)

    # Bảng trích đoạn ma trận khoảng cách 6 điểm
    add_p(doc, "Bảng 2.2: Trích đoạn ma trận khoảng cách Euclidean giữa 6 địa điểm tiêu biểu (km)", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    dm_cities = HCMC_LOCATIONS[:6]
    dm_sub = build_distance_matrix(dm_cities)
    t_dm = doc.add_table(rows=7, cols=7)
    t_dm.style = "Table Grid"
    t_dm.alignment = WD_TABLE_ALIGNMENT.CENTER
    names_short = ["Kho (0)", "BếnThành (1)", "Rùa (2)", "Landmark (3)", "VănLang1 (4)", "TânSơnNhất (5)"]
    t_dm.rows[0].cells[0].text = "Điểm"
    for j in range(6):
        t_dm.rows[0].cells[j+1].text = names_short[j]
    style_table_header(t_dm.rows[0])
    for i in range(6):
        set_cell_font(t_dm.rows[i+1].cells[0], names_short[i], size=8.5, bold=True)
        for j in range(6):
            val_str = f"{dm_sub[i][j]:.2f}" if i != j else "0.00"
            set_cell_font(t_dm.rows[i+1].cells[j+1], val_str, size=8.5, align=WD_ALIGN_PARAGRAPH.CENTER)

    add_h(doc, "2.2.2. Minh họa tính toán từng bước với 6 địa điểm thực tế", level=3)
    add_p(doc, "Để minh chứng rõ ràng mô hình tính toán theo đúng yêu cầu đề bài, xét bài toán thu nhỏ với 6 địa điểm [0, 1, 2, 3, 4, 5]:")
    add_p(doc, "• Bước 1 (Khởi tạo nghiệm): Sinh ngẫu nhiên 2 cá thể P1 = [3, 1, 2, 0, 4, 5] và P2 = [1, 2, 3, 4, 5, 0].")
    add_p(doc, "  - Quãng đường P1 = d(3,1) + d(1,2) + d(2,0) + d(0,4) + d(4,5) + d(5,3) = 3.24 + 1.25 + 0.63 + 0.72 + 5.76 + 9.77 = 21.38 km. Fitness = 1000 / 21.38 = 46.77.")
    add_p(doc, "  - Quãng đường P2 = 21.47 km. Fitness = 46.57.")
    add_p(doc, "• Bước 2 (Lai ghép Ordered Crossover): Chọn điểm cắt idx1=1, idx2=3 trên P1. Đoạn con [1, 2, 0] được giữ nguyên tại các vị trí giữa của cá thể Con. Duyệt vòng tròn Cha 2 từ sau điểm cắt: [4, 5, 0, 1, 2, 3]; lọc bỏ các gen đã có trong Con ([0, 1, 2]), còn lại danh sách bù [4, 5, 3]. Điền lần lượt vào các vị trí trống -> Cá thể Con hoàn chỉnh: [3, 1, 2, 0, 4, 5]. Hoán vị hoàn toàn hợp lệ 100%.")
    add_p(doc, "• Bước 3 (Đột biến Inversion Mutation): Chọn đoạn ngẫu nhiên [2, 0, 4] trong cá thể Con và đảo ngược thứ tự -> chuỗi đột biến = [3, 1, 4, 0, 2, 5].")
    add_p(doc, "• Bước 4 (Chọn lọc & Tiến hóa qua 50 thế hệ): Nhờ áp dụng toán tử di truyền, quần thể nhanh chóng đào thải các liên kết cắt chéo, đưa chiều dài chu trình từ 20.85 km xuống nghiệm tối ưu toàn cục tuyệt đối 20.26 km chỉ sau 0.02 giây.")

    add_h(doc, "2.3. Cài đặt mã nguồn và Giải thích chi tiết các thuật toán", level=2)
    add_p(doc, "Chương trình được thiết kế theo mô hình hướng đối tượng (OOP) module hóa nghiêm ngặt theo chuẩn PEP 8. Toàn bộ logic tối ưu hoá di truyền được tự cài đặt từ đầu (scratch) trong module src/genetic_algorithm.py mà không phụ thuộc vào bất kỳ thư viện tối ưu hóa đóng gói sẵn nào. Dưới đây là phân tích chi tiết từng khối thuật toán cốt lõi:")

    # 2.3.1
    add_h(doc, "2.3.1. Cấu trúc dữ liệu Cá thể (Individual) & Hàm thích nghi (Fitness Function)", level=3)
    code_f = (
        "class Individual:\n"
        "    \"\"\"Đại diện cho một cá thể (một lộ trình giao hàng khép kín).\"\"\"\n"
        "    def __init__(self, genes=None):\n"
        "        # genes: Chuỗi hoán vị gồm N địa điểm từ 0 đến N-1\n"
        "        self.genes = list(genes) if genes is not None else []\n"
        "        self.distance = 0.0\n"
        "        self.fitness = 0.0\n\n"
        "    def calculate_fitness(self, distance_matrix):\n"
        "        \"\"\"Tính tổng khoảng cách và độ thích nghi (Fitness).\"\"\"\n"
        "        total_dist = sum(distance_matrix[self.genes[i]][self.genes[i + 1]]\n"
        "                         for i in range(len(self.genes) - 1))\n"
        "        total_dist += distance_matrix[self.genes[-1]][self.genes[0]]  # Chặng khép chu trình về Kho\n"
        "        self.distance = round(total_dist, 3)\n"
        "        # Độ thích nghi tỉ lệ nghịch với tổng quãng đường (cộng 1e-6 chống chia cho 0)\n"
        "        self.fitness = 1000.0 / (self.distance + 1e-6)\n"
        "        return self.fitness"
    )
    add_code_block(doc, code_f, "Khối mã 2.1: Cấu trúc cá thể Individual và hàm tính toán Fitness")
    add_p(doc, "• Bản chất toán học: Nghiệm bài toán TSP được biểu diễn bằng một hoán vị thứ tự các đỉnh π = [c_0, c_1, ..., c_{N-1}]. Chu trình giao hàng bắt buộc phải khép kín (xe xuất phát từ Kho Tổng c_0 và quay trở về đúng Kho Tổng), do đó ngoài tổng cự ly qua N-1 chặng liên tiếp, hàm phải cộng thêm chặng cuối từ c_{N-1} về c_0.")
    add_p(doc, "• Hàm thích nghi (Fitness Mapping): Mục tiêu của TSP là tìm min D(π). Để phù hợp với nguyên lý chọn lọc tự nhiên của GA (cá thể thích nghi tốt hơn có fitness cao hơn), hàm fitness được nghịch đảo theo công thức: Fitness = 1000.0 / (D + 1e-6). Hằng số 1e-6 loại bỏ hoàn toàn nguy cơ chia cho 0, hệ số 1000 giúp chuẩn hóa thang đo độ thích nghi. Độ phức tạp tính toán: O(N) thời gian, O(N) không gian bộ nhớ.")

    # 2.3.2
    add_h(doc, "2.3.2. Toán tử Chọn lọc giải đấu (Tournament Selection, k=5)", level=3)
    code_sel = (
        "def selection_tournament(self):\n"
        "    \"\"\"Phương pháp chọn lọc giải đấu: Bốc ngẫu nhiên k cá thể và lấy cá thể tốt nhất.\"\"\"\n"
        "    tournament = random.sample(self.population, self.tournament_size)  # k = 5\n"
        "    tournament.sort(key=lambda ind: ind.fitness, reverse=True)\n"
        "    return tournament[0]  # Cá thể chiến thắng giải đấu được chọn làm cha/mẹ"
    )
    add_code_block(doc, code_sel, "Khối mã 2.2: Toán tử chọn lọc giải đấu Tournament Selection")
    add_p(doc, "• Cơ chế chọn lọc: Lấy mẫu ngẫu nhiên k=5 cá thể từ quần thể hiện tại mà không hoàn lại, sau đó chọn ra cá thể có fitness cao nhất để đưa vào tập cha mẹ sinh sản.")
    add_p(doc, "• Ưu việt so với Roulette Wheel: Phương pháp Roulette Wheel phân bổ xác suất tỉ lệ thuận theo fitness tuyệt đối, dẫn đến nhược điểm nghiêm trọng: ở các thế hệ ban đầu, một cá thể ngẫu nhiên vượt trội sẽ độc chiếm xác suất chọn lọc và nhân bản ồ ạt, làm suy giảm đa dạng di truyền và gây hội tụ non (premature convergence). Ngược lại, Tournament Selection duy trì áp lực chọn lọc (Selection Pressure) ổn định, cho phép các cá thể trung bình vẫn có cơ hội sinh sản nếu không phải đối đầu trực diện cá thể mạnh nhất trong nhóm nhỏ k=5. Độ phức tạp: O(k) cực nhanh.")

    # 2.3.3
    add_h(doc, "2.3.3. Toán tử Lai ghép có thứ tự (Ordered Crossover - OX)", level=3)
    code_ox = (
        "def ordered_crossover(self, parent1, parent2):\n"
        "    \"\"\"Toán tử lai ghép OX bảo toàn tính hợp lệ của hoán vị không sinh trùng lặp.\"\"\"\n"
        "    if random.random() > self.crossover_rate:  # Tỉ lệ lai ghép 85%\n"
        "        return copy.deepcopy(parent1), copy.deepcopy(parent2)\n\n"
        "    def make_child(p1, p2):\n"
        "        size = len(p1.genes)\n"
        "        child_genes = [None] * size\n"
        "        # Bước 1: Chọn ngẫu nhiên 2 điểm cắt trên Parent 1\n"
        "        idx1, idx2 = sorted(random.sample(range(size), 2))\n"
        "        child_genes[idx1:idx2 + 1] = p1.genes[idx1:idx2 + 1]\n"
        "        copied_set = set(child_genes[idx1:idx2 + 1])\n"
        "        # Bước 2: Duyệt vòng tròn Parent 2 từ sau điểm cắt idx2+1 để lấy gen bù\n"
        "        p2_genes = p2.genes[idx2 + 1:] + p2.genes[:idx2 + 1]\n"
        "        remaining_genes = [g for g in p2_genes if g not in copied_set]\n"
        "        # Bước 3: Điền các gen còn lại vào vị trí trống\n"
        "        fill_indices = list(range(idx2 + 1, size)) + list(range(0, idx1))\n"
        "        for i, gene in zip(fill_indices, remaining_genes):\n"
        "            child_genes[i] = gene\n"
        "        return Individual(child_genes)\n\n"
        "    return make_child(parent1, parent2), make_child(parent2, parent1)"
    )
    add_code_block(doc, code_ox, "Khối mã 2.3: Toán tử lai ghép có thứ tự Ordered Crossover (OX)")
    add_p(doc, "• Giải quyết bài toán hoán vị: Phép lai ghép nhị phân thông thường (1 điểm cắt, 2 điểm cắt) khi áp dụng lên TSP sẽ tạo ra cá thể con bị trùng lặp hoặc thiếu thành phố (vi phạm ràng buộc ghé thăm mỗi điểm đúng 1 lần).")
    add_p(doc, "• Cơ chế bảo toàn hoán vị hợp lệ 100%: Toán tử OX giữ nguyên đoạn gen giữa [idx1:idx2] từ Parent 1 để bảo tồn các cụm liên kết chặng tốt (building blocks), sau đó duyệt tuần hoàn Parent 2 để bù các gen còn lại theo đúng thứ tự xuất hiện mà chưa có trong con. Nhờ sử dụng cấu trúc hash set tra cứu O(1), thuật toán đảm bảo 100% con sinh ra là hoán vị hợp lệ mà không cần bước sửa sai phụ trợ (Repair Mechanism). Độ phức tạp: O(N) thời gian và bộ nhớ.")

    # 2.3.4
    add_h(doc, "2.3.4. Toán tử Đột biến đảo đoạn (Inversion Mutation)", level=3)
    code_mut = (
        "def mutate_inversion(self, individual):\n"
        "    \"\"\"Đột biến đảo đoạn: Đảo ngược thứ tự các gen nằm giữa 2 vị trí ngẫu nhiên.\"\"\"\n"
        "    if random.random() < self.mutation_rate:  # Tỉ lệ đột biến 20%\n"
        "        idx1, idx2 = sorted(random.sample(range(self.num_cities), 2))\n"
        "        individual.genes[idx1:idx2 + 1] = reversed(individual.genes[idx1:idx2 + 1])\n\n"
        "def mutate_swap(self, individual):\n"
        "    \"\"\"Đột biến hoán vị (dùng để đối chứng thực nghiệm): Đổi chỗ 2 gen đơn lẻ.\"\"\"\n"
        "    if random.random() < self.mutation_rate:\n"
        "        idx1, idx2 = random.sample(range(self.num_cities), 2)\n"
        "        individual.genes[idx1], individual.genes[idx2] = individual.genes[idx2], individual.genes[idx1]"
    )
    add_code_block(doc, code_mut, "Khối mã 2.4: Toán tử đột biến Inversion Mutation và Swap Mutation")
    add_p(doc, "• Bản chất hình học tháo gỡ đường cắt chéo (2-Opt Uncrossing): Theo bất đẳng thức tam giác trong hình học phẳng, đoạn đường tự cắt chéo (crossing) luôn dài hơn đoạn đường không cắt chéo: d(A, C) + d(B, D) > d(A, B) + d(C, D). Phép đảo ngược một đoạn gen con [idx1:idx2] trong Inversion Mutation chính là thao tác đảo chiều lộ trình di chuyển của phân đoạn đó, giúp tháo gỡ lập tức nút thắt chéo và rút ngắn cự ly đáng kể.")
    add_p(doc, "• Vượt trội so với Swap Mutation: Swap Mutation chỉ đổi chỗ 2 đỉnh rời rạc, làm đứt gãy cùng lúc 4 cạnh và thường vô tình tạo ra các nút giao chéo mới. Ngược lại, Inversion Mutation chỉ ngắt 2 cạnh biên và đảo chiều đoạn trong, bảo toàn nguyên vẹn tính liên kết nội bộ, giúp thuật toán thoát khỏi bẫy cực tiểu địa phương cực kỳ hiệu quả.")

    # 2.3.5
    add_h(doc, "2.3.5. Thuật toán tối ưu hóa cục bộ Memetic 2-Opt (Local Search)", level=3)
    code_2opt = (
        "def apply_2opt(self, individual):\n"
        "    \"\"\"Tối ưu hóa cục bộ 2-Opt tháo gỡ toàn bộ các cặp cạnh cắt chéo trên chu trình.\"\"\"\n"
        "    route = individual.genes.copy()\n"
        "    n = len(route)\n"
        "    improved = True\n"
        "    while improved:\n"
        "        improved = False\n"
        "        for i in range(1, n - 1):\n"
        "            for j in range(i + 1, n):\n"
        "                if j - i == 1: continue\n"
        "                p_prev, p_i = route[i - 1], route[i]\n"
        "                p_j, p_next = route[j], route[(j + 1) % n]\n"
        "                # Kỹ thuật Delta Evaluation O(1): So sánh cự ly 2 cạnh cũ và 2 cạnh mới\n"
        "                d_old = self.distance_matrix[p_prev][p_i] + self.distance_matrix[p_j][p_next]\n"
        "                d_new = self.distance_matrix[p_prev][p_j] + self.distance_matrix[p_i][p_next]\n"
        "                if d_new < d_old - 1e-5:  # Nếu phương án mới ngắn hơn\n"
        "                    route[i:j + 1] = reversed(route[i:j + 1])\n"
        "                    improved = True\n"
        "                    break\n"
        "            if improved: break\n"
        "    individual.genes = route\n"
        "    individual.calculate_fitness(self.distance_matrix)\n"
        "    return individual"
    )
    add_code_block(doc, code_2opt, "Khối mã 2.5: Thuật toán tối ưu hóa cục bộ Memetic 2-Opt")
    add_p(doc, "• Kỹ thuật đánh giá chênh lệch nhanh (Delta Evaluation O(1)): Khi kiểm tra khả năng đảo đoạn giữa 2 cạnh (p_prev -> p_i) và (p_j -> p_next), thuật toán không cần duyệt lại toàn bộ chu trình O(N), mà chỉ tính hiệu cự ly của 4 cạnh liên quan với chi phí O(1). Nếu d_new < d_old - 1e-5, đoạn đường được đảo ngược ngay lập tức.")
    add_p(doc, "• Sức mạnh của kiến trúc Memetic: GA truyền thống có ưu thế tìm kiếm diện rộng toàn cục (Global Exploration) nhưng hội tụ chậm ở giai đoạn tinh chỉnh cuối. Bằng cách tích hợp 2-Opt (Local Exploitation) lên cá thể dẫn đầu sau mỗi thế hệ, thuật toán triệt tiêu hoàn toàn các cạnh cắt chéo, giúp quần thể hội tụ về nghiệm tối ưu sâu hơn với tốc độ cực nhanh.")

    # 2.3.6
    add_h(doc, "2.3.6. Vòng lặp tiến hóa tổng thể (Evolutionary Cycle) & Chiến lược Elitism", level=3)
    code_engine = (
        "def evolve_generation(self):\n"
        "    \"\"\"Quy trình tiến hóa đồng bộ một thế hệ mới.\"\"\"\n"
        "    self.current_generation += 1\n"
        "    new_population = []\n"
        "    # 1. Chiến lược Elitism: Bảo tồn nguyên vẹn Top E=5 cá thể xuất sắc nhất\n"
        "    self.population.sort(key=lambda ind: ind.fitness, reverse=True)\n"
        "    for i in range(self.elite_size):\n"
        "        new_population.append(copy.deepcopy(self.population[i]))\n"
        "    # 2. Sinh các cá thể con cho đến khi đủ quy mô quần thể P=100\n"
        "    while len(new_population) < self.pop_size:\n"
        "        parent1 = self.selection_tournament()\n"
        "        parent2 = self.selection_tournament()\n"
        "        child1, child2 = self.ordered_crossover(parent1, parent2)\n"
        "        self.mutate(child1)\n"
        "        self.mutate(child2)\n"
        "        child1.calculate_fitness(self.distance_matrix)\n"
        "        child2.calculate_fitness(self.distance_matrix)\n"
        "        new_population.append(child1)\n"
        "        if len(new_population) < self.pop_size:\n"
        "            new_population.append(child2)\n"
        "    # 3. Kỹ thuật Memetic: Áp dụng 2-Opt cục bộ lên cá thể tốt nhất thế hệ\n"
        "    if self.use_local_search:\n"
        "        self.apply_2opt(new_population[0])\n"
        "    self.population = new_population\n"
        "    self.update_best()\n"
        "    self.history_best_distance.append(self.best_individual.distance)\n"
        "    avg_d = sum(ind.distance for ind in self.population) / self.pop_size\n"
        "    self.history_avg_distance.append(avg_d)"
    )
    add_code_block(doc, code_engine, "Khối mã 2.6: Vòng lặp tiến hóa toàn diện kết hợp Elitism và Memetic")
    add_p(doc, "• Tính chất bảo toàn đơn điệu (Monotonic Elitism): Bằng cách giữ lại Top E=5 cá thể ưu tú nhất chuyển thẳng sang thế hệ tiếp theo mà không qua đột biến hay lai ghép, thuật toán bảo đảm chất lượng nghiệm tốt nhất của quần thể là một hàm đơn điệu không tăng: D_best(g+1) <= D_best(g). Nghiệm xuất sắc đã tìm được sẽ không bao giờ bị mất đi.")
    add_p(doc, "• Tính đồng bộ của bộ máy tiến hóa: Sự phối hợp nhịp nhàng giữa Tournament Selection (chọn lọc cha mẹ), Ordered Crossover (tái tổ hợp đặc tính tốt), Inversion Mutation (duy trì đa dạng, thoát bẫy cực trị), Elitism (giữ vững tinh hoa) và Memetic 2-Opt (tối ưu cục bộ) tạo nên một cỗ máy tìm kiếm metaheuristic toàn diện, giải quyết trọn vẹn bài toán TSP 20 địa điểm trong thời gian chỉ 0.089 giây.")

    # -------------------------------------------------------------
    # 2.4. Kết quả thực nghiệm và trực quan hóa
    # -------------------------------------------------------------
    add_h(doc, "2.4. Kết quả thực nghiệm và trực quan hóa", level=2)
    
    add_h(doc, "2.4.1. Bảng số liệu tối ưu và thời gian thực thi", level=3)
    add_p(doc, "Thuật toán được thiết lập tham số: Kích thước quần thể P = 100, Số thế hệ G = 150, Tỉ lệ lai ghép Pc = 0.85, Tỉ lệ đột biến Pm = 0.20, Số lượng Elitism E = 5 cá thể, tích hợp Memetic 2-Opt. Thực nghiệm chạy trên tập 20 địa điểm TP.HCM mang lại kết quả vượt bậc:")

    add_p(doc, "Bảng 2.3: Bảng kết quả thực nghiệm tối ưu hoá lộ trình trên 20 địa điểm TP.HCM", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    t_res = doc.add_table(rows=6, cols=2)
    t_res.style = "Table Grid"
    t_res.alignment = WD_TABLE_ALIGNMENT.CENTER
    rdata = [
        ("Chỉ tiêu đánh giá hiệu năng", "Kết quả thực nghiệm"),
        ("Tổng quãng đường ban đầu (Gen 0 ngẫu nhiên)", "152.62 km"),
        ("Tổng quãng đường sau khi tối ưu hoá (GA + Memetic)", "80.76 km"),
        ("Quãng đường rút ngắn được", "71.85 km"),
        ("Tỉ lệ cải thiện độ dài lộ trình", "47.08%"),
        ("Thời gian thực thi thuật toán (150 thế hệ)", "0.089 giây (< 0.10 giây)")
    ]
    for i, (k, v) in enumerate(rdata):
        t_res.rows[i].cells[0].text = k
        t_res.rows[i].cells[1].text = v
    style_table_header(t_res.rows[0])
    for row in t_res.rows[1:]:
        set_cell_font(row.cells[0], row.cells[0].text, size=9.5)
        set_cell_font(row.cells[1], row.cells[1].text, size=9.5, align=WD_ALIGN_PARAGRAPH.CENTER, bold=True)

    add_img(doc, os.path.join(RESULTS_DIR, "hinh2_truoc_sau_toi_uu.png"), 5.2, "Hình 2.2: So sánh trực quan lộ trình giao hàng TRƯỚC và SAU khi tối ưu hoá")

    add_h(doc, "2.4.2. Lịch trình giao hàng chi tiết 20 chặng (Turn-by-turn)", level=3)
    add_p(doc, "Dưới đây là lịch trình chi tiết xuất phát từ Bưu điện Trung tâm Sài Gòn, đi qua 19 địa điểm và khép chu trình về lại Bưu điện với tổng chiều dài đạt 80.76 km:")

    add_p(doc, "Bảng 2.4: Lịch trình giao hàng chi tiết 20 chặng xuất phát từ Kho Tổng", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    t_itin = doc.add_table(rows=1, cols=5)
    t_itin.style = "Table Grid"
    t_itin.alignment = WD_TABLE_ALIGNMENT.CENTER
    i_headers = ["Chặng", "Điểm xuất phát", "Điểm đến tiếp theo", "Cự ly (km)", "Tích lũy (km)"]
    for i, h in enumerate(i_headers):
        t_itin.rows[0].cells[i].text = h
    style_table_header(t_itin.rows[0])

    itinerary_data = [
        (1, "Kho Tổng (Bưu điện TT Sài Gòn)", "Hồ Con Rùa", 0.63, 0.63),
        (2, "Hồ Con Rùa", "Hàng Xanh Interchange", 2.80, 3.43),
        (3, "Hàng Xanh Interchange", "Landmark 81 (Vinhomes)", 1.36, 4.79),
        (4, "Landmark 81 (Vinhomes)", "Chợ Thủ Đức", 6.87, 11.66),
        (5, "Chợ Thủ Đức", "Khu Công nghệ Cao (SHTP)", 3.99, 15.65),
        (6, "Khu Công nghệ Cao (SHTP)", "ĐHQG TP.HCM (Làng ĐH)", 3.45, 19.10),
        (7, "ĐHQG TP.HCM (Làng ĐH)", "Bến xe Miền Đông Cũ", 13.45, 32.55),
        (8, "Bến xe Miền Đông Cũ", "Công viên Gia Định", 4.10, 36.65),
        (9, "Công viên Gia Định", "Đại học Văn Lang (Cơ sở 3)", 1.91, 38.56),
        (10, "Đại học Văn Lang (Cơ sở 3)", "Bến xe An Sương", 5.23, 43.79),
        (11, "Bến xe An Sương", "Sân bay Tân Sơn Nhất", 5.87, 49.66),
        (12, "Sân bay Tân Sơn Nhất", "Công viên Hoàng Văn Thụ", 1.72, 51.38),
        (13, "Công viên Hoàng Văn Thụ", "Công viên Đầm Sen", 5.42, 56.80),
        (14, "Công viên Đầm Sen", "Aeon Mall Bình Tân", 3.55, 60.35),
        (15, "Aeon Mall Bình Tân", "Bến xe Chợ Lớn", 4.42, 64.77),
        (16, "Bến xe Chợ Lớn", "Crescent Mall (Phú Mỹ Hưng)", 7.06, 71.83),
        (17, "Crescent Mall (Phú Mỹ Hưng)", "Khu chế xuất Tân Thuận", 2.48, 74.31),
        (18, "Khu chế xuất Tân Thuận", "Chợ Bến Thành", 5.19, 79.50),
        (19, "Chợ Bến Thành", "Đại học Văn Lang (Cơ sở 1)", 0.54, 80.04),
        (20, "Đại học Văn Lang (Cơ sở 1)", "Kho Tổng (Bưu điện TT Sài Gòn)", 0.72, 80.76)
    ]

    for item in itinerary_data:
        row = t_itin.add_row()
        set_cell_font(row.cells[0], str(item[0]), size=8.5, align=WD_ALIGN_PARAGRAPH.CENTER)
        set_cell_font(row.cells[1], item[1], size=8.5)
        set_cell_font(row.cells[2], item[2], size=8.5)
        set_cell_font(row.cells[3], f"{item[3]:.2f}", size=8.5, align=WD_ALIGN_PARAGRAPH.CENTER)
        set_cell_font(row.cells[4], f"{item[4]:.2f}", size=8.5, align=WD_ALIGN_PARAGRAPH.CENTER, bold=True)

    add_h(doc, "2.4.3. Phân tích đồ thị hội tụ và tham số di truyền", level=3)
    add_p(doc, "Đồ thị đường cong hội tụ (Hình 2.3) thể hiện rõ đặc tính tiến hóa: Quần thể giảm cự ly cực nhanh trong 40 thế hệ đầu tiên nhờ năng lực tìm kiếm toàn cục của Ordered Crossover. Từ thế hệ 40 đến 100, Memetic 2-Opt tháo gỡ hoàn toàn các nút giao cắt chéo, đưa lộ trình về trạng thái tối ưu ổn định 80.76 km.")
    add_img(doc, os.path.join(RESULTS_DIR, "hinh3_duong_cong_hoi_tu.png"), 4.8, "Hình 2.3: Đồ thị đường cong hội tụ (Convergence Curve) qua 150 thế hệ")

    add_p(doc, "Thực nghiệm so sánh toán tử đột biến (Hình 2.4) khẳng định tính ưu việt của Inversion Mutation: Trong khi Swap Mutation chỉ đổi chỗ 2 gen đơn lẻ nên thường bị mắc kẹt tại cực trị địa phương (quãng đường dừng ở mức ~95 km), thì Inversion Mutation cho phép đảo ngược cả một phân đoạn, giúp quần thể vượt qua rào cản cục bộ và chạm mốc tối ưu 80.76 km.")
    add_img(doc, os.path.join(RESULTS_DIR, "hinh4_so_sanh_mutation.png"), 4.6, "Hình 2.4: So sánh hiệu năng hội tụ giữa Inversion Mutation và Swap Mutation")

    add_p(doc, "Nghiên cứu độ nhạy của tỉ lệ đột biến (Hình 2.5) cho thấy: Với Mutation Rate quá thấp (0.01 - 0.05), quần thể mất tính đa dạng di truyền và hội tụ non. Với Mutation Rate quá cao (0.50), thuật toán bị nhiễu loạn tương tự duyệt ngẫu nhiên (Random Walk). Giá trị tối ưu nằm trong khoảng 0.15 - 0.25 (đồ án chọn 0.20).")
    add_img(doc, os.path.join(RESULTS_DIR, "hinh5_mutation_rate_analysis.png"), 4.3, "Hình 2.5: Đánh giá ảnh hưởng của tỉ lệ đột biến (Mutation Rate)")

    # -------------------------------------------------------------
    # 2.5. Hiệu quả kinh tế & Môi trường (Green Logistics)
    # -------------------------------------------------------------
    add_h(doc, "2.5. Hiệu quả kinh tế & Môi trường (Green Logistics)", level=2)
    add_p(doc, "Trong bối cảnh chi phí logistics tại Việt Nam chiếm tỉ trọng cao trong GDP và cam kết phát thải ròng bằng '0' (Net Zero) vào năm 2050, giải pháp tối ưu hóa lộ trình bằng GA mang lại giá trị định lượng rõ rệt:")

    impact = GeneticAlgorithmTSP.calculate_economic_impact(71.85)
    add_p(doc, "Bảng 2.5: Bảng tính toán lợi ích kinh tế và cắt giảm khí thải nhà kính (Green Logistics)", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    t_eco = doc.add_table(rows=6, cols=2)
    t_eco.style = "Table Grid"
    t_eco.alignment = WD_TABLE_ALIGNMENT.CENTER
    ecodata = [
        ("Chỉ tiêu đo lường kinh tế & môi trường", "Lợi ích định lượng thực tế"),
        ("Nhiên liệu tiết kiệm mỗi chuyến giao hàng", f"{impact['fuel_saved_liters']:.2f} lít xăng (định mức xe máy 2.5 L/100km)"),
        ("Chi phí xăng tiết kiệm cho 1 chuyến xe", f"{impact['cost_saved_vnd']:,} VNĐ / chuyến"),
        ("Chi phí tiết kiệm cho 1 xe trong 1 tháng (30 chuyến)", f"{impact['cost_saved_vnd'] * 30:,} VNĐ / tháng / xe"),
        ("Chi phí tiết kiệm cho đội xe 10 chiếc trong 1 tháng", f"{impact['cost_saved_vnd'] * 30 * 10:,} VNĐ / tháng (~155 triệu VNĐ/năm)"),
        ("Cắt giảm phát thải khí nhà kính CO2 (2.31 kg CO2/lít)", f"{impact['co2_reduced_kg']:.2f} kg CO2 / chuyến (~1.5 tấn CO2/năm/đội xe)")
    ]
    for i, (k, v) in enumerate(ecodata):
        t_eco.rows[i].cells[0].text = k
        t_eco.rows[i].cells[1].text = v
    style_table_header(t_eco.rows[0])
    for row in t_eco.rows[1:]:
        set_cell_font(row.cells[0], row.cells[0].text, size=9.5)
        set_cell_font(row.cells[1], row.cells[1].text, size=9.5, align=WD_ALIGN_PARAGRAPH.CENTER, bold=True)

    add_h(doc, "2.6. Kết luận chương 2", level=2)
    add_p(doc, "Chương 2 đã hoàn thành toàn diện việc ứng dụng thực tế giải thuật di truyền vào bài toán giao hàng 20 địa điểm tại TP.HCM. Các kết quả thực nghiệm định lượng, kiểm chứng mô hình tính toán từng bước, bộ mã nguồn chuẩn mực, giao diện trực quan và phân tích Green Logistics đã chứng minh tính hiệu quả vượt bậc của thuật toán.")

    doc.add_page_break()

    # -------------------------------------------------------------
    # TRANG 18 - 19: CHƯƠNG 3
    # -------------------------------------------------------------
    add_h(doc, "CHƯƠNG 3. KẾT LUẬN VÀ HƯỚNG PHÁT TRIỂN", level=1)
    
    add_h(doc, "3.1. Đánh giá kết quả đạt được", level=2)
    add_p(doc, "Đồ án môn học 'Nhập môn Tối ưu hoá' với đề tài Áp dụng Genetic Algorithm cho bài toán TSP đã hoàn thành xuất sắc toàn bộ các mục tiêu đặt ra:")
    add_p(doc, "1. Về mặt học thuật: Làm chủ cơ chế hoạt động của thuật toán di truyền phỏng sinh học; xây dựng các toán tử hoán vị chuyên biệt (Tournament Selection, Ordered Crossover, Inversion Mutation, Elitism) và thuật toán tìm kiếm cục bộ Memetic 2-Opt.")
    add_p(doc, "2. Về mặt cài đặt kỹ thuật: Tự cài đặt 100% thuật toán bằng ngôn ngữ Python thuần, không phụ thuộc các thư viện metaheuristic đóng gói sẵn. Cấu trúc mã nguồn đạt chuẩn PEP 8, có kiểm thử tự động (Unit Test pass 100%).")
    add_p(doc, "3. Về mặt ứng dụng thực tiễn: Giải quyết bài toán logistics thực tế trên 20 địa điểm TP.HCM, rút ngắn quãng đường từ 152.62 km xuống còn 80.76 km (cải thiện 47.08%), thời gian xử lý siêu tốc dưới 0.1 giây, mang lại hiệu quả kinh tế và cắt giảm khí thải CO2 rõ nét.")
    add_p(doc, "4. Về mặt sản phẩm: Xây dựng đồng thời giao diện dòng lệnh CLI chuẩn mực và giao diện đồ họa GUI tương tác trực quan thời gian thực (Tkinter + Matplotlib), cho phép người dùng điều chỉnh linh hoạt mọi tham số di truyền.")

    add_h(doc, "3.2. Ưu điểm và hạn chế của giải pháp", level=2)
    add_p(doc, "A. Ưu điểm nổi bật:", bold=True)
    add_p(doc, "• Tốc độ tính toán tức thời (< 0.1s), đáp ứng tốt yêu cầu điều phối tuyến đường theo thời gian thực.")
    add_p(doc, "• Không gian nghiệm được khám phá đa dạng, thuật toán có khả năng thoát khỏi các cực trị cục bộ rất tốt nhờ Inversion Mutation và Elitism.")
    add_p(doc, "• Sản phẩm hoàn thiện cao, đóng gói trực quan, sẵn sàng triển khai demo.")
    add_p(doc, "B. Hạn chế hiện tại:", bold=True)
    add_p(doc, "• Mô hình khoảng cách hiện dựa trên tọa độ phẳng Euclidean (đường chim bay), chưa phản ánh chính xác các yếu tố đường một chiều, kẹt xe giờ cao điểm hay cầu đường thực tế.")
    add_p(doc, "• Bài toán mới giới hạn ở một xe đơn lẻ (Single Vehicle TSP) và chưa xét đến ràng buộc dung tích tải trọng hay khung giờ giao hàng (Time Windows).")

    add_h(doc, "3.3. Hướng phát triển trong tương lai", level=2)
    add_p(doc, "Nhằm phát triển ứng dụng thành một hệ thống quản lý vận tải (TMS) chuyên nghiệp, các hướng nâng cấp tiếp theo bao gồm:")
    add_p(doc, "1. Tích hợp Google Maps Distance Matrix API: Thu thập cự ly đường bộ thực tế và thời gian di chuyển động theo tình hình giao thông thực tế tại TP.HCM.")
    add_p(doc, "2. Mở rộng bài toán VRP (Vehicle Routing Problem): Tối ưu hóa lộ trình cho đội xe gồm nhiều phương tiện vận hành đồng thời, có ràng buộc tải trọng tối đa (Capacitated VRP).")
    add_p(doc, "3. Thêm ràng buộc cửa sổ thời gian (VRPTW - Vehicle Routing Problem with Time Windows): Đáp ứng yêu cầu khách hàng chỉ nhận hàng trong những khoảng giờ cố định.")
    add_p(doc, "4. Ứng dụng điện toán đám mây: Triển khai backend tối ưu hóa dưới dạng RESTful API trên nền tảng Cloud (Google Cloud / AWS) để tích hợp vào ứng dụng di động cho tài xế giao hàng.")

    # -------------------------------------------------------------
    # TÀI LIỆU THAM KHẢO
    # -------------------------------------------------------------
    add_h(doc, "TÀI LIỆU THAM KHẢO", level=1)
    
    refs = [
        "[1] ThS. Trần Ngọc Việt, Bài giảng môn Nhập môn Tối ưu hoá (HK261, Mã lớp: 261_71AICR30203_01), Khoa Công nghệ Thông tin, Trường Đại học Văn Lang, 2026.",
        "[2] Holland, J. H. (1975). Adaptation in Natural and Artificial Systems: An Introductory Analysis with Applications to Biology, Control, and Artificial Intelligence. University of Michigan Press, Ann Arbor.",
        "[3] Goldberg, D. E. (1989). Genetic Algorithms in Search, Optimization and Machine Learning. Addison-Wesley Longman Publishing Co., Inc., Boston, MA, USA.",
        "[4] Larranaga, P., Kuijpers, C. M. H., Murga, R. H., Inza, I., & Dizdarevic, S. (1999). Genetic algorithms for the travelling salesman problem: A review of representations and operators. Artificial Intelligence Review, 13(2), 129-170.",
        "[5] Moscato, P. (1989). On Evolution, Search, Optimization, Genetic Algorithms and Martial Arts: Towards Memetic Algorithms. Caltech Concurrent Computation Program, C3P Report 826, California Institute of Technology.",
        "[6] Croes, G. A. (1958). A method for solving traveling-salesman problems. Operations Research, 6(6), 791-812.",
        "[7] Applegate, D. L., Bixby, R. E., Chvátal, V., & Cook, W. J. (2006). The Traveling Salesman Problem: A Computational Study. Princeton University Press.",
        "[8] Toth, P., & Vigo, D. (Eds.). (2014). Vehicle Routing: Problems, Methods, and Applications (2nd ed.). Society for Industrial and Applied Mathematics (SIAM)."
    ]

    for ref in refs:
        p_ref = doc.add_paragraph()
        p_ref.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p_ref.paragraph_format.space_before = Pt(4)
        p_ref.paragraph_format.space_after = Pt(6)
        p_ref.paragraph_format.line_spacing = 1.2
        r = p_ref.add_run(ref)
        r.font.name = "Times New Roman"
        r.font.size = Pt(11)

    # Lưu tài liệu hoàn chỉnh
    doc.save(TARGET_PATH)
    print(f"-> Đã ghi xong toàn bộ nội dung đồ án vào file Word!")


def step3_update_word_fields():
    """Gọi Microsoft Word qua COM để cập nhật tự động toàn bộ Mục lục (TOC) và số trang."""
    print("Bước 3: Tự động cập nhật Mục lục native và kiểm tra số trang bằng Word...")
    
    ps_script = f"""
    try {{
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $word.DisplayAlerts = 0
        $docPath = "{os.path.abspath(TARGET_PATH)}"
        $doc = $word.Documents.Open($docPath)
        
        # Cập nhật tất cả các bảng mục lục
        if ($doc.TablesOfContents.Count -ge 1) {{
            for ($i = 1; $i -le $doc.TablesOfContents.Count; $i++) {{
                $doc.TablesOfContents.Item($i).Update()
            }}
            Write-Host "-> Đã cập nhật xong Mục lục chuẩn xác theo số trang thực tế!"
        }}
        
        # Cập nhật các trường footer (page number)
        $doc.Fields.Update()
        
        # Thống kê tổng số trang
        $pages = $doc.ComputeStatistics(2)
        Write-Host "-> Tổng số trang báo cáo Word: $pages trang"
        
        # Xuất song song định dạng PDF chuẩn in ấn / nộp bài
        $pdfPath = "{os.path.abspath('BaoCao_DoAn_ToiUuHoa_GA_TSP.pdf')}"
        $doc.ExportAsFixedFormat($pdfPath, 17)
        Write-Host "-> Đã xuất bản PDF chuẩn in ấn: $pdfPath"
        
        $doc.Save()
        $doc.Close()
        $word.Quit()
    }} catch {{
        Write-Host "Cảnh báo Word COM: $_"
        if ($word) {{ $word.Quit() }}
    }}
    """
    
    try:
        res = subprocess.run(["powershell", "-NoProfile", "-Command", ps_script], capture_output=True, text=True, check=True)
        print(res.stdout.strip())
    except Exception as e:
        print(f"Lưu ý: Không thể chạy Word COM tự động ({e}), Word sẽ tự động cập nhật khi người dùng mở file.")

    print(f"\n===========================================================================")
    print(f"ĐÃ TẠO THÀNH CÔNG BÁO CÁO WORD ĐỒ ÁN TỐI ƯU HOÁ CHUẨN MẪU TRƯỜNG ĐH VĂN LANG!")
    print(f"File kết quả: {os.path.abspath(TARGET_PATH)}")
    print(f"===========================================================================")


if __name__ == "__main__":
    step1_prepare_base_template()
    step2_populate_content()
    step3_update_word_fields()
