# ĐỒ ÁN MÔN HỌC: NHẬP MÔN TỐI ƯU HOÁ (HK261)
**Trường Đại học Văn Lang – Khoa Công nghệ Thông tin**  
**Giảng viên hướng dẫn:** ThS. Trần Ngọc Việt  
**Sinh viên thực hiện:** Wendy Huy (MSSV: 2474802010135)  
**Đề tài 4:** Tìm hiểu và cài đặt ứng dụng bài toán tối ưu Genetic Algorithm (GA)  
**Bài toán ứng dụng thực tế:** Tối ưu hoá lộ trình giao hàng (TSP) tại TP. Hồ Chí Minh kết hợp Green Logistics

---

## 1. Điểm nổi bật & Cải tiến của Đồ án

1. **Thuật toán cốt lõi tối ưu chuyên biệt cho TSP:**
   * Cài đặt hoàn chỉnh từ đầu (from scratch) bằng Python thuần.
   * Toán tử lai ghép có thứ tự **Ordered Crossover (OX)** đảm bảo 100% nghiệm hoán vị hợp lệ.
   * Toán tử đột biến đảo đoạn **Inversion Mutation (2-Opt style)** có khả năng tháo gỡ nút giao chéo vượt trội so với Swap Mutation.
   * Cơ chế **Elitism (Bảo tồn ưu tú)** duy trì các cá thể xuất sắc nhất qua từng thế hệ.

2. **Cải tiến nâng cao – Thuật toán Memetic (Hybrid GA + 2-Opt Local Search):**
   * Kết hợp tìm kiếm toàn cục của GA với tìm kiếm cực tiểu cục bộ của 2-Opt.
   * Đẩy nhanh tốc độ hội tụ và đạt chất lượng nghiệm tối ưu sâu hơn (~80.76 km).

3. **Gắn liền bài toán kinh tế & Môi trường (Green Logistics):**
   * Tự động định lượng: Lượng xăng tiết kiệm (lít), chi phí tiết kiệm (VNĐ/chuyến & theo tháng), lượng khí thải $\text{CO}_2$ cắt giảm, và thời gian di chuyển giảm thiểu.
   * Tính năng xuất file lịch trình giao hàng chi tiết từng chặng (*Turn-by-turn itinerary*).

4. **Trực quan hoá đồ hoạ đa dạng:**
   * **Giao diện GUI (Tkinter + Matplotlib):** Đồ họa động theo thời gian thực (realtime animation), thanh trượt tùy chỉnh tham số, nút bấm Tạm dừng/Tiếp tục/Xuất lịch trình.
   * **Giao diện CLI:** Đáp ứng đầy đủ tiêu chí Rubric chấm điểm (Input, Output, Bảng tiến trình).

---

## 2. Cấu trúc thư mục dự án

```text
D:\DOAN_toiuu\
│
├── src\
│   ├── cities_data.py                  # Dữ liệu tọa độ 20 điểm TP.HCM + Ma trận khoảng cách
│   ├── genetic_algorithm.py            # Thuật toán GA, Memetic 2-Opt, Green Logistics
│   └── __init__.py
│
├── app_gui.py                          # Giao diện đồ họa tương tác (GUI)
├── run_cli.py                          # Chương trình dòng lệnh đầy đủ Input/Output (CLI)
├── step_by_step_demo.py                # Minh họa 7 bước tính toán bằng số liệu cụ thể
├── test_ga.py                          # Kịch bản kiểm thử thuật toán tự động
├── run_app.bat                         # Menu chạy nhanh 1-click trên Windows
│
├── BaoCao_DoAn_ToiUuHoa_GA_TSP.docx    # Báo cáo Word chính thức hoàn chỉnh (~18-20 trang)
├── README.md                           # Tài liệu hướng dẫn sử dụng và giới thiệu đồ án
│
├── results\                            # Các biểu đồ kết quả độ nét cao & Lịch trình
│   ├── hinh1_ban_do_diem_giao.png      # Bản đồ phân bố 20 điểm
│   ├── hinh2_truoc_sau_toi_uu.png      # So sánh lộ trình trước và sau tối ưu
│   ├── hinh3_duong_cong_hoi_tu.png     # Đồ thị đường cong hội tụ qua 200 thế hệ
│   ├── hinh4_so_sanh_mutation.png      # So sánh thực nghiệm Inversion vs Swap
│   ├── hinh5_mutation_rate_analysis.png# Phân tích ảnh hưởng tỉ lệ đột biến
│   ├── hinh6_luu_do_thuat_toan.png     # Lưu đồ thuật toán GA
│   ├── ket_qua_toi_uu_ga_tsp.png       # Biểu đồ lộ trình & hội tụ tổng hợp
│   └── lich_trinh_chi_tiet.txt         # Bảng lịch trình chi tiết từng chặng (Turn-by-turn)
│
├── BM-006- Đồ án MH - NM Tối ưu hoá- HK 261.docx.pdf # Đề bài & rubric môn học
└── TRANG BÌA - ĐAMH -TỐI ƯU HÓA HK261.docx           # Mẫu bìa gốc của trường
```

---

## 3. Hướng dẫn chạy chương trình

### Cách 1: Sử dụng menu nhanh 1-click (Khuyên dùng)
Nhấp đúp chuột vào file [`run_app.bat`](file:///D:/DOAN_toiuu/run_app.bat):
* **Phím 1:** Mở giao diện đồ họa trực quan (GUI).
* **Phím 2:** Chạy chế độ dòng lệnh (CLI).
* **Phím 3:** Chạy kiểm thử tự động thuật toán.

### Cách 2: Chạy trực tiếp bằng dòng lệnh
```powershell
python app_gui.py            # Chạy giao diện đồ họa tương tác
python run_cli.py            # Chạy giao diện console Input/Output
python step_by_step_demo.py  # Chạy minh họa từng bước tính toán chi tiết
```
