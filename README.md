# ĐỒ ÁN MÔN HỌC: NHẬP MÔN TỐI ƯU HOÁ (HK261)
**Giảng viên hướng dẫn:** ThS. Trần Ngọc Việt  
**Đề tài 4:** Tìm hiểu và cài đặt ứng dụng bài toán tối ưu Genetic Algorithm (GA)  
**Bài toán thực tế:** Tối ưu hoá lộ trình giao hàng (TSP) tại TP. Hồ Chí Minh

---

## 1. Cấu trúc thư mục đồ án hoàn chỉnh

- `run_app.bat`: Menu chạy nhanh 1-click trên Windows (chọn GUI, CLI hoặc Test).
- `app_gui.py`: Giao diện đồ họa tương tác (Tkinter + Matplotlib) theo dõi hội tụ và đường đi thời gian thực.
- `run_cli.py`: Chương trình dòng lệnh có đầy đủ Input, Output và các bước tính toán theo Rubric.
- `step_by_step_demo.py`: Minh họa chi tiết 7 bước tính toán bằng số liệu cụ thể (phục vụ thuyết minh/báo cáo).
- `test_ga.py`: Kịch bản kiểm thử thuật toán tự động.
- `src/`:
  - `cities_data.py`: Tọa độ thực tế 20 địa điểm tại TP.HCM và hàm tính ma trận khoảng cách.
  - `genetic_algorithm.py`: Cài đặt đầy đủ thuật toán GA (Khởi tạo, Fitness, Tournament, Ordered Crossover, Inversion Mutation, Elitism).
- `results/`: Các biểu đồ kết quả độ phân giải cao phục vụ chèn vào báo cáo và slide.
- `BaoCao_DoAn_ToiUuHoa_GA_TSP.docx`: Báo cáo đồ án môn học hoàn chỉnh (đúng cấu trúc, font Times New Roman 13pt).
- `BM-006- Đồ án MH - NM Tối ưu hoá- HK 261.docx.pdf`: Đề bài và rubric gốc của môn học.
- `TRANG BÌA - ĐAMH -TỐI ƯU HÓA HK261.docx`: Mẫu trang bìa gốc của trường.

---

## 2. Hướng dẫn chạy chương trình

### Cách 1: Chạy bằng file Menu (Khuyên dùng)
Nhấp đúp chuột vào file `run_app.bat`:
- Nhấn phím `1`: Mở giao diện đồ họa trực quan (GUI).
- Nhấn phím `2`: Mở chế độ dòng lệnh (CLI).
- Nhấn phím `3`: Chạy kiểm thử thuật toán.

### Cách 2: Chạy trực tiếp qua dòng lệnh Python
```powershell
python app_gui.py            # Giao diện đồ họa (GUI)
python run_cli.py            # Dòng lệnh Input/Output (CLI)
python step_by_step_demo.py  # Minh họa 7 bước tính toán
```

---

## 3. Lưu ý trước khi nộp bài
Mở file `BaoCao_DoAn_ToiUuHoa_GA_TSP.docx` để điền chính xác họ tên sinh viên và MSSV tại trang bìa trước khi nộp.
