# -*- coding: utf-8 -*-
"""
app_gui.py - Giao diện đồ họa (GUI) trực quan hóa Giải thuật Di truyền cho bài toán TSP.
Sử dụng thư viện chuẩn Tkinter và Matplotlib (có sẵn trong Python).
Cho phép người dùng tương tác, điều chỉnh tham số, quan sát lộ trình và đồ thị hội tụ động theo thời gian thực.
Tích hợp:
  - Tùy chọn thuật toán Memetic (Hybrid GA + 2-Opt Local Search)
  - Phân tích hiệu quả kinh tế & môi trường (Green Logistics)
  - Xuất lịch trình giao hàng chi tiết turn-by-turn ra file văn bản
"""

import sys
import os
import time
import random
import threading
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

# Đảm bảo UTF-8
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

# Thiết lập thư mục làm việc
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src.cities_data import HCMC_LOCATIONS, build_distance_matrix
from src.genetic_algorithm import GeneticAlgorithmTSP

import matplotlib
matplotlib.use("TkAgg")
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure

class GA_TSP_App:
    def __init__(self, root):
        self.root = root
        self.root.title("Tối ưu hoá Lộ trình Giao hàng (TSP TP.HCM) - Giải thuật Di truyền (GA)")
        self.root.geometry("1320x820")
        self.root.minsize(1100, 700)
        
        # Biến trạng thái
        self.is_running = False
        self.is_paused = False
        self.cities = HCMC_LOCATIONS
        self.dist_matrix = build_distance_matrix(self.cities)
        self.ga = None
        self.start_time = 0
        
        # Cấu hình giao diện
        self._setup_ui()
        self.reset_simulation()

    def _setup_ui(self):
        # Frame chính chia làm 2 cột: Trái (Điều khiển & Thống kê), Phải (Đồ thị)
        main_frame = ttk.Frame(self.root, padding=8)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # ==================== CỘT TRÁI: ĐIỀU KHIỂN & SỐ LIỆU ====================
        left_container = ttk.Frame(main_frame, width=380)
        left_container.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))

        # Canvas có thanh cuộn cho bảng điều khiển bên trái
        canvas_scroll = tk.Canvas(left_container, borderwidth=0, highlightthickness=0, width=360)
        scrollbar = ttk.Scrollbar(left_container, orient=tk.VERTICAL, command=canvas_scroll.yview)
        control_panel = ttk.Frame(canvas_scroll)
        
        control_panel.bind("<Configure>", lambda e: canvas_scroll.configure(scrollregion=canvas_scroll.bbox("all")))
        canvas_scroll.create_window((0, 0), window=control_panel, anchor="nw")
        canvas_scroll.configure(yscrollcommand=scrollbar.set)
        
        canvas_scroll.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 1. Chọn nguồn dữ liệu
        lbl_data = ttk.Label(control_panel, text="1. Nguồn dữ liệu bài toán:", font=("Segoe UI", 9, "bold"))
        lbl_data.pack(anchor=tk.W, pady=(0, 2))

        self.dataset_var = tk.StringVar(value="hcmc")
        r1 = ttk.Radiobutton(control_panel, text="20 Địa điểm thực tế tại TP.HCM", value="hcmc", 
                             variable=self.dataset_var, command=self.on_dataset_change)
        r1.pack(anchor=tk.W, padx=8)
        
        r2 = ttk.Radiobutton(control_panel, text="Sinh ngẫu nhiên N địa điểm", value="random", 
                             variable=self.dataset_var, command=self.on_dataset_change)
        r2.pack(anchor=tk.W, padx=8, pady=(2, 4))

        # Khung nhập N điểm ngẫu nhiên
        self.rand_frame = ttk.Frame(control_panel)
        ttk.Label(self.rand_frame, text="Số điểm N:").pack(side=tk.LEFT, padx=(8, 4))
        self.rand_n_entry = ttk.Spinbox(self.rand_frame, from_=8, to=60, width=5)
        self.rand_n_entry.set(15)
        self.rand_n_entry.pack(side=tk.LEFT)
        btn_gen = ttk.Button(self.rand_frame, text="Sinh lại", command=self.generate_random_dataset)
        btn_gen.pack(side=tk.LEFT, padx=4)

        ttk.Separator(control_panel, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=6)

        # 2. Tham số Giải thuật Di truyền
        lbl_params = ttk.Label(control_panel, text="2. Cấu hình thuật toán & tham số:", font=("Segoe UI", 9, "bold"))
        lbl_params.pack(anchor=tk.W, pady=(0, 2))

        # Tùy chọn Memetic (Hybrid GA + 2-Opt)
        self.memetic_var = tk.BooleanVar(value=False)
        chk_memetic = ttk.Checkbutton(control_panel, text="Kích hoạt Memetic (GA + 2-Opt Local Search)", 
                                      variable=self.memetic_var, command=self.reset_simulation)
        chk_memetic.pack(anchor=tk.W, padx=4, pady=(0, 4))

        # Số thế hệ (Generations)
        ttk.Label(control_panel, text="Số thế hệ tối đa:").pack(anchor=tk.W)
        self.gen_scale = tk.Scale(control_panel, from_=50, to=500, orient=tk.HORIZONTAL, resolution=10)
        self.gen_scale.set(150)
        self.gen_scale.pack(fill=tk.X, pady=(0, 4))

        # Kích thước quần thể (Population Size)
        ttk.Label(control_panel, text="Kích thước quần thể (Pop Size):").pack(anchor=tk.W)
        self.pop_scale = tk.Scale(control_panel, from_=20, to=300, orient=tk.HORIZONTAL, resolution=10)
        self.pop_scale.set(100)
        self.pop_scale.pack(fill=tk.X, pady=(0, 4))

        # Tỉ lệ lai ghép (Crossover Rate)
        ttk.Label(control_panel, text="Tỉ lệ lai ghép (Crossover Rate):").pack(anchor=tk.W)
        self.cross_scale = tk.Scale(control_panel, from_=0.4, to=1.0, orient=tk.HORIZONTAL, resolution=0.05)
        self.cross_scale.set(0.85)
        self.cross_scale.pack(fill=tk.X, pady=(0, 4))

        # Tỉ lệ đột biến (Mutation Rate)
        ttk.Label(control_panel, text="Tỉ lệ đột biến (Mutation Rate):").pack(anchor=tk.W)
        self.mut_scale = tk.Scale(control_panel, from_=0.01, to=0.50, orient=tk.HORIZONTAL, resolution=0.01)
        self.mut_scale.set(0.20)
        self.mut_scale.pack(fill=tk.X, pady=(0, 4))

        # Toán tử đột biến
        ttk.Label(control_panel, text="Toán tử đột biến:").pack(anchor=tk.W)
        self.mut_type_combo = ttk.Combobox(control_panel, values=["Inversion (2-Opt / Khuyên dùng)", "Swap (Hoán đổi 2 gen)"], state="readonly")
        self.mut_type_combo.current(0)
        self.mut_type_combo.pack(fill=tk.X, pady=(0, 6))

        ttk.Separator(control_panel, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=6)

        # 3. Các nút bấm hành động
        self.btn_start = tk.Button(control_panel, text="▶  BẮT ĐẦU TỐI ƯU", font=("Segoe UI", 10, "bold"),
                                   bg="#27ae60", fg="white", activebackground="#2ecc71", relief=tk.FLAT,
                                   command=self.start_optimization, height=2)
        self.btn_start.pack(fill=tk.X, pady=2)

        self.btn_pause = tk.Button(control_panel, text="⏸  TẠM DỪNG", font=("Segoe UI", 9),
                                   bg="#f39c12", fg="white", activebackground="#e67e22", relief=tk.FLAT,
                                   command=self.toggle_pause, state=tk.DISABLED)
        self.btn_pause.pack(fill=tk.X, pady=2)

        self.btn_reset = tk.Button(control_panel, text="🔄  ĐẶT LẠI (RESET)", font=("Segoe UI", 9),
                                   bg="#7f8c8d", fg="white", activebackground="#95a5a6", relief=tk.FLAT,
                                   command=self.reset_simulation)
        self.btn_reset.pack(fill=tk.X, pady=2)

        self.btn_save = tk.Button(control_panel, text="💾  XUẤT ẢNH BIỂU ĐỒ", font=("Segoe UI", 9),
                                  bg="#2980b9", fg="white", activebackground="#3498db", relief=tk.FLAT,
                                  command=self.save_figure)
        self.btn_save.pack(fill=tk.X, pady=2)

        self.btn_export = tk.Button(control_panel, text="📋  XUẤT LỊCH TRÌNH CHI TIẾT (.TXT)", font=("Segoe UI", 9),
                                    bg="#8e44ad", fg="white", activebackground="#9b59b6", relief=tk.FLAT,
                                    command=self.export_itinerary)
        self.btn_export.pack(fill=tk.X, pady=2)

        ttk.Separator(control_panel, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=6)

        # 4. Bảng số liệu thống kê tối ưu
        stat_frame = ttk.LabelFrame(control_panel, text=" THỐNG KÊ KẾT QUẢ TỐI ƯU ", padding=6)
        stat_frame.pack(fill=tk.X, pady=(0, 6))

        self.lbl_gen = ttk.Label(stat_frame, text="Thế hệ: 0 / 0", font=("Segoe UI", 9))
        self.lbl_gen.pack(anchor=tk.W, pady=1)

        self.lbl_init_dist = ttk.Label(stat_frame, text="Quãng đường ban đầu: -- km", font=("Segoe UI", 9))
        self.lbl_init_dist.pack(anchor=tk.W, pady=1)

        self.lbl_best_dist = ttk.Label(stat_frame, text="Quãng đường tốt nhất: -- km", font=("Segoe UI", 10, "bold"), foreground="#c0392b")
        self.lbl_best_dist.pack(anchor=tk.W, pady=1)

        self.lbl_improve = ttk.Label(stat_frame, text="Mức độ cải thiện: 0.00 %", font=("Segoe UI", 9, "bold"), foreground="#27ae60")
        self.lbl_improve.pack(anchor=tk.W, pady=1)

        self.lbl_time = ttk.Label(stat_frame, text="Thời gian chạy: 0.000 s", font=("Segoe UI", 9))
        self.lbl_time.pack(anchor=tk.W, pady=1)

        # 5. Phân tích Green Logistics (Kinh tế & Môi trường)
        green_frame = ttk.LabelFrame(control_panel, text=" HIỆU QUẢ LOGISTICS XANH & TIẾT KIỆM ", padding=6)
        green_frame.pack(fill=tk.X, pady=(0, 6))

        self.lbl_fuel = ttk.Label(green_frame, text="⛽ Xăng tiết kiệm: 0.00 L", font=("Segoe UI", 9))
        self.lbl_fuel.pack(anchor=tk.W, pady=1)

        self.lbl_cost = ttk.Label(green_frame, text="💰 Tiền tiết kiệm: 0 VNĐ / chuyến", font=("Segoe UI", 9, "bold"), foreground="#d35400")
        self.lbl_cost.pack(anchor=tk.W, pady=1)

        self.lbl_co2 = ttk.Label(green_frame, text="🌱 Giảm phát thải CO2: 0.00 kg", font=("Segoe UI", 9))
        self.lbl_co2.pack(anchor=tk.W, pady=1)

        self.lbl_time_saved = ttk.Label(green_frame, text="⏱ Giảm thời gian lái xe: 0 phút", font=("Segoe UI", 9))
        self.lbl_time_saved.pack(anchor=tk.W, pady=1)

        # ==================== CỘT PHẢI: KHU VỰC VẼ ĐỒ THỊ ====================
        chart_panel = ttk.Frame(main_frame)
        chart_panel.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        # Khởi tạo Matplotlib Figure (2 biểu đồ con song song)
        self.fig = Figure(figsize=(9, 6), dpi=100)
        self.ax_route = self.fig.add_subplot(121)
        self.ax_curve = self.fig.add_subplot(122)
        self.fig.tight_layout(pad=3.0)

        self.canvas = FigureCanvasTkAgg(self.fig, master=chart_panel)
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

    def on_dataset_change(self):
        """Xử lý khi người dùng đổi loại tập dữ liệu."""
        if self.dataset_var.get() == "random":
            self.rand_frame.pack(anchor=tk.W, pady=2)
            self.generate_random_dataset()
        else:
            self.rand_frame.pack_forget()
            self.cities = HCMC_LOCATIONS
            self.dist_matrix = build_distance_matrix(self.cities)
            self.reset_simulation()

    def generate_random_dataset(self):
        """Sinh tập dữ liệu ngẫu nhiên."""
        try:
            n = int(self.rand_n_entry.get())
            n = max(6, min(n, 60))
        except ValueError:
            n = 15
        
        cities = [{"id": 0, "name": "Kho Tổng (Depot)", "district": "Trung tâm", "x": 50.0, "y": 50.0}]
        for i in range(1, n):
            cities.append({
                "id": i,
                "name": f"Điểm {i}",
                "district": f"Vùng {chr(65 + (i % 6))}",
                "x": round(random.uniform(5, 95), 1),
                "y": round(random.uniform(5, 95), 1)
            })
        self.cities = cities
        self.dist_matrix = build_distance_matrix(self.cities)
        self.reset_simulation()

    def reset_simulation(self):
        """Đặt lại trạng thái mô phỏng ban đầu."""
        self.is_running = False
        self.is_paused = False
        self.btn_start.config(state=tk.NORMAL, text="▶  BẮT ĐẦU TỐI ƯU")
        self.btn_pause.config(state=tk.DISABLED, text="⏸  TẠM DỪNG")
        
        # Khởi tạo GA mới
        mut_type = "inversion" if "Inversion" in self.mut_type_combo.get() else "swap"
        self.ga = GeneticAlgorithmTSP(
            cities=self.cities,
            distance_matrix=self.dist_matrix,
            pop_size=self.pop_scale.get(),
            elite_size=5,
            crossover_rate=self.cross_scale.get(),
            mutation_rate=self.mut_scale.get(),
            tournament_size=5,
            mutation_type=mut_type,
            use_local_search=self.memetic_var.get()
        )
        self.initial_distance = self.ga.best_individual.distance
        
        # Cập nhật số liệu
        self.lbl_gen.config(text=f"Thế hệ: 0 / {self.gen_scale.get()}")
        self.lbl_init_dist.config(text=f"Quãng đường ban đầu: {self.initial_distance:.2f} km")
        self.lbl_best_dist.config(text=f"Quãng đường tốt nhất: {self.initial_distance:.2f} km")
        self.lbl_improve.config(text="Mức độ cải thiện: 0.00 %")
        self.lbl_time.config(text="Thời gian chạy: 0.000 s")
        
        # Reset số liệu Green Logistics
        self.lbl_fuel.config(text="⛽ Xăng tiết kiệm: 0.00 L")
        self.lbl_cost.config(text="💰 Tiền tiết kiệm: 0 VNĐ / chuyến")
        self.lbl_co2.config(text="🌱 Giảm phát thải CO2: 0.00 kg")
        self.lbl_time_saved.config(text="⏱ Giảm thời gian lái xe: 0 phút")
        
        # Vẽ biểu đồ ban đầu
        self.draw_plots()

    def draw_plots(self):
        """Vẽ lộ trình và đường cong hội tụ trên Canvas Matplotlib."""
        self.ax_route.clear()
        self.ax_curve.clear()
        
        # 1. Vẽ lộ trình
        best_route = self.ga.best_individual.genes
        route_coords = [self.cities[idx] for idx in best_route]
        route_coords.append(self.cities[best_route[0]]) # Khép vòng
        
        xs = [c["x"] for c in route_coords]
        ys = [c["y"] for c in route_coords]
        
        # Vẽ các cạnh lộ trình
        self.ax_route.plot(xs, ys, color="#2980b9", linestyle="-", linewidth=1.8, alpha=0.85, zorder=1)
        
        # Vẽ các điểm thành phố
        for i, city in enumerate(self.cities):
            if city["id"] == 0:
                self.ax_route.scatter(city["x"], city["y"], color="#e74c3c", s=150, marker="*", edgecolors="black", linewidths=1.2, zorder=3, label="Kho Tổng (Depot)")
                self.ax_route.annotate(f" {city['name']}", (city["x"], city["y"]), fontsize=8, fontweight="bold", color="#c0392b")
            else:
                self.ax_route.scatter(city["x"], city["y"], color="#27ae60", s=55, marker="o", edgecolors="black", zorder=2)
                # Chỉ vẽ tên nếu số điểm <= 25 để tránh rối
                if len(self.cities) <= 25:
                    self.ax_route.annotate(f" {city['id']}", (city["x"], city["y"]), fontsize=7, alpha=0.9)
                    
        mode_str = " (Memetic GA)" if self.memetic_var.get() else ""
        self.ax_route.set_title(f"Lộ trình tối ưu{mode_str} (Dist: {self.ga.best_individual.distance:.2f} km)", fontsize=11, fontweight="bold")
        self.ax_route.set_xlabel("Tọa độ X (km)", fontsize=9)
        self.ax_route.set_ylabel("Tọa độ Y (km)", fontsize=9)
        self.ax_route.grid(True, linestyle="--", alpha=0.4)
        self.ax_route.legend(loc="upper right", fontsize=8)
        
        # 2. Vẽ đường cong hội tụ
        gens = list(range(len(self.ga.history_best_distance)))
        self.ax_curve.plot(gens, self.ga.history_best_distance, color="#e74c3c", linewidth=2.0, label="Best Distance")
        self.ax_curve.plot(gens, self.ga.history_avg_distance, color="#3498db", linestyle="--", linewidth=1.4, label="Avg Distance")
        
        self.ax_curve.set_title("Đường cong hội tụ qua các thế hệ", fontsize=11, fontweight="bold")
        self.ax_curve.set_xlabel("Thế hệ (Generation)", fontsize=9)
        self.ax_curve.set_ylabel("Tổng quãng đường (km)", fontsize=9)
        self.ax_curve.grid(True, linestyle="--", alpha=0.4)
        self.ax_curve.legend(fontsize=8)
        
        self.canvas.draw_idle()

    def start_optimization(self):
        """Bắt đầu tiến trình tối ưu hoá."""
        if not self.is_running:
            self.is_running = True
            self.is_paused = False
            self.btn_start.config(state=tk.DISABLED)
            self.btn_pause.config(state=tk.NORMAL, text="⏸  TẠM DỪNG")
            self.start_time = time.time()
            # Khởi động luồng chạy ngầm để không đơ giao diện
            threading.Thread(target=self._run_loop, daemon=True).start()

    def toggle_pause(self):
        """Tạm dừng hoặc tiếp tục."""
        if self.is_paused:
            self.is_paused = False
            self.btn_pause.config(text="⏸  TẠM DỪNG", bg="#f39c12")
        else:
            self.is_paused = True
            self.btn_pause.config(text="▶  TIẾP TỤC", bg="#27ae60")

    def _run_loop(self):
        """Vòng lặp tiến hóa chạy trong background thread."""
        max_gens = self.gen_scale.get()
        while self.is_running and self.ga.current_generation < max_gens:
            if not self.is_paused:
                self.ga.evolve_generation()
                
                # Cập nhật GUI mỗi thế hệ (hoặc cách nhịp nếu số thế hệ lớn)
                if self.ga.current_generation % 2 == 0 or self.ga.current_generation == max_gens:
                    self.root.after(0, self._update_ui_state)
                    time.sleep(0.015) # Giãn nhịp để mắt người nhìn thấy hiệu ứng động
            else:
                time.sleep(0.1)
                
        self.is_running = False
        self.root.after(0, self._on_finish)

    def _update_ui_state(self):
        """Cập nhật các nhãn và đồ thị trên luồng chính của Tkinter."""
        cur_gen = self.ga.current_generation
        max_gens = self.gen_scale.get()
        cur_best = self.ga.best_individual.distance
        improve = ((self.initial_distance - cur_best) / self.initial_distance) * 100
        elapsed = time.time() - self.start_time
        
        self.lbl_gen.config(text=f"Thế hệ: {cur_gen} / {max_gens}")
        self.lbl_best_dist.config(text=f"Quãng đường tốt nhất: {cur_best:.2f} km")
        self.lbl_improve.config(text=f"Mức độ cải thiện: {improve:.2f} %")
        self.lbl_time.config(text=f"Thời gian chạy: {elapsed:.2f} s")
        
        # Cập nhật Green Logistics
        dist_saved = max(0.0, self.initial_distance - cur_best)
        impact = GeneticAlgorithmTSP.calculate_economic_impact(dist_saved)
        self.lbl_fuel.config(text=f"⛽ Xăng tiết kiệm: {impact['fuel_saved_liters']:.2f} L")
        self.lbl_cost.config(text=f"💰 Tiền tiết kiệm: {impact['cost_saved_vnd']:,} VNĐ / chuyến")
        self.lbl_co2.config(text=f"🌱 Giảm phát thải CO2: {impact['co2_reduced_kg']:.2f} kg")
        self.lbl_time_saved.config(text=f"⏱ Giảm thời gian lái xe: {impact['time_saved_minutes']} phút")
        
        self.draw_plots()

    def _on_finish(self):
        """Xử lý khi hoàn tất tất cả thế hệ."""
        self.btn_start.config(state=tk.NORMAL, text="▶  CHẠY TIẾP")
        self.btn_pause.config(state=tk.DISABLED)
        dist_saved = max(0.0, self.initial_distance - self.ga.best_individual.distance)
        impact = GeneticAlgorithmTSP.calculate_economic_impact(dist_saved)
        
        messagebox.showinfo("Hoàn tất", f"Quá trình tối ưu hoá đã kết thúc!\n\n"
                                        f"Quãng đường ban đầu: {self.initial_distance:.2f} km\n"
                                        f"Quãng đường tối ưu:  {self.ga.best_individual.distance:.2f} km\n"
                                        f"Tiết kiệm được:       {((self.initial_distance - self.ga.best_individual.distance)/self.initial_distance)*100:.2f}%\n\n"
                                        f"--- HIỆU QUẢ LOGISTICS XANH ---\n"
                                        f"Xăng tiết kiệm:       {impact['fuel_saved_liters']:.2f} lít / chuyến\n"
                                        f"Chi phí tiết kiệm:    {impact['cost_saved_vnd']:,} VNĐ / chuyến\n"
                                        f"Ước tính tháng (30 ngày): {impact['cost_saved_vnd'] * 30:,} VNĐ / xe\n"
                                        f"Giảm phát thải CO2:   {impact['co2_reduced_kg']:.2f} kg CO2\n"
                                        f"Thời gian tiết kiệm:  {impact['time_saved_minutes']} phút")

    def save_figure(self):
        """Lưu biểu đồ hiện tại ra file ảnh."""
        file_path = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=[("PNG Image", "*.png"), ("JPEG Image", "*.jpg"), ("PDF File", "*.pdf")],
            initialfile="do_thi_lo_trinh_ga_tsp.png"
        )
        if file_path:
            self.fig.savefig(file_path, dpi=250, bbox_inches="tight")
            messagebox.showinfo("Thành công", f"Đã lưu biểu đồ thành công tại:\n{file_path}")

    def export_itinerary(self):
        """Xuất lịch trình chi tiết ra file văn bản."""
        file_path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text File", "*.txt"), ("CSV File", "*.csv"), ("All Files", "*.*")],
            initialfile="lich_trinh_chi_tiet_tsp.txt"
        )
        if file_path:
            itinerary = self.ga.get_detailed_itinerary()
            dist_saved = max(0.0, self.initial_distance - self.ga.best_individual.distance)
            impact = GeneticAlgorithmTSP.calculate_economic_impact(dist_saved)
            
            with open(file_path, "w", encoding="utf-8") as f:
                f.write("=" * 75 + "\n")
                f.write("             LỊCH TRÌNH GIAO HÀNG CHI TIẾT (TURN-BY-TURN)\n")
                f.write("      TỐI ƯU HOÁ BẰNG GIẢI THUẬT DI TRUYỀN (GENETIC ALGORITHM)\n")
                f.write("=" * 75 + "\n\n")
                f.write(f"Quãng đường ban đầu:          {self.initial_distance:.2f} km\n")
                f.write(f"Quãng đường sau tối ưu:       {self.ga.best_individual.distance:.2f} km\n")
                f.write(f"Quãng đường tiết kiệm:        {impact['distance_saved_km']:.2f} km\n")
                f.write(f"Nhiên liệu tiết kiệm:         {impact['fuel_saved_liters']:.2f} lít xăng\n")
                f.write(f"Chi phí tiết kiệm (chuyến):   {impact['cost_saved_vnd']:,} VNĐ\n")
                f.write(f"Chi phí tiết kiệm (30 ngày):  {impact['cost_saved_vnd'] * 30:,} VNĐ / xe\n")
                f.write(f"Lượng CO2 giảm thiểu:         {impact['co2_reduced_kg']:.2f} kg CO2\n")
                f.write(f"Thời gian lưu thông tiết kiệm:{impact['time_saved_minutes']} phút\n\n")
                f.write("-" * 75 + "\n")
                f.write(f"{'Chặng':<6} | {'Điểm xuất phát':<30} -> {'Điểm đến':<30} | {'Độ dài':<8} | {'Tổng tích lũy'}\n")
                f.write("-" * 75 + "\n")
                for leg in itinerary:
                    f.write(f"{leg['step']:<6d} | {leg['from_name']:<30} -> {leg['to_name']:<30} | {leg['distance_km']:>5.2f} km | {leg['cumulative_km']:>6.2f} km\n")
                f.write("-" * 75 + "\n")
                f.write(f"TỔNG QUÃNG ĐƯỜNG TOÀN CHU TRÌNH: {self.ga.best_individual.distance:.2f} km\n")
                f.write("=" * 75 + "\n")
                
            messagebox.showinfo("Thành công", f"Đã xuất lịch trình chi tiết thành công tại:\n{file_path}")

def main():
    root = tk.Tk()
    app = GA_TSP_App(root)
    root.mainloop()

if __name__ == "__main__":
    main()
