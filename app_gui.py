# -*- coding: utf-8 -*-
"""
app_gui.py - Giao diện đồ họa (GUI) trực quan hóa Giải thuật Di truyền cho bài toán TSP.
Sử dụng thư viện chuẩn Tkinter và Matplotlib (có sẵn trong Python).
Cho phép người dùng tương tác, điều chỉnh tham số, quan sát lộ trình và đồ thị hội tụ động theo thời gian thực.
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
        self.root.title("Ứng dụng Tối ưu hoá Lộ trình Giao hàng (TSP) - Giải thuật Di truyền (GA)")
        self.root.geometry("1280x760")
        self.root.minsize(1050, 650)
        
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
        # Frame chính chia làm 2 cột: Trái (Điều khiển), Phải (Đồ thị)
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # ==================== CỘT TRÁI: ĐIỀU KHIỂN ====================
        control_panel = ttk.LabelFrame(main_frame, text=" BẢNG ĐIỀU KHIỂN & THAM SỐ ", padding=12)
        control_panel.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))

        # 1. Chọn nguồn dữ liệu
        lbl_data = ttk.Label(control_panel, text="1. Nguồn dữ liệu bài toán:", font=("Segoe UI", 9, "bold"))
        lbl_data.pack(anchor=tk.W, pady=(0, 4))

        self.dataset_var = tk.StringVar(value="hcmc")
        r1 = ttk.Radiobutton(control_panel, text="20 Địa điểm thực tế tại TP.HCM", value="hcmc", 
                             variable=self.dataset_var, command=self.on_dataset_change)
        r1.pack(anchor=tk.W, padx=10)
        
        r2 = ttk.Radiobutton(control_panel, text="Sinh ngẫu nhiên N địa điểm", value="random", 
                             variable=self.dataset_var, command=self.on_dataset_change)
        r2.pack(anchor=tk.W, padx=10, pady=(2, 6))

        # Khung nhập N điểm ngẫu nhiên
        self.rand_frame = ttk.Frame(control_panel)
        ttk.Label(self.rand_frame, text="Số điểm N:").pack(side=tk.LEFT, padx=(10, 5))
        self.rand_n_entry = ttk.Spinbox(self.rand_frame, from_=8, to=60, width=5)
        self.rand_n_entry.set(15)
        self.rand_n_entry.pack(side=tk.LEFT)
        btn_gen = ttk.Button(self.rand_frame, text="Sinh lại", command=self.generate_random_dataset)
        btn_gen.pack(side=tk.LEFT, padx=5)

        ttk.Separator(control_panel, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=10)

        # 2. Tham số Giải thuật Di truyền
        lbl_params = ttk.Label(control_panel, text="2. Cấu hình tham số GA:", font=("Segoe UI", 9, "bold"))
        lbl_params.pack(anchor=tk.W, pady=(0, 4))

        # Số thế hệ (Generations)
        ttk.Label(control_panel, text="Số thế hệ tối đa:").pack(anchor=tk.W)
        self.gen_scale = tk.Scale(control_panel, from_=50, to=500, orient=tk.HORIZONTAL, resolution=10)
        self.gen_scale.set(150)
        self.gen_scale.pack(fill=tk.X, pady=(0, 6))

        # Kích thước quần thể (Population Size)
        ttk.Label(control_panel, text="Kích thước quần thể (Pop Size):").pack(anchor=tk.W)
        self.pop_scale = tk.Scale(control_panel, from_=20, to=300, orient=tk.HORIZONTAL, resolution=10)
        self.pop_scale.set(100)
        self.pop_scale.pack(fill=tk.X, pady=(0, 6))

        # Tỉ lệ lai ghép (Crossover Rate)
        ttk.Label(control_panel, text="Tỉ lệ lai ghép (Crossover Rate):").pack(anchor=tk.W)
        self.cross_scale = tk.Scale(control_panel, from_=0.4, to=1.0, orient=tk.HORIZONTAL, resolution=0.05)
        self.cross_scale.set(0.85)
        self.cross_scale.pack(fill=tk.X, pady=(0, 6))

        # Tỉ lệ đột biến (Mutation Rate)
        ttk.Label(control_panel, text="Tỉ lệ đột biến (Mutation Rate):").pack(anchor=tk.W)
        self.mut_scale = tk.Scale(control_panel, from_=0.01, to=0.50, orient=tk.HORIZONTAL, resolution=0.01)
        self.mut_scale.set(0.20)
        self.mut_scale.pack(fill=tk.X, pady=(0, 6))

        # Toán tử đột biến
        ttk.Label(control_panel, text="Toán tử đột biến:").pack(anchor=tk.W)
        self.mut_type_combo = ttk.Combobox(control_panel, values=["Inversion (2-Opt / Khuyên dùng)", "Swap (Hoán đổi 2 gen)"], state="readonly")
        self.mut_type_combo.current(0)
        self.mut_type_combo.pack(fill=tk.X, pady=(0, 10))

        ttk.Separator(control_panel, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=8)

        # 3. Các nút bấm hành động
        self.btn_start = tk.Button(control_panel, text="▶  BẮT ĐẦU TỐI ƯU", font=("Segoe UI", 10, "bold"),
                                   bg="#27ae60", fg="white", activebackground="#2ecc71", relief=tk.FLAT,
                                   command=self.start_optimization, height=2)
        self.btn_start.pack(fill=tk.X, pady=3)

        self.btn_pause = tk.Button(control_panel, text="⏸  TẠM DỪNG", font=("Segoe UI", 9),
                                   bg="#f39c12", fg="white", activebackground="#e67e22", relief=tk.FLAT,
                                   command=self.toggle_pause, state=tk.DISABLED)
        self.btn_pause.pack(fill=tk.X, pady=3)

        self.btn_reset = tk.Button(control_panel, text="🔄  ĐẶT LẠI (RESET)", font=("Segoe UI", 9),
                                   bg="#7f8c8d", fg="white", activebackground="#95a5a6", relief=tk.FLAT,
                                   command=self.reset_simulation)
        self.btn_reset.pack(fill=tk.X, pady=3)

        self.btn_save = tk.Button(control_panel, text="💾  XUẤT ẢNH BIỂU ĐỒ", font=("Segoe UI", 9),
                                  bg="#2980b9", fg="white", activebackground="#3498db", relief=tk.FLAT,
                                  command=self.save_figure)
        self.btn_save.pack(fill=tk.X, pady=3)

        ttk.Separator(control_panel, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=10)

        # 4. Bảng số liệu thống kê thời gian thực
        stat_frame = ttk.LabelFrame(control_panel, text=" THỐNG KÊ KẾT QUẢ ", padding=8)
        stat_frame.pack(fill=tk.BOTH, expand=True)

        self.lbl_gen = ttk.Label(stat_frame, text="Thế hệ: 0 / 0", font=("Segoe UI", 9))
        self.lbl_gen.pack(anchor=tk.W, pady=2)

        self.lbl_init_dist = ttk.Label(stat_frame, text="Quãng đường ban đầu: -- km", font=("Segoe UI", 9))
        self.lbl_init_dist.pack(anchor=tk.W, pady=2)

        self.lbl_best_dist = ttk.Label(stat_frame, text="Quãng đường tốt nhất: -- km", font=("Segoe UI", 10, "bold"), foreground="#c0392b")
        self.lbl_best_dist.pack(anchor=tk.W, pady=2)

        self.lbl_improve = ttk.Label(stat_frame, text="Mức độ cải thiện: 0.00 %", font=("Segoe UI", 9, "bold"), foreground="#27ae60")
        self.lbl_improve.pack(anchor=tk.W, pady=2)

        self.lbl_time = ttk.Label(stat_frame, text="Thời gian chạy: 0.000 s", font=("Segoe UI", 9))
        self.lbl_time.pack(anchor=tk.W, pady=2)

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
            mutation_type=mut_type
        )
        self.initial_distance = self.ga.best_individual.distance
        
        # Cập nhật số liệu
        self.lbl_gen.config(text=f"Thế hệ: 0 / {self.gen_scale.get()}")
        self.lbl_init_dist.config(text=f"Quãng đường ban đầu: {self.initial_distance:.2f} km")
        self.lbl_best_dist.config(text=f"Quãng đường tốt nhất: {self.initial_distance:.2f} km")
        self.lbl_improve.config(text="Mức độ cải thiện: 0.00 %")
        self.lbl_time.config(text="Thời gian chạy: 0.000 s")
        
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
                    
        self.ax_route.set_title(f"Lộ trình tối ưu (Dist: {self.ga.best_individual.distance:.2f} km)", fontsize=11, fontweight="bold")
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
        
        self.draw_plots()

    def _on_finish(self):
        """Xử lý khi hoàn tất tất cả thế hệ."""
        self.btn_start.config(state=tk.NORMAL, text="▶  CHẠY TIẾP")
        self.btn_pause.config(state=tk.DISABLED)
        messagebox.showinfo("Hoàn tất", f"Quá trình tối ưu hoá đã kết thúc!\n"
                                        f"Quãng đường ban đầu: {self.initial_distance:.2f} km\n"
                                        f"Quãng đường tối ưu:  {self.ga.best_individual.distance:.2f} km\n"
                                        f"Tiết kiệm được:       {((self.initial_distance - self.ga.best_individual.distance)/self.initial_distance)*100:.2f}%")

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

def main():
    root = tk.Tk()
    app = GA_TSP_App(root)
    root.mainloop()

if __name__ == "__main__":
    main()
