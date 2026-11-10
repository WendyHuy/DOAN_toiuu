# -*- coding: utf-8 -*-
"""
run_cli.py - Chương trình dòng lệnh mô phỏng và tối ưu bài toán TSP bằng Genetic Algorithm.
Đáp ứng chuẩn tiêu chí Rubric:
  - Input: Dữ liệu tọa độ, khoảng cách, tham số GA.
  - Output: Lộ trình tối ưu, quãng đường ngắn nhất, bảng tiến trình, xuất biểu đồ ảnh.
  - Mô hình tính toán: Trình bày rõ các bước (Khởi tạo -> Đánh giá -> Chọn lọc -> Lai ghép -> Đột biến).
"""

import sys
import os
import random
import time

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8")

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src.cities_data import HCMC_LOCATIONS, build_distance_matrix, calculate_distance
from src.genetic_algorithm import GeneticAlgorithmTSP
import matplotlib.pyplot as plt

def generate_random_cities(num_cities=15, width=100, height=100):
    """Sinh ngẫu nhiên danh sách thành phố."""
    cities = []
    cities.append({"id": 0, "name": "Kho Trung Tâm (Depot)", "district": "Trung Tâm", "x": width / 2, "y": height / 2})
    for i in range(1, num_cities):
        cities.append({
            "id": i,
            "name": f"Điểm giao {i}",
            "district": f"Khu vực {chr(65 + (i % 6))}",
            "x": round(random.uniform(5, width - 5), 1),
            "y": round(random.uniform(5, height - 5), 1)
        })
    return cities

def plot_and_save_results(ga, cities, output_dir="results"):
    """Vẽ và lưu kết quả trực quan hóa ra file ảnh PNG để đưa vào báo cáo."""
    os.makedirs(output_dir, exist_ok=True)
    
    # 1. Vẽ đồ thị lộ trình tối ưu
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(15, 6))
    
    # Lấy thứ tự các thành phố theo cá thể tốt nhất
    best_route = ga.best_individual.genes
    route_coords = [cities[idx] for idx in best_route]
    # Thêm điểm đầu để khép kín chu trình
    route_coords.append(cities[best_route[0]])
    
    xs = [c["x"] for c in route_coords]
    ys = [c["y"] for c in route_coords]
    
    # Vẽ các đoạn đường nối
    ax1.plot(xs, ys, color="#2980b9", linestyle="-", linewidth=2, alpha=0.8, zorder=1)
    
    # Vẽ các điểm giao hàng
    for i, city in enumerate(cities):
        if city["id"] == 0:
            # Kho trung tâm
            ax1.scatter(city["x"], city["y"], color="#e74c3c", s=180, marker="*", edgecolors="black", linewidths=1.5, zorder=3, label="Kho Tổng (Depot)")
            ax1.annotate(f" {city['name']}", (city["x"], city["y"]), fontsize=9, fontweight="bold", color="#c0392b")
        else:
            ax1.scatter(city["x"], city["y"], color="#27ae60", s=70, marker="o", edgecolors="black", zorder=2)
            ax1.annotate(f" {city['id']}. {city['name']}", (city["x"], city["y"]), fontsize=8, alpha=0.85)
            
    ax1.set_title(f"Lộ trình tối ưu tìm được (Tổng: {ga.best_individual.distance:.2f} km)", fontsize=12, fontweight="bold")
    ax1.set_xlabel("Tọa độ X (km)")
    ax1.set_ylabel("Tọa độ Y (km)")
    ax1.grid(True, linestyle="--", alpha=0.5)
    ax1.legend(loc="upper right")
    
    # 2. Vẽ biểu đồ đường cong hội tụ (Convergence Curve)
    gens = list(range(len(ga.history_best_distance)))
    ax2.plot(gens, ga.history_best_distance, color="#e74c3c", linewidth=2.5, label="Quãng đường tốt nhất (Best)")
    ax2.plot(gens, ga.history_avg_distance, color="#3498db", linestyle="--", linewidth=1.8, label="Quãng đường trung bình (Avg)")
    
    ax2.set_title("Biểu đồ đường cong hội tụ qua các thế hệ", fontsize=12, fontweight="bold")
    ax2.set_xlabel("Thế hệ (Generation)")
    ax2.set_ylabel("Tổng quãng đường (km)")
    ax2.grid(True, linestyle="--", alpha=0.5)
    ax2.legend()
    
    plt.tight_layout()
    chart_path = os.path.join(output_dir, "ket_qua_toi_uu_ga_tsp.png")
    plt.savefig(chart_path, dpi=200)
    plt.close()
    return chart_path

def main():
    print("=" * 70)
    print("   CHƯƠNG TRÌNH TỐI ƯU HOÁ LỘ TRÌNH GIAO HÀNG (TSP)")
    print("      BẰNG GIẢI THUẬT DI TRUYỀN (GENETIC ALGORITHM)")
    print("   Đồ án môn học: Nhập môn Tối ưu hoá - GVHD: ThS. Trần Ngọc Việt")
    print("=" * 70)
    
    # Chọn bộ dữ liệu
    print("\n[BƯỚC 1: THIẾT LẬP INPUT VÀ BỘ DỮ LIỆU]")
    print("1. Sử dụng tập dữ liệu thực tế 20 địa điểm tại TP. Hồ Chí Minh (Mặc định)")
    print("2. Tự động sinh ngẫu nhiên N địa điểm")
    choice = input("Lựa chọn của bạn [1/2, Enter = 1]: ").strip()
    
    if choice == "2":
        try:
            n_cities = int(input("Nhập số lượng địa điểm cần giao (từ 10 đến 50) [15]: ").strip() or "15")
        except ValueError:
            n_cities = 15
        cities = generate_random_cities(n_cities)
        print(f"-> Đã sinh ngẫu nhiên {len(cities)} điểm giao hàng.")
    else:
        cities = HCMC_LOCATIONS
        print(f"-> Đã tải thành công {len(cities)} địa điểm thực tế tại TP. Hồ Chí Minh.")
        
    dist_matrix = build_distance_matrix(cities)
    
    # Thiết lập tham số GA
    print("\n[BƯỚC 2: CẤU HÌNH THAM SỐ THUẬT TOÁN GENETIC ALGORITHM]")
    try:
        pop_size = int(input("Kích thước quần thể (Population Size) [100]: ").strip() or "100")
        generations = int(input("Số thế hệ tối đa (Generations) [150]: ").strip() or "150")
        cross_rate = float(input("Tỉ lệ lai ghép (Crossover Rate) [0.85]: ").strip() or "0.85")
        mut_rate = float(input("Tỉ lệ đột biến (Mutation Rate) [0.20]: ").strip() or "0.20")
    except ValueError:
        pop_size, generations, cross_rate, mut_rate = 100, 150, 0.85, 0.20
        
    print("\n-> Cấu hình xác nhận:")
    print(f"   + Số lượng điểm: {len(cities)}")
    print(f"   + Quần thể: {pop_size} cá thể | Số thế hệ: {generations}")
    print(f"   + Lai ghép: Ordered Crossover (OX) - Tỉ lệ: {cross_rate * 100:.1f}%")
    print(f"   + Đột biến: Inversion Mutation (2-Opt) - Tỉ lệ: {mut_rate * 100:.1f}%")
    print(f"   + Chọn lọc: Tournament Selection (k=5) & Elitism (top 5)")
    
    # Khởi tạo thuật toán
    print("\n[BƯỚC 3: TIẾN TRÌNH TỐI ƯU HOÁ VÀ TÍNH TOÁN]")
    start_time = time.time()
    
    ga = GeneticAlgorithmTSP(
        cities=cities,
        distance_matrix=dist_matrix,
        pop_size=pop_size,
        elite_size=5,
        crossover_rate=cross_rate,
        mutation_rate=mut_rate,
        tournament_size=5,
        mutation_type="inversion"
    )
    
    initial_dist = ga.best_individual.distance
    print(f"-> Quãng đường ngẫu nhiên ban đầu (Gen 0): {initial_dist:.2f} km")
    print("-" * 70)
    print(f"{'Thế hệ':<10} | {'Quãng đường tốt nhất':<24} | {'Quãng đường trung bình':<24}")
    print("-" * 70)
    
    log_interval = max(1, generations // 10)
    for g in range(1, generations + 1):
        ga.evolve_generation()
        if g % log_interval == 0 or g == generations:
            print(f"Thế hệ {g:<4} | {ga.best_individual.distance:18.2f} km | {ga.history_avg_distance[-1]:18.2f} km")
            
    elapsed_time = time.time() - start_time
    final_dist = ga.best_individual.distance
    improvement = ((initial_dist - final_dist) / initial_dist) * 100
    
    print("-" * 70)
    print("\n[BƯỚC 4: KẾT QUẢ ĐẦU RA (OUTPUT)]")
    print(f"Thời gian thực thi:           {elapsed_time:.3f} giây")
    print(f"Quãng đường ban đầu:          {initial_dist:.2f} km")
    print(f"Quãng đường tối ưu tìm được:  {final_dist:.2f} km")
    print(f"Tỉ lệ tối ưu hoá rút ngắn:    {improvement:.2f}% (Rút ngắn được {initial_dist - final_dist:.2f} km)")
    
    print("\nChi tiết thứ tự lộ trình tối ưu:")
    route = ga.best_individual.genes
    for step, c_idx in enumerate(route):
        c = cities[c_idx]
        print(f"  {step + 1:2d}. {c['name']:<35} ({c.get('district', '')})")
    print(f"  --> Quay về điểm xuất phát: {cities[route[0]]['name']}")
    
    # Xuất ảnh đồ thị
    print("\nĐang xuất biểu đồ phân tích và lộ trình...")
    img_path = plot_and_save_results(ga, cities)
    print(f"-> Đã lưu biểu đồ kết quả tại: {os.path.abspath(img_path)}")
    print("=" * 70)
    print("CHƯƠNG TRÌNH ĐÃ HOÀN TẤT THÀNH CÔNG!")

if __name__ == "__main__":
    main()
