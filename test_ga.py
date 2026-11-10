# -*- coding: utf-8 -*-
"""
test_ga.py - Kiểm tra chức năng thuật toán Genetic Algorithm cho TSP
"""
import sys
import os

# Đảm bảo hiển thị tiếng Việt UTF-8 chuẩn trên Windows Console
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8")

# Thêm thư mục gốc vào PYTHONPATH
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src.cities_data import HCMC_LOCATIONS, build_distance_matrix
from src.genetic_algorithm import GeneticAlgorithmTSP

def test_run():
    print("=" * 60)
    print("BẮT ĐẦU KIỂM THỬ THUẬT TOÁN GENETIC ALGORITHM (TSP TP.HCM)")
    print("=" * 60)
    
    # Chuẩn bị ma trận khoảng cách
    dist_matrix = build_distance_matrix(HCMC_LOCATIONS)
    print(f"-> Số lượng điểm giao hàng: {len(HCMC_LOCATIONS)}")
    
    # Khởi tạo GA
    ga = GeneticAlgorithmTSP(
        cities=HCMC_LOCATIONS,
        distance_matrix=dist_matrix,
        pop_size=100,
        elite_size=5,
        crossover_rate=0.85,
        mutation_rate=0.20,
        tournament_size=5,
        mutation_type="inversion"
    )
    
    initial_best = ga.best_individual.distance
    print(f"-> Khoảng cách ngẫu nhiên ban đầu (Gen 0): {initial_best:.2f} km")
    
    # Chạy 150 thế hệ
    print("\nTiến trình tối ưu hoá qua các thế hệ:")
    ga.run(max_generations=150, verbose=True, log_interval=25)
    
    final_best = ga.best_individual.distance
    improvement = ((initial_best - final_best) / initial_best) * 100
    
    print("\n" + "=" * 60)
    print("KẾT QUẢ TỐI ƯU HOÁ:")
    print(f"Quãng đường ban đầu: {initial_best:.2f} km")
    print(f"Quãng đường sau GA:  {final_best:.2f} km")
    print(f"Tỉ lệ cải thiện:     {improvement:.2f}%")
    print("Lộ trình tối ưu tìm được:")
    route_names = [HCMC_LOCATIONS[idx]["name"] for idx in ga.best_individual.genes]
    print(" -> ".join(route_names[:5]) + " -> ... -> " + route_names[0])
    print("=" * 60)
    print("KIỂM THỬ THÀNH CÔNG!")

if __name__ == "__main__":
    test_run()
