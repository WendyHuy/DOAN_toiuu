# -*- coding: utf-8 -*-
"""
genetic_algorithm.py
Cài đặt giải thuật di truyền (Genetic Algorithm - GA) cho bài toán Người giao hàng (TSP).
Bao gồm:
  - Khởi tạo quần thể (Population Initialization)
  - Hàm đánh giá độ thích nghi (Fitness Evaluation)
  - Chọn lọc (Tournament Selection / Roulette Wheel Selection)
  - Lai ghép có thứ tự (Ordered Crossover - OX)
  - Đột biến (Inversion Mutation & Swap Mutation)
  - Chiến lược bảo tồn cá thể ưu tú (Elitism)
  - Mở rộng: Thuật toán Memetic (Hybrid GA kết hợp 2-Opt Local Search)
  - Đánh giá tác động kinh tế & môi trường (Green Logistics)
"""

import random
import copy
import math

class Individual:
    """Đại diện cho một cá thể (một lộ trình di chuyển)."""
    def __init__(self, genes=None):
        # genes: Danh sách các chỉ số địa điểm (hoán vị từ 0 đến N-1)
        self.genes = genes if genes is not None else []
        self.distance = 0.0
        self.fitness = 0.0

    def calculate_fitness(self, distance_matrix):
        """Tính tổng khoảng cách di chuyển và giá trị độ thích nghi (Fitness)."""
        total_dist = 0.0
        n = len(self.genes)
        for i in range(n - 1):
            total_dist += distance_matrix[self.genes[i]][self.genes[i + 1]]
        # Quay trở về điểm xuất phát
        total_dist += distance_matrix[self.genes[-1]][self.genes[0]]
        
        self.distance = round(total_dist, 3)
        # Độ thích nghi tỉ lệ nghịch với tổng khoảng cách
        # Tránh chia cho 0 bằng hằng số nhỏ 1e-6
        self.fitness = 1000.0 / (self.distance + 1e-6)
        return self.fitness

    def __repr__(self):
        return f"Route(Dist={self.distance:.2f}km, Fitness={self.fitness:.4f})"


class GeneticAlgorithmTSP:
    """
    Lớp điều khiển toàn bộ tiến trình tối ưu hoá Genetic Algorithm cho TSP.
    Hỗ trợ cả GA tiêu chuẩn (Pure GA) và GA lai ghép (Memetic Algorithm với 2-Opt).
    """
    def __init__(self, cities, distance_matrix, 
                 pop_size=100, 
                 elite_size=5, 
                 crossover_rate=0.85, 
                 mutation_rate=0.15, 
                 tournament_size=5,
                 mutation_type="inversion",
                 use_local_search=False):
        self.cities = cities
        self.num_cities = len(cities)
        self.distance_matrix = distance_matrix
        
        # Các siêu tham số của GA
        self.pop_size = pop_size
        self.elite_size = min(elite_size, pop_size)
        self.crossover_rate = crossover_rate
        self.mutation_rate = mutation_rate
        self.tournament_size = tournament_size
        self.mutation_type = mutation_type # "inversion" hoặc "swap"
        self.use_local_search = use_local_search # Memetic Algorithm (2-Opt)
        
        # Trạng thái thuật toán
        self.population = []
        self.best_individual = None
        self.history_best_distance = []
        self.history_avg_distance = []
        self.current_generation = 0
        
        self.initialize_population()

    def initialize_population(self):
        """Khởi tạo quần thể ban đầu với các lộ trình ngẫu nhiên."""
        self.population = []
        base_genes = list(range(self.num_cities))
        
        for _ in range(self.pop_size):
            genes = base_genes.copy()
            random.shuffle(genes)
            ind = Individual(genes)
            ind.calculate_fitness(self.distance_matrix)
            self.population.append(ind)
            
        self.update_best()
        self.current_generation = 0
        self.history_best_distance = [self.best_individual.distance]
        avg_d = sum(ind.distance for ind in self.population) / self.pop_size
        self.history_avg_distance = [avg_d]

    def update_best(self):
        """Cập nhật cá thể tốt nhất hiện tại trong quần thể."""
        self.population.sort(key=lambda ind: ind.fitness, reverse=True)
        if self.best_individual is None or self.population[0].fitness > self.best_individual.fitness:
            self.best_individual = copy.deepcopy(self.population[0])

    def selection_tournament(self):
        """Phương pháp chọn lọc giải đấu (Tournament Selection)."""
        tournament = random.sample(self.population, self.tournament_size)
        tournament.sort(key=lambda ind: ind.fitness, reverse=True)
        return tournament[0]

    def selection_roulette_wheel(self):
        """Phương pháp chọn lọc bánh xe Roulette."""
        total_fitness = sum(ind.fitness for ind in self.population)
        if total_fitness <= 0:
            return random.choice(self.population)
        pick = random.uniform(0, total_fitness)
        current = 0
        for ind in self.population:
            current += ind.fitness
            if current >= pick:
                return ind
        return self.population[-1]

    def ordered_crossover(self, parent1, parent2):
        """
        Toán tử lai ghép có thứ tự (Ordered Crossover - OX).
        Toán tử tiêu chuẩn và tối ưu nhất cho bài toán hoán vị TSP:
        1. Chọn 2 điểm cắt ngẫu nhiên trên Cha 1, sao chép đoạn con này sang Con.
        2. Điền các gen còn lại theo thứ tự xuất hiện từ Cha 2 chưa có trong đoạn con.
        """
        if random.random() > self.crossover_rate:
            return copy.deepcopy(parent1), copy.deepcopy(parent2)

        def make_child(p1, p2):
            size = len(p1.genes)
            child_genes = [None] * size
            
            # Chọn 2 điểm cắt
            idx1, idx2 = sorted(random.sample(range(size), 2))
            
            # Sao chép đoạn gen từ p1
            child_genes[idx1:idx2 + 1] = p1.genes[idx1:idx2 + 1]
            copied_set = set(child_genes[idx1:idx2 + 1])
            
            # Điền các vị trí còn lại từ p2 theo vòng tròn
            p2_genes = p2.genes[idx2 + 1:] + p2.genes[:idx2 + 1]
            remaining_genes = [g for g in p2_genes if g not in copied_set]
            
            fill_indices = list(range(idx2 + 1, size)) + list(range(0, idx1))
            for i, gene in zip(fill_indices, remaining_genes):
                child_genes[i] = gene
                
            return Individual(child_genes)

        child1 = make_child(parent1, parent2)
        child2 = make_child(parent2, parent1)
        return child1, child2

    def mutate_swap(self, individual):
        """Đột biến hoán vị (Swap Mutation): Đổi chỗ 2 gen ngẫu nhiên."""
        if random.random() < self.mutation_rate:
            idx1, idx2 = random.sample(range(self.num_cities), 2)
            individual.genes[idx1], individual.genes[idx2] = individual.genes[idx2], individual.genes[idx1]

    def mutate_inversion(self, individual):
        """
        Đột biến đảo đoạn (Inversion Mutation / 2-Opt style):
        Đảo ngược thứ tự của một đoạn gen con ngẫu nhiên.
        Toán tử này cực kỳ hiệu quả trong việc gỡ các nút thắt giao nhau trong TSP.
        """
        if random.random() < self.mutation_rate:
            idx1, idx2 = sorted(random.sample(range(self.num_cities), 2))
            individual.genes[idx1:idx2 + 1] = reversed(individual.genes[idx1:idx2 + 1])

    def mutate(self, individual):
        """Thực hiện biến dị dựa trên cấu hình toán tử."""
        if self.mutation_type == "swap":
            self.mutate_swap(individual)
        else:
            self.mutate_inversion(individual)

    def apply_2opt(self, individual):
        """
        Tối ưu hoá cục bộ 2-Opt (Local Search) cho một cá thể.
        Gỡ bỏ các cạnh giao nhau để tinh chỉnh nghiệm đạt cực tiểu địa phương.
        """
        route = individual.genes.copy()
        n = len(route)
        improved = True
        
        while improved:
            improved = False
            for i in range(1, n - 1):
                for j in range(i + 1, n):
                    if j - i == 1:
                        continue
                    # Chi phí cũ: (i-1 -> i) + (j -> j+1)
                    p_prev, p_i = route[i - 1], route[i]
                    p_j, p_next = route[j], route[(j + 1) % n]
                    d_old = self.distance_matrix[p_prev][p_i] + self.distance_matrix[p_j][p_next]
                    d_new = self.distance_matrix[p_prev][p_j] + self.distance_matrix[p_i][p_next]
                    
                    if d_new < d_old - 1e-5:
                        route[i:j + 1] = reversed(route[i:j + 1])
                        improved = True
                        break
                if improved:
                    break
                    
        individual.genes = route
        individual.calculate_fitness(self.distance_matrix)
        return individual

    def evolve_generation(self):
        """Tiến hành tiến hóa 1 thế hệ mới."""
        self.current_generation += 1
        new_population = []
        
        # 1. Bảo tồn cá thể ưu tú (Elitism)
        self.population.sort(key=lambda ind: ind.fitness, reverse=True)
        for i in range(self.elite_size):
            new_population.append(copy.deepcopy(self.population[i]))
            
        # 2. Sinh các cá thể con cho đến khi đủ kích thước quần thể
        while len(new_population) < self.pop_size:
            # Chọn lọc cha mẹ
            parent1 = self.selection_tournament()
            parent2 = self.selection_tournament()
            
            # Lai ghép
            child1, child2 = self.ordered_crossover(parent1, parent2)
            
            # Đột biến
            self.mutate(child1)
            self.mutate(child2)
            
            # Đánh giá fitness cá thể mới
            child1.calculate_fitness(self.distance_matrix)
            child2.calculate_fitness(self.distance_matrix)
            
            new_population.append(child1)
            if len(new_population) < self.pop_size:
                new_population.append(child2)
                
        # 3. Kỹ thuật Memetic (Local Search trên cá thể tốt nhất nếu được kích hoạt)
        if self.use_local_search:
            # Tối ưu cục bộ cho cá thể dẫn đầu để đẩy nhanh hội tụ
            best_in_pop = new_population[0]
            self.apply_2opt(best_in_pop)
                
        self.population = new_population
        self.update_best()
        
        # Ghi nhận lịch sử hội tụ
        self.history_best_distance.append(self.best_individual.distance)
        avg_d = sum(ind.distance for ind in self.population) / self.pop_size
        self.history_avg_distance.append(avg_d)

    def run(self, max_generations=200, verbose=False, log_interval=20):
        """Chạy thuật toán qua số lượng thế hệ quy định."""
        for gen in range(1, max_generations + 1):
            self.evolve_generation()
            if verbose and (gen % log_interval == 0 or gen == max_generations):
                print(f"[Thế hệ {gen:3d}] Quãng đường tốt nhất: {self.best_individual.distance:.2f} km | Trung bình: {self.history_avg_distance[-1]:.2f} km")
        return self.best_individual

    @staticmethod
    def calculate_economic_impact(distance_saved, fuel_consumption=0.025, fuel_price=24000, avg_speed=25.0):
        """
        Tính toán hiệu quả kinh tế và môi trường (Green Logistics):
          - Nhiên liệu tiết kiệm (lít xăng): Định mức xe máy 2.5 lít / 100 km (0.025 L/km)
          - Chi phí tiết kiệm (VNĐ): Đơn giá xăng RON 95 ~24.000 VNĐ/L
          - Giảm phát thải CO2 (kg): 1 lít xăng tiêu thụ sinh ~2.31 kg CO2
          - Thời gian tiết kiệm (phút): Vận tốc trung bình trong đô thị ~25 km/h
        """
        dist = max(0.0, float(distance_saved))
        fuel_liters = dist * fuel_consumption
        cost_vnd = fuel_liters * fuel_price
        co2_kg = fuel_liters * 2.31
        time_minutes = (dist / avg_speed) * 60.0
        
        return {
            "distance_saved_km": round(dist, 2),
            "fuel_saved_liters": round(fuel_liters, 2),
            "cost_saved_vnd": int(round(cost_vnd, -2)),
            "co2_reduced_kg": round(co2_kg, 2),
            "time_saved_minutes": int(round(time_minutes))
        }

    def get_detailed_itinerary(self):
        """Trả về danh sách các chặng đi chi tiết (Turn-by-turn)."""
        route = self.best_individual.genes
        n = len(route)
        legs = []
        cumulative = 0.0
        
        for i in range(n):
            from_idx = route[i]
            to_idx = route[(i + 1) % n]
            c_from = self.cities[from_idx]
            c_to = self.cities[to_idx]
            leg_dist = self.distance_matrix[from_idx][to_idx]
            cumulative += leg_dist
            
            legs.append({
                "step": i + 1,
                "from_name": c_from["name"],
                "from_district": c_from.get("district", ""),
                "to_name": c_to["name"],
                "to_district": c_to.get("district", ""),
                "distance_km": round(leg_dist, 2),
                "cumulative_km": round(cumulative, 2)
            })
            
        return legs
