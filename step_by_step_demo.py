# -*- coding: utf-8 -*-
"""
step_by_step_demo.py
Minh hoa chi tiet tung buoc cua Giai thuat Di truyen (GA) cho bai toan TSP.
Module nay duoc su dung de in ra cac buoc tinh toan cu the, giup viet bao cao.
"""

import sys
import os
import random
import copy

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src.cities_data import HCMC_LOCATIONS, build_distance_matrix
from src.genetic_algorithm import Individual


def demo_step_by_step():
    """Minh hoa tung buoc cua GA voi tap du lieu nho (6 diem dau tien)."""
    
    # Su dung 6 diem dau de minh hoa de hieu
    small_cities = HCMC_LOCATIONS[:6]
    dist_matrix = build_distance_matrix(small_cities)
    num_cities = len(small_cities)
    
    print("=" * 70)
    print("   MINH HOA TUNG BUOC GIAI THUAT DI TRUYEN (GA) CHO BAI TOAN TSP")
    print("=" * 70)
    
    # === BUOC 0: Du lieu dau vao ===
    print("\n" + "=" * 70)
    print("BUOC 0: DU LIEU DAU VAO (INPUT)")
    print("=" * 70)
    print(f"\nSo luong diem giao hang: {num_cities}")
    print("\nDanh sach cac diem:")
    for c in small_cities:
        print(f"  ID={c['id']}: {c['name']:<35} Toa do: ({c['x']:.1f}, {c['y']:.1f})")
    
    print("\nMa tran khoang cach (km):")
    header = "      " + "  ".join([f"  D{i:<3}" for i in range(num_cities)])
    print(header)
    for i in range(num_cities):
        row = f"  D{i}  "
        for j in range(num_cities):
            row += f" {dist_matrix[i][j]:5.2f} "
        print(row)
    
    # === BUOC 1: Khoi tao quan the ===
    print("\n" + "=" * 70)
    print("BUOC 1: KHOI TAO QUAN THE BAN DAU (Population Initialization)")
    print("=" * 70)
    
    pop_size = 6
    random.seed(42)  # Seed co dinh de ket qua lap lai duoc
    population = []
    base = list(range(num_cities))
    
    for k in range(pop_size):
        genes = base.copy()
        random.shuffle(genes)
        ind = Individual(genes)
        ind.calculate_fitness(dist_matrix)
        population.append(ind)
    
    print(f"\nKich thuoc quan the: {pop_size}")
    print("Moi ca the la 1 hoan vi cua {0, 1, 2, ..., N-1} dai dien cho thu tu giao hang.\n")
    
    for k, ind in enumerate(population):
        route_str = " -> ".join([str(g) for g in ind.genes]) + f" -> {ind.genes[0]}"
        ten_route = " -> ".join([small_cities[g]["name"][:12] for g in ind.genes])
        print(f"  Ca the {k+1}: [{', '.join(map(str, ind.genes))}]")
        print(f"    Lo trinh: {route_str}")
        print(f"    Dia diem: {ten_route}")
        print(f"    Tong quang duong: {ind.distance:.2f} km")
        print(f"    Fitness = 1000 / {ind.distance:.2f} = {ind.fitness:.4f}")
        print()
    
    # === BUOC 2: Danh gia va xep hang ===
    print("=" * 70)
    print("BUOC 2: DANH GIA VA XEP HANG DO THICH NGHI (Fitness Evaluation)")
    print("=" * 70)
    
    population.sort(key=lambda x: x.fitness, reverse=True)
    print("\nXep hang tu tot nhat den xau nhat:")
    for rank, ind in enumerate(population):
        print(f"  Hang {rank+1}: Genes={ind.genes}  Dist={ind.distance:.2f} km  Fitness={ind.fitness:.4f}")
    
    best = population[0]
    print(f"\n=> Ca the tot nhat hien tai: {best.genes}  ({best.distance:.2f} km)")
    
    # === BUOC 3: Chon loc Tournament ===
    print("\n" + "=" * 70)
    print("BUOC 3: CHON LOC (Tournament Selection, k=3)")
    print("=" * 70)
    
    random.seed(100)
    tournament_size = 3
    
    # Chon parent 1
    t1_indices = random.sample(range(pop_size), tournament_size)
    t1 = [population[i] for i in t1_indices]
    t1.sort(key=lambda x: x.fitness, reverse=True)
    parent1 = t1[0]
    
    print(f"\nChon Cha 1 (Parent 1):")
    print(f"  Chon ngau nhien {tournament_size} ca the tu quan the:")
    for i, idx in enumerate(t1_indices):
        ind = population[idx]
        winner = " <-- THANG" if ind is parent1 else ""
        print(f"    Ca the hang {idx+1}: Genes={ind.genes}  Fitness={ind.fitness:.4f}{winner}")
    print(f"  => Parent 1: {parent1.genes}  (Fitness={parent1.fitness:.4f})")
    
    # Chon parent 2
    t2_indices = random.sample(range(pop_size), tournament_size)
    t2 = [population[i] for i in t2_indices]
    t2.sort(key=lambda x: x.fitness, reverse=True)
    parent2 = t2[0]
    
    print(f"\nChon Cha 2 (Parent 2):")
    print(f"  Chon ngau nhien {tournament_size} ca the tu quan the:")
    for i, idx in enumerate(t2_indices):
        ind = population[idx]
        winner = " <-- THANG" if ind is parent2 else ""
        print(f"    Ca the hang {idx+1}: Genes={ind.genes}  Fitness={ind.fitness:.4f}{winner}")
    print(f"  => Parent 2: {parent2.genes}  (Fitness={parent2.fitness:.4f})")
    
    # === BUOC 4: Lai ghep OX ===
    print("\n" + "=" * 70)
    print("BUOC 4: LAI GHEP CO THU TU (Ordered Crossover - OX)")
    print("=" * 70)
    
    p1_genes = parent1.genes.copy()
    p2_genes = parent2.genes.copy()
    size = len(p1_genes)
    
    # Chon 2 diem cat co dinh de minh hoa ro
    idx1, idx2 = 1, 3
    
    print(f"\n  Parent 1 (P1): {p1_genes}")
    print(f"  Parent 2 (P2): {p2_genes}")
    print(f"\n  Diem cat: vi tri {idx1} den {idx2}")
    
    # Tao child 1
    child1_genes = [None] * size
    child1_genes[idx1:idx2+1] = p1_genes[idx1:idx2+1]
    copied = set(child1_genes[idx1:idx2+1])
    
    print(f"\n  --- Tao Child 1 ---")
    print(f"  B1. Sao chep doan gen tu P1[{idx1}:{idx2}]: {child1_genes}")
    print(f"      Cac gen da sao chep: {copied}")
    
    p2_order = p2_genes[idx2+1:] + p2_genes[:idx2+1]
    remaining = [g for g in p2_order if g not in copied]
    print(f"  B2. Thu tu gen tu P2 (bat dau sau diem cat): {p2_order}")
    print(f"      Loc bo cac gen da co: {remaining}")
    
    fill_pos = list(range(idx2+1, size)) + list(range(0, idx1))
    for pos, gene in zip(fill_pos, remaining):
        child1_genes[pos] = gene
    
    print(f"  B3. Dien vao cac vi tri trong: {fill_pos}")
    print(f"  => Child 1: {child1_genes}")
    
    child1_ind = Individual(child1_genes)
    child1_ind.calculate_fitness(dist_matrix)
    print(f"     Quang duong Child 1: {child1_ind.distance:.2f} km")
    
    # === BUOC 5: Dot bien Inversion ===
    print("\n" + "=" * 70)
    print("BUOC 5: DOT BIEN DAO DOAN (Inversion Mutation)")
    print("=" * 70)
    
    before_mut = child1_genes.copy()
    m1, m2 = 2, 4
    
    print(f"\n  Truoc dot bien:  {before_mut}")
    print(f"  Chon doan gen tu vi tri {m1} den {m2}: {before_mut[m1:m2+1]}")
    
    child1_genes[m1:m2+1] = list(reversed(child1_genes[m1:m2+1]))
    print(f"  Dao nguoc doan: {child1_genes[m1:m2+1]}")
    print(f"  Sau dot bien:    {child1_genes}")
    
    child1_mut = Individual(child1_genes)
    child1_mut.calculate_fitness(dist_matrix)
    print(f"\n  Quang duong truoc dot bien: {child1_ind.distance:.2f} km")
    print(f"  Quang duong sau dot bien:  {child1_mut.distance:.2f} km")
    change = child1_ind.distance - child1_mut.distance
    if change > 0:
        print(f"  => Dot bien GIAM duoc {change:.2f} km (tot hon!)")
    else:
        print(f"  => Dot bien TANG {abs(change):.2f} km (se bi chon loc loai bo o the he sau)")
    
    # === BUOC 6: Elitism ===
    print("\n" + "=" * 70)
    print("BUOC 6: CHIEN LUOC UU TU (Elitism)")
    print("=" * 70)
    
    elite_size = 2
    print(f"\n  Giu lai {elite_size} ca the tot nhat khong thay doi sang the he moi:")
    for i in range(elite_size):
        ind = population[i]
        print(f"    Elite {i+1}: {ind.genes}  Dist={ind.distance:.2f} km")
    print(f"\n  Cac vi tri con lai ({pop_size - elite_size} ca the) duoc dien bang con lai")
    print(f"  (duoc sinh tu lai ghep va dot bien nhu tren).")
    
    # === BUOC 7: Lap lai qua nhieu the he ===
    print("\n" + "=" * 70)
    print("BUOC 7: LAP LAI QUA NHIEU THE HE (Iterations)")
    print("=" * 70)
    print(f"""
  Chu trinh tren (Buoc 2 -> 6) duoc lap lai qua nhieu the he (Generations).
  Sau moi the he:
    - Quan the duoc cap nhat voi cac ca the moi (con) va ca the uu tu (elite).
    - Ca the tot nhat (best) duoc ghi nhan.
    - Quang duong trung binh cua quan the giam dan (hoi tu).
  
  Dieu kien dung: Dat so the he toi da hoac quang duong khong cai thien.

  Vi du ket qua qua 50 the he voi 6 diem:""")
    
    from src.genetic_algorithm import GeneticAlgorithmTSP
    random.seed(42)
    ga_small = GeneticAlgorithmTSP(
        cities=small_cities,
        distance_matrix=dist_matrix,
        pop_size=20,
        elite_size=2,
        crossover_rate=0.85,
        mutation_rate=0.20,
        tournament_size=3,
        mutation_type="inversion"
    )
    
    init_d = ga_small.best_individual.distance
    for g in range(1, 51):
        ga_small.evolve_generation()
        if g % 10 == 0:
            print(f"    The he {g:3d}: Best = {ga_small.best_individual.distance:.2f} km | Avg = {ga_small.history_avg_distance[-1]:.2f} km")
    
    final_d = ga_small.best_individual.distance
    imp = ((init_d - final_d) / init_d) * 100
    
    print(f"\n  => Quang duong ban dau:  {init_d:.2f} km")
    print(f"  => Quang duong toi uu:  {final_d:.2f} km")
    print(f"  => Ti le cai thien:     {imp:.1f}%")
    print(f"  => Lo trinh toi uu:     {ga_small.best_individual.genes}")
    
    route_names = [small_cities[i]["name"] for i in ga_small.best_individual.genes]
    print(f"     " + " -> ".join(route_names) + " -> " + route_names[0])
    
    print("\n" + "=" * 70)
    print("KET THUC MINH HOA TUNG BUOC. Xem them chuong trinh chinh: run_cli.py")
    print("=" * 70)


if __name__ == "__main__":
    demo_step_by_step()
