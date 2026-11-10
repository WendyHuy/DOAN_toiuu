# -*- coding: utf-8 -*-
"""
cities_data.py
Định nghĩa tập dữ liệu các địa điểm giao hàng thực tế tại TP. Hồ Chí Minh.
Tọa độ được tính theo km (gốc tọa độ tại Kho Trung tâm - Bưu điện Trung tâm Sài Gòn).
"""

import math

# Danh sách 20 điểm giao hàng thực tế tại TP. Hồ Chí Minh
# Mỗi điểm gồm: Tên địa điểm, Quận, Tọa độ X (km), Tọa độ Y (km) tính từ tâm Quận 1
HCMC_LOCATIONS = [
    {"id": 0,  "name": "Kho Tổng (Bưu điện TT Sài Gòn)", "district": "Quận 1",       "x": 0.0,   "y": 0.0,   "lat": 10.7798, "lon": 106.6999},
    {"id": 1,  "name": "Chợ Bến Thành",                "district": "Quận 1",       "x": -0.8,  "y": -0.9,  "lat": 10.7725, "lon": 106.6980},
    {"id": 2,  "name": "Hồ Con Rùa",                    "district": "Quận 3",       "x": -0.2,  "y": 0.6,   "lat": 10.7828, "lon": 106.6958},
    {"id": 3,  "name": "Landmark 81 (Vinhomes)",        "district": "Bình Thạnh",   "x": 3.2,   "y": 1.4,   "lat": 10.7950, "lon": 106.7218},
    {"id": 4,  "name": "Sân bay Tân Sơn Nhất",         "district": "Tân Bình",     "x": -4.2,  "y": 4.5,   "lat": 10.8185, "lon": 106.6588},
    {"id": 5,  "name": "Công viên Gia Định",           "district": "Gò Vấp",       "x": -2.5,  "y": 4.8,   "lat": 10.8190, "lon": 106.6770},
    {"id": 6,  "name": "Khu chế xuất Tân Thuận",       "district": "Quận 7",       "x": 3.5,   "y": -3.8,  "lat": 10.7480, "lon": 106.7330},
    {"id": 7,  "name": "Crescent Mall (Phú Mỹ Hưng)",  "district": "Quận 7",       "x": 1.8,   "y": -5.6,  "lat": 10.7295, "lon": 106.7198},
    {"id": 8,  "name": "Bến xe Miền Đông Cũ",          "district": "Bình Thạnh",   "x": 1.5,   "y": 3.9,   "lat": 10.8143, "lon": 106.7127},
    {"id": 9,  "name": "Đại học Văn Lang (Cơ sở 3)",    "district": "Gò Vấp",       "x": -3.8,  "y": 6.2,   "lat": 10.8276, "lon": 106.6852},
    {"id": 10, "name": "Khu Công nghệ Cao (SHTP)",      "district": "TP. Thủ Đức",  "x": 11.5,  "y": 8.0,   "lat": 10.8540, "lon": 106.7900},
    {"id": 11, "name": "ĐHQG TP.HCM (Làng Đại học)",   "district": "TP. Thủ Đức",  "x": 12.8,  "y": 11.2,  "lat": 10.8752, "lon": 106.8007},
    {"id": 12, "name": "Công viên Đầm Sen",            "district": "Quận 11",      "x": -6.5,  "y": -1.2,  "lat": 10.7680, "lon": 106.6375},
    {"id": 13, "name": "Bến xe Chợ Lớn",               "district": "Quận 5",       "x": -4.8,  "y": -3.1,  "lat": 10.7516, "lon": 106.6534},
    {"id": 14, "name": "Aeon Mall Bình Tân",           "district": "Bình Tân",     "x": -9.2,  "y": -3.5,  "lat": 10.7435, "lon": 106.6125},
    {"id": 15, "name": "Bến xe An Sương",              "district": "Quận 12",      "x": -8.5,  "y": 8.5,   "lat": 10.8492, "lon": 106.6186},
    {"id": 16, "name": "Chợ Thủ Đức",                  "district": "TP. Thủ Đức",  "x": 7.8,   "y": 6.5,   "lat": 10.8505, "lon": 106.7583},
    {"id": 17, "name": "Đại học Văn Lang (Cơ sở 1)",    "district": "Quận 1",       "x": -0.6,  "y": -0.4,  "lat": 10.7672, "lon": 106.6938},
    {"id": 18, "name": "Hàng Xanh Interchange",        "district": "Bình Thạnh",   "x": 2.1,   "y": 2.2,   "lat": 10.8016, "lon": 106.7118},
    {"id": 19, "name": "Công viên Hoàng Văn Thụ",       "district": "Tân Bình",     "x": -3.2,  "y": 3.1,   "lat": 10.8012, "lon": 106.6625}
]

def calculate_distance(city_a, city_b):
    """Tính khoảng cách Euclidean (km) giữa 2 điểm."""
    dx = city_a["x"] - city_b["x"]
    dy = city_a["y"] - city_b["y"]
    return math.sqrt(dx * dx + dy * dy)

def build_distance_matrix(cities):
    """Xây dựng ma trận khoảng cách giữa tất cả các cặp điểm."""
    n = len(cities)
    matrix = [[0.0] * n for _ in range(n)]
    for i in range(n):
        for j in range(n):
            if i != j:
                matrix[i][j] = calculate_distance(cities[i], cities[j])
    return matrix
