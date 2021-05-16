import sys
from pathlib import Path
p = Path(__file__).parents[1].absolute()

sys.path.append(str(p.absolute()))
from hydraulic import lib


# Квадрат
def test_poly_area_1():
    assert lib.poly_area([0, 0, 1, 1], [0, 1, 1, 0]) == 1


# Прямоугольник
def test_poly_area_2():
    assert lib.poly_area([0, 0, 10, 10], [0, 5, 5, 0]) == 50


# Трапеция
def test_poly_area_3():
    assert lib.poly_area([0, 2.5, 7.5, 10], [0, 5, 5, 0]) == 37.5
