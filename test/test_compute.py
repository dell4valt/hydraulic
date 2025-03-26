import sys
import pytest
import numpy as np
from pathlib import Path
from unittest.mock import patch

# Добавляем родительскую директорию в путь поиска модулей
p = Path(__file__).parents[1].absolute()
sys.path.append(str(p.absolute()))

from hydraulic.profile import ComputeCVQ
from hydraulic import config


class TestComputeCVQ:
    """Тесты для класса ComputeCVQ"""

    def setup_method(self):
        """Создание базового объекта для тестов"""
        # Создаем базовый объект для тестов с типичными параметрами
        self.default_compute = ComputeCVQ(
            n=0.035,  # Коэффициент шероховатости
            i=3.0,  # Уклон в промилле
            h=1.5,  # Средняя глубина
            a=10.0,  # Площадь водного сечения
            p=9.0,  # Смоченный периметр
            r=1.11,  # Гидравлический радиус
        )

    def test_initialization(self):
        """Тест корректной инициализации объекта"""
        # Проверяем, что объект правильно инициализируется с заданными параметрами
        compute = ComputeCVQ(n=0.04, i=2.5, h=1.2, a=8.0, p=8.0, r=1.0)

        assert compute.n == 0.04
        assert compute.i == 2.5
        assert compute.h == 1.2
        assert compute.a == 8.0
        assert compute.p == 8.0
        assert compute.r == 1.0
        assert compute._g == pytest.approx(9.80665)
        # assert compute.type__ == "Не определен"
        assert compute.v != 0  # Скорость должна быть рассчитана
        assert compute.q != 0  # Расход должен быть рассчитан
        assert compute.shezi != 0  # Коэффициент Шези должен быть рассчитан

    def test_calc_consumption(self):
        """Тест расчета расхода воды"""
        # Проверяем функцию расчета расхода
        result = self.default_compute._calc_consumption(10.0, 2.0)
        assert result == 20.0  # Расход = площадь * скорость

    @patch.object(config, "USE_H_INSTEAD_R", True)
    def test_use_h_instead_r(self):
        """Тест использования глубины вместо гидравлического радиуса"""
        # Проверяем, что при USE_H_INSTEAD_R=True используется h вместо r
        compute = ComputeCVQ(n=0.035, i=3.0, h=1.5, a=10.0, p=9.0, r=1.11)

        # При USE_H_INSTEAD_R=True, r должен быть равен h
        assert compute.r == 1.5

    def test_calc_velocity_regular_water(self):
        """Тест расчета скорости для обычной воды"""
        # Параметры: уклон 3 промилле, коэффициент Шези 40, гидравлический радиус 1.2
        with patch.object(config, "CALC_TYPE", 1):
            result = self.default_compute._calc_velocity(3.0, 40.0, 1.2, calc_type=1)
            expected = 40.0 * np.sqrt(1.2 * 0.003)
            assert result == pytest.approx(expected)

    def test_calc_velocity_nanosovodny_sel(self):
        """Тест расчета скорости для наносоводного селя"""
        with patch.object(config, "CALC_TYPE", 2):
            result = self.default_compute._calc_velocity(3.0, 40.0, 1.2, calc_type=2)
            expected = 4.5 * 1.2**0.67 * (3.0 / 1000) ** 0.17
            assert result == pytest.approx(expected)

    def test_calc_velocity_mudstone_sel(self):
        """Тест расчета скорости для грязекаменного селя"""
        with patch.object(config, "CALC_TYPE", 3):
            result = self.default_compute._calc_velocity(3.0, 40.0, 1.2, calc_type=3)
            expected = 3.75 * 1.2**0.50 * (3.0 / 1000) ** 0.17
            assert result == pytest.approx(expected)

    def test_calc_velocity_invalid_type(self):
        """Тест обработки неверного типа расчета"""
        with patch.object(config, "CALC_TYPE", 999), pytest.raises(ValueError):
            self.default_compute._calc_velocity(3.0, 40.0, 1.2, calc_type=999)

    def test_calc_shezi_manning(self):
        """Тест расчета коэффициента Шези по формуле Маннинга"""
        with patch.object(config, "SHEZI_TYPE", "Маннинга"):
            compute = ComputeCVQ(n=0.035, i=3.0, h=1.5, a=10.0, p=9.0, r=1.11)
            expected = (1 / 0.035) * 1.11 ** (1 / 6)
            assert compute.shezi == pytest.approx(expected)
            assert "Маннинга" in compute.type__

    def test_calc_shezi_pavlovskij(self):
        """Тест расчета коэффициента Шези по формуле Павловского"""
        with patch.object(config, "SHEZI_TYPE", "Павловского"):
            compute = ComputeCVQ(n=0.035, i=3.0, h=1.5, a=10.0, p=9.0, r=1.11)
            sqrt_n = np.sqrt(0.035)
            sqrt_r = np.sqrt(1.11)
            y = 2.5 * sqrt_n - 0.13 - 0.75 * sqrt_r * (sqrt_n - 0.10)
            expected = (1 / 0.035) * 1.11**y
            assert compute.shezi == pytest.approx(expected)
            assert "Павловского" in compute.type__

    def test_calc_shezi_zheleznjakov(self):
        """Тест расчета коэффициента Шези по формуле Железнякова"""
        with patch.object(config, "SHEZI_TYPE", "Железнякова"):
            compute = ComputeCVQ(n=0.035, i=3.0, h=1.5, a=10.0, p=9.0, r=1.11)
            sqrt_g = np.sqrt(9.80665)
            log_r = np.log10(1.11)

            term1 = 1 / 2 * ((1 / 0.035) - (sqrt_g / 0.13) * (1 - log_r))
            term2 = (1 / 4) * (1 / 0.035 - (sqrt_g / 0.13) * (1 - log_r)) ** 2
            term3 = (sqrt_g / 0.13) * ((1 / 0.035) + (sqrt_g * log_r))

            expected = term1 + np.sqrt(term2 + term3)
            assert compute.shezi == pytest.approx(expected)
            assert "Железнякова" in compute.type__

    def test_calc_shezi_pavlovskij_zheleznjakov(self):
        """Тест расчета коэффициента Шези по формуле Павловского-Железнякова"""
        with patch.object(config, "SHEZI_TYPE", "Павловского-Железнякова"):
            compute = ComputeCVQ(n=0.035, i=3.0, h=1.5, a=10.0, p=9.0, r=1.11)
            sqrt_g = np.sqrt(9.80665)
            log_r = np.log10(1.11)

            term1 = 1 / 2 - (0.035 * sqrt_g / 0.26) * (1 - log_r)
            term2 = 1 / 4 * (1 / 0.035 - sqrt_g / 0.13 * (1 - log_r)) ** 2
            term3 = sqrt_g / 0.13 * (1 / 0.035 + sqrt_g * log_r)

            y = (1 / log_r) * np.log10(term1 + 0.035 * np.sqrt(term2 + term3))

            expected = (1 / 0.035) * 1.11**y
            assert compute.shezi == pytest.approx(expected)
            assert "Павловского" in compute.type__
            assert "Железнякова" in compute.type__

    def test_calc_shezi_agroskina(self):
        """Тест расчета коэффициента Шези по формуле Агроскина"""
        with patch.object(config, "SHEZI_TYPE", "Агроскина"):
            compute = ComputeCVQ(n=0.035, i=3.0, h=1.5, a=10.0, p=9.0, r=1.11)
            expected = (1 / 0.035) + 17.72 * np.log(1.11)
            assert compute.shezi == pytest.approx(expected)
            assert "Агроскина" in compute.type__

    def test_calc_shezi_sp(self):
        """Тест расчета коэффициента Шези по формуле СП"""
        with patch.object(config, "SHEZI_TYPE", "СП"):
            compute = ComputeCVQ(n=0.035, i=3.0, h=1.5, a=10.0, p=9.0, r=1.11)
            sqrt_n = np.sqrt(0.035)
            sqrt_r = np.sqrt(1.11)
            y = 2.5 * sqrt_n - 0.13 - 0.75 * sqrt_r * (sqrt_n - 0.10)
            expected = (1 / 0.035) * 1.11**y
            assert compute.shezi == pytest.approx(expected)
            assert "Павловского" in compute.type__

    def test_calc_shezi_gidroraschet(self):
        """Тест расчета коэффициента Шези по формуле Гидрорасчеты"""
        # Тест для r <= 3 (должна использоваться формула Павловского)
        with patch.object(config, "SHEZI_TYPE", "Гидрорасчеты"):
            compute = ComputeCVQ(n=0.035, i=3.0, h=1.5, a=10.0, p=9.0, r=2.5)
            sqrt_n = np.sqrt(0.035)
            sqrt_r = np.sqrt(2.5)
            y = 2.5 * sqrt_n - 0.13 - 0.75 * sqrt_r * (sqrt_n - 0.10)
            expected = (1 / 0.035) * 2.5**y
            assert compute.shezi == pytest.approx(expected)
            assert "Павловского" in compute.type__

        # Тест для r > 3 (должна использоваться формула Павловского-Железнякова)
        with patch.object(config, "SHEZI_TYPE", "Гидрорасчеты"):
            compute = ComputeCVQ(n=0.035, i=3.0, h=1.5, a=10.0, p=9.0, r=3.5)
            sqrt_g = np.sqrt(9.80665)
            log_r = np.log10(3.5)

            term1 = 1 / 2 - (0.035 * sqrt_g / 0.26) * (1 - log_r)
            term2 = 1 / 4 * (1 / 0.035 - sqrt_g / 0.13 * (1 - log_r)) ** 2
            term3 = sqrt_g / 0.13 * (1 / 0.035 + sqrt_g * log_r)

            y = (1 / log_r) * np.log10(term1 + 0.035 * np.sqrt(term2 + term3))

            expected = (1 / 0.035) * 3.5**y
            assert compute.shezi == pytest.approx(expected)
            assert "Павловского" in compute.type__
            assert "Железнякова" in compute.type__

    def test_calc_shezi_invalid_type(self):
        """Тест обработки неверного типа формулы расчета коэффициента Шези"""
        # Тестируем с несуществующим типом формулы
        with patch.object(config, "SHEZI_TYPE", "Несуществующая формула"), pytest.raises(ValueError):
            ComputeCVQ(n=0.035, i=3.0, h=1.5, a=10.0, p=9.0, r=1.11)

    def test_complete_calculation_regular_water(self):
        """Интеграционный тест полного расчета для обычной воды"""
        with patch.object(config, "CALC_TYPE", 1), patch.object(config, "SHEZI_TYPE", "Маннинга"):
            compute = ComputeCVQ(n=0.035, i=3.0, h=1.5, a=10.0, p=9.0, r=1.11)

            # Проверяем Шези по Маннингу
            expected_shezi = (1 / 0.035) * 1.11 ** (1 / 6)
            assert compute.shezi == pytest.approx(expected_shezi)

            # Проверяем скорость
            expected_v = expected_shezi * np.sqrt(1.11 * 0.003)
            assert compute.v == pytest.approx(expected_v)

            # Проверяем расход
            expected_q = 10.0 * expected_v
            assert compute.q == pytest.approx(expected_q)

    def test_complete_calculation_nanosovodny_sel(self):
        """Интеграционный тест полного расчета для наносоводного селя"""
        with patch.object(config, "CALC_TYPE", 2), patch.object(config, "SHEZI_TYPE", "Маннинга"):
            compute = ComputeCVQ(n=0.035, i=3.0, h=1.5, a=10.0, p=9.0, r=1.11)

            # Проверяем скорость для наносоводного селя
            expected_v = 4.5 * 1.11**0.67 * (3.0 / 1000) ** 0.17
            assert compute.v == pytest.approx(expected_v)

            # Проверяем расход
            expected_q = 10.0 * expected_v
            assert compute.q == pytest.approx(expected_q)

    def test_complete_calculation_mudstone_sel(self):
        """Интеграционный тест полного расчета для грязекаменного селя"""
        with patch.object(config, "CALC_TYPE", 3), patch.object(config, "SHEZI_TYPE", "Маннинга"):
            compute = ComputeCVQ(n=0.035, i=3.0, h=1.5, a=10.0, p=9.0, r=1.11)

            # Проверяем скорость для грязекаменного селя
            expected_v = 3.75 * 1.11**0.50 * (3.0 / 1000) ** 0.17
            assert compute.v == pytest.approx(expected_v)

            # Проверяем расход
            expected_q = 10.0 * expected_v
            assert compute.q == pytest.approx(expected_q)
