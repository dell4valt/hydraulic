import sys
from pathlib import Path
from unittest.mock import patch

import pytest

p = Path(__file__).parents[1].absolute()

sys.path.append(str(p.absolute()))
from hydraulic import lib


# Тесты для функции poly_area
def test_poly_area_1():
    assert lib.poly_area([0, 0, 1, 1], [0, 1, 1, 0]) == 1


def test_poly_area_2():
    assert lib.poly_area([0, 0, 10, 10], [0, 5, 5, 0]) == 50


def test_poly_area_3():
    assert lib.poly_area([0, 2.5, 7.5, 10], [0, 5, 5, 0]) == 37.5


# Тесты для функции question_continue_app
@patch("builtins.input", return_value="да")
def test_question_continue_app_yes(mock_input):
    assert lib.question_continue_app() is True


@patch("builtins.input", side_effect=["invalid", "д"])
def test_question_continue_app_invalid_then_yes(mock_input):
    assert lib.question_continue_app() is True


@patch("builtins.input", return_value="нет")
def test_question_continue_app_no(mock_input):
    with pytest.raises(SystemExit):
        lib.question_continue_app()


# Тесты для функции chunk_list
def test_chunk_list_even_split():
    result = lib.chunk_list([1, 2, 3, 4, 5, 6], 3)
    assert result == [[1, 2], [3, 4], [5, 6]]


def test_chunk_list_uneven_split():
    result = lib.chunk_list([1, 2, 3, 4, 5], 2)
    assert result == [[1, 2, 3], [4, 5]] or result == [[1, 2], [3, 4, 5]]


def test_chunk_list_empty():
    result = lib.chunk_list([], 3)
    assert result == []


# Тесты для функции text_sanitize
def test_text_sanitize_int():
    assert lib.text_sanitize(42) == "42"


def test_text_sanitize_float():
    assert lib.text_sanitize(3.14) == "3.14"


def test_text_sanitize_zero_decimal():
    assert lib.text_sanitize(42.0) == "42"


def test_text_sanitize_with_prefix_suffix():
    assert lib.text_sanitize(42, suffix=" units", prefix="Value: ") == "Value: 42 units"


def test_text_sanitize_with_num_suffix():
    assert lib.text_sanitize(42, num_suffix=" м") == "42 м"


def test_text_sanitize_text():
    assert lib.text_sanitize("Hello") == "Hello"


# Тесты для функции get_pk
def test_get_pk_default():
    assert lib.get_pk(256) == "2+56"


def test_get_pk_with_different_divider():
    assert lib.get_pk(1256, divider=1000) == "1+256"


def test_get_pk_with_decimal():
    assert lib.get_pk(256.75, decimal=True) == "2+56.75"


def test_get_pk_zero():
    assert lib.get_pk(0) == "0+00"


# Тесты для функции floor_float
def test_floor_float_integer():
    assert lib.floor_float(3.75) == 3.0


def test_floor_float_one_decimal():
    assert lib.floor_float(3.75, 1) == 3.7


def test_floor_float_negative():
    assert lib.floor_float(-1.23, 1) == -1.3


def test_floor_float_zero():
    assert lib.floor_float(0, 2) == 0.0


# Тесты для функции closest_upper_multiple
def test_closest_upper_multiple_basic():
    assert lib.closest_upper_multiple(10, 3) == 12.0


def test_closest_upper_multiple_exact():
    assert lib.closest_upper_multiple(10, 5) == 10.0


def test_closest_upper_multiple_zero():
    assert lib.closest_upper_multiple(0, 5) == 0.0


def test_closest_upper_multiple_negative():
    assert lib.closest_upper_multiple(-10, 3) == -9.0


def test_closest_upper_multiple_invalid():
    with pytest.raises(ValueError):
        lib.closest_upper_multiple(10, 0)


# Тесты для функции split_list_by_min_value
def test_split_list_by_min_value_middle():
    assert lib.split_list_by_min_value([3, 1, 4, 2]) == [[3], [1, 4, 2]]


def test_split_list_by_min_value_start():
    assert lib.split_list_by_min_value([1, 3, 4, 2]) == [[1], [3, 4, 2]]


def test_split_list_by_min_value_end():
    assert lib.split_list_by_min_value([3, 4, 2, 1]) == [[3, 4, 2], [1]]


def test_split_list_by_min_value_too_short():
    with pytest.raises(ValueError):
        lib.split_list_by_min_value([1])


# Тесты для функции calculate_line_length
def test_calculate_line_length_straight_line():
    assert lib.calculate_line_length([0, 3], [0, 4]) == 5.0


def test_calculate_line_length_zigzag():
    assert round(lib.calculate_line_length([0, 3, 5], [0, 4, 7]), 4) == 8.6056


def test_calculate_line_length_zigzag_2():
    assert round(lib.calculate_line_length([0, 0, 1, 2], [0, 1, 1, 2]), 4) == 3.4142


def test_calculate_line_length_single_point():
    assert lib.calculate_line_length([1], [1]) == 0.0


def test_calculate_line_length_empty():
    assert lib.calculate_line_length([], []) == 0.0


def test_calculate_line_length_unequal_arrays():
    assert lib.calculate_line_length([1, 2], [1]) == 0.0


# Тест для функции rmdir с использованием временной директории
@pytest.fixture
def temp_dir(tmp_path):
    dir_path = tmp_path / "test_dir"
    dir_path.mkdir()

    # Создаем подпапку и файлы
    subdir = dir_path / "subdir"
    subdir.mkdir()

    test_file = dir_path / "test.txt"
    test_file.write_text("test content")

    subdir_file = subdir / "subfile.txt"
    subdir_file.write_text("subfile content")

    return dir_path


def test_rmdir(temp_dir):
    # Проверяем что директория существует
    assert temp_dir.exists()

    # Удаляем директорию
    lib.rmdir(str(temp_dir))

    # Проверяем что директория была удалена
    assert not temp_dir.exists()
