# -*- coding: utf-8 -*-
"""
Библиотека вспомогательных функций

Этот модуль содержит набор функций, используемых в гидравлических расчетах, обработке
данных поперечных профилей водных объектов и формирования отчетной документации.

Функции:
    question_continue_app: Запрашивает у пользователя продолжение работы программы.
    poly_area: Вычисляет площадь многоугольника по координатам вершин.
    chunk_list: Разбивает список на указанное количество частей.
    insert_summary_QV_tables: Формирует и вставляет таблицы с расчетными данными в отчет.
    text_sanitize: Форматирует текст или числовые значения для отображения.
    rmdir: Рекурсивно удаляет директорию и все её содержимое.
    get_pk: Возвращает строку пикетажа в формате 'км+м'.
    floor_float: Округляет число вниз с заданной точностью.
    closest_upper_multiple: Находит ближайшее большее кратное число.
    split_list_by_min_value: Разделяет список на две части по минимальному значению.
    get_water_sections: Определяет секторы воды при заданном уровне.
    calculate_line_length: Вычисляет длину ломаной линии по координатам точек.
"""

import sys
from pathlib import Path

import numpy as np


def question_continue_app() -> bool:
    """
    Запрашивает у пользователя решение о продолжении работы программы.

    Функция выводит запрос пользователю и ожидает ответа. В зависимости от
    полученного ответа продолжает выполнение программы или завершает её.

    Допустимые ответы для продолжения: "да", "д", "yes", "y", "ага".
    Допустимые ответы для завершения: "нет", "н", "no", "n".

    При получении недопустимого ответа запрос повторяется.

    Returns:
        bool: Функция возвращает True при положительном ответе.

    Raises:
        SystemExit: При отрицательном ответе программа завершается.
    """
    while True:
        answer = input("Продолжить расчет? (да/нет)")
        if answer.lower() in ["да", "д", "yes", "y", "ага"]:
            print("Хорошо. Продолжаем расчет.\n")
            return True
        elif answer.lower() in ["no", "нет", "n", "н"]:
            print("Программа будет завершена.\n")
            sys.exit()
        else:
            continue


def poly_area(x: list, y: list) -> float:
    """
    Вычисляет площадь многоугольника по координатам вершин.

    Функция использует формулу площади Гаусса (известную также как формула шнурования),
    которая вычисляет площадь многоугольника через координаты его вершин.

    Args:
        x (list, numpy.ndarray): Список или массив координат x вершин многоугольника.
        y (list, numpy.ndarray): Список или массив координат y вершин многоугольника.

    Returns:
        float: Площадь многоугольника.

    Note:
        Координаты вершин должны быть упорядочены либо по часовой, либо против
        часовой стрелки. Функция работает для выпуклых и невыпуклых многоугольников.

    Examples:
        >>> poly_area([0, 1, 1, 0], [0, 0, 1, 1])
        1.0
        >>> poly_area([0, 2, 2, 0], [0, 0, 2, 2])
        4.0
    """
    return 0.5 * np.abs(np.dot(x, np.roll(y, 1)) - np.dot(y, np.roll(x, 1)))


def chunk_list(seq, num):
    """
    Разбивает заданный список на указанное количество примерно равных частей.

    Функция делит список seq на num подсписков с приблизительно одинаковым
    количеством элементов в каждом. Размер каждого подсписка определяется
    средним значением от общего количества элементов, деленного на num.

    Args:
        seq (list): Исходный список для разбиения.
        num (int): Количество частей, на которое нужно разбить список.

    Returns:
        list: Список подсписков, содержащих элементы исходного списка.
            Количество подсписков равно num.

    Examples:
        >>> chunk_list([1, 2, 3, 4, 5, 6], 3)
        [[1, 2], [3, 4], [5, 6]]
        >>> chunk_list([1, 2, 3, 4, 5], 2)
        [[1, 2, 3], [4, 5]]
    """
    avg = len(seq) / float(num)
    out = []
    last = 0.0

    while last < len(seq):
        out.append(seq[int(last) : int(last + avg)])
        last += avg

    return out


def text_sanitize(text, suffix="", prefix="", num_suffix=""):
    """Возвращает входной параметр text. В случае если число целое,
    возвращает без десятичных нулей. Если не целое, с указанием десятых.
    Можно задать префикс и суффикс соответствующими параметрами.

    Args:
        text (str, int, float): Входящая строка или число
        suffix (str, optional): Окончание возвращаемой строки. По умолчанию ''.
        prefix (str, optional): Начало возвращаемой строки. По умолчанию ''.
        num_suffix (str, optional): Окончание возвращаемой строки
        только если на входе число. По умолчанию ''.

    Returns:
        _type_: Возвращаемая строка
    """

    try:
        return f"{prefix}{text:g}{num_suffix}{suffix}"
    except ValueError:
        return f"{prefix}{str(text)}{suffix}"


def rmdir(dir_path: str) -> None:
    """
    Рекурсивно удаляет директорию и все её содержимое.

    Функция проходит по всем файлам и подпапкам в указанной директории,
    удаляет их, а затем удаляет саму директорию.

    Args:
        dir_path (str): Путь к директории, которую необходимо удалить.

    Returns:
        None
    """
    directory = Path(str(dir_path))

    for item in directory.iterdir():
        if item.is_dir():
            rmdir(str(item))
        else:
            item.unlink()
    directory.rmdir()


def get_pk(distance: float, divider=100, decimal=False) -> str:
    """Возвращает строку пикетажа в формате 'км+мм'.

    Args:
        distance (float): Расстояние в метрах.
        divider (int, optional): Делитель расстояния. По умолчанию 100.
        decimal (bool, optional): Флаг отображения десятичной части. По умолчанию False.

    Returns:
        str: Строка пикетажа.
    """

    # Переводим метры в километры и метры
    first = distance // divider
    second = distance % divider

    if decimal:
        # Извлекаем десятичную часть числа
        decimal_part = distance - int(distance)
        decimal = f".{int(decimal_part * 100):02d}"
    else:
        decimal = ""

    # Форматируем строку пикетажа
    if divider < 1000:
        return f"{int(first)}+{int(second):02d}{decimal}"
    return f"{int(first)}+{int(second):03d}{decimal}"


def floor_float(n, precision=0):
    """
    Округляет число вниз с заданной точностью.

    Функция округляет число 'n' вниз (в сторону отрицательной бесконечности)
    с заданной точностью, определяемой количеством десятичных знаков.

    Args:
        n (float): Число, которое нужно округлить.
        precision (int, optional): Количество десятичных знаков после запятой.
            По умолчанию 0, что означает округление до целого числа.

    Returns:
        float: Округленное число с заданной точностью.

    Examples:
        >>> floor_float(3.75)
        3.0
        >>> floor_float(3.75, 1)
        3.7
        >>> floor_float(-1.23, 1)
        -1.3
    """
    return np.true_divide(np.floor(n * 10**precision), 10**precision)


def closest_upper_multiple(n, k):
    """
    Находит ближайшее большее кратное число.

    Функция вычисляет ближайшее число, которое больше или равно n и кратно k.

    Args:
        n (float, int): Исходное число, для которого ищется ближайшее кратное.
        k (float, int): Число, на которое должен делиться результат без остатка.

    Returns:
        float: Ближайшее большее кратное k, которое >= n.

    Raises:
        ValueError: Если k <= 0, так как кратное число должно быть положительным.

    Examples:
        >>> closest_upper_multiple(10, 3)
        12.0
        >>> closest_upper_multiple(7, 2)
        8.0
        >>> closest_upper_multiple(5, 5)
        5.0
    """
    # Проверяем, что k больше нуля
    if k <= 0:
        raise ValueError("Второй аргумент (k) должен быть больше нуля.")

    # Вычисляем ближайшее старшее кратное
    result = np.ceil(n / k) * k
    return result


def split_list_by_min_value(values: list) -> list:
    """
    Разделяет список на две части по минимальному значению.

    Функция находит минимальное значение в списке и использует его индекс
    для разделения списка на две части. Если минимальное значение находится
    в начале или конце списка, индекс разделения корректируется для
    предотвращения создания пустых списков.

    Args:
        values (list): Список для разделения.

    Returns:
        list: Список из двух непустых списков.

    Raises:
        ValueError: Если входной список содержит менее 2 элементов.

    Examples:
        >>> split_list_by_min_value([3, 1, 4, 2])
        [[3], [1, 4, 2]]
        >>> split_list_by_min_value([1, 3, 4, 2])
        [[1], [3, 4, 2]]
        >>> split_list_by_min_value([3, 4, 2, 1])
        [[3, 4, 2], [1]]
    """
    if len(values) < 2:
        raise ValueError(
            "Input list must have at least 2 elements to split without empty lists"
        )

    min_val = min(values)
    min_index = values.index(min_val)

    # Если минимальный элемент находится в начале списка,
    # сдвигаем разделение на следующий индекс
    if min_index == 0:
        min_index = 1
    # Если минимальный элемент находится в конце списка,
    # сдвигаем разделение так, чтобы последний элемент оказался во второй части
    elif min_index == len(values) - 1:
        min_index = len(values) - 1

    return [values[:min_index], values[min_index:]]


def get_water_sections(morfostvor, water_level: float, overflow: bool = False) -> tuple:
    """
    Определяет секторы воды при заданном уровне на основе морфоствора.

    Функция анализирует заданный морфоствор и определяет участки, которые
    будут заполнены водой при указанном уровне воды. Может работать в двух
    режимах: с учетом перелива и без него.

    Args:
        morfostvor (Morfostvor): Объект морфоствора, содержащий информацию о геометрии русла.
        water_level (float): Заданный уровень воды для анализа.
        overflow (bool, optional): Флаг учета перелива между секторами.
            Если True, происходит моделирование заполнения с учетом возможного
            перелива воды между секторами. Если False, заполнение рассматривается
            независимо для каждого участка. По умолчанию False.

    Returns:
        tuple: Кортеж из двух списков:
            - Список объектов WaterSection, представляющих секторы воды.
            - Список секторов морфоствора, соответствующих водным секторам.

    Note:
        При overflow=True расчет начинается с сектора, имеющего минимальную отметку,
        и рекурсивно распространяется на соседние секторы, если уровень воды
        превышает максимальные отметки на границах секторов.
    """
    from hydraulic.models import WaterSection

    # Участок с минимальной отметкой дна
    min_sector = morfostvor.get_min_sector()

    result_sections = []
    result_sectors = []

    if overflow:
        # Исходные сектора для расчёта (сектор, содержащий минимальную отметку)
        calc_sectors = [min_sector[0]]

        # Копируем список, чтобы избежать модификации во время итерации
        sectors_to_process = calc_sectors.copy()

        while sectors_to_process:
            i = sectors_to_process.pop(0)

            if i < 0 or i >= len(morfostvor.sectors):
                continue

            sector = morfostvor.sectors[i]
            x = sector.coord[0]
            y = sector.coord[1]

            try:
                # Максимальная отметка слева и справа
                previous_min_ele = max(split_list_by_min_value(y)[0])
                next_min_ele = max(split_list_by_min_value(y)[1])

                # Проверка на перелив левой границы участка
                if (
                    water_level >= previous_min_ele
                    and (i - 1) not in calc_sectors
                    and (i - 1) >= 0
                ):
                    # Проверка что вода дошла до левой границы участка
                    # костыль через try, чтобы избежать ошибки определения границы
                    try:
                        water = WaterSection(
                            x,
                            y,
                            water_level,
                            start_point=morfostvor.x[sector.end_point],
                        )
                    except ValueError:
                        water = WaterSection(
                            x,
                            y,
                            water_level,
                            start_point=morfostvor.x[sector.end_point] - 1,
                        )
                    if water.water_section_x[0] == x[0]:
                        calc_sectors.append(i - 1)
                        sectors_to_process.append(i - 1)
                # Проверка на перелив правой границы участка
                if (
                    water_level >= next_min_ele
                    and (i + 1) not in calc_sectors
                    and (i + 1) < len(morfostvor.sectors)
                ):
                    calc_sectors.append(i + 1)
                    sectors_to_process.append(i + 1)
            except (ValueError, IndexError) as e:
                # Обработка ошибки, если список слишком короткий для разделения
                print(f"Ошибка при обработке сектора {i}: {e}")
                continue

            # Создаем водный сектор в зависимости от его положения
            water = None

            # Расчетный участок является участком с минимальными отметками
            if sector.id == min_sector[1].id:
                min_y_index = sector.coord[1].index(min(sector.coord[1]))
                water = WaterSection(
                    x, y, water_level, start_point=sector.coord[0][min_y_index]
                )
            # Расчетный участок находится слева от начального
            elif sector.id < min_sector[1].id:
                water = WaterSection(
                    x, y, water_level, start_point=morfostvor.x[sector.end_point]
                )
            # Расчетный участок находится справа от начального
            elif sector.id > min_sector[1].id:
                water = WaterSection(
                    x, y, water_level, start_point=morfostvor.x[sector.start_point]
                )

            if water is not None:
                result_sections.append(water)
                result_sectors.append(sector)
    else:
        # Отрисовка с заполнением по участкам
        for sector in morfostvor.sectors:
            x = sector.coord[0]
            y = sector.coord[1]

            if min(y) < water_level:
                # Сектор воды и основные его параметры
                water = WaterSection(x, y, water_level)
                result_sections.append(water)
                result_sectors.append(sector)

    return result_sections, result_sectors


def calculate_line_length(x_coords: list[float], y_coords: list[float]) -> float:
    """
    Вычисляет длину ломаной линии по координатам точек.

    Функция принимает два списка с координатами точек X и Y,
    и рассчитывает суммарную длину ломаной линии, соединяющей эти точки.

    Args:
        x_coords (list): Список координат X точек ломаной линии.
        y_coords (list): Список координат Y точек ломаной линии.

    Returns:
        float: Длина ломаной линии. Если списки имеют разную длину или
               содержат менее 2 точек, возвращает 0.0.

    Examples:
        >>> calculate_line_length([0, 3, 5], [0, 4, 7])
        8.0
        >>> calculate_line_length([1, 1], [1, 2])
        1.0
        >>> calculate_line_length([1], [1])
        0.0
    """
    # Проверяем, что списки координат имеют одинаковую длину и минимум 2 точки
    if len(x_coords) != len(y_coords) or len(x_coords) < 2:
        return 0.0

    total_length = 0.0
    # Проходим по всем точкам, кроме последней
    for i in range(len(x_coords) - 1):
        # Вычисляем расстояние между текущей и следующей точкой
        dx = x_coords[i + 1] - x_coords[i]
        dy = y_coords[i + 1] - y_coords[i]
        total_length += (dx**2 + dy**2) ** 0.5

    return total_length
