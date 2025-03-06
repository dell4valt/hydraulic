from dataclasses import dataclass, field
from typing import ClassVar, Dict, List

import numpy as np
from scipy import interpolate

from hydraulic import config
from hydraulic.lib import poly_area


@dataclass
class ProfileSector:
    """Класс участка профиля (пойма, русло и т.д.).

    :param id: Номер участка
    :param name: Описание (название) участка
    :param start_point: Номер первой точки участка
    :param end_point: Номер последней точки участка
    :param roughness: Коэффициент шероховатости n
    :param slope: Уклон данного участка I, ‰
    :param coord: Кортеж с двумя списками координат (x и y) участка
    """

    id: int
    name: str
    start_point: int
    end_point: int
    roughness: float
    slope: float
    coord: tuple[list[float], list[float]]

    consumption: float = field(default=np.nan)
    depth: float = field(default=np.nan)
    speed: float = field(default=np.nan)
    area: float = field(default=np.nan)
    width: float = field(default=np.nan)
    color: list[float] = field(init=False)

    def __post_init__(self):
        self.color = self.get_color()
        self.__validate_types()

    def get_color(self) -> list[float]:
        """Определяет цвет участка в зависимости от его типа."""
        name_lower = self.name.lower()

        if "русло" in name_lower:
            return [0, 0.5, 1]
        elif "протока" in name_lower:
            return [0, np.random.uniform(0, 0.5), np.random.uniform(0.5, 1)]
        elif "пойма" in name_lower:
            return [np.random.uniform(0.3, 1), 0, 0]

        return np.random.uniform(0, 1, 3).tolist()

    @property
    def length(self) -> float:
        return round(self.coord[0][-1] - self.coord[0][0], 3)

    def __validate_types(self):
        """Проверяет соответствие типов атрибутов."""
        expected_types = {
            "id": int,
            "name": str,
            "start_point": int,
            "end_point": int,
            "roughness": float,
            "slope": float,
            "coord": tuple,
        }

        for field_name, expected_type in expected_types.items():
            value = getattr(self, field_name)

            if not isinstance(value, expected_type):
                type_map = {
                    int: "целое число",
                    float: "десятичное число",
                    str: "строка",
                    list: "список",
                    tuple: "кортеж"
                }

                readable_name = {
                    "slope": "уклон",
                    "roughness": "коэффициент шероховатости",
                }.get(field_name, field_name)

                raise TypeError(
                    f"Ошибка типа данных: {readable_name} должен быть '{type_map[expected_type]}', "
                    f"но получено '{type_map[type(value)]}' ({value})."
                )


@dataclass
class SituationSector:
    id: int
    type: str
    start_point: int
    end_point: int

    COLOR_MAPPING: ClassVar[Dict[str, str]] = {
        'grass': 'honeydew',
        'concrete': 'gainsboro',
        'field': 'burlywood',
        'wood': 'limegreen',
        'water': 'deepskyblue',
        'sand': 'lemonchiffon',
        'gravel': 'tan',
        'reed': 'cadetblue',
        'bush': 'darkkhaki',
    }

    CATEGORIES: ClassVar[Dict[str, List[str]]] = {
        'grass': ['трава', 'луг', 'газон'],
        'concrete': ['бетон', 'асфальт'],
        'field': ['пашня', 'поле'],
        'reed': ['камыш', 'кам', 'кам.', 'осока'],
        'wood': ['лес', 'редкий лес', 'поросль'],
        'bush': ['кустарник', 'кусты'],
        'water': ['вода', 'ув', 'протока', 'ручей'],
        'sand': ['песок'],
        'gravel': ['гравий', 'галька', 'аллювий'],
    }

    def get_color(self) -> str:
        """Возвращает цвет сектора на основе его типа.

        Returns:
            str: Название цвета в CSS-формате или 'white' если тип не распознан
        """
        normalized_type = self._normalize_type(self.type)

        for category, keywords in self.CATEGORIES.items():
            if normalized_type in keywords:
                return self.COLOR_MAPPING[category]
        return 'white'

    def _normalize_type(self, type_str: str) -> str:
        """Приводит строку типа к стандартному виду для сравнения"""
        return type_str.strip().lower()


@dataclass
class SituationBorder:
    id: int
    type: str
    point: int


@dataclass
class WaterSection:
    """Класс водного сечения

    :param x: Точки x всего профиля
    :param y: Точки y всего профиля
    :param water_level: Уровень воды
    :param water_section_x: Точки x водного сечения
    :param water_section_y: Точки y водного сечения
    :param width: Ширина водного сечения
    :param area: Площадь водного сечения
    :param average_depth: Средняя глубина
    :param max_depth: Максимальная глубина
    :param wet_perimeter: Смоченный периметр
    :param r_hydraulic: Гидравлический радиус
    :param start_point: Точка начала расчёта [point_index, y]

    """

    x: float
    y: float
    water_level: float
    water_section_x: list = field(default_factory=list)
    water_section_y: list = field(default_factory=list)
    width: float = 0.0
    area: float = 0.0
    average_depth: float = 0.0
    max_depth: float = 0.0
    wet_perimeter: float = 0.0
    r_hydraulic: float = 0.0
    start_point: list = field(default_factory=list)

    def __post_init__(self):
        # start_point=[self.y.index(min(self.y)), min(self.y)]
        boundary = self.boundary()
        if len(boundary) > 1:
            for water_boundary in boundary:
                try:
                    self._calculate_parameters(water_boundary)
                except IndexError:
                    print(
                        "Ошибка в определении границ урезов! Программа будет завершена."
                    )
                    sys.exit(2)

            # Вычисления если урезов несколько
            self.width = sum(self.width)
            self.area = sum(self.area)
            self.average_depth = np.average(self.average_depth)
            self.max_depth = max(self.max_depth)
            self.wet_perimeter = sum(self.wet_perimeter)
            self.r_hydraulic = sum(self.r_hydraulic)

        else:
            try:
                self._calculate_parameters(boundary[0])
            except IndexError:
                print("Ошибка в определении границ урезов! Программа будет завершена.")
                sys.exit(2)

    def boundary(self):
        x = self.x
        y = self.y
        water_level = self.water_level  # Отметка уреза воды
        water_boundary_x, water_boundary_y, water_boundary_points = [], [], []
        result = []
        start_point = self.start_point

        if not start_point:
            start_point = [y.index(min(y)), min(y)]

        # Проверка на ошибку расположения уреза под поверхностью дна
        if water_level < min(y):
            print(
                "Ошибка! Уровень воды ниже низшей точки дна. Программа будет завершена с ошибкой."
            )
            sys.exit(1)
        else:
            # Цикл влево от стартовой точки
            for i in range(start_point[0], -1, -1):
                # Если индекс минимальной отметки совпадает с левой правой участка
                if start_point[0] == 0 and y[start_point[0]] <= water_level:
                    water_boundary_x.append(x[0])
                    water_boundary_y.append(water_level)
                    water_boundary_points.append(0)
                    break

                # Условие пересечения уреза с дном
                if y[i - 1] >= water_level and y[i] <= water_level:
                    x1, x2 = x[i - 1], x[i]
                    y1, y2 = y[i - 1], y[i]

                    # Нахождение координаты x уреза между точками дна
                    f = interpolate.interp1d([y1, y2], [x1, x2])
                    # Находим координату x, зная y (точка пересечения уреза с дном)
                    water_boundary_x.append(float(f(water_level)))
                    water_boundary_y.append(water_level)
                    # Присоединяем номер точки дна с границей воды
                    water_boundary_points.append(i - 1)
                    break  # Прерываем поиск если нашли пересечение

                # Условие отсутствия пересечения с дном и дохождения до начала участка
                elif i - 1 == 0 and y[i - 1] <= water_level:
                    water_boundary_x.append(x[i - 1])
                    water_boundary_y.append(water_level)
                    water_boundary_points.append(i - 1)
                    break  # Прерываем поиск если нашли пересечение

            # Цикл вправо от стартовой точки
            for i in range(start_point[0], len(y) - 1):
                # Условие пересечения уреза с дном
                if y[i] <= water_level and y[i + 1] >= water_level:
                    x1, x2 = x[i], x[i + 1]
                    y1, y2 = y[i], y[i + 1]

                    # Нахождение координаты x уреза между точками дна
                    f = interpolate.interp1d([y1, y2], [x1, x2])
                    # Находим координату x, зная y (точка пересечения уреза с дном)
                    water_boundary_x.append(float(f(water_level)))
                    water_boundary_y.append(water_level)
                    # Присоединяем номер точки дна с границей воды
                    water_boundary_points.append(i)
                    break  # Прерываем поиск если нашли пересечение

                elif i + 1 == len(y) - 1 and y[len(y) - 1] <= water_level:
                    water_boundary_x.append(x[len(x) - 1])
                    water_boundary_y.append(water_level)
                    water_boundary_points.append(i + 1)
                    break  # Прерываем поиск если нашли пересечение

            # Если индекс минимальной отметки совпадает с правой границей участка
            if start_point[0] == len(y) - 1 and y[start_point[0]] <= water_level:
                water_boundary_x.append(x[len(y) - 1])
                water_boundary_y.append(water_level)
                water_boundary_points.append(len(y) - 1)

            result.append(
                [water_boundary_x, water_boundary_y, water_boundary_points, 0]
            )
        return result

    # Функция выполняющая основные вычисления по данному водному сечению
    def _calculate_parameters(self, water_boundary):
        sum_sqr = 0
        water_level = self.water_level
        x = self.x
        y = self.y
        depth = []

        # Обрабатываем урезы по две точки (со второй до третьей пропускам)
        # Вводим служебные координаты (первая и последняя точки)
        x1, x2 = water_boundary[0][0], water_boundary[0][1]
        y1, y2 = water_boundary[1][0], water_boundary[1][1]

        # Точки смоченного периметра (номера точек под урезом)
        water_section_x = x[water_boundary[2][0] + 1: water_boundary[2][1] + 1]
        water_section_y = y[water_boundary[2][0] + 1: water_boundary[2][1] + 1]

        water_section_x.insert(0, x1)
        water_section_x.insert(len(water_section_x), x2)

        water_section_y.insert(0, y1)
        water_section_y.insert(len(water_section_y), y2)

        # Если первая точка УВ выше первой точки дна, вставляем точку дна на второе место
        # TODO: Костыль для определения полигона водной поверхности для расчёта с переливом
        #  и одновременным заполнением, нужно продумать как исправить
        if config.OVERFLOW:  # исходные данные точек x и y по всему профилю
            if water_level > y[water_boundary[2][0]]:
                water_section_x.insert(1, x[0])
                water_section_y.insert(1, y[0])
        else:  # исходные данные точек x и y по участкам
            if water_level > y[0]:
                water_section_x.insert(1, x[0])
                water_section_y.insert(1, y[0])

        # Если последняя точка УВ выше последней точки дна, вставляем точку на предпоследнее место
        if water_boundary[3] > 1 and water_level > y[-1]:
            water_section_x.insert(len(water_section_x) - 1, x[-1])
            water_section_y.insert(len(water_section_y) - 1, y[-1])

        # Координаты x и y смоченного периметра
        self.water_section_x = water_section_x
        self.water_section_y = water_section_y

        # Определяем ширину водной поверхности
        self.width = x2 - x1

        # Площадь воды
        self.area = poly_area(water_section_x, water_section_y)

        # Глубины
        for i in range(len(water_section_y)):
            depth.append(water_level - water_section_y[i])

        # Средняя глубина
        if self.area > 0 and self.width > 0:
            self.average_depth = self.area / self.width
        else:
            self.average_depth = 0

        if self.average_depth == 0:  # Костыль
            self.average_depth = 0.00001

        # Максимальная глубина
        self.max_depth = max(depth)

        # Смоченный периметр
        for i in range(len(water_section_x) - 1):
            sum_sqr += (water_section_x[i + 1] - water_section_x[i]) ** 2
        self.w_perimeter = np.sqrt(sum_sqr)

        # Гидравлический радиус
        if self.area > 0 and self.w_perimeter > 0:
            self.r_hydraulic = self.area / self.w_perimeter
        else:
            self.r_hydraulic = 0

        if self.r_hydraulic == 0:  # Костыль
            self.r_hydraulic = 0.00001

