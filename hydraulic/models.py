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
    """
    Класс водного сечения

    :param x: Список x-координат всего профиля
    :param y: Список y-координат всего профиля
    :param water_level: Уровень воды
    :param water_section_x: Список x-координат водного сечения (будет заполнен)
    :param water_section_y: Список y-координат водного сечения (будет заполнен)
    :param width: Ширина водного сечения
    :param area: Площадь водного сечения
    :param average_depth: Средняя глубина
    :param max_depth: Максимальная глубина
    :param wet_perimeter: Смочённый периметр
    :param r_hydraulic: Гидравлический радиус
    :param start_point: Точка начала расчёта [index, y] (необязательный параметр)
    """
    profile_x_coords: List[float] = field(default_factory=list)
    profile_y_coords: List[float] = field(default_factory=list)
    water_level: float = 0.0
    overflow: bool = config.OVERFLOW
    start_point: float = np.nan

    water_section_x: list = field(default_factory=list, init=False)
    water_section_y: list = field(default_factory=list, init=False)
    width: float = field(default=0.0, init=False)
    area: float = field(default=0.0, init=False)
    average_depth: float = field(default=0.0, init=False)
    max_depth: float = field(default=0.0, init=False)
    wet_perimeter: float = field(default=0.0, init=False)
    r_hydraulic: float = field(default=0.0, init=False)

    def __post_init__(self):
        # Определяем все водные сечения
        self.segments = self.get_water_sections()
        if not self.segments:
            raise ValueError("Ошибка! Не удалось определить сечения.")

        # Если задан start_point, выбираем только сегмент, содержащий его
        if self.start_point is not np.nan:
            self.segments = [seg for seg in self.segments if self.start_point in seg[0]]

            if not self.segments:
                raise ValueError("Ошибка! Заданная стартовая точка не попадает ни в один сегмент.")

        # Списки для хранения параметров по каждому сечению
        widths = []
        areas = []
        avg_depths = []
        max_depths = []
        perimeters = []
        r_hydraulics = []
        combined_ws_x = []
        combined_ws_y = []

        # Вычисляем параметры для каждого сечения
        for seg in self.segments:
            params, ws_x, ws_y = self._calculate_parameters(seg)
            widths.append(params['width'])
            areas.append(params['area'])
            avg_depths.append(params['average_depth'])
            max_depths.append(params['max_depth'])
            perimeters.append(params['wet_perimeter'])
            r_hydraulics.append(params['r_hydraulic'])
            combined_ws_x.extend(ws_x)
            combined_ws_y.extend(ws_y)

        # Комбинируем результаты по всем сечениям
        self.width = sum(widths)
        self.area = sum(areas)
        self.average_depth = np.average(avg_depths) if avg_depths else 0
        self.max_depth = max(max_depths) if max_depths else 0
        self.wet_perimeter = sum(perimeters)
        self.r_hydraulic = sum(r_hydraulics)
        self.water_section_x = combined_ws_x
        self.water_section_y = combined_ws_y

    def get_water_sections(self):
        """Определяет все непрерывные водные сечения в профиле.
        Алгоритм:
          1. Если water_level ниже минимального y – завершаем работу.
          2. Проходим по всем точкам профиля с фиксацией состояния "в сечении"/"не в сечении":
             - При переходе из надводного (y > water_level) в подводное (y <= water_level)
               интерполируем точку пересечения – это левая граница сечения.
             - При переходе из подводного в надводное интерполируем правую границу,
               собираем все точки сечения и сохраняем найденное сечение.
          3. Если профиль заканчивается под водой, завершаем последнее сечение.

        Returns:
          list: Список сечений, где каждое сечение представлено списком:
                [seg_x, seg_y, seg_indices, 0]
                    - seg_x: список x-координат (включая интерполированные границы),
                    - seg_y: список y-координат (границы – water_level, внутренние – реальные),
                    - seg_indices: индексы исходных точек,
                    - 0: служебное значение.
        """
        x = self.profile_x_coords
        y = self.profile_y_coords
        water_level = self.water_level

        if water_level < min(y):
            raise ValueError("Ошибка! Уровень воды ниже низшей точки дна.")

        segments = []
        in_segment = False
        segment_start_index = None
        segment_start_x = None

        n = len(y)

        # Если первая точка уже под водой, начинаем сечение с начала
        if y[0] <= water_level:
            in_segment = True
            segment_start_index = 0
            segment_start_x = x[0]

        for i in range(n - 1):
            # Переход из надводного в подводное – начало сечения
            if not in_segment and y[i] > water_level and y[i+1] <= water_level:
                f = interpolate.interp1d([y[i], y[i+1]], [x[i], x[i+1]])
                segment_start_x = float(f(water_level))
                segment_start_index = i
                in_segment = True

            # Переход из подводного в надводное – конец сечения
            elif in_segment and y[i] <= water_level and y[i+1] > water_level:
                f = interpolate.interp1d([y[i], y[i+1]], [x[i], x[i+1]])
                segment_end_x = float(f(water_level))
                segment_end_index = i + 1

                # Собираем точки сечения: начинаем с левой границы
                seg_x = [segment_start_x]
                seg_y = [water_level]
                seg_indices = [segment_start_index]
                # Добавляем все точки между началом и концом, которые находятся под или на water_level
                for j in range(segment_start_index + 1, i + 1):
                    if y[j] <= water_level:
                        seg_x.append(x[j])
                        seg_y.append(y[j])
                        seg_indices.append(j)
                # Добавляем правую границу
                seg_x.append(segment_end_x)
                seg_y.append(water_level)
                seg_indices.append(segment_end_index)

                segments.append([seg_x, seg_y, seg_indices])
                in_segment = False

        # Если профиль заканчивается под водой, завершаем последнее сечение
        if in_segment:
            if y[-1] <= water_level:
                segment_end_x = x[-1]
                segment_end_index = n - 1
            else:
                f = interpolate.interp1d([y[-2], y[-1]], [x[-2], x[-1]])
                segment_end_x = float(f(water_level))
                segment_end_index = n - 2

            seg_x = [segment_start_x]
            seg_y = [water_level]
            seg_indices = [segment_start_index]
            for j in range(segment_start_index + 1, n):
                if y[j] <= water_level:
                    seg_x.append(x[j])
                    seg_y.append(y[j])
                    seg_indices.append(j)
            seg_x.append(segment_end_x)
            seg_y.append(water_level)
            seg_indices.append(segment_end_index)
            segments.append([seg_x, seg_y, seg_indices])
        return segments

    def _calculate_parameters(self, water_boundary):
        """
        Вычисляет параметры водного сечения для одного сегмента.

        Args:
            water_boundary: список вида [seg_x, seg_y, seg_indices]

        Returns:
            list: список вида [width, area, average_depth, max_depth, wet_perimeter, r_hydraulic]
            list: список x-координат водного сечения (с границами)
            list: список y-координат водного сечения (с границами)
        """
        seg_x, seg_y, seg_indices = water_boundary
        water_level = self.water_level

        # Функция для линейной интерполяции x по дну, зная y
        def interpolate_x(y_target, idx1, idx2):
            """ Интерполирует x-координату для заданного уровня y по дну. """
            x1, x2 = self.profile_x_coords[idx1], self.profile_x_coords[idx2]
            y1, y2 = self.profile_y_coords[idx1], self.profile_y_coords[idx2]
            if y1 == y2:
                return x1  # На случай горизонтального участка дна
            f = interpolate.interp1d([y1, y2], [x1, x2], fill_value="extrapolate")
            return float(f(y_target))

        # Проверяем что левая граница находится на уровне воды
        # и при необходимости добавляем точку чтобы избежать
        # срезания углов левой границы сегмента
        if seg_y[0] == water_level:
            first_index = seg_indices[0] if seg_indices[0] == 0 else seg_indices[0] - 1
            if self.profile_y_coords[first_index] < water_level:  # Дно ниже уровня воды
                x_interp = interpolate_x(self.profile_y_coords[first_index], first_index, first_index + 1)

                # Вставляем точки чтобы избежать срезания углов
                seg_x.insert(1, x_interp)  
                seg_y.insert(1, self.profile_y_coords[first_index])
                seg_indices.insert(1, first_index)

        # Проверяем правую границу сегмента
        if seg_y[-1] == water_level:
            last_index = seg_indices[-2] if seg_indices[-1] == self.profile_x_coords else seg_indices[-1]
            if self.profile_y_coords[last_index] < water_level:  # Дно ниже уровня воды
                x_interp = interpolate_x(self.profile_y_coords[last_index], last_index - 1, last_index)
                # Проверяем не ровное ли дно на последних точках
                if self.profile_y_coords[last_index] == self.profile_y_coords[last_index - 1]:
                    x_interp = self.profile_x_coords[last_index]
                # Вставляем точки чтобы избежать срезания углов
                seg_x.insert(-1, x_interp)  
                seg_y.insert(-1, self.profile_y_coords[last_index])
                seg_indices.insert(-1, last_index)

        # Вычисление ширины сечения как разность между правой и левой границей
        width = seg_x[-1] - seg_x[0]
        # Площадь сечения
        area = poly_area(seg_x, seg_y)
        # Вычисляем глубины в каждой точке сечения (разница между water_level и y)
        depths = [water_level - val for val in seg_y]
        average_depth = area / width if area > 0 and width > 0 else 0
        if average_depth == 0:
            average_depth = 0.00001
        max_depth = max(depths) if depths else 0

        # Вычисляем смочённый периметр как сумму расстояний между соседними точками
        sum_sqr = 0
        for i in range(len(seg_x) - 1):
            sum_sqr += (seg_x[i+1] - seg_x[i]) ** 2
        wet_perimeter = np.sqrt(sum_sqr)
        r_hydraulic = area / wet_perimeter if area > 0 and wet_perimeter > 0 else 0
        if r_hydraulic == 0:
            r_hydraulic = 0.00001

        params = {
            'width': width,
            'area': area,
            'average_depth': average_depth,
            'max_depth': max_depth,
            'wet_perimeter': wet_perimeter,
            'r_hydraulic': r_hydraulic
        }
        return params, seg_x, seg_y
