import os
import sys
import time
import warnings
from dataclasses import dataclass, field
from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from scipy import interpolate
from openpyxl import load_workbook
from report.utils import get_xls_sheet_quantity

from hydraulic import config
from hydraulic.graph import (GraphFH, GraphProfile, GraphQF, GraphQH, GraphQHV,
                             GraphQV, GraphQWVH, GraphVH)
from hydraulic.lib import (chunk_list, insert_summary_QV_tables,
                           question_continue_app)
from hydraulic.models import (ProfileSector, SituationBorder, SituationSector,
                              WaterSection)
from hydraulic.profile_report import generate_morfostvor_report, save_graphic

# Отключаем все UserWarning (предупреждения от labellines)
warnings.filterwarnings("ignore", category=UserWarning)
# Закрывает все открытые графики
plt.close('all')


@dataclass
class Calculation:
    """
    Класс гидравлических расчётов скорости, расхода воды и коэффициента Шези для водного объекта.

    :param n: Коэффициент шероховатости
    :param i: Уклон, промилле
    :param h: Средняя глубина водного сечения
    :param a: Площадь водного сечения

    """

    n: float  # Коэффициент шероховатости
    i: float  # Уклон
    h: float  # Средняя глубина
    a: float  # Площадь водного сечений
    v: float = 0  # Скорость
    q: float = 0  # Расход
    _g: float = 9.80665  # Ускорение свободного падения
    shezi: float = 0  # Коэффициент Шези
    type__: str = "Не определен"

    def __post_init__(self):
        # В зависимости от глубины считаем по разным формулам
        # до 3-х метров по Павловскому, свыше 3-х метров по
        # Павловскому-Железнякову
        if self.h >= 0 and self.h <= 3:
            self.__shezi_pavlovskij()
        else:
            self.__shezi_pavlovskij_zheleznjakov()

        # Тип расчёта, обычная вода или селевой поток
        if config.CALC_TYPE == 1:
            # Расчёт скорости воды
            self.v = self.shezi * np.sqrt(self.h * (self.i / 1000))
        elif config.CALC_TYPE == 2:
            # Расчёт скорости воды для наносоводных селей
            self.v = 4.5 * self.h ** 0.67 * (self.i / 1000) ** 0.17
        elif config.CALC_TYPE == 3:
            # Расчёт скорости воды для грязекаменных селей селей
            self.v = 3.75 * self.h ** 0.50 * (self.i / 1000) ** 0.17
        else:
            print(
                "Ошибка выбора формулы расчёта скорости потока. Программа будет завершена."
            )
            sys.exit(1)
        # Расчёт расхода воды
        self.q = self.a * self.v

    # Коэффициент Шези по формуле Н. Н. Павловского, степенной коэффициент по формуле Железнякова
    def __shezi_pavlovskij_zheleznjakov(self):
        # Показатель степени по формуле Г. В. Железнякова
        y = (
            1
            / np.log10(self.h)
            * np.log10(
                (1 / 2 - (self.n * np.sqrt(self._g) / 0.26) * (1 - np.log10(self.h)))
                + self.n
                * np.sqrt(
                    1
                    / 4
                    * (1 / self.n - np.sqrt(self._g) / 0.13 * (1 - np.log10(self.h)))
                    ** 2
                    + np.sqrt(self._g)
                    / 0.13
                    * (1 / self.n + np.sqrt(self._g) * np.log10(self.h))
                )
            )
        )

        self.shezi = (1 / self.n) * self.h ** y
        self.type__ = "Коэффициент Шези определён по формуле Павловского, \
                       показатель степени определён по формуле Железнякова"

    # Коэффициент шези по формуле Маннинга
    def __shezi_manning(self):
        self.shezi = (1 / self.n) * self.h ** (1 / 6)
        self.type__ = "Коэффициент Шези определён по формуле Маннинга"

    # Коэффициент Шези по формуле Павловского
    # для глубин 0.1 < h < 3 (Гидрорасчёты считают по этой формуле)
    def __shezi_pavlovskij(self):
        y = (
            2.5 * np.sqrt(self.n)
            - 0.13
            - 0.75 * np.sqrt(self.h) * (np.sqrt(self.n) - 0.10)
        )
        self.shezi = (1 / self.n) * self.h ** y
        self.type__ = (
            "Коэффициент шези определён по формуле Павловского для глубин 0.1 < h < 3 м"
        )

    # Коэффициент шези по формуле Железнякова
    def __shezi_zheleznjakov(self):
        self.shezi = 1 / 2 * (
            (1 / self.n) - (np.sqrt(self._g) / 0.13) * (1 - np.log10(self.h))
        ) + np.sqrt(
            (1 / 4)
            * (1 / self.n - (np.sqrt(self._g) / 0.13) * (1 - np.log10(self.h))) ** 2
            + (np.sqrt(self._g) / 0.13)
            * ((1 / self.n) + (np.sqrt(self._g) * np.log10(self.h)))
        )
        self.type__ = "Коэффициент шези определён по формуле Железнякова"


@dataclass
class Morfostvor:

    """Класс описывающий морфоствор."""

    # Основные параметры морфоствора
    title: str = ""
    x: list = field(default_factory=list)
    y: list = field(default_factory=list)
    situation: list = field(default_factory=list)
    situation_borders: list = field(default_factory=list)
    sectors: list = field(default_factory=list)
    ele_max: float = 0
    ele_min: float = 0
    date: str = ""
    dH: int = 5
    waterline: float = 0
    erosion_limit: float = 0
    erosion_limit_coord: list = field(default_factory=list)
    top_limit: float = 0
    top_limit_description: str = ""

    probability: list = field(default_factory=list)
    design_water_level_index: int = 0
    coords: list = field(default_factory=list)
    strings: dict = field(default_factory=dict)

    levels_result: pd.DataFrame = pd.DataFrame
    hydraulic_result: pd.DataFrame = pd.DataFrame
    sectors_result: pd.DataFrame = pd.DataFrame
    hydraulic_table: pd.DataFrame = pd.DataFrame

    def __post_init__(self):
        # Выбор варианта расчёта
        if config.CALC_TYPE == 1:
            self.strings["type"] = "воды"
        elif config.CALC_TYPE == 2:
            self.strings["type"] = "наносоводного селевого потока"
        elif config.CALC_TYPE == 3:
            self.strings["type"] = "грязекаменного селевого потока"
        else:
            print(
                "Неверно выбран тип расчёта в конфигурационном файле. Программа будет завершена."
            )
            sys.exit(0)

        self.qh_title = f"Кривая расхода {self.strings['type']} Q = f(H)"

    def read_xls(self, file_path, page=0):
        """Функция чтения из xls файла."""
        try:
            data_file = load_workbook(file_path, data_only=True)  # Открываем xls файл
        except FileNotFoundError:
            print(f"Ошибка! Файл {file_path} не найден. Программа будет завершена.")
            sys.exit(33)

        try:
            # Открываем лист по заданному номеру
            sheet = data_file.worksheets[page]
        except IndexError:
            print(
                "Неверно указан индекс листа .xls файла. Проверьте параметры запуска расчёта."
            )
            sys.exit(34)

        print(
            f"\n----- Считываем исходные данные из .xls файла: "
            f"{file_path}, страница {page} ({sheet.title}) -----\n"
        )

        __raw_data = []  # Сырые строки xls файла
        i = 0

        # Позиционирование столбцов с данными в .xls файле
        __x_coord_col = 0
        __y_coord_col = 1
        __sector_name_col = 2
        __roughness_col = 3
        __slope_col = 4
        __situation_col = 5
        __description_col = 8

        def get_situation(self):
            """Функция считывания участков ситуации из исходных файлов."""

            print("    — Определяем участки ситуации ... ", end="")

            lines_num = 0

            # Считываем количество строк с не пустыми координатами
            for line in __raw_data:
                if not isinstance(line[__x_coord_col], str):
                    lines_num += 1

            situation = self.situation
            situation_borders = self.situation_borders
            x = self.x  # Координаты профиля X
            num = 1  # Порядковый участка ситуации
            bnum = 1  # Порядковый номер границы

            for line in range(lines_num):
                try:
                    s1 = __raw_data[line][__situation_col].split(",")[0]
                    s2 = __raw_data[line][__situation_col].split(",")[1]

                    situation_borders.append(
                        SituationBorder(bnum, s2.strip().lower(), line)
                    )
                    bnum += 1
                except IndexError:
                    s1 = __raw_data[line][__situation_col]

                if line == 0:
                    situation.append(
                        SituationSector(num, s1, line, line)
                    )

                elif s1 != situation[num - 1].type:
                    if situation[num - 1].id == 1:
                        situation[num - 1].end_point = line
                    else:
                        situation[num - 1].end_point = line

                    num += 1

                    situation.append(
                        SituationSector(
                            num,
                            s1,
                            situation[num - 2].end_point,
                            line
                        )
                    )
            situation[-1].end_point = len(x) - 1

            print("успешно.\n")
            return situation

        def get_sectors(self):
            """Функция считывания участков и их параметров из исходных файлов."""

            print("    — Определяем морфометрические участки ... ", end="")
            # №, Описание участка, номер первой точки, номер последней точки,
            # коэффициент шероховатости, уклон ‰, координата x, координаты y
            lines_num = 0

            # Считываем количество строк с не пустыми координатами
            for line in __raw_data:
                if not isinstance(line[__x_coord_col], str):
                    lines_num += 1

            sectors = self.sectors  # Список участков
            x = self.x  # Координаты профиля X
            y = self.y  # Координаты профиля Y

            num = 1  # Номера участков

            ###
            # Перебираем все строки xls файла и ищем участки
            for line in range(lines_num):
                name = __raw_data[line][__sector_name_col].strip()  # Название участка
                # Коэффициент шероховатости
                roughness = __raw_data[line][__roughness_col]
                # Костыль обхода типа данных, нужен float
                try:
                    slope = float(__raw_data[line][__slope_col])  # Уклон
                except ValueError:
                    slope = __raw_data[line][__slope_col]

                # По первой строке создаём первый сектор
                if line == 0:
                    coord = ()
                    sectors.append(
                        ProfileSector(num, name, line, line, roughness, slope, coord)
                    )

                # Сравниваем имя предыдущего участка с текущим,
                # если не совпадают то создаем новый сектор:
                elif name.lower() != sectors[num - 1].name.lower():

                    # TODO: Проверить это условие
                    if sectors[num - 1].id == 1:  # Если первый участок
                        # Записываем номер последний точки - 1
                        sectors[num - 1].end_point = line
                    else:  # Если все остальные участки
                        # Записываем номер последний точки
                        # в предыдущий участок для всех остальных участков
                        sectors[num - 1].end_point = line

                    num += 1  # Увеличиваем номер сектора на 1
                    sectors.append(
                        ProfileSector(
                            num,
                            name,
                            sectors[num - 2].end_point,
                            line,
                            roughness,
                            slope,
                            coord
                        )
                    )

            # Проверка участков
            for sector in sectors:
                if sector.roughness == '':
                    print()
                    print('-----------------------------------------------------------')
                    print(f"Ошибка! В участке №{sector.id} «{sector.name}» "
                          "не задан коэффициент шероховатости n.")
                    print('Программа будет завершена.\n')
                    sys.exit()
                elif sector.slope == '':
                    print()
                    print('-----------------------------------------------------------')
                    print(f'Ошибка! В участке №{sector.id} «{sector.name}» не задан уклон i.')
                    print('Программа будет завершена.\n')
                    sys.exit()

                if sector.roughness < 0.02 or sector.roughness > 0.2:
                    print()
                    print('-----------------------------------------------------------')
                    print(f'Обнаружен подозрительный коэффициент шероховатости\
                            на участке №{sector.id} «{sector.name}» — {sector.roughness}.')
                    question_continue_app()

                if sector.slope <= 0 or sector.slope > 900:
                    print()
                    print('-----------------------------------------------------------')
                    print("Обнаружен подозрительный уклон "
                          f"на участке №{sector.id} «{sector.name}» — {sector.slope}‰.")
                    question_continue_app()

            # Номер последней точки в последнем секторе
            sectors[-1].end_point = len(x) - 1

            # Записываем координаты и длины участков
            for sector in sectors:
                sector.coord = (
                    x[sector.start_point: sector.end_point + 1],
                    y[sector.start_point: sector.end_point + 1],
                )  # Координаты из начальной и конечной точек

            try:
                # Максимальная отметка участка слева
                self.max_l = max(chunk_list(sector.coord[1], 2)[0])
                # Максимальная отметка участка справа
                self.max_r = max(chunk_list(sector.coord[1], 2)[1])
            except:
                print("\n\nОшибка в определении участков. Список участков:\n")
                for sector in sectors:
                    print(sector)

                print("Завершаем программу.")
                raise SystemExit

            print(f"успешно, найдено {len(sectors)} участка.")
            return sectors

        # Перебираем все строки, начиная со второй (min_row=2)
        # И получаем список сырых данных
        for row in sheet.iter_rows(min_row=2, values_only=True):
            # Проверяем, что строка не пустая
            if any(cell is not None and cell != "" for cell in row):
                # Заменяем None на пустую строку
                processed_row = [cell if cell is not None else "" for cell in row]
                __raw_data.append(processed_row)

        # Устанавливаем основные параметры морфоствора
        print("    — Устанавливаем основные параметры морфоствора ... ", end="")
        self.title = __raw_data[2][__description_col]  # Заголовок профиля
        self.date = __raw_data[3][__description_col]  # Дата профиля

        self.waterline = __raw_data[4][__description_col]  # Отметка уреза воды
        # Проверяем задан ли урез текстом, если нет округляем до 2 знаков
        if not isinstance(self.waterline, str):
            self.waterline = round(self.waterline, 2)

        self.dH = __raw_data[5][__description_col]  # Расчётный шаг по глубине
        self.coords = __raw_data[6][__description_col]  # Координаты

        # Считываем отметку предела размыва (в скобках можно указать границы)
        try:
            erosion_limit_list = [
                float(x.strip()) for x in __raw_data[7][__description_col].split(",")
            ]
            # Предел размыва
            self.erosion_limit = erosion_limit_list[0]
            # координаты предела размыва
            self.erosion_limit_coord = erosion_limit_list[1:]
        except:
            self.erosion_limit = __raw_data[7][__description_col]

        self.top_limit = __raw_data[8][__description_col]  # Верхняя граница
        self.top_limit_description = __raw_data[9][
            __description_col
        ]  # Описание верхней границы
        print("успешно!")

        # Считываем и записываем все точки x и y профиля
        print("    — Считываем координаты профиля ... ", end="")
        for i in range(len(__raw_data)):
            if not isinstance(__raw_data[i][__x_coord_col], str):
                self.x.append(__raw_data[i][__x_coord_col])
                self.y.append(__raw_data[i][__y_coord_col])
        print(f"успешно, найдено {len(self.x)} точки, длина профиля {self.x[-1]:.2f} м")

        self.ele_min = min(self.y)  # Минимальная отметка профиля
        self.ele_max = max(self.y)  # Максимальная отметка профиля

        # Заполнения таблицы обеспеченностей
        print("    — Считываем обеспеченности ... ", end="")
        for i in range(6, len(__raw_data[0])):
            prob_ind = __raw_data[0][i]
            prob_val = __raw_data[1][i]

            # Определяем РУВВ
            if str(prob_ind).endswith('*'):
                try:
                    self.probability.append([float(prob_ind[:-1]), prob_val])
                except ValueError:
                    self.probability.append([prob_ind[:-1], prob_val])
                # Устанавливаем индекс РУВВ из таблицы обеспеченностей
                self.design_water_level_index = i - 6
            else:
                self.probability.append([prob_ind, prob_val])

        # Удаляем пустые обеспеченности из списка обеспеченностей
        self.probability = [x for x in self.probability if x != ["", ""]]

        print(f"успешно, найдено {len(self.probability)} обеспеченностей.")

        # Обработка и получение данных по секторам из "сырых" данных
        self.sectors = get_sectors(self)
        self.situation = get_situation(self)

    def get_sectors_result(self):
        df = self.hydraulic_table.swaplevel(0, 1, axis=0)
        wl = self.levels_result.iloc[self.design_water_level_index]['H']
        q, h, v, b, f = np.nan, np.nan, np.nan, np.nan, np.nan

        result = pd.DataFrame(columns=[
            'name', 'slope', 'roughness', 'consumption',
            'depth', 'speed', 'width', 'area'])

        for sector in self.sectors:
            try:
                if wl >= df.loc[sector.name].index.min() and wl <= df.loc[sector.name].index.max():
                    fQ = interpolate.interp1d(
                        df.loc[(sector.name), "Q"].index,
                        df.loc[(sector.name), "Q"].values,
                    )
                    fV = interpolate.interp1d(
                        df.loc[(sector.name), "V"].index,
                        df.loc[(sector.name), "V"].values,
                    )
                    fH = interpolate.interp1d(
                        df.loc[(sector.name), "Hср"].index,
                        df.loc[(sector.name), "Hср"].values,
                    )
                    fB = interpolate.interp1d(
                        df.loc[(sector.name), "B"].index,
                        df.loc[(sector.name), "B"].values,
                    )
                    fF = interpolate.interp1d(
                        df.loc[(sector.name), "F"].index,
                        df.loc[(sector.name), "F"].values,
                    )

                    q = float(fQ(wl))
                    h = float(fH(wl))
                    v = float(fV(wl))
                    b = float(fB(wl))
                    f = float(fF(wl))
            except KeyError:
                q, h, v, b, f = np.nan, np.nan, np.nan, np.nan, np.nan

            row = {
                'name': sector.name,
                'slope': sector.slope,
                'roughness': sector.roughness,
                'consumption': q,
                'depth': h,
                'speed': v,
                'width': b,
                'area': f
            }

            sector.consumption = q
            sector.speed = v
            sector.width = b
            sector.area = f
            sector.depth = h
            # Удаляем столбцы полностью состоящие из NaN для избежания предупреждения
            # Pandas: FutureWarning concatenation with empty or all-NA entries is deprecated
            result.dropna(axis=1, how='all', inplace=True)
            result = pd.concat([result, pd.DataFrame.from_records([row])], ignore_index=True)
            q, h, v, b, f = np.nan, np.nan, np.nan, np.nan, np.nan

        # Подбираем параметры суммирующей кривой
        sum_text = 'Сумма'

        fQ = interpolate.interp1d(df.loc[(sum_text), 'Q'].index, df.loc[(sum_text), 'Q'].values)
        fV = interpolate.interp1d(df.loc[(sum_text), 'V'].index, df.loc[(sum_text), 'V'].values)
        fH = interpolate.interp1d(df.loc[(sum_text), 'Hср'].index, df.loc[(sum_text), 'Hср'].values)
        fB = interpolate.interp1d(df.loc[(sum_text), 'B'].index, df.loc[(sum_text), 'B'].values)
        fF = interpolate.interp1d(df.loc[(sum_text), 'F'].index, df.loc[(sum_text), 'F'].values)

        q = round(float(fQ(wl)), 3)
        h = round(float(fH(wl)), 3)
        v = round(float(fV(wl)), 3)
        b = round(float(fB(wl)), 3)
        f = round(float(fF(wl)), 3)

        sum_row = {
            'name': "Все участки",
            'slope': np.nan,
            'roughness': np.nan,
            'consumption': q,
            'depth': h,
            'speed': v,
            'width': b,
            'area': f
        }

        result = pd.concat([result, pd.DataFrame.from_records([sum_row])], ignore_index=True)
        return result

    def get_min_sector(self):
        """
        Функция нахождения участка с наименьшей отметкой дна.

        :return: [Номер по списку, [Участок]]
        """

        id = 0
        i = 0
        min_sector = self.sectors[0]

        for sector in self.sectors:
            if min(sector.coord[1]) < min(min_sector.coord[1]):
                min_sector = sector
                id = i
            i += 1
        return (id, min_sector)

    def get_q_max(self):
        """
        Функция нахождения максимальной обеспеченности и расхода воды по исходным данным.

        :return: [Обеспеченность, Расход]
        """
        q_max = float(self.probability[0][1])
        obsp = self.probability[0][0]
        for Q in self.probability:
            if q_max <= Q[1]:
                q_max = Q[1]
                obsp = Q[0]

        return (obsp, q_max)

    def calculate(self):
        # Значение расхода до которого необходимо
        # считать (максимальной введенная обеспеченности + 20%)
        consumption_check = self.get_q_max()[1] + (self.get_q_max()[1] * 0.20)

        # Проверяем задан ли расчётный шаг в исходных данных
        if isinstance(self.dH, str) or self.dH == 0:
            self.dH = 1
            dH = self.dH
        else:
            dH = self.dH

        # Переводим сантиметры приращения в метры
        dH = dH / 100

        min_sector = self.get_min_sector()

        # Исходные сектора для расчёта (сектор, содержащий минимальную отметку)
        calc_sectors = [min_sector[0]]

        # Уровень воды, с минимальным отступом
        water_level = min(self.y) + dH

        # Обнулённые переменные
        consumption_summ = 0
        area_summ = 0
        n = 0

        col = ["Участок", "УВ", "F", "B", "Hср", "Hмакс", "V", "Q", "Shezi"]
        df = pd.DataFrame(columns=col, dtype=float)
        # Первый расчётный элемент суммирующей кривой со всеми нулями
        df = pd.concat(
            [
                df,
                pd.DataFrame.from_records(
                    [dict(zip(col, ["Сумма", self.ele_min, 0, 0, 0, 0, 0, 0, 0]))]
                ),
            ],
            ignore_index=True,
        )

        # Цикл расчёта до максимальной обеспеченности + 20% из исходных данных
        while consumption_summ < consumption_check:
            print(f"Выполняем расчёты для уровня {water_level:.2f}", end="\r")

            consumption_summ = 0
            wc_list = list()
            area_list = list()

            if config.OVERFLOW:
                for i in calc_sectors:
                    sector = self.sectors[i]
                    x = sector.coord[0]
                    y = sector.coord[1]

                    # Максимальная отметка слева
                    previous_min_ele = max(chunk_list(y, 2)[0])
                    # Максимальная отметка справа
                    next_min_ele = max(chunk_list(y, 2)[1])

                    # Проверка на перелив через границы участка
                    if (
                        (water_level >= previous_min_ele)
                        and (i - 1 not in calc_sectors)
                        and (i - 1 >= 0)
                    ):
                        calc_sectors.append(i - 1)
                    if (
                        (water_level >= next_min_ele)
                        and (i + 1 not in calc_sectors)
                        and (i + 1 <= len(self.sectors) - 1)
                    ):
                        calc_sectors.append(i + 1)

                    # Сектор воды и основные его параметры
                    # Расчетный участок является участком с минимальными отметками
                    # либо расчёт выполняется с одновременным заполнением
                    # начинаем заполнять с точки с минимальной отметкой
                    if sector.id == min_sector[1].id:
                        water = WaterSection(x, y, water_level)

                    # Расчетный участок находится слева от начального
                    # начинаем заполнять с крайней правой точки
                    elif sector.id < min_sector[1].id:
                        water = WaterSection(
                            x, y, water_level, start_point=[len(y) - 1, y[-1]]
                        )

                    # Расчетный участок находится справа от начального
                    # начинаем заполнять с крайней левой точки
                    elif sector.id > min_sector[1].id:
                        water = WaterSection(x, y, water_level, start_point=[0, y[0]])

                    # Расчёт параметров для воды
                    calc = Calculation(
                        h=water.average_depth,
                        n=sector.roughness,
                        i=sector.slope,
                        a=water.area,
                    )

                    wc_list.append(calc.q)

                    r = dict(
                        zip(
                            col,
                            [
                                sector.name,
                                round(water_level, 2),
                                water.area,
                                water.width,
                                water.average_depth,
                                water.max_depth,
                                calc.v,
                                calc.q,
                                calc.shezi,
                            ],
                        )
                    )

                    # Добавляем в список с результирующими значениями значения по секторам
                    # для последующего суммирования/вычисления средних значений
                    df = df._append(r, ignore_index=True)

            else:
                # Расчёт с заполнением по участкам
                for sector in self.sectors:
                    x = sector.coord[0]
                    y = sector.coord[1]

                    if min(y) < water_level:
                        # Сектор воды и основные его параметры
                        water = WaterSection(x, y, water_level)

                        # Расчёт параметров для воды
                        calc = Calculation(
                            h=water.average_depth,
                            n=sector.roughness,
                            i=sector.slope,
                            a=water.area,
                        )

                        wc_list.append(calc.q)

                        # Добавляем в список с значения по секторам
                        r = dict(
                            zip(
                                col,
                                [
                                    sector.name,
                                    round(water_level, 2),
                                    water.area,
                                    water.width,
                                    water.average_depth,
                                    water.max_depth,
                                    calc.v,
                                    calc.q,
                                    calc.shezi,
                                ],
                            )
                        )

                        # Добавляем в список с результирующими значениями значения по секторам
                        # для последующего суммирования/вычисления средних значений
                        df = pd.concat([df, pd.DataFrame.from_records([r])], ignore_index=True)

            consumption_summ += sum(wc_list)
            area_summ += sum(area_list)

            # Пустые значения для суммирующей кривой
            r_sum = dict(
                zip(col, ["Сумма", round(water_level, 2), 0, 0, 0, 0, 0, 0, 0])
            )
            df = pd.concat([df, pd.DataFrame.from_records([r_sum])], ignore_index=True)

            water_level += dH
            n += 1

        # TODO: remake to use one dataframe
        df = df.set_index(["УВ", "Участок"])
        water_levels = df.index.levels[0]

        # Заполняем суммирующие данные
        df.loc[(water_levels, "Сумма"), "F"] = df.groupby(level=0)["F"].transform("sum")
        df.loc[(water_levels, "Сумма"), "B"] = df.groupby(level=0)["B"].transform("sum")
        df.loc[(water_levels, "Сумма"), "Hср"] = df.groupby(level=0)["F"].transform(
            "sum"
        ) / df.groupby(level=0)["B"].transform("sum")
        df.loc[(water_levels, "Сумма"), "Hмакс"] = df.groupby(level=0)[
            "Hмакс"
        ].transform("max")
        df.loc[(water_levels, "Сумма"), "Q"] = df.groupby(level=0)["Q"].transform("sum")
        df.loc[(water_levels, "Сумма"), "V"] = df.groupby(level=0)["Q"].transform(
            "sum"
        ) / df.groupby(level=0)["F"].transform("sum")
        df.loc[(water_levels, "Сумма"), "Shezi"] = df.groupby(level=0)[
            "Shezi"
        ].transform("sum") / (df.groupby(level=0)["Shezi"].transform("count") - 1)
        df = df.fillna(0)

        # Интерполируем значения гидравлической кривой
        # для необходимых обеспеченностей и обновляем таблицу
        p_table = df.loc[(water_levels, "Сумма"), :].droplevel(1)
        self.levels_result = self.get_prob_table(p_table)
        self.hydraulic_table = df
        self.sectors_result = self.get_sectors_result()

        self.fig_profile = GraphProfile(self)

        self.fig_QH = GraphQH(self)
        self.fig_QHV = GraphQHV(self)
        self.fig_QV = GraphQV(self)
        self.fig_VH = GraphVH(self)
        self.fig_QF = GraphQF(self)
        self.fig_FH = GraphFH(self)
        self.fig_QWVH = GraphQWVH(morfostvor=self)

        return df

    def get_prob_table(self, df: pd.DataFrame):
        result = pd.DataFrame(columns=["P", "Q", "H", "F"])

        for prob in self.probability:
            fQ = interpolate.interp1d(df["Q"], df.index)
            fV = interpolate.interp1d(df["Q"], df["V"])
            fF = interpolate.interp1d(df["Q"], df["F"])
            h = float(fQ(prob[1]))
            v = float(fV(prob[1]))
            f = float(fF(prob[1]))

            # Удаляем столбцы полностью состоящие из NaN для избежания предупреждения
            # Pandas: FutureWarning concatenation with empty or all-NA entries is deprecated
            result.dropna(axis=1, how='all', inplace=True)
            result = pd.concat(
                [
                    result,
                    pd.DataFrame.from_records(
                        [{"P": prob[0], "H": h, "Q": prob[1], "V": v, "F": f}]
                    ),
                ],
                ignore_index=True,
            )

        return result

    def get_topography_table(self):
        # Создаем базовый словарь с координатами
        topography_data = {
            "x": self.x,
            "h": self.y,
        }

        # Инициализируем списки для данных о секторах
        sectors = []
        roughness = []
        slope = []

        # Заполняем списки данными из секторов
        for sector in self.sectors:
            for _ in range(sector.start_point, sector.end_point):
                sectors.append(sector.name)
                roughness.append(sector.roughness)
                slope.append(sector.slope)
            # Дописываем данные для последней точки
            if sector == self.sectors[-1]:
                sectors.append(sector.name)
                roughness.append(sector.roughness)
                slope.append(sector.slope)

        # Добавляем данные в основной словарь
        topography_data['sectors'] = sectors
        topography_data['roughness'] = roughness
        topography_data['slope'] = slope

        # Создаем DataFrame и возвращаем его
        return pd.DataFrame(topography_data)


def xls_calculate_hydraulic(in_filename, out_filename, page=None):
    """
    Выполнение гидравлических расчетов и создание отчета по результатам расчетов.
    Исходные данные представлены в in_filename (xls файл).
    По умолчанию расчеты производятся для всех листов xls файла.
    Если задан параметр page, расчет производится только
    для указанной страницы. По результат создается out_filename
    (результирующий отчет в формате docx).

        :param in_filename: Входные данные по створам (.xls файл)
        :param out_filename: Результаты расчетов  (.docx файл)
        :param page=None: Номер страницы в xls файле,
    по умолчанию None (расчеты производятся для всего документа)
    """
    __start_time = time.time()
    # Создаем родительскую папку, если она не существует
    Path(out_filename).parents[0].mkdir(parents=True, exist_ok=True)

    # Удаляем предыдущий отчет, если включена перезапись файла
    if config.REWRITE_DOC_FILE:
        try:
            os.remove(out_filename)
        except FileNotFoundError:
            pass
        except PermissionError:
            print(f"\nОшибка! Программа не может получить доступ "
                  f"к файлу {out_filename}, возможно он открыт?")
            print('Программа будет завершена.')
            sys.exit(35)

    page_quantity = get_xls_sheet_quantity(in_filename)
    stvors = []

    def single_page(in_filename, out_filename, page):
        """Выполнение расчета одной страницы исходных данных из xls файла.

        Args:
            in_filename (str): _description_
            out_filename (str): _description_
            page (ind): _description_

        Returns:
            Morfostvor(): Возвращает объект морфоствора с выполненными расчетами, и сохраняет отчет
        """
        __start_time = time.time()
        stvor = Morfostvor()
        stvor.read_xls(in_filename, page)
        stvor.calculate()
        __compute_time = time.time() - __start_time
        __report_start_time = time.time()
        generate_morfostvor_report(stvor, out_filename)
        if config.PROFILE_SAVE_PICTURES or config.CURVE_SAVE_PICTURES:
            save_graphic(stvor, str(Path(out_filename).parents[0]))

        print(
            f"\n------------------------ "
            f"Файл {out_filename} сохранён успешно "
            f"------------------------\n"
        )
        if config.DEBUG:
            print(f"--- Расчеты: {__compute_time:.4f} секунд ---")
            print(
                f"--- Сборка отчета: "
                f"{time.time() - __report_start_time:.4f} секунд ---"
            )
            print(f"--- Всего: {time.time() - __start_time:.4f} секунд ---\n")
        return stvor

    # Расчет для всех листов xls файла
    if page is None:
        for i in range(page_quantity):
            stvors.append(single_page(in_filename, out_filename, i))

        # Вставка сводных таблиц
        __summary_start_time = time.time()
        insert_summary_QV_tables(stvors, out_filename)
        if config.DEBUG:
            print(
                f"\n--- Вставка сводных таблиц: "
                f"{time.time() - __summary_start_time:.4f} секунд ---"
            )

    # Расчет только одного листа xls файла
    elif isinstance(page, int):
        single_page(in_filename, out_filename, page)

    else:
        print("Номер листа должен быть целым числом.")
        sys.exit(0)

    if config.DEBUG:
        print(f"--- Итого: {time.time() - __start_time:.4f} секунд ---\n")
