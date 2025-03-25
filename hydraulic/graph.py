from dataclasses import dataclass

import matplotlib
import matplotlib.figure
import matplotlib.patches as patches
import matplotlib.patheffects as path_effects
import matplotlib.pyplot as plt
import numpy as np
import scipy.interpolate as interpolate
from labellines import labelLines
from matplotlib import gridspec
from matplotlib.patches import Rectangle

# from hydraulic.profile import Morfostvor
import hydraulic.config as config
from hydraulic.lib import closest_upper_multiple, get_pk, get_water_sections, text_sanitize
from hydraulic.models import WaterSection


@dataclass
class Graph:
    _fig_size = (16.5, 9)
    _y_limits = []

    _x_label_text = ""
    _y_label_text = ""
    _ax_title_text = ""

    morfostvor: object
    fig: plt.Figure = plt.figure(figsize=_fig_size)
    ax: plt.subplot = fig.add_subplot(111)

    def __post_init__(self):
        self.clean()
        morfostvor = self.morfostvor

        # Вытягиваем цвета
        self.sector_colors = {}
        for sector in morfostvor.sectors:
            self.sector_colors[sector.name] = sector.color

        # Выполняем отрисовку содержимого
        self.draw()
        self.set_style()

    def draw(self):
        pass

    def set_style(self):
        fig = self.fig
        ax = self.ax

        fig.subplots_adjust(bottom=0.08, left=0.08, right=0.92)

        # Устанавливаем заголовки графиков
        if config.GRAPHICS_TITLES:
            ax.set_title(
                self._ax_title_text,
                color=config.COLOR["title_text"],
                fontsize=config.FONT_SIZE["title"],
                y=1.05,
            )

        # Настраиваем границы и толщину линий границ
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        ax.spines["left"].set_linewidth(config.LINE_WIDTH["ax_border"])
        ax.spines["bottom"].set_linewidth(config.LINE_WIDTH["ax_border"])

        # Включаем отображение второстепенных засечек на осях
        self.ax.minorticks_on()

        # Устанавливаем параметры засечек на основных осях
        ax.tick_params(
            which="major",
            direction="out",
            width=2,
            length=5,
            pad=config.PADDING["ax_tick_labels"],
            labelcolor=config.COLOR["ax_value_text"],
            labelsize=config.FONT_SIZE["ax_major"],
        )

        ax.tick_params(
            which="minor",
            direction="out",
            width=1.5,
            length=3.5,
            pad=config.PADDING["ax_tick_labels"],
            labelcolor=config.COLOR["ax_value_text"],
            labelsize=config.FONT_SIZE["ax_minor"],
        )

        # Устанавливаем параметры подписей осей
        ax.set_xlabel(
            self._x_label_text,
            color=config.COLOR["ax_label_text"],
            fontsize=config.FONT_SIZE["ax_label"],
            fontstyle="italic",
            weight="normal",
        )
        ax.xaxis.set_label_coords(0.5, -0.06)
        ax.set_ylabel(
            self._y_label_text,
            color=config.COLOR["ax_label_text"],
            fontsize=config.FONT_SIZE["ax_label"],
            fontstyle="italic",
            weight="normal",
        )
        ax.yaxis.set_label_coords(-0.065, 0.5)

        # Устанавливает параметры вывода значений осей
        ax.yaxis.set_major_formatter(matplotlib.ticker.FormatStrFormatter("%.10g"))

        # Настройка параметров отображение сетки
        ax.grid(
            which="major",
            color=config.COLOR["ax_grid"],
            linestyle=":",
            linewidth=1,
            alpha=0.9,
        )
        ax.grid(
            which="minor",
            color=config.COLOR["ax_grid_sub"],
            linestyle=":",
            linewidth=1,
            alpha=0.9,
        )

        # Устанавливаем отступы в графиках
        ax.margins(0.025)

        # Скрыть сетку
        if config.CURVE_HIDE_GRID:
            ax.grid(visible=False, which="Both")

        # Установка параметров полей графика
        self.fig.subplots_adjust(left=0.08, bottom=0.08, right=0.93, top=0.9)

    def clean(self):
        """Очистка осей графика и обнуление связанных переменных"""
        # Очищаем все оси
        plt.close(self.fig)

        for attr in vars(self):
            if attr.startswith("ax"):
                getattr(self, attr).cla()

        # Очистка осей скоростей на совмещенном графике
        if hasattr(self, "ax_secondary"):
            self.ax_secondary.cla()

        # Обнуляем границы y
        self._y_limits = []
        self._y_limits = []


@dataclass
class GraphCurve(Graph):
    def draw_water_levels(self, morfostvor, ax: plt.subplot, x="Q", y="H", y_min=0):
        """Функция выводит на график ax отметку и линии пересечения
           x и y.

        Args:
            morfostvor (Morfostvor): Объект из которого необходимо брать данные.
            ax (plt.subplot): График для нанесения отметок.
            x (str, optional): Ось x. Defaults to 'Q'.
            y (str, optional): Ось y. Defaults to 'H'.
        """
        try:
            if config.HYDRAULIC_CURVE_LEVELS:
                for _, row in morfostvor.levels_result.iterrows():
                    x1, x2 = 0, row[x]
                    y1, y2 = row[y], row[y]

                    # Вывод значений округленных, проверка на содержание значений
                    try:
                        water_level_text = ax.text(
                            0.002,
                            row[y],
                            f"$P_{{{row['P']:.2g}\\%}} = {row[y]:.2f}$",
                            color=config.COLOR["water_level_text"],
                            fontsize=config.FONT_SIZE["water_level"],
                            weight="bold",
                        )

                        water_level_text.set_path_effects(
                            [
                                path_effects.Stroke(linewidth=3, foreground="white", alpha=0.55),
                                path_effects.Normal(),
                            ]
                        )

                    except ValueError:
                        water_level_text = ax.text(
                            0.002,
                            row[y],
                            f"${row['P']} = {row[y]:.2f}$",
                            color=config.COLOR["water_level_text"],
                            fontsize=config.FONT_SIZE["water_level"],
                            weight="bold",
                        )

                        water_level_text.set_path_effects(
                            [
                                path_effects.Stroke(linewidth=3, foreground="white", alpha=0.55),
                                path_effects.Normal(),
                            ]
                        )

                    ax.plot(
                        [x1, x2, x2, x2],
                        [y1, y2, y_min, y_min],
                        linestyle="-",
                        color="mediumturquoise",
                        marker="o",
                        linewidth=1,
                        markersize=1,
                    )
        except:
            print("Внимание! Вывод расчётных уровней на график не возможен!")

    def draw_curve(self, morfostvor, ax: plt.subplot, x="Q", y="УВ"):
        """Отрисовка кривой на графике по заданным из морфоствора параметрам.

        Args:
            morfostvor (Morfostvor): Объект морфоствора из которого получаем данные
            ax (plt.subplot): Ось на которой строится график
            x (str, optional): Значения по оси x. Defaults to 'Q'.
            y (str, optional): Значения по оси y. Defaults to 'УВ'.
        """

        df = morfostvor.hydraulic_table
        df = df.reset_index(level=0)  # Переводим индекс уровня воды в столбец

        sectors = set(df.index)  # Удаляем дублирующиеся записи
        sectors.remove("Сумма")  # Удаляем запись суммирующего участка

        # Отрисовка суммирующей кривой на графике
        ax.plot(
            df.loc[("Сумма"), x],
            df.loc[("Сумма"), y],
            label="Общая",
            linewidth=3,
            color="red",
        )

        # Отрисовка кривых по участкам
        for sector in sectors:
            ax.plot(
                df.loc[(sector), x],
                df.loc[(sector), y],
                "--",
                label=sector,
                color=self.sector_colors[sector],
            )
        # Подписи на линиях
        try:
            labelLines(ax.get_lines(), zorder=2.5, fontsize=12, shrink_factor=0.05)
        except:
            print("Внимание! Подписи на линиях не возможны!")

    def draw_legend(self, ax: plt.subplot):
        # Отрисовка легенды
        ax.legend(loc="lower right", fontsize=config.FONT_SIZE["legend"])


class GraphQWVH(GraphCurve):
    vertical_offset = [0, -0.08, -0.08]

    def __init__(self, morfostvor):
        self.clean()
        self.morfostvor = morfostvor
        self._initialize_colors()
        self._initialize_figure()
        self._draw_graphs()

    def _initialize_colors(self):
        """Извлекает цвета секторов из morfostvor."""
        self.sector_colors = {sector.name: sector.color for sector in self.morfostvor.sectors}

    def _initialize_figure(self):
        """Создаёт фигуру и оси для трёх графиков."""
        self.fig, (self.ax1, self.ax2, self.ax3) = plt.subplots(
            1, 3, figsize=(16, 8), sharey=True, gridspec_kw={"wspace": -0.08}
        )
        plt.subplots_adjust(left=0.075, bottom=0.10, right=0.97, top=0.95)

    def _draw_graphs(self):
        """Отрисовывает три графика Q(H), W(H) и V(H)."""
        df = self._prepare_dataframe()

        graph_specs = [
            (self.ax1, "Q", "red", r"$Q=f(H), м³/сек$"),
            (self.ax2, "F", "green", r"$F=f(H), м²$"),
            (self.ax3, "V", "blue", r"$V=f(H), м/сек$"),
        ]

        for i, (ax, x, color, label) in enumerate(graph_specs):
            self._draw_graph(ax, df, x, color, label, self.vertical_offset[i])

        self.fig.legend(loc="upper center", ncols=6)

    def _draw_graph(self, ax, df, x, color, label, offset):
        """Отрисовка одного графика графика."""
        self.style_axis(ax, offset, label, color)

        # Подписи линий
        main_label = f"{x}_{{общ.}}"
        secondary_label = f"{x}_{{русл.}}"
        if ax == self.ax3:
            main_label = f"{x}_{{ср. общ.}}"
            secondary_label = f"{x}_{{ср. русл.}}"

        # Отрисовка линий
        self._plot_graph(ax, df, x, "УВ", "сумма", main_label, color, zorder=10, alpha=0.5)
        self._plot_graph(
            ax,
            df,
            x,
            "УВ",
            "русло",
            secondary_label,
            color,
            linestyle="--",
            linewidth=1,
            zorder=11,
        )

        # Отрисовка линий пересечений и подписей
        self._draw_water_levels_x(self.morfostvor, ax, x, "H", color, offset)
        self._draw_water_vertical_labels(self.morfostvor, ax, x, color, offset)

        if ax == self.ax1:
            self.ax1.set_ylabel(
                f"H, м {config.ALTITUDE_SYSTEM}",
                color="black",
                fontsize=config.FONT_SIZE["ax_label"],
                fontstyle="italic",
                weight="normal",
            )
            # Расположение подписи оси Y
            self.ax1.yaxis.set_label_coords(-0.18, 0.61)
            ax.tick_params(axis="y", which="both", color="black", labelcolor="black")
            ax.spines[["left"]].set_color("black")
            self._draw_water_horizontal_labels(self.morfostvor, ax, "H")

        else:
            ax.spines["left"].set_visible(False)
            ax.tick_params(which="both", axis="y", length=0, labelleft=False)

    def _prepare_dataframe(self):
        """Подготавливает и нормализует таблицу гидравлических данных."""
        df = self.morfostvor.hydraulic_table.reset_index(level=0)
        df.index = df.index.str.lower()
        return df

    def _draw_water_levels_x(
        self,
        morfostvor,
        ax: plt.subplot,
        x="Q",
        y="H",
        color="black",
        x_ax_v_offset=0,
    ):
        x_min = 0
        x_max = ax.get_xlim()[1]
        y_limits = ax.get_ylim()

        # Минимальное значение по оси Y кривых
        min_value = morfostvor.hydraulic_table.index.get_level_values(0).min()

        # Получаем текущие засечки оси y
        yticks = ax.get_yticks()

        # Фильтруем засечки, чтобы оставить только те, которые больше или равны min_value
        new_yticks = yticks[yticks >= min_value]

        # Удаляем первую засечку если она ниже оси X
        if new_yticks[0] < min_value:
            new_yticks = np.delete(new_yticks, 0)

        # Применяем новые основные засечки
        ax.set_yticks(new_yticks)

        # Скрываем все значения на оси y, которые меньше оси x
        ax.set_yticklabels([f"{y_tick:.2f}" if y_tick >= min_value else "" for y_tick in new_yticks])

        # Включаем основные засечки
        ax.minorticks_on()

        # Расстояние между засечками
        minor_tick_dist = (new_yticks[1] - new_yticks[0]) / 5

        # Определяем границы отрисовки засечек
        minor_tick_min_y = closest_upper_multiple(min_value, minor_tick_dist)
        minor_tick_max_y = new_yticks[-1]

        # Включаем вспомогательные засечки
        ax.set_yticks(np.arange(minor_tick_min_y, minor_tick_max_y, minor_tick_dist), minor=True)

        # Вычисляем фактическое положение оси X с учетом смещения
        x_ax_position_on_y = y_limits[0] + (y_limits[1] - y_limits[0]) * x_ax_v_offset

        # Принудительно устанавливаем y_limits так, чтобы ось X не обрезалась
        ax.set_ylim(min(y_limits[0], x_ax_position_on_y), y_limits[1])

        # Сдвигаем ось X выше, чтобы вертикальные линии доходили до неё
        ax.spines["bottom"].set_position(("data", x_ax_position_on_y))

        # Обрезаем линию оси Y ниже оси X
        clip_rect = patches.Rectangle(
            (-1e6, x_ax_position_on_y),  # Координаты (широкий прямоугольник слева)
            2e6,  # Ширина (очень большая, чтобы покрыть всю область)
            1e6,  # Высота (очень большая вверх)
            transform=ax.transData,  # Учитываем систему координат оси
            clip_on=False,  # Обрезаем только элементы оси
        )
        ax.spines["left"].set_clip_path(clip_rect)

        # Отрисовка линий
        for _, row in morfostvor.levels_result.iterrows():
            if x == "V":
                x_max = row[x]

            x1, x2 = x_min, x_max
            y1, y2 = row[y], row[y]

            # Горизонтальная линия
            ax.plot(
                [x1, x2],
                [y1, y2],
                linestyle="-",
                color="lightgray",
                linewidth=1,
                markersize=1,
                alpha=1,
            )

            # Вертикальная линия
            ax.plot(
                [row[x], row[x]],
                [y2, x_ax_position_on_y],
                linestyle="-",
                color=color,
                marker="o",
                linewidth=1,
                markersize=4,
                alpha=0.3,
            )

    def _draw_water_horizontal_labels(
        self,
        morfostvor,
        ax: plt.subplot,
        y_col="H",
    ):
        # Отрисовка подписей отметок по расчетным обеспеченностям
        x_lim = ax.get_xlim()
        y_lim = ax.get_ylim()
        step_x = (x_lim[-1] - x_lim[0]) / 10
        step_y = (y_lim[-1] - y_lim[0]) / 10

        for _, row in morfostvor.levels_result.iterrows():
            x = step_x / 5
            y = row[y_col] + step_y / 20

            # Вывод значений округленных, проверка на содержание значений
            try:
                water_level_text = ax.text(
                    x,
                    y,
                    f"$P_{{{row['P']:.2g}\\%}}={row[y_col]:.2f}$",
                    color="gray",
                    fontsize=9,
                    weight="black",
                    alpha=0.8,
                    zorder=25,
                )

                water_level_text.set_path_effects(
                    [
                        path_effects.Stroke(linewidth=3, foreground="white", alpha=0.55),
                        path_effects.Normal(),
                    ]
                )

            except ValueError:
                water_level_text = ax.text(
                    x,
                    y,
                    f"${row['P']} = {row[y_col]:.2f}$",
                    color="gray",
                    fontsize=9,
                    weight="black",
                    alpha=0.8,
                    zorder=16,
                )

                water_level_text.set_path_effects(
                    [
                        path_effects.Stroke(
                            linewidth=2,
                            foreground="white",  # , alpha=0.8
                        ),
                        path_effects.Normal(),
                    ]
                )

    def _draw_water_vertical_labels(
        self,
        morfostvor,
        ax: plt.subplot,
        x_col="Q",
        color="black",
        x_ax_v_offset=0.0,
    ):
        # Отрисовка подписей параметров по расчетным обеспеченностям
        x_lim = ax.get_xlim()
        y_lim = ax.get_ylim()
        step_x = (x_lim[-1] - x_lim[0]) / 10
        step_y = (y_lim[-1] - y_lim[0]) / 10

        # Вычисляем фактическое положение оси X с учетом смещения
        x_ax_position_on_y = y_lim[0] + (y_lim[1] - y_lim[0]) * x_ax_v_offset - step_y * 10 * x_ax_v_offset

        for _, row in morfostvor.levels_result.iterrows():
            x = row[x_col]
            y = x_ax_position_on_y

            # Вывод значений округленных, проверка на содержание значений
            try:
                water_level_text = ax.text(
                    x + step_x / 25,
                    y + step_y / 5,
                    f"$P_{{{row['P']:.2g}\\%}}={row[x_col]:.2f}$",
                    color=color,
                    fontsize=9,
                    weight="black",
                    alpha=0.8,
                    zorder=25,
                    rotation="vertical",
                    ha="left",
                )

                water_level_text.set_path_effects(
                    [
                        path_effects.Stroke(linewidth=4, foreground="white", alpha=0.9),
                        path_effects.Normal(),
                    ]
                )

            except ValueError:
                water_level_text = ax.text(
                    x + step_x / 25,
                    y + step_y / 8,
                    f"${row['P']} = {row[x_col]:.2f}$",
                    color=color,
                    fontsize=9,
                    weight="black",
                    alpha=0.8,
                    zorder=25,
                    rotation="vertical",
                    ha="left",
                )

                water_level_text.set_path_effects(
                    [
                        path_effects.Stroke(linewidth=4, foreground="white", alpha=0.9),
                        path_effects.Normal(),
                    ]
                )

    def _plot_graph(self, ax, df, x_col, y_col, sector_name="сумма", label="", color="b", **kwargs):
        try:
            ax.plot(
                df.loc[(sector_name), x_col],
                df.loc[(sector_name), y_col],
                label=rf"${label}$",
                color=color,
                **kwargs,
            )
        except KeyError:
            print(f"Нет данных по участку: {sector_name}")

    # Функция для стилизации графиков
    @staticmethod
    def style_axis(ax, x_ax_v_offset, x_label, color):
        ax.spines[["top", "right"]].set_visible(False)
        ax.spines["bottom"].set_position(("axes", x_ax_v_offset))
        ax.spines[["bottom", "left"]].set_linewidth(config.LINE_WIDTH["ax_border"])
        ax.spines[["bottom", "left"]].set_color(color)

        ax.set_facecolor("none")
        ax.grid(False)
        ax.minorticks_on()

        ax.tick_params(
            which="both",
            direction="out",
            width=2,
            length=5,
            pad=config.PADDING["ax_tick_labels"],
            labelsize=config.FONT_SIZE["ax_major"],
            labelcolor=color,
            colors=color,
        )

        ax.tick_params(which="minor", width=1.5, length=3)
        ax.xaxis.set_tick_params(labelcolor=color)
        ax.yaxis.set_major_formatter(matplotlib.ticker.FormatStrFormatter("%.10g"))
        ax.set_xlabel(
            x_label,
            color=color,
            fontsize=config.FONT_SIZE["ax_label"],
            fontstyle="italic",
            weight="normal",
        )
        ax.margins(0.0)


@dataclass
class GraphQHV(GraphCurve):
    # Номер рисунка
    _fig_num = 2
    _fig_size = (16.5, 9)
    fig: plt.figure = plt.figure(_fig_num, figsize=_fig_size)
    ax: plt.subplot = fig.add_subplot(111)
    ax_secondary = ax.twinx()

    # Подписи осей
    _x_label_text = "Q, м³/с"
    _y_label_text = f"H, м {config.ALTITUDE_SYSTEM}"
    _y2_label_text = "V, м/с"
    _ax_title_text = "Гидравлическая кривая Q=f(H) с наложением Q=f(V)"

    def draw_curve(self, morfostvor, ax: plt.subplot, ax_secondary, x="Q", y="УВ", yy="V"):
        """Отрисовка кривой на графике по заданным из морфоствора параметрам.

        Args:
            morfostvor (Morfostvor): Объект морфоствора из которого получаем данные
            ax (plt.subplot): Ось на которой строится график
            x (str, optional): Значения по оси x. Defaults to 'Q'.
            y (str, optional): Значения по оси y. Defaults to 'УВ'.
            yy (str, optional): Вторые значения по оси y. Defaults to 'V'.
        """

        df = morfostvor.hydraulic_table
        df = df.reset_index(level=0)  # Переводим индекс уровня воды в столбец

        sectors = set(df.index)  # Удаляем дублирующиеся записи
        sectors.remove("Сумма")  # Удаляем запись суммирующего участка

        ax_secondary.set_label(self._y2_label_text)

        # Отрисовка суммирующей кривой на графике
        ax.plot(
            df.loc[("Сумма"), x],
            df.loc[("Сумма"), y],
            label="Сумма",
            linewidth=3,
            color="red",
        )

        ax_secondary.plot(
            df.loc[("Сумма"), x],
            df.loc[("Сумма"), yy],
            label="Сумма",
            linewidth=3,
            color="navy",
            linestyle="-.",
        )

        ax_secondary.set_ylim(df.loc[("Сумма"), yy].min(), df.loc[("Сумма"), yy].max() + 0.5)

        # Отрисовка кривых по участкам
        for sector in sectors:
            ax.plot(
                df.loc[(sector), x],
                df.loc[(sector), y],
                "--",
                label=sector,
                color=self.sector_colors[sector],
            )

        # Отрисовка легенды
        ax.legend(
            loc="lower right",
            fontsize=config.FONT_SIZE["legend"],
            title="Q = f(H)",
            title_fontsize=14,
        )
        ax_secondary.legend(fontsize=config.FONT_SIZE["legend"], title="Q = f(V)", title_fontsize=14)

        # Настраиваем границы и толщину линий границ
        ax_secondary.spines["top"].set_linewidth(config.LINE_WIDTH["ax_border"])
        ax_secondary.spines["right"].set_linewidth(config.LINE_WIDTH["ax_border"])
        ax_secondary.spines["left"].set_visible(False)
        ax_secondary.spines["bottom"].set_visible(False)

        # Включаем отображение второстепенных засечек на осях
        ax_secondary.minorticks_on()

        # Устанавливаем параметры засечек на основных осях
        ax_secondary.tick_params(
            which="major",
            direction="out",
            width=2,
            length=5,
            pad=config.PADDING["ax_tick_labels"],
            labelcolor=config.COLOR["ax_value_text"],
            labelsize=config.FONT_SIZE["ax_major"],
        )

        ax_secondary.tick_params(
            which="minor",
            direction="out",
            width=1.5,
            length=3.5,
            pad=config.PADDING["ax_tick_labels"],
            labelcolor=config.COLOR["ax_value_text"],
            labelsize=config.FONT_SIZE["ax_minor"],
        )

        # Устанавливаем параметры подписей осей
        ax_secondary.set_ylabel(
            self._y2_label_text,
            color=config.COLOR["ax_label_text"],
            fontsize=config.FONT_SIZE["ax_label"],
            fontstyle="italic",
        )
        ax_secondary.yaxis.set_label_coords(1.05, 0.5)

        # Устанавливает параметры вывода значений осей
        ax_secondary.yaxis.set_major_formatter(matplotlib.ticker.FormatStrFormatter("%.10g"))

        # Устанавливаем отступы в графиках
        ax_secondary.margins(0.025)

    def draw(self):
        y_min = min(self.morfostvor.hydraulic_table.reset_index(0).loc["Сумма"].reset_index(drop=True)["УВ"])
        self.draw_curve(self.morfostvor, self.ax, self.ax_secondary, "Q", "УВ", "V")
        self.draw_water_levels(self.morfostvor, self.ax, "Q", "H", y_min)


@dataclass
class GraphQH(GraphCurve):
    # Номер рисунка
    _fig_num = 3
    _fig_size = (16.5, 9)
    fig: plt.figure = plt.figure(_fig_num, figsize=_fig_size)
    ax: plt.subplot = fig.add_subplot(111)

    # Подписи осей
    _x_label_text = "Q, м³/с"
    _y_label_text = f"H, м {config.ALTITUDE_SYSTEM}"
    _ax_title_text = "Гидравлическая кривая Q=f(H)"

    def draw(self):
        y_min = min(self.morfostvor.hydraulic_table.reset_index(0).loc["Сумма"].reset_index(drop=True)["УВ"])
        self.draw_curve(self.morfostvor, self.ax, "Q", "УВ")
        self.draw_water_levels(self.morfostvor, self.ax, "Q", "H", y_min)
        self.draw_legend(self.ax)


@dataclass
class GraphQV(GraphCurve):
    # Номер рисунка
    _fig_num = 4
    _fig_size = (16.5, 9)
    fig: plt.figure = plt.figure(_fig_num, figsize=_fig_size)
    ax: plt.subplot = fig.add_subplot(111)

    # Подписи осей
    _x_label_text = "Q, м³/с"
    _y_label_text = "V, м/c"
    _ax_title_text = "Кривая скоростей V=f(Q)"

    def draw(self):
        self.draw_curve(self.morfostvor, self.ax, "Q", "V")
        self.draw_water_levels(self.morfostvor, self.ax, "Q", "V")
        self.draw_legend(self.ax)


@dataclass
class GraphVH(GraphCurve):
    # Номер рисунка
    _fig_num = 6
    _fig_size = (16.5, 9)
    fig: plt.figure = plt.figure(_fig_num, figsize=_fig_size)
    ax: plt.subplot = fig.add_subplot(111)

    # Подписи осей
    _x_label_text = "V, м/c"
    _y_label_text = f"H, м {config.ALTITUDE_SYSTEM}"
    _ax_title_text = "Кривая скоростей V=f(H)"

    def draw(self):
        self.draw_curve(self.morfostvor, self.ax, "V", "УВ")
        y_min = self.ax.get_ylim()[0]
        # self.draw_water_levels(self.morfostvor, self.ax, "V", "H", y_min=y_min, )
        self.draw_legend(self.ax)


@dataclass
class GraphFH(GraphCurve):
    # Номер рисунка
    _fig_num = 7
    _fig_size = (16.5, 9)
    fig: plt.figure = plt.figure(_fig_num, figsize=_fig_size)
    ax: plt.subplot = fig.add_subplot(111)

    # Подписи осей
    _x_label_text = "F, м²"
    _y_label_text = f"H, м {config.ALTITUDE_SYSTEM}"
    _ax_title_text = "Кривая площадей F=f(H)"

    def draw(self):
        self.draw_curve(self.morfostvor, self.ax, "F", "УВ")
        # self.draw_water_levels(self.morfostvor, self.ax, "F", "УВ")
        self.draw_legend(self.ax)


@dataclass
class GraphQF(GraphCurve):
    # Номер рисунка
    _fig_num = 5
    _fig_size = (16.5, 9)
    fig: plt.figure = plt.figure(_fig_num, figsize=_fig_size)
    ax: plt.subplot = fig.add_subplot(111)

    # Подписи осей
    _x_label_text = "Q, м³/с"
    _y_label_text = "F, м²"
    _ax_title_text = "Кривая площадей F=f(Q)"

    def draw(self):
        self.draw_curve(self.morfostvor, self.ax, "Q", "F")
        self.draw_water_levels(self.morfostvor, self.ax, "Q", "F")
        self.draw_legend(self.ax)


@dataclass
class GraphProfile(Graph):
    _fig_size = config.PROFILE_SIZE

    fig: plt.figure = plt.figure(figsize=_fig_size)

    __gs = gridspec.GridSpec(80, 3)

    ax_top: plt.subplot = fig.add_subplot(__gs[0, :], frame_on=False)
    ax: plt.subplot = fig.add_subplot(__gs[1:57, :])
    ax_bottom: plt.subplot = fig.add_subplot(__gs[57:, :])
    ax_bottom_overlay: plt.subplot = fig.add_subplot(__gs[57:, :], frame_on=False)

    footers_num: int = 0
    _footer_y: int = 0

    def __post_init__(self):
        self.clean()

        # Добавляем в список границ максимальную и минимальную отметки
        self._y_limits.append(max(self.morfostvor.y))
        self._y_limits.append(min(self.morfostvor.y))

        self._update_limit()
        self.set_style()

        self.draw_profile_footer()
        self.draw_sectors()
        self.draw_profile_bottom()

    def draw_profile_bottom(self):
        """
        Отрисовка дна профиля.

        :return: Отрисовывает дно профиля на графике ax_profile.
        """
        if config.PROFILE_SECTOR_BOTTOM_LINE is False:
            self.ax.plot(
                self.morfostvor.x,
                self.morfostvor.y,
                color=config.COLOR["profile_bottom"],
                linewidth=config.LINE_WIDTH["profile_bottom"],
                linestyle="solid",
                zorder=10,
            )

    def draw_profile_footer(self):
        """
        Отрисовка подвала с информацией о профиле.

            :param self:
        """
        hs = 10  # Стандартная высота ячейки подвала
        hs_small = 7.5  # Уменьшенная высота ячейки подвала
        hs_big = 13  # Увеличенная высота ячейки подвала
        x1 = self.morfostvor.x[0]
        x2 = self.morfostvor.x[-1]

        def __draw_borders(x1, x2, y_top, y_bot):
            # Верхняя граница
            self.ax_bottom_overlay.plot(
                (x1, x2),
                (y_top, y_top),
                color=config.COLOR["border"],
                linewidth=config.LINE_WIDTH["profile_bottom"],
                linestyle="solid",
            )

            # Нижняя граница
            self.ax_bottom_overlay.plot(
                (x1, x2),
                (y_bot, y_bot),
                color=config.COLOR["border"],
                linewidth=config.LINE_WIDTH["profile_bottom"],
                linestyle="solid",
            )

        def __draw_label(x2, y_mid, label):
            self.ax_bottom_overlay.text(
                x2,
                y_mid,
                "   " + label,
                color=config.COLOR["bottom_text_secondary"],
                fontsize=config.FONT_SIZE["bottom_description"],
                horizontalalignment="left",
                verticalalignment="center",
            )

        def __draw_sectors(morfostvor, parameter, y_mid, y_bot, y_top, float_precision=2):
            i = 0
            # Цикл по участкам
            for sector in morfostvor.sectors:
                x = morfostvor.x[sector.start_point]
                x1 = morfostvor.x[sector.end_point]

                x_mid = x1 - ((x1 - x) / 2)
                # Подписи коэффициентов шероховатости по участкам
                value = getattr(sector, parameter)
                if value is np.nan:
                    value = 0
                try:
                    self.ax_bottom.text(
                        x_mid,
                        y_mid,
                        f"{value:.{float_precision}f}",
                        color=config.COLOR["bottom_text"],
                        fontsize=config.FONT_SIZE["bottom_main"],
                        verticalalignment="center",
                        horizontalalignment="center",
                    )

                except ValueError:
                    raise (
                        "\nОшибка в указании параметров участков (коэффициент шероховатости \
                        или разделение на участки). Проверить данные."
                    )

                # Разделители коэффициентов шероховатости
                # Левая граница
                self.ax_bottom.plot(
                    (x, x),
                    (y_bot, y_top),
                    color=config.COLOR["border"],
                    linewidth=config.LINE_WIDTH["profile_footer_divider"],
                    linestyle="solid",
                    alpha=config.TRANSPARENCY["profile_footer_divider"],
                )

                # Отрисовка правой границы на последнем участке
                if i == len(morfostvor.sectors) - 1:
                    # Правая граница
                    self.ax_bottom.plot(
                        (x1, x1),
                        (y_bot, y_top),
                        color=config.COLOR["border"],
                        linewidth=config.LINE_WIDTH["profile_footer_divider"],
                        linestyle="solid",
                        alpha=config.TRANSPARENCY["profile_footer_divider"],
                    )
                i += 1

        def setup_box():
            y_top = self._footer_y

            # Технический разделитель (для увеличения размера границ)
            self.ax_bottom_overlay.plot((x1, x2), (y_top, y_top), alpha=0, color="red")

            self.ax_bottom.plot((x1, x1), (0, y_top), alpha=0, color="red")

        def draw_pk():
            """Отрисовывает нижнюю границу для ПК в подвале,
            сами значения ПК отрисовываются отдельно
            """
            y_bot = self._footer_y
            y_top = self._footer_y + hs
            self._footer_y = y_top
            y_mid = y_top - ((y_top - y_bot) / 2)
            x2 = self.morfostvor.x[-1]

            # Подпись ячейки
            label = "Пикеты"
            __draw_label(x2, y_mid, label)

            self.ax_bottom_overlay.plot(
                (x1, x2),
                (y_bot, y_bot),
                color=config.COLOR["border"],
                linewidth=config.LINE_WIDTH["profile_footer_divider"],
                linestyle="solid",
                alpha=config.TRANSPARENCY["profile_footer_divider"],
            )

        def draw_h():
            hs = hs_big
            y_bot = self._footer_y
            y_top = self._footer_y + hs
            self._footer_y = y_top
            y_mid = y_top - ((y_top - y_bot) / 2)
            x2 = self.morfostvor.x[-1]
            # Длина вертикальной засечки
            divider_length = 1

            # Подпись ячейки
            label = "Отм. земли"
            __draw_label(x2, y_mid, label)
            __draw_borders(x1, x2, y_top, y_bot)

            # Цикл по всем точкам
            for i in range(len(self.morfostvor.x)):
                x = self.morfostvor.x[i]
                y = self.morfostvor.y[i]

                # Подписи отметок
                self.ax_bottom.text(
                    x,
                    y_mid,
                    f"{y:.2f}",
                    color=config.COLOR["bottom_text"],
                    fontsize=config.FONT_SIZE["bottom_small"],
                    verticalalignment="center",
                    horizontalalignment="center",
                    rotation="vertical",
                )

                # засечка низ
                self.ax_bottom.plot(
                    (x, x),
                    (y_top, y_top - divider_length),
                    color=config.COLOR["border"],
                    linewidth=config.LINE_WIDTH["profile_footer_divider"],
                    linestyle="solid",
                    alpha=config.TRANSPARENCY["profile_footer_divider"],
                )

                # засечки
                self.ax_bottom.plot(
                    (x, x),
                    (y_bot, y_bot + divider_length),
                    color=config.COLOR["border"],
                    linewidth=config.LINE_WIDTH["profile_footer_divider"],
                    linestyle="solid",
                    alpha=config.TRANSPARENCY["profile_footer_divider"],
                )

            self.footers_num += 1

        def draw_dist():
            y_bot = self._footer_y
            y_top = self._footer_y + hs
            self._footer_y = y_top
            y_mid = y_top - ((y_top - y_bot) / 2)

            x1 = self.morfostvor.x[0]
            x2 = self.morfostvor.x[-1]

            # Подпись ячейки
            label = "Расстояние"
            __draw_borders(x1, x2, y_top, y_bot)
            __draw_label(x2, y_mid, label)

            # Цикл по всем точкам
            for i in range(len(self.morfostvor.x)):
                x = self.morfostvor.x[i]

                # Разделители расстояний между точками
                self.ax_bottom.plot(
                    (x, x),
                    (y_bot, y_top),
                    color=config.COLOR["border"],
                    linewidth=config.LINE_WIDTH["profile_footer_divider"],
                    linestyle="solid",
                    alpha=config.TRANSPARENCY["profile_footer_divider"],
                )

                # Подписи расстояний между точками
                if i < len(self.morfostvor.x) - 1:
                    x1_ = self.morfostvor.x[i + 1]
                    # Подписи расстояний между точками
                    self.ax_bottom.text(
                        (x + x1_) / 2,
                        y_mid,
                        f"{round(x1_ - x):d}",
                        color=config.COLOR["bottom_text"],
                        fontsize=config.FONT_SIZE["bottom_main"],
                        verticalalignment="center",
                        horizontalalignment="center",
                    )

            self.footers_num += 1

        def draw_rough():
            hs = hs_small
            y_bot = self._footer_y
            y_top = self._footer_y + hs
            self._footer_y = y_top
            y_mid = y_top - ((y_top - y_bot) / 2)

            x1 = self.morfostvor.x[0]
            x2 = self.morfostvor.x[-1]

            label = "Коэфф. n"
            __draw_borders(x1, x2, y_top, y_bot)
            __draw_label(x2, y_mid, label)
            __draw_sectors(self.morfostvor, "roughness", y_mid, y_bot, y_top, float_precision=3)
            self.footers_num += 1

        def draw_depth():
            hs = hs_small
            y_bot = self._footer_y
            y_top = self._footer_y + hs
            self._footer_y = y_top
            y_mid = y_top - ((y_top - y_bot) / 2)

            x1 = self.morfostvor.x[0]
            x2 = self.morfostvor.x[-1]

            label = "$H_{ср}$ при РУВВ"
            __draw_borders(x1, x2, y_top, y_bot)
            __draw_label(x2, y_mid, label)
            __draw_sectors(self.morfostvor, "depth", y_mid, y_bot, y_top)
            self.footers_num += 1

        def draw_speed():
            hs = hs_small
            y_bot = self._footer_y
            y_top = self._footer_y + hs
            self._footer_y = y_top
            y_mid = y_top - ((y_top - y_bot) / 2)

            x1 = self.morfostvor.x[0]
            x2 = self.morfostvor.x[-1]

            label = "$V_{ср}$ при РУВВ"
            __draw_borders(x1, x2, y_top, y_bot)
            __draw_label(x2, y_mid, label)
            __draw_sectors(self.morfostvor, "speed", y_mid, y_bot, y_top)
            self.footers_num += 1

        def draw_area():
            hs = hs_small
            y_bot = self._footer_y
            y_top = self._footer_y + hs
            self._footer_y = y_top
            y_mid = y_top - ((y_top - y_bot) / 2)

            x1 = self.morfostvor.x[0]
            x2 = self.morfostvor.x[-1]

            label = "$F$ при РУВВ"
            __draw_borders(x1, x2, y_top, y_bot)
            __draw_label(x2, y_mid, label)
            __draw_sectors(self.morfostvor, "area", y_mid, y_bot, y_top)
            self.footers_num += 1

        def draw_consumption():
            hs = hs_small
            y_bot = self._footer_y
            y_top = self._footer_y + hs
            self._footer_y = y_top
            y_mid = y_top - ((y_top - y_bot) / 2)

            x1 = self.morfostvor.x[0]
            x2 = self.morfostvor.x[-1]

            label = "$Q$ при РУВВ"
            __draw_borders(x1, x2, y_top, y_bot)
            __draw_label(x2, y_mid, label)
            __draw_sectors(self.morfostvor, "consumption", y_mid, y_bot, y_top)
            self.footers_num += 1

        def draw_situation():
            y_bot = self._footer_y
            y_top = self._footer_y + hs
            self._footer_y = y_top
            y_mid = y_top - ((y_top - y_bot) / 2)

            x1 = self.morfostvor.x[0]
            x2 = self.morfostvor.x[-1]

            # Подпись ряда
            label = "Ситуация"
            __draw_label(x2, y_mid, label)
            __draw_borders(x1, x2, y_top, y_bot)

            for sector in self.morfostvor.situation:
                x1 = self.morfostvor.x[sector.start_point]
                x2 = self.morfostvor.x[sector.end_point]
                x_mid = x2 - ((x2 - x1) / 2)

                # Отрисовка прямоугольника с заливкой
                if config.SITUATION_COLORS:
                    self.ax_bottom.add_patch(
                        Rectangle(
                            (x1, y_bot),
                            (x2 - x1),
                            hs,
                            facecolor=sector.get_color(),
                            fill=True,
                        )
                    )

                # Определение толщины и типа вертикальных линий
                if sector.type == "УВ":
                    linestyle = "solid"
                else:
                    linestyle = "--"

                linewidth = config.LINE_WIDTH["profile_footer_divider"]
                alpha = config.TRANSPARENCY["profile_footer_divider_situation"]

                # Подпись в ситуации
                self.ax_bottom.text(
                    x_mid,
                    y_mid,
                    f"{sector.type}",
                    style="italic",
                    color=config.COLOR["bottom_text"],
                    fontsize=config.FONT_SIZE["bottom_medium"],
                    verticalalignment="center",
                    horizontalalignment="center",
                )

                # Левая граница
                self.ax_bottom.plot(
                    (x1, x1),
                    (y_bot, y_top),
                    color=config.COLOR["border"],
                    linewidth=linewidth,
                    linestyle=linestyle,
                    alpha=alpha,
                )

                # Правая граница
                self.ax_bottom.plot(
                    (x2, x2),
                    (y_bot, y_top),
                    color=config.COLOR["border"],
                    linewidth=linewidth,
                    linestyle=linestyle,
                    alpha=alpha,
                )

            # Отрисовка границ специальными линиями
            for border in self.morfostvor.situation_borders:
                xb1 = self.morfostvor.x[border.point]
                n = 5  # Количество маркеров
                xb = self.morfostvor.x[border.point] * np.ones(n)
                yb = np.linspace(y_bot, y_top - 1.5, n)

                # Определяем сторону бровки и выбираем тип маркера
                if border.type == "бровка левая":
                    linesymbol = 9  # Маркер |>
                elif border.type == "бровка правая":
                    linesymbol = 8  # Маркер <|
                else:
                    if border.id % 2 == 0:
                        linesymbol = 8
                    else:
                        linesymbol = 9

                # Маркеры
                self.ax_bottom.plot(
                    xb,
                    yb,
                    linestyle=linestyle,
                    linewidth=linewidth,
                    marker=linesymbol,
                    color=(0.0, 0.0, 0, 0),
                    ms=8,
                    mfc=(0.0, 0.0, 0, 1),
                    mec=(0, 0, 0, 1),
                    clip_on=True,
                )

                # Линия
                self.ax_bottom.plot((xb1, xb1), (y_bot, y_top), ls="solid", lw=2, color=(0.0, 0.0, 0, 1))
            self.footers_num += 1

        draw_situation()
        draw_consumption()
        draw_depth()
        draw_speed()
        # draw_area()
        draw_rough()
        draw_dist()
        draw_h()
        draw_pk()
        setup_box()

    def draw_sectors(self):
        """
        Отрисовка различной информации связанной с участками профиля.

        :param fill: [bool] - заливка полигонов участков на профиле соответствующими цветами
        :param bottom: [bool] - заливка линии дна соответствующими участкам цветами
        :param label: [bool] - отрисовка названий участков,
        их длин и стрелок обозначающих границы участков
        :return: Отрисовка графической информации по участкам профиля на графике ax_profile.
        """

        h_max = np.floor(max(self.morfostvor.y)) + 1

        for sector in self.morfostvor.sectors:
            points = []

            for i in range(len(sector.coord[0])):
                points.append((sector.coord[0][i], sector.coord[1][i]))

            points.insert(0, (sector.coord[0][0], h_max))
            points.append((sector.coord[0][-1], h_max))

            polygon = matplotlib.patches.Polygon(points, alpha=0.04, linestyle="--", label=sector.name)
            polygon.set_color(sector.color)

            # Подписи названий и длин участков со стрелками
            if config.PROFILE_SECTOR_LABEL:
                p0 = 1
                p1 = 2
                p3 = 3

                # Расчёт середины участка (для центровки текста)
                cent_x = sector.coord[0][-1] - ((sector.coord[0][-1] - sector.coord[0][0]) / 2)

                # Вывод ширины участка
                self.ax_top.text(
                    cent_x,
                    p1,
                    f"{round(sector.length):d} м",
                    color=config.COLOR["sector_text"],
                    verticalalignment="center",
                    horizontalalignment="center",
                    bbox={
                        "facecolor": "white",
                        "edgecolor": "white",
                        "alpha": 1,
                        "pad": 2.5,
                    },
                )

                self.ax_top.text(
                    cent_x,
                    6,
                    sector.name,
                    color=config.COLOR["sector_text"],
                    verticalalignment="center",
                    horizontalalignment="center",
                )

                # Вывод разделителя участков профиля
                self.ax_top.plot(
                    [sector.coord[0][0], sector.coord[0][0]],
                    [p0, p3],
                    color=config.COLOR["sector_line"],
                    linestyle="-",
                    linewidth=config.LINE_WIDTH["sector_line"],
                )  # Горизонтальная слева

                self.ax_top.plot(
                    [sector.coord[0][-1], sector.coord[0][-1]],
                    [p0, p3],
                    color=config.COLOR["sector_line"],
                    linestyle="-",
                    linewidth=config.LINE_WIDTH["sector_line"],
                )  # Горизонтальная справа

                self.ax_top.plot(
                    [sector.coord[0][0], cent_x],
                    [p1, p1],
                    color=config.COLOR["sector_line"],
                    linestyle="-",
                    linewidth=config.LINE_WIDTH["sector_line"],
                )  # Вертикальная слева

                self.ax_top.plot(
                    [cent_x, sector.coord[0][-1]],
                    [p1, p1],
                    color=config.COLOR["sector_line"],
                    linestyle="-",
                    linewidth=config.LINE_WIDTH["sector_line"],
                )  # Вертикальная справа

            # Заливка на профиле участков
            if config.PROFILE_SECTOR_FILL:
                self.ax.add_patch(polygon)

            # Цвет линии дна по участкам
            if config.PROFILE_SECTOR_BOTTOM_LINE:
                self.ax.plot(
                    sector.coord[0],
                    sector.coord[1],
                    linewidth=config.LINE_WIDTH["profile_bottom"],
                    linestyle="solid",
                    color=sector.color,
                    zorder=15,
                )

    def set_style(self):
        # Устанавливаем заголовки графиков
        if config.GRAPHICS_TITLES:
            self.ax.set_title(
                self.morfostvor.title,
                color=config.COLOR["title_text"],
                fontsize=config.FONT_SIZE["title"],
                y=1.1,
            )

        self.ax.set_ylim(self._y_lim)

        # Настраиваем границы и толщину линий границ
        self.ax.spines["top"].set_visible(False)
        self.ax.spines["right"].set_visible(False)
        self.ax.spines["left"].set_linewidth(config.LINE_WIDTH["ax_border"])
        self.ax.spines["bottom"].set_linewidth(config.LINE_WIDTH["ax_border"])

        self.ax_bottom.spines["top"].set_visible(False)
        self.ax_bottom.spines["right"].set_linewidth(config.LINE_WIDTH["ax_border"])
        self.ax_bottom.spines["left"].set_linewidth(config.LINE_WIDTH["ax_border"])
        self.ax_bottom.spines["bottom"].set_linewidth(config.LINE_WIDTH["ax_border"])

        # Устанавливаем отступы в графиках
        self.ax.margins(0.025)
        self.ax_top.margins(0.025)
        self.ax_bottom.margins(0.025, 0)
        self.ax_bottom_overlay.margins(0)

        # Устанавливаем прозрачность заливки фона
        self.ax_top.patch.set_alpha(0)
        self.ax_bottom.patch.set_alpha(0)
        self.ax_bottom_overlay.patch.set_alpha(0)

        # Включаем отображение сетки
        self.ax.grid(True, which="both")

        # Включаем отображение второстепенных засечек на осях
        self.ax.minorticks_on()

        # Устанавливаем параметры засечек на основных осях
        self.ax.tick_params(
            which="major",
            direction="out",
            width=2,
            length=5,
            pad=config.PADDING["ax_profile_tick_labels"],
            labelcolor=config.COLOR["ax_label_text"],
            labelsize=config.FONT_SIZE["ax_major"],
        )

        self.ax.tick_params(
            which="minor",
            direction="out",
            width=1.5,
            length=3,
            pad=config.PADDING["ax_profile_tick_labels"],
            labelcolor=config.COLOR["ax_label_text"],
            labelsize=config.FONT_SIZE["ax_minor"],
        )

        # Отключаем засечки и подписи на осях вспомогательных графиков
        self.ax_bottom.set_xticks([])
        self.ax_bottom.set_yticks([])
        self.ax_bottom_overlay.set_xticks([])
        self.ax_bottom_overlay.set_yticks([])
        self.ax_top.set_xticks([])
        self.ax_top.set_yticks([])

        def format_picket(x, pos):
            """Функция для форматирования значений оси x в пикетаж."""
            return get_pk(x)

        # Устанавливаем параметры подписей осей
        self.ax.set_ylabel(
            f"H, м {config.ALTITUDE_SYSTEM}",
            color=config.COLOR["ax_label_text"],
            fontsize=config.FONT_SIZE["ax_label"],
            fontstyle="italic",
            rotation="horizontal",
        )

        # Настраиваем вывод значений оси x в виде пикетажа
        self.ax.xaxis.set_major_formatter(matplotlib.ticker.FuncFormatter(format_picket))

        self.ax.yaxis.set_label_coords(-0.03, 1.03)

        # Устанавливает параметры вывода значений осей
        self.ax.yaxis.set_major_formatter(matplotlib.ticker.FormatStrFormatter("%.10g"))

        # Настройка параметров отображение сетки
        self.ax.grid(
            which="major",
            color=config.COLOR["ax_grid"],
            linestyle=":",
            linewidth=1,
            alpha=0.9,
            zorder=0,
        )

        self.ax.grid(
            which="minor",
            color=config.COLOR["ax_grid_sub"],
            linestyle=":",
            linewidth=1,
            alpha=0.9,
            zorder=0,
        )

        self.ax.set_axisbelow(True)

        # Установка параметров полей графика
        self.fig.subplots_adjust(left=0.065, bottom=0.02, right=0.88, top=0.9)

    def draw_profile_point_lines(self):
        """
        Отрисовка вертикальных линий от точек до подвала.

        """
        for i in range(len(self.morfostvor.x)):
            self.ax.plot(
                (self.morfostvor.x[i], self.morfostvor.x[i]),
                (self.morfostvor.y[i], self._y_lim[0]),
                color=config.COLOR["profile_point_line"],
                linewidth=config.LINE_WIDTH["profile_point_line"],
                linestyle="solid",
                zorder=1,
            )

    def draw_erosion_limit(self, h, x1=None, x2=None, x3=None, x4=None, text="▼$H_{{разм.}} = {h:.2f}$"):
        """Функция отрисовки линии предельного профиля размыва.

        Arguments:
            h {[float]} -- Отметка линии предельного размыва

        Keyword Arguments:
            x1 {[float]} -- Координата начала линии (default: {None})
            x2 {[float]} -- Координата конца линии (default: {None})
            x3 {[float]} -- Координата начала линии профиля по поверхности  (default: {None})
            x4 {[float]} -- Координата конца линии профиля по поверхности (default: {None})
            text {[string]} -- Текст подписи линии (default: {'▼$H_{{разм.}} = {h:.2f}$'})
        """
        if config.PROFILE_EROSION_LIMIT and not isinstance(self.morfostvor.erosion_limit, str):
            # Ограничение линии предельного размыва
            # по всему профилю если параметр config.PROFILE_EROSION_LIMIT_FULL = true
            if config.PROFILE_EROSION_LIMIT_FULL:
                x1 = min(self.morfostvor.x)
                x2 = max(self.morfostvor.x)
            # Если координаты начала и конца линии не заданы, устанавливаем по границе профиля
            # если есть участки 'Левая пойма', 'Правая пойма' задаем границы линии по участкам
            else:
                if x1 is None:
                    x1 = min(self.morfostvor.x)
                    for sector in self.morfostvor.sectors:
                        if sector.name == "Левая пойма":
                            x1 = sector.coord[0][-1]
                if x2 is None:
                    x2 = max(self.morfostvor.x)
                    for sector in self.morfostvor.sectors:
                        if sector.name == "Правая пойма":
                            x2 = sector.coord[0][0]

            # Подпись текста
            erosion_limit_text = self.ax.text(
                x2 - 1,
                h + 0.01,
                text.format(h=h),
                color=config.COLOR["erosion_limit_text"],
                fontsize=config.FONT_SIZE["erosion_limit"],
                weight="bold",
                zorder=20,
            )
            # Обводка текста
            erosion_limit_text.set_path_effects(
                [
                    path_effects.Stroke(linewidth=3, foreground="white", alpha=0.95),
                    path_effects.Normal(),
                ]
            )

            y3 = None
            y4 = None

            # Функция интерполяции координат профиля
            f = interpolate.interp1d(self.morfostvor.x, self.morfostvor.y)

            # Интерполяция отметок высоты по x и исключение для 0 пикета
            if x3:
                y3 = f(float(x3))
            elif x3 == 0:
                y3 = self.morfostvor.y[0]

            if x4:
                y4 = f(float(x4))

            # Отрисовка линии предельного размыва
            self.ax.plot(
                [x3, x1, x2, x4],
                [y3, h, h, y4],
                color=config.COLOR["erosion_limit_line"],
                linestyle="--",
                linewidth=config.LINE_WIDTH["erosion_limit_line"],
            )
            # Добавляем в список границ отметку
            self._y_limits.append(h)
            self._update_limit()

    def draw_top_limit(self, h, x1=None, x2=None, text="{}\nH = {:.2f}"):
        # Если координаты начала и конца линии не заданы, устанавливаем по границе профиля
        # если есть участки 'Левая пойма', 'Правая пойма' задаем границы линии по участкам
        if x1 is None:
            x1 = min(self.morfostvor.x)
            for sector in self.morfostvor.sectors:
                if sector.name == "Левая пойма":
                    x1 = sector.coord[0][-1]
        if x2 is None:
            x2 = max(self.morfostvor.x)
            for sector in self.morfostvor.sectors:
                if sector.name == "Правая пойма":
                    x2 = sector.coord[0][0]
        y_step = self.ax.get_yticks()[1] - self.ax.get_yticks()[0]
        cent_x = x2 - ((x2 - x1) / 2)

        top_limit_text = self.ax.text(
            cent_x,
            h + (y_step * 0.2),
            f"{self.morfostvor.top_limit_description}\nH = {h:.2f}",
            color=config.COLOR["top_limit_text"],
            fontsize=config.FONT_SIZE["top_limit"],
            weight="bold",
            horizontalalignment="center",
            verticalalignment="center",
            zorder=20,
        )

        self.ax.plot(
            [x1, x2],
            [h, h],
            color=config.COLOR["top_limit_line"],
            linestyle="-.",
            linewidth=config.LINE_WIDTH["top_limit_line"],
        )

        self._y_limits.append(h)
        self._update_limit()

    def draw_waterline(
        self,
        h,
        color=config.COLOR["water_line"],
        linestyle="--",
        linewidth=config.LINE_WIDTH["water_line"],
    ):
        """
        Функция отрисовки уреза воды по границам водного объекта.

        :param water: Исходный водный объект, содержащий координаты границ воды.
        :return: урез на графике профиля (ax_profile).
        """

        def draw_line(self, water):
            for segment in water.segments:
                self.ax.plot(
                    [segment[0][0], segment[0][-1]],
                    [segment[1][0], segment[1][-1]],
                    color=color,
                    linestyle=linestyle,
                    linewidth=linewidth,
                    zorder=5,
                )

                # Заливка урезов в русле
                if config.PROFILE_WATER_FILL:
                    self.ax.fill(
                        segment[0],
                        segment[1],
                        facecolor=config.COLOR["water_fill"],
                        alpha=0.2,
                    )

        if config.OVERFLOW:
            min_sector = self.morfostvor.get_min_sector()
            # Исходные сектора для расчёта (сектор, содержащий минимальную отметку)
            calc_sectors = [min_sector[0]]

            for i in calc_sectors:
                waters, sectors = get_water_sections(self.morfostvor, h, config.OVERFLOW)
                for water in waters:
                    draw_line(self, water)

        else:
            waters, sectors = get_water_sections(self.morfostvor, h, config.OVERFLOW)
            for water in waters:
                draw_line(self, water)

        self._update_limit()
        self.set_style()

    def draw_levels_on_profile(self, levels):
        """
        Функция отрисовки полученных расчётных уровней воды на поперечном профиле.

        :param levels: DataFrame содержащий столбцы P, Q, H
        :return:
        """
        label = []

        ylim = self.ax.get_ylim()
        xlim = self.ax.get_xlim()
        y_step = (ylim[1] - ylim[0]) / 100
        x_step = (xlim[1] - xlim[0]) / 100
        prev_y1 = None
        prev_x0 = None

        def insert_water_levels_label(self, x, y, padding):
            try:
                # Если обеспеченность записана цифрами
                waterline_text = self.ax.text(
                    x,
                    y,
                    f"$P_{{{row['P']:2g}\\%}} = {row['H']:.2f}$",
                    color=config.COLOR["water_level_text"],
                    fontsize=config.FONT_SIZE["water_level"],
                    zorder=20,
                )
                waterline_text.set_path_effects(
                    [
                        path_effects.Stroke(linewidth=2.5, foreground="white", alpha=0.55),
                        path_effects.Normal(),
                    ]
                )
            except ValueError:
                # Если обеспеченность записана строкой
                waterline_text = self.ax.text(
                    x,
                    y,
                    f"${row['P']} = {row['H']:.2f}$",
                    color=config.COLOR["water_level_text"],
                    fontsize=config.FONT_SIZE["water_level"],
                    zorder=20,
                )

                waterline_text.set_path_effects(
                    [
                        path_effects.Stroke(linewidth=2.5, foreground="white", alpha=0.55),
                        path_effects.Normal(),
                    ]
                )

        # Сортируем по уровняем и проходим по каждому уровню
        levels_sorted = levels.sort_values(by="H", ascending=False)
        for index, row in levels_sorted.iterrows():
            # Отрисовка уреза
            water_level = row["H"]

            self.draw_waterline(water_level)

            if config.PROFILE_LEVELS_TITLE:
                # Подпись уровня воды на профиле
                water = WaterSection(self.morfostvor.x, self.morfostvor.y, water_level)
                try:
                    water = WaterSection(self.morfostvor.x, self.morfostvor.y, water_level)
                except:
                    print("Ошибка! При отрисовке расчётных уровней на профиле. \n")

                padding = 0.01
                x = water.water_section_x[0] + 2 * padding
                y = water_level + padding

                if config.PROFILE_LEVELS_TABLE_LINES is False:
                    insert_water_levels_label(self, x, y, padding)

            # Подпись уровня воды в таблице справа
            try:
                if self.morfostvor.levels_result["H"][self.morfostvor.design_water_level_index] == water_level:
                    label.append(
                        f"$\\mathbf{{ P_{{ {row['P']:2g}\\% }} = {water_level:.2f}\\ м\\ {config.ALTITUDE_SYSTEM} }}$\n"
                    )
                else:
                    label.append(f"$P_{{{row['P']:2g}\\%}} = {water_level:.2f}$ м {config.ALTITUDE_SYSTEM}\n")
            except ValueError:
                if self.morfostvor.levels_result["H"][self.morfostvor.design_water_level_index] == water_level:
                    label.append(f"$\\mathbf{{ {row['P']} = {water_level:.2f}\\ м\\ {config.ALTITUDE_SYSTEM} }}$\n")
                else:
                    label.append(f"${row['P']} = {water_level:.2f}$ м {config.ALTITUDE_SYSTEM}\n")

            # Вывод линий сносок от уровней воды к таблице
            if config.PROFILE_LEVELS_TABLE_LINES:
                # Определяем параметры минимального сектора
                # для корректного размещения выносок подписей уровней воды
                min_sector = self.morfostvor.get_min_sector()[1]
                min_sec_x = min_sector.coord[0]
                min_sec_y = min_sector.coord[1]
                min_sec_start_point = min_sector.coord[0][min_sector.coord[1].index(min(min_sector.coord[1]))]
                # Сечение воды для корректного размещения выносок подписей уровней воды
                __water = WaterSection(min_sec_x, min_sec_y, water_level, True, min_sec_start_point)

                # Устанавливаем координаты по умолчанию
                y0 = water_level
                y1 = y0 + y_step * 30
                x0 = __water.water_section_x[0] + (x_step * (index + 2))

                # Проверяем чтобы отметки урезов не выходили за пределы графика
                if y1 > max(self.ax.get_ylim()):
                    y1 = max(self.ax.get_ylim()) - y_step * 4

                # Проверяем координаты на пересечение
                if prev_x0:
                    x0 = prev_x0 + (x_step * 1)

                # если координата x0 выходит за пределы сечения воды слева
                if x0 < min(__water.water_section_x):
                    x0 = min(__water.water_section_x) + x_step

                # если координата x0 выходит за пределы сечения воды справа
                if x0 > max(__water.water_section_x):
                    x0 = max(__water.water_section_x) - x_step

                # контрольная проверка
                if x0 < min(__water.water_section_x) or x0 > max(__water.water_section_x):
                    x0 = min(__water.water_section_x)

                x1 = x0
                if prev_x0 and x1 < prev_x0:
                    x1 = prev_x0 + x_step * 2

                # Множитель для расчета длины выноски аннотации
                if water_level < 100:
                    wl_multiplier = 7.3
                elif 100 <= water_level < 1000:
                    wl_multiplier = 8.3
                else:
                    wl_multiplier = 10

                x2 = x1 + x_step * wl_multiplier

                if prev_y1:
                    y1 = prev_y1 - y_step * 4

                # Устанавливаем параметры отображения линий сносок
                color = config.COLOR["water_reference_line"]
                linestyle = "-"
                linewidth = config.LINE_WIDTH["water_line"] / 1.75
                alpha = 0.6

                # Рисуем линию аннотации
                self.ax.plot(
                    [x0, x1, x2],
                    [y0, y1, y1],
                    color=color,
                    linestyle=linestyle,
                    linewidth=linewidth,
                    alpha=alpha,
                    marker="o",
                    markevery=[0],
                    markersize=3,
                    zorder=45,
                )
                # Вставляем подписи урезов в аннотацию
                insert_water_levels_label(self, x1, y1 + y_step * 0.9, 0)

                # Устанавливаем предыдущие координаты для следующего цикла
                prev_y1 = y1
                prev_x0 = x0

        if self.morfostvor.waterline and type(self.morfostvor.waterline) is not str:
            label.append(f"\nУВ = {self.morfostvor.waterline:.2f} м {config.ALTITUDE_SYSTEM}\n")

            if self.morfostvor.date:
                label.append(f"({self.morfostvor.date})\n")

        if config.PROFILE_WATER_LEVEL_NOTE:
            if self.morfostvor.waterline == "-" or self.morfostvor.waterline == "":
                label.append("\nПримечание: на\nмомент съёмки\nсток отсутствует\n")

        # Вывод параметров РУВВ в таблицу справа
        if isinstance(
            self.morfostvor.probability[self.morfostvor.design_water_level_index][0],
            (float, int),
        ):
            prob_text = rf" $P_{{{
                text_sanitize(self.morfostvor.probability[self.morfostvor.design_water_level_index][0], suffix='\\%')
            }}}$"
        else:
            prob_text = rf"$ {self.morfostvor.probability[self.morfostvor.design_water_level_index][0]}$"

        prob_el = self.morfostvor.levels_result.iloc[self.morfostvor.design_water_level_index]["H"]

        label.append(f"\nРУВВ = {prob_el:.2f} м {config.ALTITUDE_SYSTEM}")
        try:
            label.append(f"\n(принят по {prob_text})")
        except ValueError:
            label.append(
                f"${self.morfostvor.probability[self.morfostvor.design_water_level_index][0]} = {water_level:.2f}$ м {config.ALTITUDE_SYSTEM}\n"
            )

        # Вывод таблицы уровней с разными обеспеченностями (справа)
        self.ax.annotate(
            "".join(label).rstrip(),
            xy=(1, 1),
            ha="left",
            va="top",
            xycoords="axes fraction",
            size=config.FONT_SIZE["levels_table"],
            color=config.COLOR["levels_table"],
            bbox={"boxstyle": "round", "fc": "white", "ec": "none"},
        )

    def draw_wet_perimeter(self):
        """Функция отрисовки смоченного периметра на графике поперечного профиля"""

        # Проверяем задан ли расчётный шаг в исходных данных
        if isinstance(self.morfostvor.dh, str) or self.morfostvor.dh == 0:
            self.morfostvor.dh = 1
            dh = self.morfostvor.dh
        else:
            dh = self.morfostvor.dh

        # Переводим сантиметры приращения в метры
        dh = dh / 100

        # Исходные сектора для расчёта (сектор, содержащий минимальную отметку)
        min_sector = self.morfostvor.get_min_sector()
        calc_sectors = [min_sector[0]]

        # Уровень воды, с минимальным отступом
        water_level = min(self.morfostvor.y) + dh

        def segment_fill(ax, segment):
            ax.fill(
                segment[0],
                segment[1],
                facecolor="red",
                edgecolor="black",
                fill=False,
                linestyle=":",
                alpha=0.6,
                linewidth=2,
                zorder=10,
            )

        # Цикл расчёта до максимального уровня воды
        while water_level <= self.morfostvor.levels_result["H"].max() + dh * config.PROFILE_WET_PERIMETER_NUM:
            if config.OVERFLOW:
                for i in calc_sectors:
                    waters, sectors = get_water_sections(self.morfostvor, water_level, config.OVERFLOW)
                    for water in waters:
                        # Отрисовка смоченного периметра на профиле на профиле
                        for segment in water.segments:
                            segment_fill(self.ax, segment)
            else:
                waters, sectors = get_water_sections(self.morfostvor, water_level, config.OVERFLOW)
                for water in waters:
                    # Отрисовка смоченного периметра для каждого сегмента
                    for segment in water.segments:
                        segment_fill(self.ax, segment)

            water_level += dh

    def _update_limit(self):
        # Шаг засечек по вертикали
        y_step = self.ax.get_yticks()[1] - self.ax.get_yticks()[0]

        # Минимальное и максимальное значения из списка границ
        min_y = min(self._y_limits)
        max_y = max(self._y_limits)

        # Нижняя граница
        self.bottom_limit = np.ceil(min_y) - y_step
        if self.morfostvor.erosion_limit:
            self.bottom_limit = np.ceil(self.morfostvor.erosion_limit) - y_step
            while (self.morfostvor.erosion_limit - self.bottom_limit) < (y_step / 3):
                self.bottom_limit -= y_step
        else:
            self.bottom_limit = np.ceil(min_y) - y_step

            while (self.morfostvor.ele_min - self.bottom_limit) < (y_step / 3):
                self.bottom_limit -= y_step

        # Верхняя граница
        if self.morfostvor.levels_result["H"].max() > max_y:
            max_y = self.morfostvor.levels_result["H"].max()

        if y_step > 0.5:
            self.top_limit = round(np.floor(max_y) + y_step, 3)
        else:
            self.top_limit = round((max_y // y_step * y_step) + y_step * 2, 3)

        # Устанавливаем границы отображения
        self._y_lim = (self.bottom_limit, self.top_limit)
        self.ax.set_ylim(self._y_lim)
        self.draw_profile_point_lines()
