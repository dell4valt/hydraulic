"""Модуль содержит процедуру отвечающую за создания отчета
в формате .docx по заданному экземпляру объекта класса
Hydraulic.Morfostvor.
"""

import os
import sys
from pathlib import Path

import numpy as np
from docx import Document
from pathvalidate import sanitize_filename

import hydraulic.config as config

from hydraulic.lib import rmdir, text_sanitize, get_pk
from report.core import Report
from hydraulic.graph import (GraphProfile, GraphQH, GraphQWVH, GraphFH, GraphQV, GraphQF, GraphQHV, GraphVH)


def generate_morfostvor_report(morfostvor, out_filename, rewrite=False):
    """Процедура создает отчет по заданному морфоствору и сохраняет
    его в указанный .docx файл.

    Args:
        morfostvor (hydraulic.Morfostvor): Экземпляр объекта класса Morfostvor
        out_filename (_type_): путь и названия файла куда будет сохранен отчет,
        на конце должно быть указание расширения .docx
        rewrite (bool, optional): Параметр позволяет включить перезапись файла отчета,
        если выключен то отчет будет добавлен в конец документа. Defaults to False.
    """
    print("\n\nФормируем doc файл: ")

    if rewrite:
        report = Report()
    else:
        report = Report(out_filename)
        # Если документ уже содержит параграфы,
        # то вставляем разрыв страницы перед добавлением нового отчета
        if len(report.doc.paragraphs) > 2:
            report.insert_page_break()

    # Создаем временную папку, и папку для графики если они не существуют
    temp_dir = Path(config.TEMP_DIR_NAME)
    temp_dir.mkdir(parents=True, exist_ok=True)

    # Отрисовка смоченного периметра
    if config.PROFILE_WET_PERIMETER:
        morfostvor.fig_profile.draw_wet_perimeter()

    # Отрисовка верхней границы сооружения
    if morfostvor.top_limit:
        morfostvor.fig_profile.draw_top_limit(
            morfostvor.top_limit, text=morfostvor.top_limit_description
        )

    # Отрисовка границы предельного размыва профиля
    if morfostvor.erosion_limit and len(morfostvor.erosion_limit_coord) == 2:
        morfostvor.fig_profile.draw_erosion_limit(
            morfostvor.erosion_limit,
            morfostvor.erosion_limit_coord[0],
            morfostvor.erosion_limit_coord[1])
    elif morfostvor.erosion_limit and len(morfostvor.erosion_limit_coord) == 4:
        morfostvor.fig_profile.draw_erosion_limit(
            morfostvor.erosion_limit,
            morfostvor.erosion_limit_coord[0],
            morfostvor.erosion_limit_coord[1],
            morfostvor.erosion_limit_coord[2],
            morfostvor.erosion_limit_coord[3])
    elif morfostvor.erosion_limit:
        morfostvor.fig_profile.draw_erosion_limit(morfostvor.erosion_limit)

    # Отрисовка расчетных уровней воды на графике профиля
    morfostvor.fig_profile.draw_levels_on_profile(morfostvor.levels_result)
    morfostvor.fig_profile._update_limit()

    # TODO: сделать отрисовку линий урезов воды по каждому
    # участку УВ из описания ситуации исходного файла
    # Отрисовка урез воды на графике профиля
    if morfostvor.waterline and type(morfostvor.waterline) != str:
        morfostvor.fig_profile.draw_waterline(
            round(morfostvor.waterline, 2), color="blue", linestyle="-"
        )

    if config.GRAPHICS_TITLES_TEXT:
        profile_title = f"{morfostvor.fig_profile.morfostvor.title}"
        qh_title = f"{morfostvor.fig_QH._ax_title_text}"
        qhv_title = f"{morfostvor.fig_QHV._ax_title_text}"
        qv_title = f"{morfostvor.fig_QV._ax_title_text}"
        vh_title = f"{morfostvor.fig_VH._ax_title_text}"
        qf_title = f"{morfostvor.fig_QF._ax_title_text}"
        fh_title = f"{morfostvor.fig_FH._ax_title_text}"
        qwvh_title = f"{morfostvor.fig_QWVH._ax_title_text}"
    else:
        profile_title = ""
        qh_title = ""
        qhv_title = ""
        qv_title = ""
        vh_title = ""
        qf_title = ""
        fh_title = ""
        qwvh_title = ""

    # Вставляем заголовок профиля
    report.add_paragraph(morfostvor.title, style="З-приложение-подзаголовок")

    # Добавляем изображения профиля и гидравлической кривой
    print("    — Вставляем графику (профиль)... ", end="")
    report.insert_mpl_figure(morfostvor.fig_profile.fig, width=16, title=profile_title)
    print("успешно!")

    # Dictionary of curve configurations with their properties
    curves = [
        {"config": "HYDRAULIC_CURVE", "fig": morfostvor.fig_QH.fig, "title": qh_title,
         "message": "Вставляем график (кривая QH)"},
        {"config": "QWVH_CURVE", "fig": morfostvor.fig_QWVH.fig, "title": qwvh_title,
         "message": "Вставляем график (кривая QWVH)"},
        {"config": "HYDRAULIC_AND_SPEED_CURVE", "fig": morfostvor.fig_QHV.fig, "title": qhv_title,
         "message": "Вставляем график (кривая QHV)"},
        {"config": "SPEED_CURVE", "fig": morfostvor.fig_QV.fig, "title": qv_title,
         "message": "Вставляем график кривой скоростей QV"},
        {"config": "SPEED_VH_CURVE", "fig": morfostvor.fig_VH.fig, "title": vh_title,
         "message": "Вставляем график кривой скоростей VH"},
        {"config": "AREA_CURVE", "fig": morfostvor.fig_QF.fig, "title": qf_title,
         "message": "Вставляем график кривой площадей от расхода воды"},
        {"config": "AREA_FH_CURVE", "fig": morfostvor.fig_FH.fig, "title": fh_title,
         "message": "Вставляем график кривой площадей от уровня"}
    ]

    # Insert all configured curves
    for curve in curves:
        if getattr(config, curve["config"]):
            print(f"    — {curve['message']}... ", end="")
            report.insert_mpl_figure(curve["fig"], width=16, title=curve["title"])
            print("успешно!")

    # Вывод таблицы расчётных уровней, скоростей и площадей воды
    print("    — Записываем таблицу уровней, скоростей и площадей воды ... ", end="")
    report.insert_df_to_table(
        morfostvor.levels_result[["P", "Q", "H", "V", "F"]],
        (
            f"Расчётные уровни, скорости и площади "
            f"к заданным расходам {morfostvor.strings['type']}"
        ),
        col_names=(
            "Обеспеченность P, %",
            "Расход Q, м³/сек",
            f"Уровень H, м {config.ALTITUDE_SYSTEM}",
            f"Средняя скорость Vср, м/сек",
            f"Площадь живого сечения F, м²",
        ),
        col_widths=(6, 6, 6, 6, 6),
        col_format=(":g", ":g", ":.2f", ":.2f", ":.2f"),
    )
    print("успешно!")

    # Вывод таблицы участков
    print("    — Записываем таблицу участков ... ", end="")
    # Заменяем пустые значения на прочерк и добавляем номер участка
    df_sectors = morfostvor.sectors_result.replace(np.nan, '-')
    df_sectors.insert(loc=0, column='N', value=df_sectors.index + 1)

    prob_text = text_sanitize(
        morfostvor.probability[morfostvor.design_water_level_index][0],
        num_suffix="% обеспеченности",
    )

    topography_table = morfostvor.get_topography_table()
    topography_table["x"] = topography_table["x"].apply(lambda x: get_pk(x, decimal=True))
    topo_table = report.insert_df_to_table(
        topography_table,
        f"Топографические данные створа",
        col_names=(
            "ПК",
            f"Отметка, м {config.ALTITUDE_SYSTEM}",
            "Участок",
            "Коэффициент шероховатости, n",
            "Уклон I, ‰",
        ),
        col_format=("", ":.2f", "", ":.3f", ":.2f"),
    )

    # Объединение ячеек участков 
    for sector in morfostvor.sectors:
        if sector == morfostvor.sectors[-1]:
            report.merge_table_cells(
                topo_table,
                sector.start_point + 1,
                sector.end_point + 1,
                2,
                2,
                sector.name,
            )
            report.merge_table_cells(
                topo_table,
                sector.start_point + 1,
                sector.end_point + 1,
                3,
                3,
                f"{sector.roughness:.3f}",
            )
            report.merge_table_cells(
                topo_table,
                sector.start_point + 1,
                sector.end_point + 1,
                4,
                4,
                f"{sector.slope:.2f}",
            )
        else:
            report.merge_table_cells(
                topo_table, sector.start_point + 1, sector.end_point, 2, 2, sector.name
            )
            report.merge_table_cells(
                topo_table,
                sector.start_point + 1,
                sector.end_point,
                3,
                3,
                f"{sector.roughness:.3f}",
            )
            report.merge_table_cells(
                topo_table,
                sector.start_point + 1,
                sector.end_point,
                4,
                4,
                f"{sector.slope:.2f}",
            )

    report.insert_df_to_table(
        df_sectors,
        f"Расчётные участки и их параметры",
        col_names=(
            "№",
            "Описание",
            "Уклон i, ‰",
            "Коэффициент шероховатости n",
            "Q при РУВВ, м³/сек",
            f"Hср при РУВВ, м {config.ALTITUDE_SYSTEM}",
            "Vср при РУВВ, м/сек",
            "B при РУВВ, м",
            "F при РУВВ, м²",
        ),
        col_widths=(1.3, 4, 4, 4, 4, 4, 4, 4, 4),
        col_format=(":d", "", ":g", ":.3f", ":.2f", ":.2f", ":.2f", ":.2f", ":.2f"),
        footer_text=(
            "Примечание: Расчетный уровень высоких вод (РУВВ) "
            f"принят по расходу {prob_text}."
        ),
    )
    print("успешно!")

    # Вывод таблицы гидравлической кривой
    print("    — Записываем таблицу кривой расхода воды ... ", end="")

    table = morfostvor.hydraulic_table.reset_index(0).loc["Сумма"].reset_index(drop=True)
    table_round = table.round(3)  # Округляем
    # Убираем столбец с коэффициентами Шези
    table_round = table_round.drop(columns=["Shezi"])

    if config.DOC_TABLE_SHORT:
        # Количество строк в таблице
        table_quant = table_round["УВ"].count()

        # Уменьшаем количество выводимых строк в таблицу
        # чтобы поместилось на один лист
        if table_quant <= 25:
            divider = 1
        elif table_quant > 25 and table_quant <= 50:
            divider = 2
        elif table_quant > 50 and table_quant <= 75:
            divider = 3
        elif table_quant > 75 and table_quant <= 100:
            divider = 4
        elif table_quant > 100 and table_quant <= 125:
            divider = 5
        else:
            divider = 10
    else:
        divider = 1

    # Записываем только чётные элементы таблицы
    table_round = table_round[table_round.index % divider == 0]

    report.insert_df_to_table(
        table_round,
        f"Параметры расчёта кривой расхода {morfostvor.strings['type']}",
        col_names=(
            f"Отм. уровня H, м {config.ALTITUDE_SYSTEM}",
            "Площадь F, м²",
            "Ширина B, м",
            "Средняя глубина Hср, м",
            "Макс. глубина Hмакс, м",
            "Средняя скорость Vср, м/сек",
            "Расход Q, м³/сек",
        ),
        col_widths=(5, 5, 5, 5, 5, 5, 5),
        col_format=(":.2f", ":.3f", ":.3f", ":.3f", ":.3f", ":.3f", ":.3f"),
        footer_text=(
            f"Расчётный шаг: {morfostvor.dH:g} см. "
            f"В таблице приведён каждый {divider}-й результат расчёта."
        ),
    )

    print("успешно!")

    try:
        report.save(out_filename)
    except PermissionError:
        print(
            "\nОшибка! Не удалось сохранить файл. "
            "Проверьте возможность записи файла по указанному пути."
        )
        print("Возможно записываемый файл уже существует и открыт.")
        sys.exit(1)

    # Удаляем временную папку со всем содержимым
    print("    — Удаляем временную папку ... ", end="")
    rmdir(Path(f"{config.TEMP_DIR_NAME}"))
    print("успешно!")


def save_graphic(morfostvor, path):
    # Создаем временную папку, и папку для графики если они не существуют
    temp_dir = Path(config.TEMP_DIR_NAME)
    temp_dir.mkdir(parents=True, exist_ok=True)

    # Проверяем имя файла
    profile_name = sanitize_filename(morfostvor.title)

    # Создаем папку для сохранения отдельных изображений
    picture_dir = Path(
        str(Path(path)) + "/" + config.GRAPHICS_DIR_NAME
    )
    picture_dir.mkdir(parents=True, exist_ok=True)

    # Сохраняем картинки в отдельные файлы в папку graphics
    if config.PROFILE_SAVE_PICTURES:
        morfostvor.fig_profile.fig.savefig(
            Path(f"{picture_dir}/{profile_name}.png", dpi=config.FIG_DPI)
        )
    if config.CURVE_SAVE_PICTURES:
        if config.HYDRAULIC_CURVE:
            morfostvor.fig_QH.fig.savefig(
                Path(f"{picture_dir}/{profile_name}_QH.png", dpi=config.FIG_DPI)
            )
        if config.HYDRAULIC_AND_SPEED_CURVE:
            morfostvor.fig_QHV.fig.savefig(
                Path(f"{picture_dir}/{profile_name}_QHV.png", dpi=config.FIG_DPI)
            )
        if config.SPEED_CURVE:
            morfostvor.fig_QV.fig.savefig(
                Path(f"{picture_dir}/{profile_name}_QV.png", dpi=config.FIG_DPI)
            )
        if config.AREA_CURVE:
            morfostvor.fig_QF.fig.savefig(
                Path(f"{picture_dir}/{profile_name}_QF.png", dpi=config.FIG_DPI)
            )
        if config.QWVH_CURVE:
            morfostvor.fig_QWVH.fig.savefig(
                Path(f"{picture_dir}/{profile_name}_QWVH.png", dpi=config.FIG_DPI)
            )
