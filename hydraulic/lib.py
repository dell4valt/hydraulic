# -*- coding: utf-8 -*-

import sys
import xlrd
import numpy as np
from docx.shared import Cm
from docx import Document
from docx.enum.text import WD_BREAK
from pathlib import Path


def insertPageBreak(Document):
    paragraphs = Document.paragraphs
    run = paragraphs[-1].add_run()
    run.add_break(WD_BREAK.PAGE)


def setLastParagraphStyle(style, Document):
    last_paragraph = Document.paragraphs[-1]
    last_paragraph.style = style


def poly_area(x, y):
    """
    Функция определения площади кривой фигуры.

        :param x: Список координат x
        :param y: Список координат y
    """
    return 0.5 * np.abs(np.dot(x, np.roll(y, 1)) - np.dot(y, np.roll(x, 1)))


def chunk_list(seq, num):
    """
    Функция разбивает заданный список на равные количество списков.

        :param seq: Исходный список
        :param num: Количество разбиваемых списков
    """
    avg = len(seq) / float(num)
    out = []
    last = 0.0

    while last < len(seq):
        out.append(seq[int(last):int(last + avg)])
        last += avg

    return out


def get_xls_sheet_quantity(file_path):
    """
    Функция считываем количество листов в xls файле
        :param file_path: Путь к xls файлу
    """
    try:
        data_file = xlrd.open_workbook(file_path)  # Открываем xls файл
    except FileNotFoundError:
        print('Ошибка! Файл {} не найден. Программа будет завершена.'.format(file_path))
        sys.exit(33)

    quantity = data_file.nsheets
    return quantity


def table_style(table, style, width=False):
    """
    Функция устанавливает стиль параграфов в таблице.

        :param table: Исходная таблица
        :param style: Название устанавливаемого стиля
    """

    for row in table.rows:
        i = 0
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                paragraph.style = style

            if width is not False:
                try:
                    success_width = Cm(width[i])
                    cell.width = success_width
                except IndexError:
                    cell.width = success_width
                i += 1


def write_table(doc, table, param, title='Таблица'):
    """
    Функция записи таблицы в док файл
    (док файл, таблица которую записываем, параметры((заголовки),
    (название столбцов), (ширина столбцов)), заголовок)

        :param doc: Document()
        :param table:
        :param param:
        :param title='Таблица':
    """
    doc.add_paragraph(title, style='Т-название')
    rows = 0  # Количество строк

    # Считаем количество не пустых строк
    for row in table:
        rows += 1
    result = doc.add_table(rows + 1, len(param[0]), style='Table Grid')

    # Устанавливаем заголовки столбцов
    for i in range(len(param[0])):
        result.cell(0, i).text = param[0][i]

    # Устанавливаем ширину столбцов
    for row in result.rows:
        for idx, width in enumerate(param[1]):
            row.cells[idx].width = Cm(width)

    i = 1
    # Записываем не пустые значения в таблицу
    for row in table:
        j = 0
        for element in row:
            # Если есть массив с форматами строк, подставляем форматы
            try:
                var = '{' + '{}'.format(param[2][j]) + '}'
                result.cell(i, j).text = var.format(element)
            except:
                result.cell(i, j).text = str(element)
            j += 1
        i += 1
    table_style(result, 'Т-таблица')
    return result


def insert_summary_QV_tables(stvors, out_filename):
    print('Формируем и вставляем сводные таблицы уровней и скоростей ... ', end='')
    # Подготовка данных для записи результирующей таблицы
    levels_table = []
    speed_table = []
    i = 1
    doc = Document(out_filename)

    insertPageBreak(doc)

    doc.add_paragraph('Сводные таблицы', style='З-приложение-подзаголовок')

    param_levels = (['№', 'Описание', 'Мин. отм', 'УВ'],
                    [0.85, 6, 1.25, 1.25],
                    ['', '', ':.2f', ':.2f'])

    param_speed = (['№', 'Описание', ],
                   [0.85, 8.5, ],
                   ['', '', ])

    title_check = False
    title2_check = False

    for stvor in stvors:
        levels_table.append([i, stvor.title, stvor.ele_min, stvor.waterline])
        speed_table.append([i, stvor.title, ])

        for obsp in stvor.levels_result.values.tolist():
            levels_table[i-1].append(round(obsp[2], 2))

            if title_check is False:
                try:
                    param_levels[0].append('P{:g}%'.format(obsp[0]))
                except ValueError:
                    param_levels[0].append('{}'.format(obsp[0]))
                param_levels[1].append(1.25)
                param_levels[2].append(':.2f')

        for obsp in stvor.levels_result.values.tolist():
            speed_table[i-1].append(round(obsp[3], 3))

            if title2_check is False:
                try:
                    param_speed[0].append('V{:g}%'.format(obsp[0]))
                except ValueError:
                    param_speed[0].append('{}'.format(obsp[0]))

                param_speed[1].append(1.25)
                param_speed[2].append(':.2f')

        title_check = True
        title2_check = True
        i += 1

    # rows = len(levels_table)
    cols = len(levels_table[0])

    #############################################
    # Таблицы расчётных уровней и скоростей воды
    doc.add_paragraph('Таблица — Расчётные уровни {}'.format(
        stvor.strings['type']), style='Т-название')
    lev_table = doc.add_table(2, cols, style='Table Grid')

    doc.add_paragraph('Таблица — Расчётные скорости {}'.format(
        stvor.strings['type']), style='Т-название')
    spd_table = doc.add_table(2, cols, style='Table Grid')

    lev_table.cell(0, 0).merge(lev_table.cell(1, 0)
                               ).text = str(param_levels[0][0])
    lev_table.cell(0, 1).merge(lev_table.cell(1, 1)
                               ).text = str(param_levels[0][1])
    lev_table.cell(0, 2).merge(lev_table.cell(1, 2)
                               ).text = str(param_levels[0][2])
    lev_table.cell(0, 3).merge(lev_table.cell(1, 3)
                               ).text = str(param_levels[0][3])
    lev_table.cell(0, 4).merge(lev_table.cell(
        0, len(param_levels[0]) - 1)).text = 'Уровни воды (м БС), обеспеченностью Р%'

    spd_table.cell(0, 0).merge(spd_table.cell(1, 0)
                               ).text = str(param_levels[0][0])
    spd_table.cell(0, 1).merge(spd_table.cell(1, 1)
                               ).text = str(param_levels[0][1])
    spd_table.cell(0, 2).merge(spd_table.cell(1, 2)
                               ).text = str(param_levels[0][2])
    spd_table.cell(0, 3).merge(spd_table.cell(1, 3)
                               ).text = str(param_levels[0][3])
    spd_table.cell(0, 4).merge(spd_table.cell(
        0, len(param_levels[0]) - 1)).text = 'Скорости воды (м/с), обеспеченностью Р%'

    stvor_num = 1
    # Подписываем вероятности
    for i in range(len(stvors[0].probability)):
        try:
            lev_table.cell(
                1, i + 4).text = '{:g}'.format(stvors[0].probability[i][0])
            spd_table.cell(
                1, i + 4).text = '{:g}'.format(stvors[0].probability[i][0])
        except ValueError:
            lev_table.cell(
                1, i + 4).text = '{}'.format(stvors[0].probability[i][0])
            spd_table.cell(
                1, i + 4).text = '{}'.format(stvors[0].probability[i][0])

    # Заполняем сводные таблицы данными
    for stvor in stvors:
        levels = stvor.levels_result[['P', 'H', 'Q']].values.tolist()
        speed = stvor.levels_result[['P', 'H', 'V']].values.tolist()

        lev_cell = lev_table.add_row().cells
        lev_cell[0].text = str(stvor_num)
        lev_cell[1].text = str(stvor.title)
        lev_cell[2].text = str('{:.2f}'.format(stvor.ele_min))

        spd_cell = spd_table.add_row().cells
        spd_cell[0].text = str(stvor_num)
        spd_cell[1].text = str(stvor.title)
        spd_cell[2].text = str('{:.2f}'.format(stvor.ele_min))

        # Проверка наличия уреза воды и вставка его в таблицу
        if type(stvor.waterline) == float:
            lev_cell[3].text = str('{:.2f}'.format(stvor.waterline))
            spd_cell[3].text = str('{:.2f}'.format(stvor.waterline))
        else:
            lev_cell[3].text = '-'
            spd_cell[3].text = '-'

        for i in range(4, len(levels) + 4):
            try:
                lev_cell[i].text = str('{:.2f}'.format(levels[i - 4][1]))
                spd_cell[i].text = str('{:.2f}'.format(speed[i - 4][2]))
            except:
                print(
                    '\n\nОшибка соответствия обеспеченностей в профилях.\
                     Обеспеченности на всех профилях должны быть одинаковые.')
                print('Сводные таблицы не будут записаны в файл.')
                sys.exit(1)

        stvor_num += 1

        table_style(lev_table, 'Т-таблица',
                    width=[0.85, 6, 1.25, 1.25, 1.25, ])
        table_style(spd_table, 'Т-таблица',
                    width=[0.85, 6, 1.25, 1.25, 1.25, ])

    print('успешно!')
    doc.save(out_filename)


def rmdir(dir):
    directory = Path(str(dir))

    for item in directory.iterdir():
        if item.is_dir():
            rmdir(item)
        else:
            item.unlink()
    directory.rmdir()
