import sys
import time
import os
import datetime

import pandas as pd
from loguru import logger

from Data.cells import *


def file_name() -> tuple:
    """нахождение файлов с 6.1 и результата просчета
    :return имена файлов"""

    file_list = os.listdir()
    file_base = 'Нет подходящих файлов'
    file_check = 'Нет подходящих файлов'
    for item in file_list:
        if item.endswith('.xlsx'):
            if item.startswith('6.1'):
                file_base = item
            elif item != 'Результат.xlsx' and item != 'Для импорта в пст(недостача).xlsx':
                file_check = item

    print('\nФайл из 6.1: {}\nФайл проверки: {}'.format(
        file_base, file_check
    ))
    return file_base, file_check


def return_int(num):
    if isinstance(num, float):
        return int(num)
    elif isinstance(num, str):
        return int(num.replace('.0', ''))
    else:
        return num


def read_file(names: tuple):
    """Запись в бд"""
    file_path = "mydatabase.db"
    if os.path.exists(file_path):
        os.remove(file_path)
    try:
        excel_data_df = pd.read_excel('{}'.format(names[0]), skiprows=13, header=1,
                                      usecols=['Местоположение', 'Код \nноменклатуры',
                                               'Описание товара', 'Физические \nзапасы', 'Передано на доставку',
                                               'Продано', 'Зарезерви\nровано',
                                               'Доступно',
                                               'Номер документа'])
        key_code = list(excel_data_df.keys()).index('Код \nноменклатуры')
        key_place = list(excel_data_df.keys()).index('Местоположение')
        key_name = list(excel_data_df.keys()).index('Описание товара')
        key_num = list(excel_data_df.keys()).index('Физические \nзапасы')
        key_num_dost = list(excel_data_df.keys()).index('Передано на доставку')
        key_num_sell = list(excel_data_df.keys()).index('Продано')
        key_num_res = list(excel_data_df.keys()).index('Зарезерви\nровано')
        key_num_free = list(excel_data_df.keys()).index('Доступно')
        key_doc = list(excel_data_df.keys()).index('Номер документа')
        excel_data_df = excel_data_df.fillna(0)
        dbhandle.connect()
        Cells.create_table()
        Check.create_table()
        print('Запись в базу...')
        for row in excel_data_df.values:
            if not isinstance(row[key_doc], str):
                place = row[key_place]
                code = row[key_code]
                name = row[key_name]
                num = return_int(row[key_num])
                num_reserve = return_int(row[key_num_res])
                num_free = return_int(row[key_num_free])
                num_sell = return_int(row[key_num_sell])
                num_dost = return_int(row[key_num_dost])
                Cells.add_art(place=place, code=code, name=name, num=num, num_dost=num_dost,
                              num_sell=num_sell, num_reserve=num_reserve, num_free=num_free)
    except Exception as ex:
        logger.debug('Ошибка записи в базу, нет файла из 6.1(название не начинается на 6.1)\n'
                     'или не хватает столбцов в таблице: {}'.format(ex))
        exit_error()

    try:
        excel_data_df = pd.read_excel('{}'.format(names[1]), header=0,
                                      usecols=['Код номенклатуры', 'Местоположение', 'Количество факт'])
        key_code = list(excel_data_df.keys()).index('Код номенклатуры')
        key_place = list(excel_data_df.keys()).index('Местоположение')
        key_num = list(excel_data_df.keys()).index('Количество факт')
        for row in excel_data_df.values:
            place = row[key_place]
            code = row[key_code]
            num = row[key_num]
            temp = Check.create(place=place, code=code, num=num)
            temp.save()

    except Exception as ex:
        logger.debug('Ошибка записи в базу, нет файла просчета\n'
                     'или не хватает столбцов в таблице: {}'.format(ex))
        exit_error()
    finally:
        dbhandle.close()


def check_data():
    """Проверка расхождений"""
    try:
        dbhandle.connect()
        place_list = list(set(item.place for item in Cells.select()))
        place_list_check = list(set(item.place for item in Check.select()))
        for i in Cells.select():
            if i.place not in place_list_check:
                i.delete_instance()

        art_dict = dict()

        for cell in place_list_check:
            art_list = list(set(item.code for item in Cells.select() if item.place == cell))
            art_dict[cell] = art_list

        not_dowload = dict()
        for i in place_list_check:
            not_dowload[i] = []

        for i in Check.select():
            print('Проверяется {} из {}'.format(i.code, i.place))
            for j in Cells.select():
                if i.place == j.place and i.code == j.code:
                    j.num_check = i.num
                elif i.place == j.place and i.code not in art_dict[i.place]:
                    Cells.add_art(i.place, i.code, i.num)
                    art_dict[i.place].append(i.code)
                elif i.place not in place_list:
                    if i.code not in not_dowload[i.place]:
                        Cells.add_art(i.place, i.code, i.num,
                                      name='Ячейка не выгружена в файле 6.1, но есть в файле просчета')
                        art_dict[i.place].append(i.code)
                        not_dowload[i.place].append(i.code)
                j.delta = j.num_check - j.num
                j.save()
    except Exception as ex:
        logger.debug('Ошибка проверки расхождений: {}'.format(ex))
        exit_error()
    finally:
        dbhandle.close()


def write_exsel():
    """Запись расхождений в Exsel с форматированием таблицы"""
    data = {'Местоположение': [],
            'Артикул': [],
            'Описание товара': [],
            'Физ.запас': [],
            'В резерве': [],
            'Продано': [],
            'Передано на доставку': [],
            'Доступно': [],
            'Посчитано': [],
            'Разница': []}
    dbhandle.connect()
    count_error = 0
    query = Cells.select()
    for i in query:
        data['Местоположение'].append(i.place)
        data['Артикул'].append(i.code)
        data['Описание товара'].append(i.name)
        data['Физ.запас'].append(i.num)
        data['В резерве'].append(i.num_reserve)
        data['Доступно'].append(i.num_free)
        data['Посчитано'].append(i.num_check)
        data['Разница'].append(i.delta)
        data['Продано'].append(i.num_sell)
        data['Передано на доставку'].append(i.num_dost)

        if i.delta != 0:
            count_error += 1

    try:
        df_marks = pd.DataFrame(data)
        len_doc = 1
        if data['Местоположение']:
            len_doc = len(data['Местоположение'])
        writer = pd.ExcelWriter('Результат.xlsx')
        sorted_df = df_marks.sort_values(by='Местоположение')
        sorted_df.to_excel(writer, sheet_name='Сверка', index=False, na_rep='NaN')

        workbook = writer.book
        worksheet = writer.sheets['Сверка']

        cell_format = workbook.add_format()
        cell_format.set_align('center')
        cell_format.set_bold()
        cell_format.set_border(1)
        cell_format.set_num_format('[Blue]General;[Red]-General;General')

        cell_format2 = workbook.add_format({'align': 'left',
                                            'valign': 'vcenter',
                                            'border': 1})

        cell_format3 = workbook.add_format({'align': 'center',
                                            'valign': 'vcenter',
                                            'border': 1})

        worksheet.set_column('A:B', 18, cell_format2)
        worksheet.set_column('C:C', 50, cell_format2)
        worksheet.set_column('D:I', 12, cell_format3)
        worksheet.set_column('G:G', 20, cell_format3)
        worksheet.set_column('J:J', 12, cell_format)

        query_all = Cells.select(Cells.code, fn.SUM(Cells.delta)).group_by(Cells.code)
        data_all_result = {
            'Артикул': [],
            'Общее количество': []
        }
        for i in query_all:
            if i.delta != 0:
                data_all_result['Артикул'].append(i.code)
                data_all_result['Общее количество'].append(i.delta)
        df_marks_all = pd.DataFrame(data_all_result)
        df_marks_all.to_excel(writer, sheet_name='Общий итог', index=False, na_rep='NaN')
        worksheet2 = writer.sheets['Общий итог']
        worksheet2.set_column('B:B', 12, cell_format)
        writer.save()
    except Exception as ex:
        logger.debug(ex)
        exit_error()
        os.remove('mydatabase.db')

    writer = pd.ExcelWriter('Для импорта в пст(недостача).xlsx')
    data_for_import = {
        'Номенклатура': [],
        'Кол-во': [],
        'Со склада': [],
        'С ячейки': [],
        'На БЮ': [],
        'На склад': [],
        'Дата отгрузки': [],
        'Промо': [],
        'С "reason code"': [],
        'На "reason code"': [],
        'С профиля учета': [],
        'На профиль учета': [],
        'В ячейку': [],
        'С сайта': [],
        'На сайт': [],
        'С владельца': [],
        'На владельца': [],
        'Из партии': [],
        'В партию': [],
        'Из ГТД': [],
        'В ГТД': [],
        'С серийного номера': [],
        'На серийный номер': []
    }
    try:
        for i in query:
            if i.delta < 0:
                data_for_import['Номенклатура'].append(i.code)
                data_for_import['Кол-во'].append(-i.delta)
                data_for_import['Со склада'].append(i.place[:7])
                data_for_import['С ячейки'].append(i.place)
                data_for_import['На БЮ'].append(i.place[4:7])
                data_for_import['На склад'].append(i.place[:7])
                data_for_import['Дата отгрузки'].append('')
                data_for_import['Промо'].append('')
                data_for_import['С "reason code"'].append('')
                data_for_import['На "reason code"'].append('')
                data_for_import['С профиля учета'].append('')
                data_for_import['На профиль учета'].append('')
                data_for_import['В ячейку'].append('{}-01-01-0'.format(i.place[:7]))
                data_for_import['С сайта'].append('')
                data_for_import['На сайт'].append('')
                data_for_import['С владельца'].append('')
                data_for_import['На владельца'].append('')
                data_for_import['Из партии'].append('')
                data_for_import['В партию'].append('')
                data_for_import['Из ГТД'].append('')
                data_for_import['В ГТД'].append('')
                data_for_import['С серийного номера'].append('')
                data_for_import['На серийный номер'].append('')

        df_marks_import = pd.DataFrame(data_for_import)
        df_marks_import.to_excel(writer, sheet_name='import', index=False, na_rep='NaN')

        print('Выявленно расхождений: {}'.format(count_error))
        writer.save()
    except Exception as ex:
        logger.debug('Ошибка записи в файл для пст\n {}'.format(ex))
        exit_error()
    finally:
        dbhandle.close()


def exit_error():
    time.sleep(3)
    exit()


if __name__ == "__main__":
    logger.add(sys.stderr, format="{time} {level} {message}", filter="my_module")
    time_start = datetime.datetime.now()
    read_file(file_name())
    check_data()
    write_exsel()
    os.remove('mydatabase.db')
    print('Время сверки: {} секунд(ы)'.format((datetime.datetime.now() - time_start).total_seconds()))
    print('Создан файл с расхождениями: Результат.xlsx')
    print('Создан файл для импорта в перенос: Для импорта в пст(недостача).xlsx')

    exit_error()
