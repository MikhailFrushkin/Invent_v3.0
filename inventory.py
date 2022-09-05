import sys
import time

from Data.cells import *

import pandas as pd
from loguru import logger


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


def read_file(names: tuple):
    """Запись в бд"""
    file_path = "Data/mydatabase.db"
    if os.path.exists(file_path):
        os.remove(file_path)
    try:
        excel_data_df = pd.read_excel('{}'.format(names[0]),
                                      sheet_name='6.1 Складские лоты', skiprows=13, header=1)
        excel_data_df = excel_data_df.fillna(0)
        dbhandle.connect()
        Cells.create_table()
        Check.create_table()
        try:
            for row in excel_data_df.values:
                if not isinstance(row[12], str):
                    place = row[1]
                    code = row[2]
                    name = row[4]
                    num = int(row[7]) if isinstance(row[7], float) else 0
                    num_reserve = int(row[10]) if isinstance(row[10], float) else 0
                    num_free = int(row[11]) if isinstance(row[11], float) else 0

                    temp = Cells.create(place=place, code=code, name=name, num=num,
                                        num_reserve=num_reserve, num_free=num_free)
                    temp.save()

        except peewee.InternalError as px:
            print(str(px))

        excel_data_df = pd.read_excel('{}'.format(names[1]),
                                      sheet_name='Sheet1', header=0)
        try:
            for row in excel_data_df.values:
                place = row[5]
                code = row[2]
                num = row[7]

                temp = Check.create(place=place, code=code, num=num)
                temp.save()

        except peewee.InternalError as px:
            print(str(px))
        finally:
            dbhandle.close()

    except Exception as ex:
        logger.debug('Ошибка записи в базу: {}'.format(ex))
        exit_error()


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
            for j in Cells.select():
                if i.place == j.place and i.code == j.code:
                    j.num_check = i.num
                elif i.place == j.place and i.code not in art_dict[i.place]:
                    Cells.add_art(i.place, i.code, i.num)
                    art_dict[i.place].append(i.code)
                elif i.place not in place_list:
                    if i.code not in not_dowload[i.place]:
                        Cells.add_art(i.place, i.code, i.num,
                                      name='Не выгружена ячейка в файле 6.1, но есть в файле просчета')
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
            'Доступно': [],
            'Посчитано': [],
            'Разница': []}
    dbhandle.connect()

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
    try:
        df_marks = pd.DataFrame(data)

        writer = pd.ExcelWriter('Результат.xlsx')
        sorted_df = df_marks.sort_values(by='Местоположение')
        print(sorted_df)
        sorted_df.to_excel(writer, sheet_name='Result', index=False, na_rep='NaN')

        workbook = writer.book
        worksheet = writer.sheets['Result']

        cell_format = workbook.add_format()
        cell_format.set_bold()
        cell_format.set_num_format('[Blue]General;[Red]-General;General')

        cell_format2 = workbook.add_format()
        cell_format2.set_align('left')

        cell_format3 = workbook.add_format()
        cell_format3.set_align('center')

        worksheet.set_column('A:B', 18, cell_format2)
        worksheet.set_column('C:C', 80, cell_format2)
        worksheet.set_column('D:H', 12, cell_format3)
        worksheet.set_column('H:H', 12, cell_format)
        writer.save()

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

        writer.save()

    except Exception as ex:
        logger.debug(ex)
    finally:
        dbhandle.close()


def exit_error():
    time.sleep(15)
    exit()


if __name__ == "__main__":
    logger.add(sys.stderr, format="{time} {level} {message}", filter="my_module")
    time_start = datetime.datetime.now()
    read_file(file_name())
    check_data()
    write_exsel()
    print('Время сверки: {} секунд(ы)'.format((datetime.datetime.now() - time_start).total_seconds()))
