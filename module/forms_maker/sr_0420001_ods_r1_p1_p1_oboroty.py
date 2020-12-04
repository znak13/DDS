"""
0420011 Отчетность об операциях с денежными средствами
некредитных финансовых организаций.
Раздел 1. Операции, совершенные с использованием банковских счетов
некредитной финансовой организации
1.1. Виды и суммы операций, совершенных по банковским счетам
некредитной  финансовой организации Обороты
http://www.cbr.ru/xbrl/nso/uk/rep/2019-12-31/tab/sr_0420001_ods_r1_p1_p1_oboroty
0420011 Отчетность об операци_2
"""
from module.copy_style import make_form
from module.functions import razdel_name_row
from openpyxl.utils import column_index_from_string  # 'B' -> 2


# -----------------------------------------------------------------------
def rows_of_razdel(df_avancor, all_razdels, col=None,
                   row_start=None, row_end=None, rows_under=True):
    """ файл-Аванкор: поиск номеров строк разделов"""
    # row_start, row_end - поиск между этих строк
    # col - колонка, в котолой расположен раздел

    rows_of_title = []
    i = 0
    for row in range(row_start, row_end):
        title = str(df_avancor.loc[row, column_index_from_string(col)])
        if title in all_razdels:
            # если данные ниже названия раздела, то добавляем единицу
            row_data = row + 1 if rows_under else row
            rows_of_title.append({'title_name': title,
                                  'row_razdel_begin': row_data,
                                  'row_razdel_end': None})

            if len(rows_of_title) > 1:
                rows_of_title[i - 1]['row_razdel_end'] = row - 1

            i += 1

    return rows_of_title


# -----------------------------------------------------------------------
def rows_with_data(df_avancor, banks, all_razdels):
    """ Находим номера строк с данными, которые будем копировать"""

    avancoreTitle = '1.1. Виды и суммы операций, совершенных по банковским счетам некредитной \n' \
                    'финансовой организации'
    col = 'B'

    # файл-Аванкор: интервал строк, в пределах которых ищем строки с разделами
    row_start = 32
    row_end = 91

    # Кол-во строк и столбцов в файле-Аванкор
    index_max = df_avancor.shape[0]
    collumn_max = df_avancor.shape[1]

    title_row = razdel_name_row(df_avancor, avancoreTitle,
                                title_col=column_index_from_string(col),
                                row_start=row_start, row_end=row_end)

    # номер строки с названиями валют
    row_currency = title_row[0] + 3

    # номера строк с разделами
    rows_of_title = rows_of_razdel(df_avancor, all_razdels, col='D',
                                   row_start=row_start, row_end=row_end)
    rows_itogi = rows_of_razdel(df_avancor, all_razdels, col='B',
                                row_start=row_start, row_end=row_end, rows_under=False)
    # устанавливаем нижнюю строку последнего раздела перед итогами
    rows_of_title[-1]['row_razdel_end'] = rows_itogi[0]['row_razdel_begin'] - 1

    return rows_of_title


# -----------------------------------------------------------------------
def sr_0420001_ods_r1_p1_p1_oboroty(df_avancor, wb, banks):
    """Формирование формы - основная функция"""

    bank_list = [bank['bank_id'] for bank in banks]
    sheet_name = '0420011 Отчетность об операци_2'
    # количество столбцов в таблице
    max_cols = 5

    # файл-Аванкор: наименование разделов
    avancoreTitle = '1.1. Виды и суммы операций, совершенных по банковским счетам некредитной \nфинансовой организации'
    razdel_names = ['1.1.1. Операции с резидентами – юридическими лицами',
                    '1.1.2. Операции с резидентами – индивидуальными предпринимателями',
                    '1.1.3. Операции с резидентами – физическими лицами',
                    '1.1.4. Операции с нерезидентами – юридическими лицами',
                    '1.1.5. Операции с нерезидентами – физическими лицами',
                    '1.1.6. Операции с неустановленными лицами']
    itogi_names = ['Всего\nобороты по\nсчету\n(счетам)',
                   'Остатки на\nначало\nотчетного\nпериода',
                   'Остатки на\nконец\nотчетного\nпериода']
    all_razdels = razdel_names + itogi_names
    # all_razdels.insert(0,avancoreTitle)

    # файл-Аванкор: колонки таблицы, в которых содержатся данные
    # код операции
    col_cod = 'D'
    # рубли
    col_rub = 'D'
    col_rub_out = 'D'
    col_rub_in = 'I'
    # доллары
    col_usd = 'K'
    col_usd_out = 'K'
    col_usd_in = 'L'
    # евро
    col_eur = 'O'
    col_eur_out = 'O'
    col_eur_in = 'Q'

    # формируем ШАПКИ таблиц для кождого банка
    make_form(bank_list, wb, sheet_name)

    # файл-Аванкор: строки с данными
    rows_of_title = {}
    for bank in banks:
        rows_of_title[bank['name']] = rows_with_data(df_avancor, banks, all_razdels)



# ---------------------------------------------------------------------------
if __name__ == '__main__':
    import pandas as pd
    from module.globals import *
    import shutil
    import openpyxl

    file_Avancore = 'Аванкор_ДДС_01.08.2020-31.08.2020.xlsx'
    path_to_report = r'c:\Users\Сотрудник\YandexDisk-atovanchov\XBRL_DDS\#Отчетность\2020.08'
    file_new_name = 'dds_08.xlsx'

    # file = path_to_report + '\\' + file_Avancore
    file = r'C:\Users\Сотрудник\YandexDisk-atovanchov\XBRL_DDS\#Отчетность\2020.08\Аванкор_ДДС_01.08.2020-31.08.2020.xlsx'

    # Загрузка файла-Аванкор-СЧА
    df_avancor = pd.read_excel(file,
                               index_col=None,
                               header=None)
    # устанавливаем начальный индекс не c 0, а c 1 (так удобнее)
    df_avancor.index += 1
    df_avancor.columns += 1
    # ----------------------------------------------------------

    # Создаем новый файл отчетности 'file_fond_name',
    # создав копию шаблона 'file_shablon'
    file_shablon = r"c:\Users\Сотрудник\YandexDisk-atovanchov\XBRL_DDS\#Шаблоны\ддс - 3.2.xlsx"
    shutil.copyfile(file_shablon, path_to_report + '\\' + file_new_name)
    # ---------------------------------------------------------
    # Загружаем данные из файла таблицы xbrl
    file_wb = path_to_report + '\\' + file_new_name
    wb = openpyxl.load_workbook(filename=file_wb)

    from module.forms_maker.sr_0420001_ods_r1_sved_ko import form_r1_sved_ko

    banks = form_r1_sved_ko(df_avancor, wb)

    sr_0420001_ods_r1_p1_p1_oboroty(df_avancor, wb, banks)

    # Сохраняем результат
    wb.save(path_to_report + '/' + file_new_name)
