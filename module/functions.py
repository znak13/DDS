"""Общие функции"""
from openpyxl.utils import column_index_from_string  # 'B' -> 2

# -------------------------------------------------------------
def razdel_name_row(df_avancor, title_name,
                    title_col=None, row_start=None, row_end = None ):
    """Поиск Номера строки с названием раздела в файле Аванкор"""

    # номер колонки, в которой ищим название раздела
    if not title_col:
        title_col = 2
    elif type(title_col) == str:
        title_col = column_index_from_string(title_col)
    elif type(title_col) != int:
        print(f'Ошибка при указании колинки')

    # номер строки, начиная с которой пеербираем строки
    if not row_start:
        row_start = 1
    # номер строки, до которой пеербираем строки
    if not row_end:
        row_end = df_avancor.shape[0] # количество строк в файле

    title_row = []
    for row in range(row_start, row_end):
        title = str(df_avancor.loc[row, title_col])
        if title == title_name or title[:20] == title_name[:20]:
            title_row.append(row)

    # если найдена одна строка, то возвращаем номер строки, а не список
    # if len(title_row) == 1:
    #     title_row = title_row[0]

    return title_row if title_row else print(f'раздел не найден')

# -------------------------------------------------------------
def file_read(file_name):
    """Чтение данных из файла"""
    # Получаем список строк
    # (file_name - имя файла, включая путь

    data_from_file = []
    # test_file = r'c:\Users\Сотрудник\YandexDisk-atovanchov\XBRL_DDS\module\currency_code.csv'
    with open(file_name) as data_file:
        for line in data_file:
            # Исключаем знак переноса строки в конце строки: line[:-1]
            # (последгяя строка в "file_name" должна быть пустой,
            # иначе в последней строке будет потерян последний символ)
            data_from_file.append(line[:-1])
    return data_from_file

# -------------------------------------------------------------