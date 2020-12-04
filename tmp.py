
def file_read(file_name):
    """Чтение данных из файла"""
    # Получаем список строк
    # (file_name - имя файла, включая путь

    data_from_file = []
    # test_file = r'c:\Users\Сотрудник\YandexDisk-atovanchov\XBRL_DDS\module\currency_code.csv'
    with open(file_name) as data_file:
        for line in data_file:
            data_from_file.append(line[:-1])
    return data_from_file

test_file = r'c:\Users\Сотрудник\YandexDisk-atovanchov\XBRL_DDS\module\operation_codes.csv'
q = file_read(test_file)