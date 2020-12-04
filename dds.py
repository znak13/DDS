import pandas as pd
import shutil
import openpyxl
from loguru import logger


from module.globals import *

from module.forms_maker import sr_0420001_ods_r1_sved_ko

# logger.add("file_{time}.log")


file_Avancore = 'Аванкор_ДДС_01.08.2020-31.08.2020.xlsx'
path_to_report = dir_reports + '/2020.08'
file_new_name = 'dds_08.xlsx'

logger.info(f'Загрузка данных из файла-Аванкор-СЧА "{file_Avancore}" в "df_avancor"')
# Загрузка файла-Аванкор-СЧА
df_avancor = pd.read_excel(path_to_report + '/' + file_Avancore,
                           index_col=None,
                           header=None)
# устанавливаем начальный индекс не c 0, а c 1 (так удобнее)
df_avancor.index += 1
df_avancor.columns += 1
# ----------------------------------------------------------
logger.info(f'Создаем новый файл отчетности: "{file_new_name}" из шаблона: "{fileShablon}"')
# Создаем новый файл отчетности 'file_fond_name',
# создав копию шаблона 'file_shablon'
shutil.copyfile(dir_shablon + fileShablon, path_to_report + '/' + file_new_name)
# ---------------------------------------------------------
logger.info(f'Создаем "work-book" из созданного файла отчетности: "{file_new_name}"')
# Загружаем данные из файла таблицы xbrl
wb = openpyxl.load_workbook(filename=(path_to_report + '/' + file_new_name))

logger.info(f'Формируем 1-ую форму: "0420011 Отчетность об операциях"')
# Формируем 1-ую форму: "0420011 Отчетность об операциях"
banks = sr_0420001_ods_r1_sved_ko.form_r1_sved_ko(df_avancor, wb)

logger.info(f'Формируем 2-ую форму: "0420011 Отчетность об операци_2"')
# Формируем 2-ую форму: "0420011 Отчетность об операци_2"



# Сохраняем результат
# wb.save(path_to_report + '/' + file_new_name)