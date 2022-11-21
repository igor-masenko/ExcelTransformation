import im_convert as cv
import openpyxl as xls
import pandas as pd
import os

HOME_DIR = os.path.abspath(r'C:\Users\Igor Masenko\Documents\Мама Excel\13.11.2022\УПД 2017-2019\П-Ю допы')
SOURCE_DIR = 'оригиналы'
TARGET_DIR = 'обработанные'

SOURCE_PATH = os.path.join(HOME_DIR, SOURCE_DIR)
TARGET_PATH = os.path.join(HOME_DIR, TARGET_DIR)

# TODO: добавить конвертацию xls -> xlsx

if not os.path.exists(TARGET_PATH):
    os.makedirs(TARGET_PATH)

if __name__ == '__main__':
    file_list = os.listdir(SOURCE_PATH)
    for file_name in file_list:
        if file_name.endswith('.xlsx'):
            dot_position = file_name.rfind('.')
            old_file_name = file_name[:dot_position]
            extension = file_name[dot_position:]

            print(f'Файл {file_name}', end='')

            source_file_name = os.path.join(SOURCE_PATH, file_name)
            wb = xls.load_workbook(filename=source_file_name)
            sheet = wb.active

            df = cv.covert_excel_to_dataframe(sheet)
            df = cv.make_calculations(df)
            df = df.sort_values(by=['Склад1', 'Склад2', 'Склад3'], ascending=(True, True, True))
            target_file_name = os.path.join(TARGET_PATH, f'{old_file_name}_о{extension}')

            writer = pd.ExcelWriter(target_file_name, engine='xlsxwriter')
            df.to_excel(writer, sheet_name='Плоская', index=False)
            writer.save()

            print(' обработан')
