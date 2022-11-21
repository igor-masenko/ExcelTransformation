import numpy as np
import openpyxl as xlsx
import pandas as pd
from openpyxl.worksheet.worksheet import Worksheet

COLNAMES_XLSX = ['Товар', 'Код', 'Акт', 'Приход']


def read_xlsx(path: str, sheet_name: str = None) -> Worksheet:
    wb = xlsx.load_workbook(path, data_only=True)
    if sheet_name is not None:
        return wb[sheet_name]
    else:
        return wb.active


def excel_to_dataframe(worksheet: Worksheet) -> pd.DataFrame:
    cols_df = ['row_index', 'productName', 'productCode', 'doc', 'income']
    d = {k: [] for k in cols_df}

    for row_ in worksheet.iter_rows(min_row=2):  # Будем считать, что 1-ая строка - заголовки
        d['row_index'].append(row_[0].row)
        d['productName'].append(row_[0].value)
        d['productCode'].append(row_[1].value)
        d['doc'].append(row_[2].value)
        d['income'].append(row_[3].value)

    df = pd.DataFrame(d)

    df.sort_values('row_index', inplace=True)
    df.drop(columns='row_index', inplace=True)
    df[['productName', 'productCode']] = df[['productName', 'productCode']].ffill()

    df[['productName', 'doc']] = df[['productName', 'doc']].astype(str)
    df[['productCode']] = df[['productCode']].astype(np.int32)
    df[['income']] = df[['income']].astype(np.float32)

    pattern = r'(\d{,2}\.\d{,2}\.\d{4} \d{,2}:\d{,2}:\d{,2})'
    df['datetimeofdoc'] = df['doc'].str.extract(pattern)
    df['datetimeofdoc'] = pd.to_datetime(df['datetimeofdoc'])

    return df