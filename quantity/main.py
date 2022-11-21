import availiablestock as avst

AVAILABLE_STOCK_FILE = r'D:\МамаФинДвижения\Фин движения товаров.xlsx'
FILE_HEADERS = [
    'Код товара', 'Фирма', 'Контрагент', 'ИНН', 'Товар/Документ',
    'Начальный остаток', 'Приход', 'Расход', 'Конечный остаток',
]


if __name__ == '__main__':
    # Получение DataFrame файла с ФинДвижениемТоваров
    ws = avst.read_xlsx(AVAILABLE_STOCK_FILE, 'TDSheet')
    header_row = avst.find_header_row(ws, FILE_HEADERS)
    ctc = avst.match_colname_to_cell(ws, FILE_HEADERS, header_row)
    available_stock_df = avst.excel_to_dataframe(ws, ctc, min_row=header_row+1)
