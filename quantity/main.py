import availiablestock as avst
import fakedoc

AVAILABLE_STOCK_FILE = r'D:\МамаФинДвижения\Фин движения товаров.xlsx'
FAKE_DOC_FILE = r'D:\МамаФинДвижения\кондитерка 43,45,47,49 - Copy.xlsx'

AVAILABLE_STOCK_FILE_HEADERS = [
    'Код товара', 'Фирма', 'Контрагент', 'ИНН', 'Товар/Документ',
    'Начальный остаток', 'Приход', 'Расход', 'Конечный остаток',
]
FAKE_DOC_FILE_HEADERS = ['Товар', 'Код', 'Акт', 'Приход']


if __name__ == '__main__':
    # Получение DataFrame файла с ФинДвижениемТоваров
    ws = avst.read_xlsx(AVAILABLE_STOCK_FILE, 'TDSheet')
    header_row = avst.find_header_row(ws, AVAILABLE_STOCK_FILE_HEADERS)
    ctc = avst.match_colname_to_cell(ws, AVAILABLE_STOCK_FILE_HEADERS, header_row)
    available_stock_df = avst.excel_to_dataframe(ws, ctc, min_row=header_row+1)

    # Получение DataFrame файла с Липовыми документами
    ws = fakedoc.read_xlsx(FAKE_DOC_FILE, 'Лист1')
    fakedoc_df = fakedoc.excel_to_dataframe(ws)

    print(available_stock_df.head(10))
    print(fakedoc_df.head(10))

