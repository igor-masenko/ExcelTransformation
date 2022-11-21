import pandas as pd


def create_dict_from_excel(excel_sheet):
    stock1 = 'Склад1'
    stock2 = 'Склад2'
    stock3 = 'Склад3'
    dict_excel = {
        stock1: [],
        stock2: [],
        stock3: [],
    }
    for column in excel_sheet.iter_cols(min_row=7, min_col=1, max_col=8):
        key = column[0].value
        if key == 'Склад' or key == 'Место хранения / Номенклатура / Документ':
            prev1 = None
            prev2 = None
            prev3 = None
            for cell in column[1:]:
                if cell.font.size == 11:
                    dict_excel[stock1].append(cell.value)
                    prev1 = dict_excel[stock1][-1]
                    prev2 = None
                    prev3 = None
                else:
                    dict_excel[stock1].append(prev1)
                if cell.font.size == 10:
                    dict_excel[stock2].append(cell.value)
                    prev2 = dict_excel[stock2][-1]
                    prev3 = None
                else:
                    dict_excel[stock2].append(prev2)
                if cell.font.size == 8:
                    dict_excel[stock3].append(cell.value)
                    prev3 = dict_excel[stock3][-1]
                else:
                    dict_excel[stock3].append(prev3)
        else:
            dict_excel[key] = []
            for cell in column[1:]:
                dict_excel[key].append(cell.value)

    return dict_excel


def covert_excel_to_dataframe(excel_sheet):
    data_frame = pd.DataFrame(data=create_dict_from_excel(excel_sheet))
    data_frame = data_frame[data_frame['Склад3'].notna()]
    data_frame = data_frame[data_frame['Склад3'].str.strip() != 'Итого']
    return data_frame


def make_calculations(data_frame):
    def make_coefficient(row, col_name):
        coefficient = (
            0.9 if -10.0 <= row[col_name] <= 0.0 else
            0.8 if -20.0 <= row[col_name] <= -10.0 else
            0.7 if -30.0 <= row[col_name] <= -20.0 else
            0.6 if -40.0 <= row[col_name] <= -30.0 else
            0.5 if -50.0 <= row[col_name] <= -40.0 else
            0.4 if -60.0 <= row[col_name] <= -50.0 else
            0.3 if -70.0 <= row[col_name] <= -60.0 else
            0.2 if -80.0 <= row[col_name] <= -70.0 else
            0.1 if -100.0 <= row[col_name] <= -80.0 else
            99999999 if row[col_name] <= -100.0 else
            1
        )
        return coefficient

    data_frame['МинНаценки'] = data_frame.groupby('Склад3')['% Наценки'].transform('min')
    data_frame['Сумма продаж - НДС = X'] = \
        data_frame['Сумма продаж'] \
        - data_frame['НДС18-20'].fillna(0) \
        - data_frame['НДС10'].fillna(0)

    data_frame['X / Кол_во = Y'] = data_frame['Сумма продаж - НДС = X'] / data_frame['Кол-во'].fillna(1)
    data_frame['Коэф_склад3'] = data_frame.apply(lambda row: make_coefficient(row, 'МинНаценки'), axis=1)
    data_frame['Коэф_склад2'] = data_frame.apply(lambda row: make_coefficient(row, '% Наценки'), axis=1)
    data_frame['X / Y = Z'] = data_frame['X / Кол_во = Y'] / data_frame['Коэф_склад3']
    data_frame['X / Z = Q'] = data_frame.apply(
        lambda row:
            round(row['Сумма продаж - НДС = X'] / (row['X / Кол_во = Y'] / row['Коэф_склад3']))
            if int(row['Кол-во'] * 1000) % 1000 == 0 else
            row['Сумма продаж - НДС = X'] / (row['X / Кол_во = Y'] / row['Коэф_склад3']),
        axis=1
    )

    return data_frame
