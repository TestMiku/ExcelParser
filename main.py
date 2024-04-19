from openpyxl import load_workbook

c = {
    "A": 1,
    "B": 2,
    "C": 3,
    "D": 4,
    "E": 5,
    "F": 6,
    "G": 7,
    "H": 8,
    "I": 9,
    "J": 10,
    "K": 11,
    "L": 12,
    "M": 13,
    "N": 14,
    "O": 15,
    "P": 16,
    "Q": 17,
    "R": 18,
    "S": 19,
    "T": 20,
    "U": 21,
    "V": 22,
    "W": 23,
    "X": 24,
    "Y": 25,
    "Z": 26,
    "AA": 27,
    "AB": 28,
    "AC": 29,
    "AD": 30,
    "AE": 31,
    "AF": 32,
    "AG": 33,
    "AH": 34,
    "AI": 35,
    "AJ": 36,
    "AK": 37,
    "AL": 38,
    "AM": 39,
    "AN": 40,
    "AO": 41,
    "AP": 42,
    "AQ": 43,
    "AR": 44,
    "AS": 45,
    "AT": 46,
    "AU": 47,
    "AV": 48,
    "AW": 49,
    "AX": 50,
    "AY": 51,
}
# Получает список строк, которые нужно найти, и возвращает список их id
def get_ids(sheet, arr):
    ids = []
    for i in arr:
        target_row_index = None
        for row_index in range(1, sheet.max_row + 1):
            if sheet.cell(row=row_index, column=1).value == i:
                target_row_index = row_index
                break
        ids.append(target_row_index)
    return ids


# Получает строку, которую нужно скопировать
def get_row_to_copy():
    return 25727


# Получает строку, в которую нужно вставить скопированную строку
def get_row_to_paste():
    return 26


# Объединяет ячейки в строке
def merge_cells_in_row(sheet, row_index):
    sheet.merge_cells(start_row=row_index, start_column=1, end_row=row_index, end_column=21)


strs_to_merge = ['Поставщики', 'Поставщики - Wisk Telecom Solutions, TOO', 'Покупатели - Wisk Telecom Solutions, TOO']


source_wb = load_workbook('excel_data/7.15.2_1 (43).xlsx')
source_sheet = source_wb.active

target_wb = load_workbook('excel_data/Реестр_договоров_Аврора_Сервис_и_Виск_телеком_на_31_12_2023_г_.xlsx')
target_sheet = target_wb.active

ids_to_merge = get_ids(target_sheet, strs_to_merge)

source_row_index = get_row_to_copy()
target_row_index = get_row_to_paste()

target_sheet.insert_rows(target_row_index)

# for col_index, cell in enumerate(source_sheet[source_row_index], start=1):
#     target_sheet.cell(row=target_row_index, column=col_index).value = cell.value


target_sheet.cell(row=target_row_index, column=c["A"]).value = source_sheet.cell(row=source_row_index, column=c["A"]).value
target_sheet.cell(row=target_row_index, column=c["B"]).value = source_sheet.cell(row=source_row_index, column=c["B"]).value
target_sheet.cell(row=target_row_index, column=c["C"]).value = source_sheet.cell(row=source_row_index, column=c["C"]).value
target_sheet.cell(row=target_row_index, column=c["D"]).value = source_sheet.cell(row=source_row_index, column=c["D"]).value
target_sheet.cell(row=target_row_index, column=c["E"]).value = source_sheet.cell(row=source_row_index, column=c["E"]).value
target_sheet.cell(row=target_row_index, column=c["F"]).value = source_sheet.cell(row=source_row_index, column=c["F"]).value
target_sheet.cell(row=target_row_index, column=c["G"]).value = source_sheet.cell(row=source_row_index, column=c["G"]).value
target_sheet.cell(row=target_row_index, column=c["H"]).value = source_sheet.cell(row=source_row_index, column=c["H"]).value
target_sheet.cell(row=target_row_index, column=c["I"]).value = source_sheet.cell(row=source_row_index, column=c["I"]).value
target_sheet.cell(row=target_row_index, column=c["J"]).value = source_sheet.cell(row=source_row_index, column=c["J"]).value
target_sheet.cell(row=target_row_index, column=c["K"]).value = source_sheet.cell(row=source_row_index, column=c["K"]).value
target_sheet.cell(row=target_row_index, column=c["L"]).value = source_sheet.cell(row=source_row_index, column=c["L"]).value
target_sheet.cell(row=target_row_index, column=c["M"]).value = source_sheet.cell(row=source_row_index, column=c["M"]).value
target_sheet.cell(row=target_row_index, column=c["N"]).value = source_sheet.cell(row=source_row_index, column=c["N"]).value
target_sheet.cell(row=target_row_index, column=c["O"]).value = source_sheet.cell(row=source_row_index, column=c["O"]).value
target_sheet.cell(row=target_row_index, column=c["P"]).value = source_sheet.cell(row=source_row_index, column=c["P"]).value
target_sheet.cell(row=target_row_index, column=c["Q"]).value = source_sheet.cell(row=source_row_index, column=c["Q"]).value
target_sheet.cell(row=target_row_index, column=c["R"]).value = source_sheet.cell(row=source_row_index, column=c["R"]).value
target_sheet.cell(row=target_row_index, column=c["S"]).value = source_sheet.cell(row=source_row_index, column=c["S"]).value
target_sheet.cell(row=target_row_index, column=c["T"]).value = source_sheet.cell(row=source_row_index, column=c["T"]).value
target_sheet.cell(row=target_row_index, column=c["U"]).value = source_sheet.cell(row=source_row_index, column=c["U"]).value

# Объединяем ячейки в целевом листе
for i in ids_to_merge:
    if i >= get_row_to_paste():
        merge_cells_in_row(target_sheet, i+1)
    else:
        pass

target_wb.save('excel_data/NEW.xlsx')