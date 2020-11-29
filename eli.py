from openpyxl import Workbook
from openpyxl import load_workbook

select_sheet = load_workbook('w47_select.xlsx').active

full_list_sheet = load_workbook('1600_list.xlsx').active

unselected_cases = Workbook()

wb = unselected_cases.active

selected_cases_list = []

for item in select_sheet.rows:
    selected_cases_list.append(item[0].value)

for row in full_list_sheet.rows:
    cell_data = [(i.value) for i in row]
    if cell_data[0] not in selected_cases_list:
        wb.append(cell_data)

unselected_cases.save('unselect_cases.xlsx')
