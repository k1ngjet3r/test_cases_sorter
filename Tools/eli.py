from openpyxl import Workbook
from openpyxl import load_workbook

select_sheet = load_workbook('W07_410.xlsx').active

full_list_sheet = load_workbook('W08_testplan.xlsx').active

wb = Workbook()

# wb = unselected_cases.active

wb.create_sheet('W07')
wb.create_sheet('W06')

selected_cases_list = []

for item in select_sheet.rows:
    selected_cases_list.append(item[0].value)

for row in full_list_sheet.rows:
    cell_data = [(i.value) for i in row]
    if cell_data[0] not in selected_cases_list:
        wb['W06'].append(cell_data)
    else:
        wb['W07'].append(cell_data)

wb.save('unselect_cases.xlsx')
