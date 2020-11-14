from openpyxl import load_workbook
from openpyxl import Workbook

original_file = load_workbook('test_example.xlsx')
sheet = original_file.active

for row in sheet.iter_rows(max_col=4, values_only=True):
    cell_data = [data for data in row]
    print(cell_data)
