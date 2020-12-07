import re
from openpyxl import load_workbook
from openpyxl import Workbook

difficult_test = load_workbook('MY22 Focus test.xlsx')
sheet = difficult_test.active
difficult_list = []

for row in sheet.iter_rows(max_col=4, values_only=True):
    cells = []
    for cell in row:
        if cell is None:
            cells.append('-')
        else:
            cells.append(cell)
    if cells[-1] != '-':
        difficult_list.append(cells[-1])
    elif cells[-1] == '-':
        break

print((difficult_list[1:]))
