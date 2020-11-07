from openpyxl import load_workbook
from openpyxl import Workbook
import re

flash_user = ['flash']
multi_user = ['multi', 'primary', 'secondary']
press_button = ['long press', 'short press', 'press "end" key']
user = ['guest', 'driver']
invalid = ['audiobook']

exceptions = [flash_user, multi_user, press_button, invalid]
expts = ['flash_user', 'multi_user', 'press_button', 'invalid']

exp_dict = {name: item for name, item in zip(expts, exceptions)}

workbook = load_workbook('Taipei_CaseList.xlsx')
sheet = workbook.active

for row in sheet.rows:
    cell_data = [cell.value for cell in row]

print(cell_data)
