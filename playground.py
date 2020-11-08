from openpyxl import load_workbook
from openpyxl import Workbook
import re

flash_user = ['flash']
multi_user = ['multi', 'primary', 'secondary']
press_button = ['long press', 'short press', 'press "end" key', 'press ptt']
invalid = ['audiobook']

exceptions = [flash_user, multi_user, press_button, invalid]
expts = ['flash_user', 'multi_user', 'press_button', 'invalid']

exp_dict = {name: item for name, item in zip(expts, exceptions)}

# print(exp_dict)

workbook = load_workbook('Untitled spreadsheet.xlsx')
sheet = workbook.active

for row in sheet.rows:
    cell_data = [cell.value.lower() for cell in row]


# print(cell_data)
for name in exp_dict:
    for i in [1, 2]:
        if name != 'press_button':
            clean_sentance = re.sub(r'[^\w]', ' ', cell_data[i])
            word_list = clean_sentance.split()
            for item in exp_dict[name]:
                if item in word_list:
                    print(name)
        else:
            sen = cell_data[i]
            for item in exp_dict[name]:
                if item in sen:
                    print(name)
