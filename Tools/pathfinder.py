import os
from shutil import copy2
from openpyxl import load_workbook



def path_finder(filename, search_path):
    print('finding file {}...'.format(filename))
    for root, direction, files in os.walk(search_path):
        if filename in files:
            # print('Found!')
            print(root+'\\'+filename)
            return root+'\\'+filename

def move_to_destination(target):
    destination = 'C:\\Users\\GM-PC-03\\Desktop\\temp\\'
    copy2(target, destination)



if __name__ == '__main__':
    case_list = load_workbook('Cases_comparison.xlsx')['Both_automated']
    search_path = 'C:\\Users\\GM-PC-03\\Desktop\\Automation\\src\\1080p_tt\\scripts'

    for case in case_list.iter_rows(max_col=1, values_only=True):
        case_name = case[0] + '.xml'
        if path_finder(case_name, search_path):
            move_to_destination(path_finder(case_name, search_path))