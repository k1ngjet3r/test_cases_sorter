from openpyxl import load_workbook
from openpyxl import Workbook
import re

flash_user = ['flash']


def matcher_slice(keywords, cell_data):
    sen = cell_data.lower()
    for key in keywords:
        if re.search(key, sen):
            return True
    return False


def matcher_split(keywords, cell_data):
    clean_sentance = re.sub(r'[^\w]', ' ', cell_data.lower())
    word_list = clean_sentance.split()
    for key in keywords:
        if key in word_list:
            return True
    return False


class Tc_sorter:
    def __init__(self, input_name, output_name):
        self.input_name = str(input_name)
        self.output_name = str(output_name)
        self.sheet = (load_workbook(self.input_name)).active
        self.wb = Workbook()
        self.wb.active

    def cell_data(self, row):
        return [cell.value for cell in row]

    def phone_type(self, cell_data):
        iphone = ['iphone', 'cp', 'wcp']
        android = ['android', 'waa', 'aa']
        phone_requirement = [0, 0]
        for cell in cell_data[1:3]:
            if matcher_split(iphone, cell):
                phone_requirement[0] = 1
            if matcher_slice(android, cell):
                phone_requirement[1] = 1
        if phone_requirement == [1, 0]:
            cell_data.append('iPhone')
        elif phone_requirement == [0, 1]:
            cell_data.append('Android')
        elif phone_requirement == [1, 1]:
            cell_data.append('Both')
        else:
            cell_data.append(' ')

    def sign_status(self, cell_data):
        sign_out = ['sign out', 'sign-out', 'signout', 'signed out']
        if matcher_slice(sign_out, cell_data[1]):
            cell_data.append('sign out')
        else:
            cell_data.append('sign in')

    def connection(self, cell_data):
        offline = ['offline']
        if matcher_split(offline, cell_data[1]):
            cell_data.append('offline')
        else:
            cell_data.append('online')

    def user(self, cell_data):
        guest = ['guest']
        others = ['secondary', 'user 1', 'user 2', 'user1', 'user2']
        if matcher_split(guest, cell_data[1]):
            cell_data.append('guest')
        elif matcher_slice(others, cell_data[1]):
            cell_data.append('others')
        else:
            cell_data.append('Driver')

    def bench_only(self, cell_data):
        ac = ['a/c', 'temperature', 'climate',
              'defroster', 'air', 'fan', 'hvac']
        press_button = ['long press', 'short press', 'press "end" key']
        cluster = ['cluster']
        bench_only_case = False
        for cell in cell_data[1:4]:
            if matcher_slice(ac, cell) or matcher_slice(press_button, cell) or matcher_slice(cluster, cell):
                bench_only_case = True
        return bench_only_case

    def sorting(self):
        sheet = self.sheet
        for row in sheet.iter_rows(max_col=4, values_only=True):
            cell_data = self.cell_data(row)
            self.phone_type(cell_data)
            self.sign_status(cell_data)
            self.connection(cell_data)
            self.user(cell_data)

        self.wb.save(self.output_name)


testing = Tc_sorter('W46.xlsx',
                    'Sorted_cases_W46.xlsx')

testing.sorting()
