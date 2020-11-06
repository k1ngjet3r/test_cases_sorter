from openpyxl import load_workbook
from openpyxl import Workbook
import re

categories = ['Flash User', 'Multi User', 'Button', 'Invalid Cases',
              'Driver/In/Off', 'Driver/In/On', 'Driver/Out/Off', 'Driver/Out/On', 'Guest/In/Off', 'Guest/In/On', 'Guest/Out/Off', 'Guest/Out/On']

# keyword for layer 1
flash_user = ['flash']
mulit_user = ['multi', 'primary' 'secondary']
press_button = ['long press', 'short press', 'press "end" key']
invalid = ['audiobook']

exceptions_names = ['flash_user', 'mulit_user', 'press_button', 'invalid']
exceptions_items = [flash_user, mulit_user, press_button, invalid]

exp_dict = {name: item for name,
            item in zip(exceptions_names, exceptions_items)}

driver = ['driver']

# Layer 2
sign_out = ["user is signed out",
            "signout the google account", "sign out the google account"]

online = ['online', 'internet']

# keyword for determining the function-related categories
navigation = ['navigat', 'go to', 'add stop', 'guidance',
              'how far', 'take me', 'address', 'traffic']

call_SMS = ['call', 'phone', 'message',
            'reply', 'text', 'sms', 'dial', 'Send', 'dail']

media = ['play', 'pause', 'next', 'previous',
         'volume', 'music', 'am', 'fm', 'radio', 'news', 'tune', 'tuned', 'plays', 'bluetooth', 'station', 'album', 'podcast', 'pauses', 'media', 'songs', 'playback', 'bt']


ac = ['a/c', 'temperature', 'climate', 'defroster', 'air', 'fan']

function_items = [navigation, call_SMS, media, ac]

function_names = ['navigation', 'call_SMS', 'media', 'ac']

functions = {name: item for name, item in zip(function_names, function_items)}


def matcher_slice(keywords, row):
    for i in range(1, 4):
        sen = (row[i].value).lower()
        for key in keywords:
            if re.search(key, sen):
                return True
    return False


def matcher_split(keywords, row):
    for i in range(1, 4):
        clean_sentance = re.sub(r'[^\w]', ' ', (row[i].value).lower())
        word_list = clean_sentance()
        for key in keywords:
            if key in word_list:
                return True
    return False


class tc_sorter:
    def __init__(self, input_name, output_name, categories):
        self.input_name = str(input_name)
        self.output_name = str(output_name)
        self.categories = categories
        self.sheet = (load_workbook(self.input_name)).active
        self.wb = Workbook().active
        for category in self.categories:
            self.wb.create_sheet(category, int(
                (self.categories).index(category)))

    # def read_file(self):
    #     workbook = load_workbook(self.input_name)
    #     sheet = workbook.active
    #     sheet.delete_rows(1)
    #     return sheet

    # def init_output_file(self):
    #     wb = Workbook()
    #     sorted_cases = wb.active
    #     for category in self.categories:
    #         sorted_cases.create_sheet(category, int(
    #             (self.categories).index(category)))
    #     return sorted_cases

    def cell_data(self, row):
        return [cell.value for cell in row]

    def sorting_exceptions(self, row, exceptions):
        cell_data = self.cell_data(row)
        wb = self.wb
        for name in exceptions:
            if name != 'press_button':
                if matcher_split(exceptions[name], row):
                    wb[name].append(cell_data)
            else:
                if matcher_slice(exceptions[name], row):
                    wb[name].append(cell_data)
        return

    def phone_type(self, row):
        if matcher_split(['iphone'], row):
            return "iPhone"
        elif matcher_split(['android'], row):
            return 'Android'
        else:
            return ' '

    def projection_type(self, row):
        if matcher_split(['carplay'], row):
            return "Apple CarPlay"
        elif matcher_slice(['android auto', 'waa'], row):
            return "Android Auto"
        else:
            return ' '

    def function_determaination(self, row, functions):
        for name in functions:
            if name == 'navigation':
                for i in range(1, 4):
                    if matcher_slice(functions[name], row[i]):
                        return name
            else:
                for j in range(1, 4):
                    if matcher_split(functions[name], row[j]):
                        return name

    def sorting(self):
        for row in (self.sheet).rows:
            cell_data = [[cell.value for cell in row]]
            # sort the exception cases first
            self.sorting_exceptions(row, exp_dict)
            # differentiate Driver/Guest, sign-in/sign-out and online/offline
            # for 'Driver/In/Off'
            if matcher_split(driver, row) and matcher_slice(sign_out, row) != True and matcher_split(online, row):
                function = self.function_determaination(row, functions)
                cell_data.append(function)
                cell_data.append(self.phone_type(row))
                projection = self.projection_type(row)
                cell_data.append(projection)
                self.wb['Driver/In/Off'].append(cell_data)

            # for Driver/In/On
            elif matcher_split(driver, row) and matcher_slice(sign_out, row) != True and matcher_split(online, row) != True:
                function = self.function_determaination(row, functions)
                cell_data.append(function)
                cell_data.append(self.phone_type(row))
                projection = self.projection_type(row)
                cell_data.append(projection)
                self.wb['Driver/In/On']
