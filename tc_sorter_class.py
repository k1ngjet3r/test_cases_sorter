from openpyxl import load_workbook
from openpyxl import Workbook
import re

categories = ['Flash User', 'Multi User', 'Button', 'Invalid Cases',
              'Driver_In_Off', 'Driver_In_On', 'Driver_Out_Off', 'Driver_Out_On',
              'Guest_In_Off', 'Guest_In_On', 'Guest_Out_Off', 'Guest_Out_On']

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


def matcher_slice(keywords, cell_data):
    for i in range(len(cell_data)):
        sen = (cell_data[i]).lower()
        for key in keywords:
            if re.search(key, sen):
                return True
    return False


def matcher_split(keywords, cell_data):
    for i in range(len(cell_data)):
        clean_sentance = re.sub(r'[^\w]', ' ', (cell_data[i]).lower())
        word_list = clean_sentance.split()
        for key in keywords:
            if key in word_list:
                return True
    return False


class Tc_sorter:
    def __init__(self, input_name, output_name, categories):
        self.input_name = str(input_name)
        self.output_name = str(output_name)
        self.categories = categories
        self.sheet = (load_workbook(self.input_name)).active
        self.wb = Workbook()
        self.wb.active
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

    def sorting_exceptions(self, cell_data, exceptions):
        wb = self.wb
        for name in exceptions:
            for i in range(len(cell_data)):
                if name != 'press_button':
                    if matcher_split(exceptions[name], cell_data[i]):
                        wb[name].append(cell_data)
                else:
                    if matcher_slice(exceptions[name], cell_data[i]):
                        wb[name].append(cell_data)
        return

    def phone_type(self, cell_data):
        for i in range(len(cell_data)):
            if matcher_split(['iphone'], cell_data[i]):
                return "iPhone"
            elif matcher_split(['android'], cell_data[i]):
                return 'Android'
            else:
                return ' '

    def projection_type(self, cell_data):
        for j in range(len(cell_data)):
            if matcher_split(['carplay'], cell_data[j]):
                return "Apple CarPlay"
            elif matcher_slice(['android auto', 'waa'], cell_data[j]):
                return "Android Auto"
            else:
                return ' '

    def function_determaination(self, cell_data, functions):
        for name in functions:
            if name == 'navigation':
                for i in range(1, 4):
                    if matcher_slice(functions[name], cell_data[i]):
                        return name
            else:
                for j in range(1, 4):
                    if matcher_split(functions[name], cell_data[j]):
                        return name

    def appending(self, name, cell_data):
        function = self.function_determaination(cell_data, functions)
        cell_data.append(function)
        cell_data.append(self.phone_type(cell_data))
        projection = self.projection_type(cell_data)
        cell_data.append(projection)
        self.wb[name].append(cell_data)

    def sorting(self):
        for row in (self.sheet).rows:
            cell_data = [cell.value for cell in row]
            # sort the exception cases first
            self.sorting_exceptions(cell_data, exp_dict)
            # differentiate Driver/Guest, sign-in/sign-out and online/offline

            # Driver
            # for 'Driver/In/Off'
            if matcher_split(driver, cell_data) and matcher_slice(sign_out, cell_data) != True and matcher_split(online, cell_data) != True:
                self.appending('Driver_In_Off', cell_data)

            # for 'Driver/In/On'
            elif matcher_split(driver, cell_data) and matcher_slice(sign_out, cell_data) != True and matcher_split(online, cell_data):
                self.appending('Driver_In_On', cell_data)

            # for 'Driver/Out/Off'
            elif matcher_split(driver, cell_data) and matcher_slice(sign_out, cell_data) and matcher_split(online, cell_data) != True:
                self.appending('Driver_In_Off', cell_data)

            # for 'Driver/Out/On'
            elif matcher_split(driver, cell_data) and matcher_slice(sign_out, cell_data) and matcher_split(online, cell_data):
                self.appending('Driver_In_Off', cell_data)

            # Guest
            # for 'Guest/In/Off'
            elif matcher_split(driver, cell_data) != True and matcher_slice(sign_out, cell_data) != True and matcher_split(online, cell_data) != True:
                self.appending('Driver_In_Off', cell_data)

            # for 'Guest/In/On'
            elif matcher_split(driver, cell_data) != True and matcher_slice(sign_out, cell_data) != True and matcher_split(online, cell_data):
                self.appending('Driver_In_On', cell_data)

            # for 'Guest/Out/Off'
            elif matcher_split(driver, cell_data) != True and matcher_slice(sign_out, cell_data) and matcher_split(online, cell_data) != True:
                self.appending('Driver_In_Off', cell_data)

            # for 'Guest/Out/On'
            elif matcher_split(driver, cell_data) != True and matcher_slice(sign_out, cell_data) and matcher_split(online, cell_data):
                self.appending('Driver_In_Off', cell_data)

        self.wb.save(self.output_name)


testing = Tc_sorter('Taipei_CaseList.xlsx',
                    'Sorted_cases_W46.xlsx', categories)

testing.sorting()
