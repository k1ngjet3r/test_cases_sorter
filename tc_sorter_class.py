from openpyxl import load_workbook
from openpyxl import Workbook
import re

categories = ['flash user', 'multi_user', 'press_button', 'invalid',
              'Driver_In_Off', 'Driver_In_On', 'Driver_Out_Off', 'Driver_Out_On',
              'Guest_In_Off', 'Guest_In_On', 'Guest_Out_Off', 'Guest_Out_On']

# keyword for layer 1
flash_user = ['flash']
multi_user = ['multi', 'primary' 'secondary']
press_button = ['long press', 'short press', 'press "end" key']
invalid = ['audiobook', 'sxm']

exceptions_names = ['flash_user', 'multi_user', 'press_button', 'invalid']
exceptions_items = [flash_user, multi_user, press_button, invalid]

exp_dict = {name: item for name,
            item in zip(exceptions_names, exceptions_items)}

guest = ['guest']

# Layer 2
sign_out = ['sign out', 'sign-out', 'signout', 'signed out']
offline = ['offline']

# keyword for determining the function-related categories
navigation = ['navigat', 'go to', 'add stop', 'guidance',
              'how far', 'take me', 'address', 'traffic']

call_SMS = ['call', 'phone', 'message', 'messages',
            'reply', 'text', 'sms', 'dial', 'Send', 'dail']

media = ['play', 'pause', 'next', 'previous',
         'volume', 'music', 'am', 'fm', 'radio', 'news', 'tune', 'tuned', 'plays', 'bluetooth', 'station', 'album', 'podcast', 'pauses', 'media', 'songs', 'playback', 'bt']


ac = ['a/c', 'temperature', 'climate', 'defroster', 'air', 'fan']

function_items = [call_SMS, media, ac, navigation]

function_names = ['call_SMS', 'media', 'ac', 'navigation']

functions = {name: item for name, item in zip(function_names, function_items)}


def matcher_slice(keywords, cell_data, index_range):
    for i in index_range:
        sen = cell_data[i]
        for key in keywords:
            if re.search(key, sen):
                return True
    return False


def matcher_split(keywords, cell_data, index_range):
    for i in index_range:
        clean_sentance = re.sub(r'[^\w]', ' ', cell_data[i])
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

    def cell_data(self, row):
        return [cell.value.lower() for cell in row]

    def sorting_exceptions(self, cell_data, exceptions):
        wb = self.wb
        index_range = [1, 2]
        for name in exceptions:
            if name != 'press_button':
                if matcher_split(exceptions[name], cell_data, index_range):
                    wb[name].append(cell_data)
            else:
                if matcher_slice(exceptions[name], cell_data, index_range):
                    wb[name].append(cell_data)
        return

    def phone_type(self, cell_data):
        if matcher_split(['iphone'], cell_data, [1]):
            return "iPhone"
        elif matcher_slice(['android phone'], cell_data, [1]):
            return 'Android'
        else:
            return ' '

    def projection_type(self, cell_data):
        if matcher_split(['carplay'], cell_data, [1]):
            return "Apple CarPlay"
        elif matcher_slice(['android auto', 'waa'], cell_data, [1]):
            return "Android Auto"
        else:
            return ' '

    def function_determaination(self, cell_data, functions):
        for name in functions:
            if name == 'navigation':
                if matcher_slice(functions[name], cell_data, [2, 3]):
                    return name
            else:
                if matcher_split(functions[name], cell_data, [2, 3]):
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
            cell_data = self.cell_data(row)

            # Filter out the exception cases
            if matcher_split(flash_user, cell_data, [1, 2]):
                self.wb['flash_user'].append(cell_data)

            elif matcher_split(multi_user, cell_data, [1, 2]):
                self.wb['multi_user'].append(cell_data)

            elif matcher_split(invalid, cell_data, [1, 2]):
                self.wb['invalid'].append(cell_data)

            elif matcher_slice('press_button', cell_data, [1, 2]):
                self.wb['press_button'].append(cell_data)

            else:
                # determine the Driver as user
                if matcher_split(guest, cell_data, [1]) == False:
                    # determine sign-in
                    if matcher_split(sign_out, cell_data, [1]) == False:
                        # determine online
                        if matcher_split(offline, cell_data, [1]) == False:
                            self.appending('Driver_In_On', cell_data)
                        # determine offline
                        elif matcher_split(offline, cell_data, [1]):
                            self.appending('Driver_In_Off', cell_data)

                    # determine sign-out
                    elif matcher_split(sign_out, cell_data, [1]):
                        # determine online
                        if matcher_split(offline, cell_data, [1]) == False:
                            self.appending('Driver_Out_On', cell_data)
                        # determine offline
                        elif matcher_split(offline, cell_data, [1]):
                            self.appending('Driver_Out_Off', cell_data)

                # determine the Guest as user
                elif matcher_split(guest, cell_data, [1]):
                    # determine sign-in
                    if matcher_split(sign_out, cell_data, [1]) == False:
                        # determine online
                        if matcher_split(offline, cell_data, [1]) == False:
                            self.appending('Guest_In_On', cell_data)
                        # determine offline
                        elif matcher_split(offline, cell_data, [1]):
                            self.appending('Guest_In_Off', cell_data)

                    # determine sign-out
                    elif matcher_split(sign_out, cell_data, [1]):
                        # determine online
                        if matcher_split(offline, cell_data, [1]) == False:
                            self.appending('Guest_Out_On', cell_data)
                        # determine offline
                        elif matcher_split(offline, cell_data, [1]):
                            self.appending('Guest_Out_Off', cell_data)

        self.wb.save(self.output_name)


testing = Tc_sorter('W46.xlsx',
                    'Sorted_cases_W46.xlsx', categories)

testing.sorting()
