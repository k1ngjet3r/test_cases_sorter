from openpyxl import load_workbook
from openpyxl import Workbook
import re

categories = ['flash user', 'multi_user', 'press_button', 'invalid',
              'Driver_In_Off', 'Driver_In_On', 'Driver_Out_Off', 'Driver_Out_On',
              'Guest_In_Off', 'Guest_In_On', 'Guest_Out_Off', 'Guest_Out_On']

# keyword for layer 1
flash_user = ['flash']
multi_user = ['multi', 'primary' 'secondary']
press_button = ['long press', 'short press', 'press "end" key', 'ignition']
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
        sen = cell_data[i].lower()
        for key in keywords:
            if re.search(key, sen):
                return True
    return False


def matcher_split(keywords, cell_data, index_range):
    for i in index_range:
        clean_sentance = re.sub(r'[^\w]', ' ', cell_data[i].lower())
        word_list = clean_sentance.split()
        for key in keywords:
            if key in word_list:
                return True
    return False


class Tc_sorter:
    def __init__(self, cell_data, categories):
        self.cell_data = cell_data
        self.categories = categories
        self.sheet = (load_workbook(self.input_name)).active
        self.wb = Workbook()
        self.wb.active
        for category in self.categories:
            catagory = []

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

    def phone_type(self, cell_data):
        if matcher_split(['iphone'], cell_data, [1]):
            return "iPhone"
        elif matcher_slice(['android'], cell_data, [1]):
            return 'Android'
        elif matcher_split(['iphone'], cell_data, [1]) and matcher_slice(['android'], cell_data, [1]):
            return 'Both'
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
                flash_user.append(cell_data)

            elif matcher_split(multi_user, cell_data, [1, 2]):
                multi_user.append(cell_data)

            elif matcher_split(invalid, cell_data, [1, 2]):
                invalid.append(cell_data)

            elif matcher_slice('press_button', cell_data, [1, 2]):
                press_button.append(cell_data)

            else:
                # determine the Driver as user
                if matcher_split(guest, cell_data, [1]) == False:
                    # determine sign-in
                    if matcher_split(sign_out, cell_data, [1]) == False:
                        # determine online
                        if matcher_split(offline, cell_data, [1]) == False:
                            Driver_In_On.append(cell_data)
                        # determine offline
                        elif matcher_split(offline, cell_data, [1]):
                            Driver_In_Off.append(cell_data)

                    # determine sign-out
                    elif matcher_split(sign_out, cell_data, [1]):
                        # determine online
                        if matcher_split(offline, cell_data, [1]) == False:
                            Driver_Out_On.append(cell_data)
                        # determine offline
                        elif matcher_split(offline, cell_data, [1]):
                            Driver_Out_Off.append(cell_data)

                # determine the Guest as user
                elif matcher_split(guest, cell_data, [1]):
                    # determine sign-in
                    if matcher_split(sign_out, cell_data, [1]) == False:
                        # determine online
                        if matcher_split(offline, cell_data, [1]) == False:
                            Guest_In_On.append(cell_data)
                        # determine offline
                        elif matcher_split(offline, cell_data, [1]):
                            Guest_In_Off.append(cell_data)

                    # determine sign-out
                    elif matcher_split(sign_out, cell_data, [1]):
                        # determine online
                        if matcher_split(offline, cell_data, [1]) == False:
                            Guest_Out_On.append(cell_data)
                        # determine offline
                        elif matcher_split(offline, cell_data, [1]):
                            Guest_Out_Off.append(cell_data)


row_1 = ['TC_Android_Auto_0001_Wireless',
         '1. System is ON. 2. Wireless Android Auto device is connected via BT.3. User is on "Phones" screen.',
         '1. Forget the Wireless Android Auto from Phones list 2. Observe Phones screen',
         'No Phones Connected screen should be displayed, "Add Phone+"option shall be displayed on left  of "Phones" screen.'
         ]

row_2 = ['TC_Android_Auto_0006_Wireless',
         '1. System is ON. 2. User is on "Phones" screen. 3. WAA was connected earlier, currently connection is disconnected.'
         '1.Tap on device in "NOT CONNECTED" section. 2.Check connection status ',
         '1.WAA device connection shall be established. 2.WAA device shall  change non-connection status to connection status.'
         ]

row_3 = ['TC_MFL_000000_GAS_Maps_0057',
         '1. system is on 2. Switch to Guest profile 3. Google Maps app is launched 4. Voice guidance option is set to Muted'
         '1. Start navigation and voice guidance is being played 2. Adjust the volume button on faceplate'
         '1. Voice guidance is being played 2. Volume adjustment quick status pane is displayed and voice guidance is suppressed'
         ]

rows = [row_1, row_2, row_3]

print(matcher_split(['guest'], row_1, [1]))
