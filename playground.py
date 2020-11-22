from openpyxl import load_workbook
from openpyxl import Workbook
import re

flash_user = []
press_button = []
invalid = []
Driver_In_Off = []
Driver_In_On = []
Driver_Out_Off = []
Driver_Out_On = []
Guest_In_Off = []
Guest_In_On = []
Guest_Out_Off = []
Guest_Out_On = []

# keyword for layer 1
flash_user = ['flash']
press_button = ['long press', 'short press', 'press "end" key']
invalid = ['audiobook', 'sxm']

exceptions_names = ['flash_user', 'press_button', 'invalid']
exceptions_items = [flash_user, press_button, invalid]

exp_dict = {name: item for name,
            item in zip(exceptions_names, exceptions_items)}

# Layer 2
sign_out = ['sign out', 'sign-out', 'signout', 'signed out']

offline = ['offline']

guest = ['guest']

ac = ['a/c', 'temperature', 'climate', 'defroster', 'air', 'fan']


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
    def __init__(self, cell_data):
        self.cell_data = cell_data

    def

    def phone_type(self, cell_data):
        iphone = ['iphone', 'cp', 'wcp']
        android = ['android', 'aa', 'waa']
        if matcher_split(iphone, cell_data, [1, 2, 3]):
            cell_data.append("iPhone")
        elif matcher_split(android, cell_data, [1, 2, 3]):
            cell_data.append('Android')
        elif matcher_split(iphone, cell_data, [1, 2, 3]) and matcher_split(android, cell_data, [1, 2, 3]):
            cell_data.append('Both')
        else:
            cell_data.append(' ')

    def sign_status(self, cell_data):
        sign_out = ['sign out', 'sign-out', 'signout', 'signed out']
        if matcher_slice(sign_out, cell_data, [1]):
            cell_data.append('sign out')
        else:
            cell_data.append('sign in')

    def connection(self, cell_data):
        offline = ['offline']
        if matcher_split(offline, cell_data, [1]):
            cell_data.append('offline')
        else:
            cell_data.append('online')

    def user(self, cell_data):
        guest = ['guest']
        others = ['secondary', 'user 1', 'user 2', 'user1', 'user2']
        if matcher_split(guest, cell_data, [1]):
            cell_data.append('guest')
        elif matcher_slice(others, cell_data, [1]):
            cell_data.append('others')
        else:
            cell_data.append('Driver')

    def sorting(self, cell_data):
        pass


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
