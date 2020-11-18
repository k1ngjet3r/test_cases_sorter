from openpyxl import load_workbook
from openpyxl import Workbook
import re

sheet_names = ['cluster', 'press_button', 'warren', 'flash', 'bench',
               'Driver_In_Off', 'Driver_In_On', 'Driver_Out_Off', 'Driver_Out_On',
               'Guest_In_Off', 'Guest_In_On', 'Guest_Out_Off', 'Guest_Out_On']

# expection cases
cluster = ['cluster', 'swc']
press_button = ['long press', 'short press', 'press "end" key', 'ignition']
warren = ['audiobook', 'sxm']
flash = ['flash']
exp_names = ['cluster', 'press_button', 'warren', 'flash']
exp_items = [cluster, press_button, warren, flash]
exp_dict = {name: item for name, item in zip(
    exp_names, exp_items)}

# bench only cases
ac = ['degree', 'temperature', 'climate',
      'defroster', 'air', 'fan', 'recirculation', 'vent']
cluster = ['cluster', 'swc']

sign_out = ['sign out', 'signed out', 'log out', 'signout', 'sign-out']

offline = ['offline']

guest = ['guest']

iphone = ['iphone', 'carplay', 'acp']
android = ['android', 'waa']
phone_names = ['iphone', 'android']
phone_keywords = [iphone, android]
phone_type = {name: keyword for name,
              keyword in zip(phone_names, phone_keywords)}


def matcher_slice(keywords, cell_data, index_range):
    for i in index_range:
        sen = cell_data[i].lower()
        for key in keywords:
            if re.search(key, sen) != None:
                return True
    return False


def matcher_split(keywords, cell_data, index_range):
    for i in index_range:
        clean_sen = re.sub(r'[^|w]', ' ', cell_data[i].lower())
        world_list = clean_sen.split()
        for key in keywords:
            if key in world_list:
                return True
    return False


class Tc_sorter:
    def __init__(self, input_name, output_name, sheet_names):
        self.input_name = str(input_name)
        self.output_name = str(output_name)
        self.sheet_names = sheet_names
        self.sheet = (load_workbook(self.input_name)).active
        self.wb = Workbook()
        self.wb.active
        for name in self.sheet_names:
            self.wb.create_sheet(name, int(
                (self.sheet_names).index(name)))

    def phone_determination(self, cell_data):
        if matcher_split(iphone, cell_data, [1, 2, 3]) and matcher_split(android, cell_data, [1, 2, 3]):
            cell_data.append('Both')

        elif matcher_split(iphone, cell_data, [1, 2, 3]):
            cell_data.append('iPhone')

        elif matcher_split(android, cell_data, [1, 2, 3]):
            cell_data.append('Android')

        else:
            cell_data.append(' ')

    def sorting(self):
        sheet = self.sheet
        for row in sheet.iter_rows(max_col=4, values_only=True):
            cell_data = [cell for cell in row]

            self.phone_determination(cell_data)

            # Filter out the exception cases
            if matcher_split(cluster, cell_data, [1, 2, 3]):
                self.wb['cluster'].append(cell_data)

            elif matcher_split(press_button, cell_data, [2]):
                self.wb['multi_user'].append(cell_data)

            elif matcher_split(warren, cell_data, [1, 2, 3]):
                self.wb['warren'].append(cell_data)

            elif matcher_slice(flash, cell_data, [1, 2]):
                self.wb['flash'].append(cell_data)

            else:
                # determine the Driver as user
                if matcher_split(guest, cell_data, [1]) == False:
                    # determine sign-in
                    if matcher_split(sign_out, cell_data, [1]) == False:
                        # determine online
                        if matcher_split(offline, cell_data, [1]) == False:
                            self.wb['Driver_In_On'].append(cell_data)
                        # determine offline
                        elif matcher_split(offline, cell_data, [1]):
                            self.wb['Driver_In_Off'].append(cell_data)

                    # determine sign-out
                    elif matcher_split(sign_out, cell_data, [1]):
                        # determine online
                        if matcher_split(offline, cell_data, [1]) == False:
                            self.wb['Driver_Out_On'].append(cell_data)
                        # determine offline
                        elif matcher_split(offline, cell_data, [1]):
                            self.wb['Driver_Out_Off'].append(cell_data)

                # determine the Guest as user
                elif matcher_split(guest, cell_data, [1]):
                    # determine sign-in
                    if matcher_split(sign_out, cell_data, [1]) == False:
                        # determine online
                        if matcher_split(offline, cell_data, [1]) == False:
                            self.wb['Guest_In_On'].append(cell_data)
                        # determine offline
                        elif matcher_split(offline, cell_data, [1]):
                            self.wb['Guest_In_Off'].append(cell_data)

                    # determine sign-out
                    elif matcher_split(sign_out, cell_data, [1]):
                        # determine online
                        if matcher_split(offline, cell_data, [1]) == False:
                            self.wb['Guest_Out_On'].append(cell_data)
                        # determine offline
                        elif matcher_split(offline, cell_data, [1]):
                            self.wb['Guest_Out_Off'].append(cell_data)

        self.wb.save(self.output_name)


sorted_file = Tc_sorter('TestCaseDetails_W46.xlsx',
                        'sorted_w46.xlsx', sheet_names)
sorted_file.sorting()
