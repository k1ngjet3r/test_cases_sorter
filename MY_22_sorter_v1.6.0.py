from openpyxl import load_workbook
from openpyxl import Workbook
import re
import json

# the index of the precondition
pre_index = 5


def json_directory(json_name):
    # directory = 'C:\\Users\\Jeter\\OneDrive\\Documents\\GitHub\\test_cases_sorter\\json_file\\'
    directory = '/Users/jeter/Documents/GitHub/test_cases_sorter/json_file/'

    with open(directory + json_name) as f:
        return json.load(f)


data_sheet = json_directory('sheet_related.json')

keywords = json_directory('keywords.json')

auto_case_list = json_directory('auto_case_id.json')

# loading other list
# logan_list_sheet = load_workbook('logan_list.xlsx').active
# logan_list = [r[0] for r in logan_list_sheet.iter_rows(
#     max_col=1, max_row=335, values_only=True)]


def matcher_slice(keywords, cell_data):
    sen = str(cell_data).replace('"', '').lower()
    for key in keywords:
        if re.search(key, sen):
            return True
    return False


def matcher_split(keywords, cell_data):
    clean_sentance = re.sub(r'[^\w]', ' ', str(cell_data).lower())
    word_list = clean_sentance.split()
    for key in keywords:
        if key in word_list:
            return True
    return False


class Tc_sorter:
    def __init__(self, test_case_list, output_name, last_week, continue_from=False):
        print('Initiallizing...')
        self.test_case_list = str(test_case_list)
        self.output_name = str(output_name)
        self.last_week = str(last_week)
        self.sheet = (load_workbook(self.test_case_list)).active
        print('{} loaded successfully'.format(self.test_case_list))

        # Loading the resut from last week
        self.last_week_result = (
            load_workbook(self.last_week))
        print('{} loaded successfully'.format(self.last_week))

        if continue_from == False:
            self.wb = Workbook()
            # self.wb.active
            for name in data_sheet['sheet_names']:
                self.wb.create_sheet(
                    name, int((data_sheet['sheet_names']).index(name)))
                self.wb[name].append(data_sheet['titles'])
            for fail_name in data_sheet['fail_case_sheet']:
                self.wb.create_sheet(fail_name, -1)
                self.wb[fail_name].append(data_sheet['fail_case_titles'])
            print('Output file initiallized')

        else:
            self.wb = load_workbook(self.output_name)

    def Automation_cases(self):
        auto_file = load_workbook('automation_cases.xlsx').active
        return [tcid[0] for tcid in auto_file.iter_rows(max_col=1, values_only=True)]

    def cell_data(self, row):
        cells = []
        for cell in row:
            if cell is None:
                cells.append('none')
            else:
                cells.append(cell)
        return cells

    def phone_type(self, cell_data):
        iphone = keywords['iphone']
        android = keywords['android']
        phone_requirement = [0, 0]
        for cell in cell_data[pre_index:pre_index+2]:
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
        sign_out = keywords['sign_out']
        if matcher_slice(sign_out, cell_data[pre_index]):
            cell_data.append('sign_out')
        else:
            cell_data.append('sign_in')

    def connection(self, cell_data):
        offline = keywords['offline']
        if matcher_split(offline, cell_data[pre_index]):
            cell_data.append('Offline')
        else:
            cell_data.append('Online')

    def formatter(self, cell_data):
        for _ in range(4):
            cell_data.insert(1, '')

    def user(self, cell_data):
        guest = keywords['guest']
        non_guest = keywords['non_guest']
        others = keywords['others']
        primary = keywords['primary']
        if (matcher_split(guest, cell_data[pre_index]) or matcher_split(guest, cell_data[pre_index+3])) and matcher_slice(non_guest, cell_data[pre_index]) is False:
            cell_data.append('Guest')
        elif matcher_slice(others, cell_data[pre_index]) or matcher_slice(non_guest, cell_data[pre_index]) or matcher_slice(others, cell_data[pre_index+3]):
            cell_data.append('Others')
        elif matcher_split(guest, cell_data[pre_index+1]) and (matcher_slice(others, cell_data[pre_index+1]) or matcher_split(primary, cell_data[pre_index+1])):
            cell_data.append('multiple')
        else:
            cell_data.append('Driver')

    def bench_only(self, cell_data):
        press_button = keywords['push_button']
        cluster = keywords['cluster']
        speed_limit = keywords['speed_limit']
        expection = keywords['expection']
        for cell in cell_data[pre_index+1:pre_index+5]:
            if (matcher_slice(press_button, cell) or matcher_slice(cluster, cell) or matcher_slice(speed_limit, cell)) and not matcher_slice(expection, cell):
                return True
            return False

    def ac_only(self, cell_data):
        ac = keywords['ac']
        ac_split = keywords['ac_split']
        for cell in cell_data[pre_index:pre_index+3]:
            if matcher_slice(ac, cell) or matcher_split(ac_split, cell):
                return True
            return False

    def tc_location_dict(self):
        tc_location = load_workbook('TC_location.xlsx').active
        # stored the data in a dictionary {test_case: location}
        return {TCID: location for (TCID, location) in tc_location.iter_rows(
            max_col=2, values_only=True) if TCID is not None}

    def last_week_result_dict(self):
        last_week_dict = {}
        last_week = self.last_week_result
        for sheet in last_week.worksheets:
            if sheet.title not in ['Fail Cases China', 'Cases need update', 'Summary']:
                for last_week_row in sheet.iter_rows(max_col=5, values_only=True):
                    last_week_cell = self.cell_data(last_week_row)
                    if last_week_cell[0] != 'Original GM TC ID':
                        last_week_dict[last_week_cell[0]] = last_week_cell[1:]
        return last_week_dict

    def nav_case(self, cell_data):
        # Finding the navigation-related cases using TCID
        # Formatting the TCID
        tcid = [i.lower() for i in cell_data[0].split('_')]
        if 'maps' in tcid:
            return True
        return False

    def call_SMS(self, cell_data):
        callsms = keywords['call_sms']
        if matcher_slice(callsms, cell_data[pre_index+1]):
            return True
        return False

    def fuel_sim(self, cell_data):
        fuel = keywords['fuel_sim']
        if matcher_slice(fuel, cell_data[pre_index+1]):
            return True
        return False

    def sorting(self):
        print('Opening a new sheet...')
        sheet = self.sheet
        print('Last week result loaded successfully')
        # difficult_cases_list = self.difficult_cases()
        print('Difficult case list generated')
        location_dict = self.tc_location_dict()
        print('Test case location dictionary generated')
        print('Iterating through the test plan......')

        k = 1

        # Iterate through the unprocessd test cases
        # Only getting the first 5 values of each row (tc, precondition, test_steps, expected_result, test_objective}
        for row in sheet.iter_rows(max_col=5, values_only=True):
            print('Iterate case no. {}'.format(k))
            # turn the data into a list
            cell_data = []
            for cell in row:
                if cell is not None:
                    cell_data.append(cell)
                else:
                    cell_data.append('none')

            # adding 'pass/fail', 'Tester', 'Automation_comment', 'bug ID', 'Note' to the list
            self.formatter(cell_data)
            # determine the phone type
            self.phone_type(cell_data)
            # determine the user type
            self.user(cell_data)
            # determine online/offline
            self.connection(cell_data)
            # determine sign-in/sign-out
            self.sign_status(cell_data)
            k += 1

            if cell_data[0] in location_dict:
                cell_data.append(location_dict[cell_data[0]])
            else:
                cell_data.append('none')

            last_week_dict = self.last_week_result_dict()

            for i in range(4):
                if cell_data[0] in last_week_dict:
                    cell_data.append(last_week_dict[cell_data[0]][i])
                else:
                    continue

            # Forming the final format
            cell_data = cell_data[:5] + [cell_data[-1]] + cell_data[5:-1]

            # Distributing the test case to the desinated sheet

            # Append the case to "difficult_cases" sheet based on last week's result
            # if cell_data[0] in logan_list:
            #     self.wb['Logan'].append(cell_data)

            if cell_data[-3] == 'Fail':
                self.wb['Difficult_cases'].append(cell_data)

            # Append the case to "auto" if the case ID is in the "auto_case_id.json"
            elif cell_data[0] in auto_case_list['auto'] or cell_data[0] in auto_case_list['fuel_sim']:
                self.wb['auto'].append(cell_data)

            elif self.nav_case(cell_data):
                self.wb['Nav'].append(cell_data)

            # elif self.fuel_sim(cell_data):
            #     self.wb['Fuel_sim'].append(cell_data)

            elif self.bench_only(cell_data):
                self.wb['Bench_only'].append(cell_data)

            elif self.call_SMS(cell_data):
                self.wb['Call&SMS'].append(cell_data)

            # elif self.ac_only(cell_data):
            #     self.wb['ac_only'].append(cell_data)

            else:
                i = 11
                if cell_data[i] == 'Driver' and cell_data[i+1] == 'Online' and cell_data[i+2] == 'sign_in':
                    self.wb['Driver_Online_In'].append(cell_data)
                elif cell_data[i] == 'Driver' and cell_data[i+1] == 'Online' and cell_data[i+2] == 'sign_out':
                    self.wb['Driver_Online_Out'].append(cell_data)
                elif cell_data[i] == 'Driver' and cell_data[i+1] == 'Offline' and cell_data[i+2] == 'sign_in':
                    self.wb['Driver_Offline_In'].append(cell_data)
                elif cell_data[i] == 'Driver' and cell_data[i+1] == 'Offline' and cell_data[i+2] == 'sign_out':
                    self.wb['Driver_Offline_Out'].append(cell_data)

                elif cell_data[i] == 'Guest' and cell_data[i+1] == 'Online' and cell_data[i+2] == 'sign_in':
                    self.wb['Guest_Online_In'].append(cell_data)
                elif cell_data[i] == 'Guest' and cell_data[i+1] == 'Online' and cell_data[i+2] == 'sign_out':
                    self.wb['Guest_Online_Out'].append(cell_data)
                elif cell_data[i] == 'Guest' and cell_data[i+1] == 'Offline' and cell_data[i+2] == 'sign_in':
                    self.wb['Guest_Offline_In'].append(cell_data)
                elif cell_data[i] == 'Guest' and cell_data[i+1] == 'Offline' and cell_data[i+2] == 'sign_out':
                    self.wb['Guest_Offline_Out'].append(cell_data)
                else:
                    self.wb['Other'].append(cell_data)

        print('Saving the file named {}'.format(self.output_name))
        self.wb.save(self.output_name)


testing = Tc_sorter('W10_cases_related.xlsx',
                    'W12_sorted.xlsx', 'W10_sorted.xlsx', continue_from=False
                    )

testing.sorting()
