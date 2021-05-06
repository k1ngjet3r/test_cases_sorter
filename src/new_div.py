from src.matcher import *
import json

def json_directory(json_name):
    with open('json\\' + json_name) as f:
        return json.load(f)

keywords = json_directory('keywords.json')
locked_tcid = json_directory('locked_tcid.json')
auto_package = json_directory('auto_packages.json')
sheetnames = json_directory('sheetname.json')
auto_cases = json_directory('auto_case_id.json')
status_list = json_directory('status.json')

class Div:
    def __init__(self, cell_data):
        self.cell_data = cell_data
        self.pre_index = 5

    def keywords_with_range(self, sheet_name, keyword_list, number, type):
        if type == 'range':
            for keyword in keyword_list:
                for kw in keywords[keyword]:
                    for i in range(1, number):
                        if matcher_slice(kw, self.cell_data[self.pre_index + i]):
                            return sheet_name

        elif type == 'single':
            for keyword in keyword_list:
                for kw in keywords[keyword]:
                    if matcher_slice(kw, self.cell_data[self.pre_index + number]):
                        return sheet_name


    def based_on_tcid(self, sheet_name):
        tcid = [i.lower() for i in self.cell_data[0].split('_')]
        if keywords[sheet_name] in tcid:
            return True

    def locked_tcid(self, id_list):
        if self.cell_data[0].lower() in [tcid.lower() for tcid in id_list]:
            return id_list

    def directing(self):
        for category in auto_package:
            if self.locked_tcid(auto_package[category]):
                return category

        for category in locked_tcid:
            if self.locked_tcid(locked_tcid[category]):
                return category

        for sheetname in sheetnames:
            if sheetnames[sheetname]['mode'] == "keywords_with_range":
                number = sheetnames[sheetname]['number']
                iter_type = sheetnames[sheetname]['type']
                keyword_list = sheetnames[sheetname]['keyword']

                if self.keywords_with_range(sheetname, keyword_list, number, iter_type) != None:
                    return self.keywords_with_range(sheetname, keyword_list, number, iter_type)

            elif sheetnames[sheetname]['mode'] == 'based_on_tcid':
                if self.based_on_tcid(sheetname) != None:
                    return self.based_on_tcid(sheetname)


    def status_det(self):
        i = 11
        for status in status_list:
            if self.cell_data[i] == status_list[status][0] and self.cell_data[i+1] == status_list[status][1] and self.cell_data[i+2] == status_list[status][2]:
                return status
            elif self.cell_data[i] == 'Guest':
                return 'Guest'
            else:
                return 'Other'

