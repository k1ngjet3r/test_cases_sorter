from src.matcher import *
import json

def json_directory(json_name):
    with open('json\\' + json_name) as f:
        return json.load(f)

keywords = json_directory('keywords.json')
locked_tcid = json_directory('locked_tcid.json')

class Div:
    def __init__(self, cell_data):
        self.cell_data = cell_data
        self.pre_index = 5

    def keywords_with_range(self, sheet_name, number, type):
        if type == 'range':
            for kw in keywords[sheet_name]:
                for i in range(1, number):
                    if matcher_slice(kw, self.cell_data[self.pre_index + i]):
                        return True

        elif type == 'single':
            for kw in keywords[sheet_name]:
                if matcher_slice(kw, self.cell_data[self.pre_index + number]):
                    return True


    def based_on_tcid(self, sheet_name):
        tcid = [i.lower() for i in self.cell_data[0].split('_')]
        if keywords[sheet_name] in tcid:
            return True

    def locked_tcid(self, id_list):
        for cata in id_list:
            if self.cell_data[0].lower() in [tcid.lower() for tcid in id_list[cata]]:
                return cata

    def directing(self):



