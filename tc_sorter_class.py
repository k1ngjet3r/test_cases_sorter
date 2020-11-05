from openpyxl import load_workbook
from openpyxl import Workbook
import re

categories = ['Flash User', 'Multi User', 'Button', 'Invalid Cases',
              'Driver/In/Off', 'Driver/In/On', 'Driver/Out/Off', 'Driver/Out/On', 'Guest/In/Off', 'Guest/In/On', 'Guest/Out/Off', 'Guest/Out/On']

# keyword for layer 1
flash_user = ['flash']
mulit_user = ['multi', 'primary' 'secondary']
press_button = ['long press', 'short press', 'press "end" key']
user = ['guest', 'driver']
invalid = ['audiobook']

exceptions = [flash_user, mulit_user, press_button, invalid]

# Layer 2
sign_out = ["user is signed out",
            "signout the google account", "sign out the google account"]

online = ['online', 'internet']

# keyword for determining the function-related categories
navigation = ['navigat', 'go to', 'add stop', 'guidance',
              'how far', 'take me', 'address', 'traffic', 'add stop']

call_SMS = ['call', 'phone', 'message',
            'reply', 'text', 'sms', 'dial', 'Send', 'dail']

media = ['play', 'pause', 'next', 'previous',
         'volume', 'music', 'am', 'fm', 'radio', 'news', 'tune', 'tuned', 'plays', 'bluetooth', 'station', 'album', 'podcast', 'pauses', 'media', 'songs', 'playback', 'bt']


ac = ['a/c', 'temperature', 'climate', 'defroster', 'air', 'fan']


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

    def read_create_file(self):
        workbook = load_workbook(self.input_name)
        sheet = workbook.active
        sheet.delete_rows(1)
        wb = Workbook()
        sorted_cases = wb.active
        for category in self.categories:
            sorted_cases.create_sheet(category, int(
                (self.categories).index(category)))
        return sheet, sorted_cases

    def remove_exceptions(self, exceptions):
        for exp in exceptions:
