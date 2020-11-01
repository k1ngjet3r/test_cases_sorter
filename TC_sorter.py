from openpyxl import load_workbook
from openpyxl import Workbook

# Defining the word matcher function


def matcher(keywords, sentance):
    for key in keywords:
        num_slices = int(len(sentance)) + 1 - int(len(key))
        for i in range(num_slices):
            if sentance[i: i + len(key)] == key:
                return True
    return False


# Load the unsorted cases file
workbook = load_workbook('cases.xlsx')
sheet = workbook.active

# Create the new file named sorted cases
wb = Workbook()
sorted_cases = wb.active
# Adding worksheets for different categories
categories = ['Call(Sign Out)', 'CallSMS', 'Media Center',
              'Projection', 'Phone as Hotspot', 'Navigation', 'HVAC', 'Others']
for category in categories:
    wb.create_sheet(category, int(categories.index(category)))

# keyword for determining the categories
sign_out = ["user is signed out",
            "signout the google account", "sign out the google account"]

call_SMS = ['call', 'phone', 'message',
            'reply', 'text', 'sms', 'phone', 'dail']

media = ['play', 'pause', 'next', 'previous']

projection = ['projection']

hotspot = ['hotspot']

navigation = ['navigation', 'go to', 'add stop', 'guidance']

ac = ['a/c', 'temperature', 'climate control', 'defroster', 'air', 'fan']

# determine the cases of the media center
