from openpyxl import load_workbook
from openpyxl import Workbook

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

total = 0
for row in sheet.rows:
    total += 1

print(total)
