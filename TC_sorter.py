from openpyxl import load_workbook
from openpyxl import Workbook

# Defining the word matcher function


def matcher(keywords, sentance):
    for key in keywords:
        num_slices = int(len(sentance.value)) + 1 - int(len(key))
        for i in range(num_slices):
            if sentance.value[i: i + len(key)] == key:
                return True
    return False


# Load the unsorted cases file
workbook = load_workbook('Taipei_W45_SelectedCaseList_1030.xlsx')
sheet = workbook.active
# Saving the titles of every column in to a list and delete it for categories determination
titles = []
for title in sheet[1]:
    titles.append(title.value)
sheet.delete_rows(1)

# Create the new file named sorted cases
wb = Workbook()
sorted_cases = wb.active
# Adding worksheets for different categories
categories = ['Button', 'Call(Sign Out)', 'CallSMS', 'Media Center',
              'Projection', 'Phone as Hotspot', 'Navigation', 'HVAC', 'Others']
for category in categories:
    wb.create_sheet(category, int(categories.index(category)))

# keyword for determining the categories
sign_out = ["user is signed out",
            "signout the google account", "sign out the google account"]

call_SMS = ['call', 'phone', 'message',
            'reply', 'text', 'sms', 'dail', 'Message', 'Text', 'Send', 'Call', 'Dail']

media = ['play', 'pause', 'next', 'previous',
         'volume', 'music', 'AM', 'FM', 'radio', 'news', 'Tune', 'Play', 'Bluetooth']

projection = ['projection', 'Projection']

hotspot = ['hotspot']

navigation = ['navigation', 'go to', 'add stop', 'guidance',
              'how far', 'take me', 'navigate', 'address', 'traffic', 'add stop']

ac = ['a/c', 'temperature', 'climate control',
      'defroster', 'air', 'fan', 'A/C', 'Air', 'Temperature']

invalid = ['Audiobook']

press_button = ['long press', 'short press',
                'Long press', 'Short press', 'press "End" Key', 'press "End" key']

# iterate through the row
for row in sheet.rows:
    # Since the cell cannot to copy to other sheet directly, we have to store the data in a list and transfer to other sheet
    cell_data = []
    for cell in row:
        cell_data.append(cell.value)

    # For buttom press related cases
    if matcher(press_button, row[2]) == True:
        wb['Button'].append(cell_data)

    # For HVAC-related cases
    elif (matcher(ac, row[1]) == True or matcher(ac, row[2]) == True) and matcher(sign_out, row[1]) != True:
        wb['HVAC'].append(cell_data)

    # For media-related cases
    elif (matcher(media, row[1]) == True or matcher(media, row[2]) == True) and matcher(sign_out, row[1]) != True:
        wb["Media Center"].append(cell_data)

    # For projection-related cases
    elif (matcher(projection, row[1]) == True or matcher(projection, row[2]) == True) and matcher(sign_out, row[1]) != True:
        wb['Projection'].append(cell_data)

    # For hotspot-related cases
    elif (matcher(hotspot, row[1]) == True or matcher(hotspot, row[2]) == True) and matcher(sign_out, row[1]) != True:
        wb['Phone as Hotspot'].append(cell_data)

    # For navigation-related cases
    elif (matcher(navigation, row[1]) == True or matcher(navigation, row[2]) == True) and matcher(sign_out, row[1]) != True:
        wb['Navigation'].append(cell_data)

    # For callsms-related cases
    elif (matcher(call_SMS, row[1]) == True or matcher(call_SMS, row[2]) == True) and matcher(sign_out, row[1]) != True:
        wb['CallSMS'].append(cell_data)

    # For call(sign out)-related cases
    elif (matcher(call_SMS, row[1]) == True or matcher(call_SMS, row[2]) == True) and matcher(sign_out, row[1]) == True:
        wb['Call(Sign Out)'].append(cell_data)

    else:
        wb['Others'].append(cell_data)

# Save the file with the name
wb.save('sorted_test_cases.xlsx')
