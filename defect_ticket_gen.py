from openpyxl import load_workbook
from docx import Document

# Initialized the excel file which containing the detail
sheet = load_workbook('Defect_Tickets.xlsx').active

for row in sheet.iter_rows(max_col=11, values_only=True):
    row_data = [value for value in row]
    if row_data[0] == 'Summary':
        pass
    elif row_data[0] is None:
        break
    else:
        # Open up a blank document with the default template
        ticket = Document()

        summary = ticket.add_paragraph('')
        summary.add_run('[Summary]').bold = True
        ticket.add_run(str(row_data[0]))
        ticket.add_paragraph(' ')

        precondiction = ticket.add_paragraph('')
        precondiction.add_run('[Precondiction]').bold = True

        testing_type = ticket.add_paragraph('')
        testing_type.add_run('Testing Type: ').bold = True
        testing_type.add_run(str(row_data[1]))

        connected_devices = ticket.add_paragraph('')
        connected_devices.add_run('Connected Devices ').bold = True
        connected_devices.add_run(str(row_data[2]))
        ticket.add_paragraph(str(row_data[3]))
        ticket.add_paragraph(' ')

        test_steps = ticket.add_paragraph('')
        test_steps.add_run('[Test Steps]').bold = True
        ticket.add_paragraph(str(row_data[4]))

        ticket.save('test.docx')
