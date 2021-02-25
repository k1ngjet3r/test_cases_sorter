from openpyxl import load_workbook, Workbook

class Auto_gen:
    def __init__(self, file_name):
        self.file_name = load_workbook(file_name)
        