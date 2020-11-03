from openpyxl import load_workbook
from openpyxl import Workbook


class tc_sorter:
    def __init__(self, input_file, output_file):
        self.input_file = input_file
        self.output_file = output_file
