from openpyxl import load_workbook, Workbook
from openpyxl.worksheet.worksheet import Worksheet

class pkg_ctrl:
    def __init__(self, package_design, case_list, output_name):
        self.package_design = load_workbook(package_design).active
        self.case_list = load_workbook(case_list).active
        self.output_file = Workbook()
        self.output_name = output_name
        self.sheetname = []
    
    def case_list_2_list(self):
        id_list = []
        for id in self.case_list.iter_rows(max_col=1, values_only=True):
            if id[0] != None:
                id_list.append(id[0].lower())
        return id_list

    def comparision(self):
        wb = self.output_file.active
        id_list = self.case_list_2_list()
        for case in self.package_design.iter_rows(max_col=6, values_only=True):
            if case[1].lower() in id_list:
                wb.append(case)
            self.output_file.save(self.output_name)
                

pkg = pkg_ctrl('Function Package Design.xlsx', 'MY22_scripts.xlsx', 'Auto_trial.xlsx')
pkg.comparision()