from openpyxl import Workbook, load_workbook


class eliminator:
    def __init__(self, wb, tcid_only=True):
        self.name = wb
        self.wb = load_workbook(wb)
        self.partial_list = self.wb['Production']
        self.full_list = self.wb['Signed_case_list']
        if tcid_only == True:
            self.max_column = 1
        else:
            self.max_column = 5

    def differentiator(self):
        wb = Workbook()
        wb.create_sheet('intersection')
        wb.create_sheet('symmetric difference')

        partial_tcid = [row[0] for row in self.partial_list.iter_rows(
            max_col=1, values_only=True)]

        for row in self.full_list.iter_rows(max_col=self.max_column, values_only=True):
            row_data = [i for i in row]
            if row_data[0] in partial_tcid:
                wb['intersection'].append(row_data)
            else:
                wb['symmetric difference'].append(row_data)

        wb.save('cases.xlsx')

    def differentiator_counter(self):
        current = 1
        matched = 0
        not_matched = 0
        partial_tcid = [row[0] for row in self.partial_list.iter_rows(
            max_col=1, values_only=True)]

        for row in self.full_list.iter_rows(max_col=1, values_only=True):
            print('iterate case number {}'.format(current))
            try:
                current +=1
                if row[0] in partial_tcid:
                    matched += 1
                else:
                    not_matched += 1
            except:
                break
        print('Matched: {}'.format(matched))
        print('Not Matched: {}'.format(not_matched))

    def differentiator_same_sheet(self):
        self.wb.create_sheet('Production_with_script')
        current = 1
        matched = 0

        partial_tcid = [row[0].lower() for row in self.partial_list.iter_rows(
            max_col=1, values_only=True)]

        for row in self.full_list.iter_rows(max_col=1, values_only=True):
            try:
                print('iterate case number {}'.format(current))
                current += 1
                if row[0].lower() in partial_tcid:
                    matched += 1
                    self.wb['Production_with_script'].append(row)
                    self.wb.save(self.name)
            except:
                break
        # self.wb.save(self.name)
        print('Matched: {}'.format(matched))
                

            

if __name__ == '__main__':
    eliminator('MY22_Scope.xlsx').differentiator_same_sheet()