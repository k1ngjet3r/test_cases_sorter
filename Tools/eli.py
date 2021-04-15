from openpyxl import Workbook, load_workbook


class eliminator:
    def __init__(self, partial_list, full_list, tcid_only=True):
        self.partial_list = load_workbook(partial_list)['Main']
        self.full_list = load_workbook(full_list)['Production']
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

if __name__ == '__main__':
    eliminator('MY22_Scope.xlsx', 'MY22_Scope.xlsx').differentiator_counter()