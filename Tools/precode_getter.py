from openpyxl import load_workbook

wb = load_workbook('MY22_intersect_cases.xlsx').active

precode_list = []

for i in wb.iter_rows(max_col=5, values_only=True):
    if i[0] == 'Phone_projection_1':
        for j in i[-1].split(','):
            if j not in precode_list:
                precode_list.append(j)

print(len(precode_list))

s = str(precode_list)

s = s.replace('[','')
s = s.replace(']','')
s = s.replace("'", '')
s = s.replace(', ', '\n')
print(s)