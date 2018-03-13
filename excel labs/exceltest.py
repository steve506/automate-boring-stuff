import openpyxl


wb = openpyxl.load_workbook('example.xlsx')

print(wb.sheetnames)

sheet = wb['Sheet1']

print(type(sheet))

# tuple(sheet['A1':'C3'])

for rowOfCellObjects in sheet['A1':'C3']:
    for cellObj in rowOfCellObjects:
        print(str(cellObj.coordinate), str(cellObj.value))
    print('--- END OF ROW ---')

for cell in sheet['B']:
    print(cell.value)
