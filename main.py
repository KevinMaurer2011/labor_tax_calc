from openpyxl import load_workbook

workbook = load_workbook(filename='rates.xlsx')


zip_code = input('Zip Code?')
# zip_code = '00501'
zipint = int(zip_code)
sheet = workbook['page1']
ziprow = 0

for i in sheet['A']:
    if i.value == zipint:
        ziprow = i.row


print('Found in spreadsheet row: ' + str(ziprow))
print('Body:  ' + str(sheet.cell(row=ziprow,column=2).value))
print('Paint: ' + str(sheet.cell(row=ziprow,column=3).value))
print('Mech:  ' + str(sheet.cell(row=ziprow,column=4).value))
print('Frame: ' + str(sheet.cell(row=ziprow,column=5).value))
print('P & M: ' + str(sheet.cell(row=ziprow,column=6).value))


