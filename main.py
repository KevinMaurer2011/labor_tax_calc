from openpyxl import load_workbook

workbook = load_workbook(filename='rates.xlsx')

zip_code = input('Zip Code?')
# zip_code = '00501'
zipint = int(zip_code)
labor = workbook['labor']
tax = workbook['tax']
ziprow = 0
taxrow = 0

for i in labor['A']:
    if i.value == zipint:
        ziprow = i.row

for i in tax['B']:
    if i.value == zipint:
        taxrow = i.row

try:
    combined_tax = tax.cell(row=taxrow, column=4).value
    tax_percent = (float(combined_tax)) * 100
except:
    print('No tax found sorry')

try:
    print('Body:  ' + str(labor.cell(row=ziprow, column=2).value))
    print('Paint: ' + str(labor.cell(row=ziprow, column=3).value))
    print('Mech:  ' + str(labor.cell(row=ziprow, column=4).value))
    print('Frame: ' + str(labor.cell(row=ziprow, column=5).value))
    print('P & M: ' + str(labor.cell(row=ziprow, column=6).value))
except:
    print('No labor rates found')

try:
    print('Tax: ' + str(tax_percent))
except:
    pass

labor_taxed_states = (
'AR', 'CT', 'FL', 'HI', 'IA', 'KS', 'MD', 'MN', 'MS', 'NE', 'NJ', 'NY', 'OH', 'PA', 'SD', 'TX', 'WV', 'WI')

if tax.cell(row=taxrow, column=1).value in labor_taxed_states:
    print('Labor Taxed: YES!')
else:
    print('Labor Taxed: No')
