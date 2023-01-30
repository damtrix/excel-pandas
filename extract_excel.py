from openpyxl import load_workbook, workbook

book = load_workbook('01092022 - 30092022.xlsx')
sheet1 = book['Blank A4 Landscape']

print("Started")

sheet1.delete_cols(8)
sheet1.delete_cols(12, 10)

book.save('01092022 - 30092022.xlsx')
print("Done")

