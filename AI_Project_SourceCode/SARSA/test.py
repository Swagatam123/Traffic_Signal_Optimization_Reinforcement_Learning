import openpyxl

file=openpyxl.load_workbook('QMatrix.xlsx')
sheet=file.get_sheet_by_name('Sheet1')
sheet['A1']=100
file.save('QMatrix.xlsx')