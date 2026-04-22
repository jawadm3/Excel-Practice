import openpyxl as xl
workbook = xl.load_workbook('raw_data.xlsx')
sheets = workbook.sheetnames

print(f'\nList of the sheets in the workbook ', sheets)
print(f'\nActive sheet in the workbook ',workbook.active.title)

print(f'\nValue at A1 {workbook['Sheet1']['A1'].value}')

print (f'\nvalue at 3rd row and 1st col',workbook['Sheet1'].cell(3,1).value)

sheet1 = workbook['Sheet1']
rows = sheet1.max_row
columns = sheet1.max_column
print(f'\nnumber of rows ',rows,'number of columns ',  columns, end='\n\n')



for i in range (1,rows+1):
     row_values = [sheet1.cell(i, j).value for j in range(1, columns + 1)]
     print(row_values)

