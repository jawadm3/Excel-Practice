from openpyxl import workbook
from openpyxl.styles import PatternFill

print('\n do you want to create a new workbook or edit an existing one: ')
option = input ('type new or existing')

def values(a,b):
       

def newinputs(rows,columns):
     workbook = workbook()
     sheetname = input('enter the name of the sheet')
     workbook['Sheet'].title = sheetname
     print(f'\n title of the workbook has been updated')
     sheet1 = workbook.active
     print ('name of the headings')
     for i in range (columns):
          headings = input(f'heading at row (1,{i}): ')
          sheet1.cell(row=1, column=i).value = headings
     print('added headings')
     print('values for each cell')
     for i in range (rows):
        for j in range (columns):
                  


if option == 'new':
      rows = int(input('how many rows do you want'))
      columns = int(input('how many columns do you want'))

elif option == 'existing':
     workbook_name=input('enter the workbook name with the path and extension')

else:
    print('entered a wrong option, enter either new or exsting')