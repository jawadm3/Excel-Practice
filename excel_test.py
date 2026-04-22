import xlwings as xw

wb = xw.Book()  # opens a new Excel workbook
sheet = wb.sheets[0]

sheet["A1"].value = "Hello from Python!"
sheet["A2"].value = 123

print("Done!")
