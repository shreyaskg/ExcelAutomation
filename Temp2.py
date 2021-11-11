import openpyxl
from GetAge import GetDays
x = openpyxl.load_workbook(filename='Week 36.xlsx')
x = x['Callsheet']
print(GetDays(x[2][].value))
