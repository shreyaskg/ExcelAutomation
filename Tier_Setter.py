import ast
import json
import openpyxl
import re
wb = openpyxl.load_workbook('output.xlsx')
open_sheet = wb['Final Data']
college = open_sheet['E']
college_list = []
for i in range(1,open_sheet.max_row):
    college_list.append(college[i].value)
print(college_list)
with open('rankings_university.txt','r') as f:
    data = f.read()
# Converting from string to python dictionary
data = ast.literal_eval(data)
data_write = {}
for keys in data.keys():
    if data[keys].isdigit():
        ranking = int(data[keys])
        if ranking <= 200:
            data_write[keys]= 'Tier 1'
        elif ranking in range(201,401):
            data_write[keys] = 'Tier 2'
        else:
            data_write[keys] = 'Tier 3'

# Let us call an api which will return us

with open('UniversityTier.txt','w') as f:
    f.write(json.dumps(data_write))
