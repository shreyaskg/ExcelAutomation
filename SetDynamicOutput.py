import openpyxl
from GetStandardCollegeNamedistance import GetStandardName
from GetStandardCollegeNamedistance import GetDistance
from GetStandardCollegeNamedistance import GetParticipants

def SetColleges(file_path):
    write = openpyxl.load_workbook(filename = file_path)
    o_sheet = write.get_sheet_by_name("Callsheet")
    data = []
    for i in range(2,o_sheet.max_row+3):
        data.append(o_sheet[i][4].value)
    for data in data:
        print(data)
    print(len(data))
SetColleges('output.xlsx')

