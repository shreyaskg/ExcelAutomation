import openpyxl
import datetime
import pandas as pd
from datetime import date

# storing today's date in the variable today
def FirstIteration(file):
    skeletal_file = 'skeletal_file.xlsx'
    output = 'output.xlsx'
    today = datetime.datetime.now()
    # filename argument should contain the file path if in different directory
    wb = openpyxl.load_workbook(filename = file)
    # opening the file to be written into
    wb_write = openpyxl.load_workbook(filename = skeletal_file)


    # Loading the information in the week2.xlsx file


    counter = 0
    # get the reference to the required sheet in the xlsx file
    o_sheet = wb.get_sheet_by_name("Callsheet")


    write_sheet = wb_write.get_sheet_by_name("Final Data")
    # Loading the information in the week1.xlsx file
    data = []
    for i in range(2,o_sheet.max_row + 3):
        individual_data = []
        for column in o_sheet[i]:
            individual_data.append(column.value)
        data.append(individual_data)
        counter += 1
    # print(data)
    #

    weeks = 1

    for i in range(1,len(data) + 1):

        if str(type(data[i-1][15])) ==  "<class 'datetime.datetime'>" and len(str(type(data[i-1][15]))) > 5:
            weeks = (today - data[i-1][15]).days
            if weeks % 7 == 0:
                weeks = weeks / 7
            else:
                # integer division
                weeks = weeks // 7 + 1
            # as per the weeks mentioned
            if weeks > 30:
                weeks = 30
            weeks = 15 + weeks*3

            # static data
            write_sheet.cell(i+1,1).value = data[i-1][0]
            write_sheet.cell(i+1,2).value = data[i - 1][1]
            write_sheet.cell(i + 1, 3).value = data[i - 1][3]
            write_sheet.cell(i + 1, 5).value = data[i - 1][4]
            write_sheet.cell(i + 1, 9).value = data[i - 1][20]
            write_sheet.cell(i + 1, 11).value = data[i - 1][22]
            write_sheet.cell(i + 1, 12).value = data[i - 1][23]
            write_sheet.cell(i + 1, 13).value = data[i - 1][24]

            # dynamic data
            write_sheet.cell(i+1,weeks + 1).value = data[i-1][13]
            write_sheet.cell(i+1,weeks + 2).value = data[i-1][17]
            write_sheet.cell(i + 1, weeks).value = data[i - 1][12]


    wb_write.save(output)



# # week 1 responses
# for cell in o_sheet['P']:
