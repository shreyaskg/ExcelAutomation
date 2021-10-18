import openpyxl
import datetime
import pandas as pd
import os.path
from datetime import date
from SecondIterationOnwards import SecondIterationOnwards
from CreateSkeleton import CreatingSkeletalFile
from FirstIteration import FirstIteration

# The program will dynamically generate number of weeks as per input
def Automate(week_number):
    files = []
    print("In progress")
    if os.path.isfile('output.xlsx'):
        os.remove("output.xlsx")
    for i in range(1,week_number+1):
        f_name = "Week " + str(i) + ".xlsx"
        if os.path.isfile(f_name):
            files.append(f_name)
    for i in range(1,len(files)):
        if i == 1:
            # Then we have to make the skeletal file,use FirstIteration
            CreatingSkeletalFile(files[i],week_number)
            FirstIteration(files[i])
        else:
            # The skeletal file will be created in the 'if' statement, upon which the file will be built
            # The 'output.xlsx' created during the if statement wll be the skeleton for the function
            SecondIterationOnwards(files[i])

    os.remove('skeletal_file.xlsx')
