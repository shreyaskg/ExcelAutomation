import openpyxl
import datetime
from datetime import date

def GetDays(date):
    if type(date) == type(datetime.datetime.now()):
        day = datetime.datetime.now()
        return (day - date).days
    return
def GetAge(date):
    if type(date) == type(datetime.datetime.now()):
        birth_year = date.date()
        current_year = datetime.datetime.now().date()
        return current_year.year - birth_year.year
    return