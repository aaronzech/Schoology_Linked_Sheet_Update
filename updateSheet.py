# Update the following Google Sheet
# https://docs.google.com/spreadsheets/d/1Ah3sU0UHcIdBi7-0TzxXKkiwHZbP7IbEgRJt0Ujqoxs/edit#gid=0
# The script will delete the current sheet, and paste in new data from the downloaded CSV parent access code file from Schoology.com

import gspread
import tkinter as tk
from tkinter import filedialog
import pandas as pd
from df2gspread import df2gspread as d2g
from datetime import date



# Clear values in "Sheet1"
def clearSheet(sh):
    worksheet = sh.sheet1;
    worksheet.clear()

def insertSchoologyData(sh):
    print("Select Schoology Parent Access Code file")
    file = filedialog.askopenfilename()
    dataframe = pd.read_csv(file)
    print("File:",file)
    worksheet = sh.sheet1;
    print(sh.id)
    wks = d2g.upload(dataframe,gfile=sh.id, wks_name='ALL_CODES',row_names=False) #made a new spreadsheet

# J1 = "Last update"
# K2 = "Date"
def addLastUpdate(sh,date):
    worksheet = sh.sheet1;
    worksheet.update('J1', 'Last update')
    worksheet.update('K1', date )

#Main
gc = gspread.service_account()
sh = gc.open("Schoology Parent Access Codes") #Name of Google Sheet
today = date.today()
date = today.strftime("%m/%d/%y")
clearSheet(sh)
insertSchoologyData(sh)
addLastUpdate(sh,date)
print("DONE")