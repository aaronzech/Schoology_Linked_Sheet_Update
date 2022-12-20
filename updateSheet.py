# Update the following Google Sheet
# https://docs.google.com/spreadsheets/d/1Ah3sU0UHcIdBi7-0TzxXKkiwHZbP7IbEgRJt0Ujqoxs/edit#gid=0
# The script will delete the current sheet, and paste in new data from the downloaded CSV parent access code file from Schoology.com

# Sometimes makes a new sheet 12/19/22

import gspread
import tkinter as tk
from tkinter import filedialog
import pandas as pd
from df2gspread import df2gspread as d2g
from datetime import date
import time


# Clear values in "Sheet1"
def clearSheet(sh):
    worksheet = sh.sheet1;
    worksheet.clear()
    time.sleep(5) # Delay for reliablility 
    print("***Sheet Cleared***")

def insertSchoologyData(sh):
    print("Select Schoology Parent Access Code file")
    file = filedialog.askopenfilename()
    dataframe = pd.read_csv(file)
    print("File:",file)
    worksheet = sh.sheet1;
    print(sh.id)
    time.sleep(5);
    wks = d2g.upload(dataframe,gfile=sh.id, wks_name='ALL_CODES',row_names=False) #made a new spreadsheet
    print("Data added to sheet",wks)

# J1 = "Last update"
# K2 = "Date"
def addLastUpdate(sh,date):
    worksheet = sh.sheet1;
    worksheet.update('J1', 'Last update')
    worksheet.update('K1', date )
def addParentLink(sh):
    worksheet = sh.sheet1;
    worksheet.update('J2','Parent Sign Up Link')
    worksheet.update('J3','https://app.schoology.com/register.php?type=parent')
def addParentResources(sh):
    worksheet = sh.sheet1;
    worksheet.update('J4','Parent Schoology Resources')
    worksheet.update('J5','https://sites.google.com/apps.district279.org/student-family-resources/technology-tools/schoology')

#Main
gc = gspread.service_account()
#sh = gc.open("Schoology Parent Access Codes") #Name of Google Sheet
sh = gc.open_by_url('https://docs.google.com/spreadsheets/d/1Ah3sU0UHcIdBi7-0TzxXKkiwHZbP7IbEgRJt0Ujqoxs/')
time.sleep(10) # Delay for reliablility 
today = date.today()
date = today.strftime("%m/%d/%y")
time.sleep(10) # Delay for reliablility 
clearSheet(sh)
time.sleep(10) # Delay for reliablility 
insertSchoologyData(sh)
time.sleep(10) # Delay for reliablility 
addLastUpdate(sh,date)
addParentLink(sh)
addParentResources(sh)
print("DONE")