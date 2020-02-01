#!/usr/bin/env python3

import os
import sys
import time
import datetime
import openpyxl

shift_path = "\\\iuser\\hubufsr\\ROC\\GCSO\\Special\\Admin\\Shifts\\"
shift = "AMS-SDD schedule 2020.xlsx"

def get_shift(f_path, file):
    
    if os.path.isfile(f_path + file):
        print("\n%s is valid file" % (file))
        print("Reading excel...\n")
    elif os.path.isfile(f_path + file) == False:
        print("\n%s does not exist or path is invalid\n\n" % (file))



def get_date():

    # validate the shift path
    get_shift(shift_path, shift)

    # create excel object with path + file 
    wb = openpyxl.load_workbook(shift_path + shift)
    # select the correct sheet
    sheet = wb["AMS_2020"]
    print(f"sheet: {sheet.title}")

    # get current date
    current_date = datetime.datetime.now().strftime("%d/%m/%Y")
    rows = [sheet.cell(row=i, column=16).coordinate for i in range(4, 22)]
    #print(rows)
    print(f"current date: {current_date}")

    #for j in sheet["B4":"B21"]:
    #    for k in j:
    #        print(k.coordinate, k.value)

    cell_coord = ""

    # for every row of cell object between P2 and BD2, iterate over the cell objects in the row
    for r in sheet["P2":"OJ2"]:
        for c in r:
            # format the datetime.datetime object into "Day/Month/Year"
            date = c.value.strftime("%d/%m/%Y")
            # if current date == date in excel, save the value to variable
            if current_date == date:
                print(f"found date in excel: {date}")
                print("saving coordinate...")
                cell_coord = c.coordinate

    print(cell_coord)
    return cell_coord

get_date()