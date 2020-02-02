#!/usr/bin/env python3

import os
import sys
import time
import datetime
import openpyxl
import win32com.client
from openpyxl.utils import get_column_letter, column_index_from_string


shift_path = "\\\iuser\\hubufsr\\ROC\\GCSO\\Special\\Admin\\Shifts\\"
shift = "AMS-SDD schedule 2020.xlsx"

date_time = datetime.datetime.now().strftime("%Y.%m.%d %H:%M:%S")
print(f"Operation started at {date_time}")

def open_schedule(f_path, file):
    
    if os.path.isfile(f_path + file):
        print("\n'%s' is valid file" % (file))
        print("Reading excel...\n")
    elif os.path.isfile(f_path + file) == False:
        print("\n'%s' does not exist or path is invalid\n\n" % (file))

def get_shifts():

    # clear screen on execution
    os.system('clear')
    # validate the shift file path
    open_schedule(shift_path, shift)

    # create excel object with path + file 
    wb = openpyxl.load_workbook(shift_path + shift)
    # select the correct sheet
    sheet = wb["AMS_2020"]
    print(f"sheet: {sheet.title}")

    # get current date
    current_date = datetime.datetime.now().strftime("%d/%m/%Y")
    print(f"current date: {current_date}")

    cell_column = ""

    # for every row of cell object between P2 and BD2, iterate over the cell objects in the row
    for r in sheet["P2":"OJ2"]:
        for c in r:
            # format the datetime.datetime object into "Day/Month/Year"
            date = c.value.strftime("%d/%m/%Y")
            # if current date == date in excel, save the value to variable
            if current_date == date:
                #print(c.row, c.column, c.coordinate)
                print(f"found date in excel: {date}")
                print("saving today's shifts...\n")

                # save the column to cell_column
                cell_column = c.column

    shifts_on_current_date = [sheet.cell(row=i, column=cell_column).value for i in range(4, 22)]
    agents = [sheet.cell(row=i, column=2).value for i in range(4, 22)]

    return shifts_on_current_date, agents

current_shifts, agent_list = get_shifts()


def create_dictionary(c_shifts, a_list):

    # create a dict with key value pairs => "Agent" : "shift"
    agents_shift_today = dict(zip(a_list, c_shifts))
    present_dict = {}
    absent_dict = {}

    for key, value in agents_shift_today.items():
        # if the shift is AL, BH, *, or empty -> save it to 'absent'
        if value == "AL" or value == "BH" or value == "*" or value == None:
            absent_dict[key] = value
        else:
            # else save the present agents into 'present'
            present_dict[key] = value
            
    print(absent_dict, "\n")
    print(present_dict, "\n")
    return absent_dict, present_dict

absent, present = create_dictionary(current_shifts, agent_list)


outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
folders = outlook.Folders

inboxes = [folders[folder].Name for folder in range(folders.Count)]
boxes = [i for i in inboxes if "adam" not in i]
print(boxes, "\n")

def generate_mailboxes(box_lst):

    for box in box_lst:
        root_f = outlook.Folders[box]
        sub_f = root_f.Folders["Inbox"]

        if str(root_f.Name) == "Servicedesk Debrecen G":
            personal_inboxes = sub_f.Folders["Team's folders"]
            processed = sub_f.Folders["Processed"]
        else:
            personal_inboxes = sub_f.Folders["+++TEAM+++"]
            processed = sub_f.Folders["+++PROCESSED+++"]
        personal_inb = personal_inboxes.Folders

        msgs = sub_f.Items
        msg = msgs.GetFirst()

    return root_f, sub_f, personal_inboxes, processed, personal_inb, msgs, msg


SDD = generate_mailboxes([boxes[0]])
AMS = generate_mailboxes([boxes[1]])
CSC = generate_mailboxes([boxes[2]])


def move_mails(item_in_mailb):

    sender = item_in_mailb[0]
    processed = item_in_mailb[3]

    for message in list(item_in_mailb[-2]):
        categories = message.categories
        categ = categories.split(',')
        
        if not categories:
            continue
        elif len(categ) > 3:
            continue
        elif str(message.Sender) == str(sender):
            #print(sender, message.Sender)
            print(f"ticket: {categ[0].strip()} => {processed}")
            print(f"\tmoving {categ[0].strip()} to folder => {processed}")
            message.Move(processed)
            time.sleep(1)
        else:
            print(f"ticket: {categ[0].strip()} => assignee: {categ[1].strip()} => mail arrived: {message.CreationTime}")
            for p_sub in item_in_mailb[4]:
                if (str(categ[1].strip()) in str(p_sub)) and (not str(categ[1].strip()) in absent):
                    print(f"\tmoving {categ[0].strip()} to folder => {p_sub}")
                    #print("\n", message.Sender,"\n", message.To, "\n ", message.Subject,"\n", message.CreationTime,"\n", "_________" * 10)
                    message.Move(p_sub)
                    messages = item_in_mailb[1].Items
                    message = messages.GetFirst()
                    time.sleep(1)

def main(mailboxes):

    for mailb in mailboxes:
        move_mails(mailb)

#main([SDD, AMS, CSC])

if __name__ == "__main__":
    main([SDD, AMS, CSC])