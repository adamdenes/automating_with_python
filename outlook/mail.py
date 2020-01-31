#!/usr/bin/env python3

import os
import sys
import time
import datetime
import win32com.client

os.system('cls')
date_time = datetime.datetime.now().strftime("%Y.%m.%d %H:%M:%S")
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
folders = outlook.Folders

inboxes = [folders[folder].Name for folder in range(folders.Count)]
boxes = [i for i in inboxes if "adam" not in i]
print(f"Operation started at {date_time}")
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
        elif str(message.Sender) == str(sender):
            #print(sender, message.Sender)
            print(f"{categ[0].strip()} => {processed}")
            message.Move(processed)
            time.sleep(1)
        else:
            print(f"ticket: {categ[0].strip()} => assignee: {categ[1].strip()} => time: {message.CreationTime}")
            for p_sub in item_in_mailb[4]:
                if str(categ[1].strip()) in str(p_sub):
                    #print(message.body)
                    print(f"\tmoving {categ[0].strip()} to folder => {p_sub}")
                    #print("\n", message.Sender,"\n", message.To, "\n ", message.Subject,"\n", message.CreationTime,"\n", "_________" * 10)
                    message.Move(p_sub)
                    messages = item_in_mailb[1].Items
                    message = messages.GetFirst()
                    time.sleep(1)

def main(mailboxes):

    for mailb in mailboxes:
        move_mails(mailb)

main([SDD, AMS, CSC])