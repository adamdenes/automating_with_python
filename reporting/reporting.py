#!/usr/bin/env python

import os
import re
import requests

#file_path = 'C:\\Users\\612561487\\desktop\\projects\\report.csv'
os.system('cls')
file_path = input('\nPlease specify the .csv file path and name (eg. C:\\reporting.csv): ')

def reporting(file):

    # check whether file is found
    if os.path.isfile(file):
        print('\n"%s" is valid file' % (file))
        print('porceeding...')
    elif os.path.isfile(csv) == False:
        print('\n"%s" does not exist or path is invalid\n\n' % (file))

    with open(file, 'r') as csv:
        csv.seek(0)
        header = next(csv).split()

        CH_CUSTOMER_NAME = []
        CH_ASSIGNEE_GROUP = []
        tup = (CH_CUSTOMER_NAME, CH_ASSIGNEE_GROUP)

        for c in csv:
            c = c.rstrip('\n')
            #print(c.split(',')[0])
            #print(c.split(',')[1])
            #regex = re.compile(r",HU_DEB\w*$", re.IGNORECASE)
            #cust = re.sub(regex, '', c)

            cust = c.split(',')[0]
            assignee = c.split(',')[1]
            CH_CUSTOMER_NAME.append(cust)
            CH_ASSIGNEE_GROUP.append(assignee)

    #print(CH_CUSTOMER_NAME)
    #print(CH_ASSIGNEE_GROUP)
    #print(tup[0])
    #return CH_CUSTOMER_NAME
    return tup

#reporting(file_path)