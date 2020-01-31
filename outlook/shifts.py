#!/usr/bin/env python3

import os
import sys
import time
import pandas

shift_path = "\\\iuser\\hubufsr\\ROC\\GCSO\\Special\\Admin\\Shifts\\AMS-SDD schedule 2020.xlsx"

def get_shift(f_path):
    try:
        if os.path.isfile(f_path):
            print('\n"%s" is valid file' % (f_path))
            print('processing...\n')
        elif os.path.isfile(f_path) == False:
            print('\n"%s" does not exist or path is invalid\n\n' % (f_path))



    except KeyboardInterrupt:
        print("\n\n* Program aborted by user. Exiting...*\n")
        sys.exit()

get_shift(shift_path)