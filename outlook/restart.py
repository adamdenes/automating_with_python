#!/usr/bin/env python3

import sys
import time
from subprocess import Popen

filename = sys.argv[1]

while True:
    try:
        #print("\nStarting " + filename)
        p = Popen("python " + filename, shell=True)
        time.sleep(60)
        p.wait()

    except KeyboardInterrupt:
        print("\n\n* Program aborted by user. Exiting... *\n")
        sys.exit()