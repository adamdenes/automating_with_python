#!/usr/bin/env python3

from subprocess import Popen
import sys
import time

filename = sys.argv[1]

while True:
    try:
        print("\nStarting " + filename)
        p = Popen("python " + filename, shell=True)
        time.sleep(120)
        p.wait()

    except KeyboardInterrupt:
        print("\n\n* Program aborted by user. Exiting...*\n")
        sys.exit()