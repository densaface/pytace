#!/usr/bin/env python
# -*- coding: utf-8 -*-

import time

log_file = ''

def log_print(mes):
    if type(mes) != str:
        mes = str(mes)
    print(mes)
    file_name = log_file
    if file_name != "":
        try:
            myfile  = open(file_name, "a")
            personal_time = time.strftime("%H:%M:%S - ", time.gmtime(time.time() + 10800))
            myfile.write(personal_time + mes + "\n")
        except:
            return mes
        myfile.close()
    return mes

