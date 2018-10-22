#!/usr/bin/python
import codecs
import json
import pprint
import os
import pandas as pd
import urllib.request as urlrequest
from urllib.parse import urlencode
from openpyxl import load_workbook

from tkinter import *
#import tkinter as tk
from tkinter import filedialog

host = 'http://api.serpstat.com/v3'

## Get filename
root = Tk()
root.withdraw()
file_path = filedialog.askopenfilename()
print(file_path)

## Get Token
#def show_entry_fields():
#   print("Token: %s\n" % (e1.get()))

master = Tk()
Label(master, text="Token").grid(row=0)
#Label(master, text="Last Name").grid(row=1)

e1 = Entry(master)
#e2 = Entry(master)

e1.grid(row=0, column=1)
#e2.grid(row=1, column=1)

Button(master, text='OK', command=master.quit).grid(row=3, column=0, sticky=W, pady=4)
#Button(master, text='Show', command=show_entry_fields).grid(row=3, column=1, sticky=W, pady=4)
mainloop( )
my_token = e1.get()

print(my_token)

#exit()

## Load in the workbook
wb = load_workbook(file_path)
sheet = wb.get_sheet_by_name('Лист1')
sheet.title


#for i in range(1, 100):
#    print(i, sheet.cell(row=i, column=1).value)
     
     
method = 'keyword_top'

for i in range(1, 100):
    qr = sheet.cell(row=i, column=1).value.replace(" ", "%20")
    params = {
        'query': qr,  # string for get info
        'se': 'g_ru',  # string search engine
        'token': my_token,  # string personal token
    }

    api_url = "{host}/{method}?{params}".format(
        host=host,
        method=method,
        params=urlencode(params)
    )

    try:
        json_data = urlrequest.urlopen(api_url).read()
    except Exception as e0:
        print("API request error: {error}".format(error=e0))

    data = json.loads(json_data)
    #pprint.pprint(data)
    #print(params)

    #getting output for requests
    print(qr)
    print(status_code)
    print(status_msg)
    print(results_found)
    print(domains)
    print(urls)