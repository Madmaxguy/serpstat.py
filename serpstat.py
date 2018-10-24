#!/usr/bin/python
import json
import urllib.request as urlrequest
from urllib.parse import urlencode
from openpyxl import load_workbook
from datetime import datetime
import openpyxl
from tkinter import *
from tkinter import filedialog
import bpy

host = 'http://api.serpstat.com/v3'
method = 'keyword_top'
se = 'g_ru'

## Get filename
root = Tk()
root.withdraw()
file_path = filedialog.askopenfilename()
print(file_path)

# Inputting Token for API
master = Tk()
Label(master, text="Token").grid(row=0)
e1 = Entry(master)
e1.grid(row=0, column=1)
Button(master, text='OK', command=master.quit).grid(row=3, column=0, sticky=W, pady=4)
mainloop( )
my_token = e1.get()
print(my_token)

# Getting Input keywords
wb1 = load_workbook(file_path)
sheet = wb1.get_sheet_by_name('Лист1')
sheet.title


# Writing output file
now = datetime.now()
output_file = 'results_' + now.strftime("%Y-%m-%d_%H:%M:%S") + '.xlsx'
wb2 = openpyxl.Workbook()
sheetname = method + "_" + se
wb2.create_sheet("Top_keyword")
sheet = wb2.get_sheet_by_name('Top_keyword')

sheet.cell(row=1, column=1).value = 'Keyword'
sheet.cell(row=1, column=1).font = openpyxl.styles.Font(bold=True)
sheet.cell(row=1, column=2).value = 'Response Message'
sheet.cell(row=1, column=2).font = openpyxl.styles.Font(bold=True)
sheet.cell(row=1, column=3).value = 'Found Results'
sheet.cell(row=1, column=3).font = openpyxl.styles.Font(bold=True)
sheet.cell(row=1, column=4).value = 'Domains'
sheet.cell(row=1, column=4).font = openpyxl.styles.Font(bold=True)
sheet.cell(row=1, column=5).value = 'Full URLs'
sheet.cell(row=1, column=5).font = openpyxl.styles.Font(bold=True)
#ws = wb.activate()



for i in range(1, 100):
    qr = sheet.cell(row=i, column=1).value.replace(" ", "%20")
    params = {
        'query': qr,  # string for get info
        'se': se ,  # string search engine
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

    # Parsing JSON
    jsondata = json.loads(json_data)
    # Extracting data from JSON
    domains = ""
    urls = ""
    left_lines = jsondata['left_lines']
    status_msg = "%i" % (jsondata['status_code']) + ", " + jsondata['status_msg']
    results_found = ""

    if jsondata['status_code'] == 200:
        results_found = jsondata['result']['results']
        # Getting domains
        for item in jsondata['result']['top']:
            domains += item.get("domain") + ","
        # Getting URLs
        for item in jsondata['result']['top']:
            urls += item.get("url") + ","

    print(qr)
    print(status_msg)
    print(results_found)
    print(domains)
    print(urls)
    print("lines left: " + results_found)
    sheet.cell(row=i + 1, column=1).value = qr
    sheet.cell(row=i + 1, column=2).value = status_msg
    sheet.cell(row=i + 1, column=3).value = results_found
    sheet.cell(row=i + 1, column=4).value = domains
    sheet.cell(row=i + 1, column=5).value = urls

wb2.save(output_file)