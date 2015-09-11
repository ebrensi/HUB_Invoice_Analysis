# -*- coding: utf-8 -*-
"""
Created on Mon Aug 17 14:23:20 2015

@author: Efrem
"""

from openpyxl import load_workbook
import pandas as pd
import re
import concurrent.futures as futures
import json
import time
import datetime
from tqdm import *


def flatten(l):
    return [item for sublist in l for item in sublist]

def find(patt):
    def search(x):
        if x:
            return patt.search(unicode(x))
        else:
            return None
    return search

file_names = ['IHO_OnGoing_InvoiceTemplate.xlsx',
               '2015 OnGoing InvoiceTemplate.xlsx']
             #'IHO_OnGoing_QuoteTemplate.xlsx'
#                ]


def parse_sheet(ws):
    info = {}

    sname = ws.title
    # make a dataframe from the current sheet
    df = pd.DataFrame([tuple([cell.value for cell in row]) for row in ws.rows])

    if any(df):
        # get the invoice number from the sheet content, if it's there
        # otherwise, try to exctract it from the sheet name
        patt = re.compile('invoice ?#(.*)', re.IGNORECASE)
        tags = df.dropna(how='all', axis=[0,1]).applymap(find(patt)).dropna(how='all', axis=[0,1])
        tags_list = flatten(tags.values.tolist())
        if tags_list:
            info['invoice#'] = tags_list[0].group(1).strip()
        else:
            match = re.search('(\d+-?\d*) ', sname)
            if match:
                info['invoice#'] = match.group(1).rstrip('.').strip()
            else:
                info['invoice#'] = sname.strip()


        # find the cell that contains rate information and parse it
        patt = re.compile('.*(rate:| rate).*', re.IGNORECASE)
        tags = df.dropna(how='all', axis=[0,1]).applymap(find(patt)).dropna(how='all', axis=[0,1])
        tags_list = flatten(tags.values.tolist())
        if tags_list:
            info['rate'] = tags_list[0].group(0).lstrip('RATE:').strip()
        else:
            info['rate'] = ''


        # find index locations of 'bill to', 'title', 'type', 'date' fields
        sep = df[df[0].str.contains('bill to|title|type',case=False, na=False)].index.tolist()

        if len(sep) < 3:
            info.update({'bill_to':'', 'title':'', 'type':''})
        else:
            # Exctract bill_to, title, and type fields
            info.update({'bill_to' : ', '.join(df[0][sep[0]+1:sep[1]-1].dropna()),
                         'title' : ', '.join(df[0][sep[1]+1:sep[2]-1].dropna()),
                         'type' : df[0][sep[2]+1] }  )

        items = {}
        sep = df[df[0].str.contains('date',case=False, na=False)].index.tolist()
        if sep:
             table_header_row = sep[0]
        sep = df[df[1].str.contains('^total',case=False, na=False)].index.tolist()
        if sep:
            last_row = sep[0]
        else:
            last_row = len(df)

        if df[5][table_header_row] and re.search('discount', df[5][table_header_row], re.IGNORECASE):
            discount_col = True
        else:
            discount_col = False

        date = ''
        for i in range(table_header_row+1, last_row):
            if df[0][i]:
                maybe_date = pd.to_datetime(str(df[0][i]), coerce=True)
                if maybe_date:
                    date = str(maybe_date.date())
                    items[date] = {}

            # if the amount field is not empty
            if df[2][i]:
                description = df[1][i]
                other = {'amount':df[2][i], 'hours':df[3][i], 'subtotal':df[4][i], 'discount':''}
                if discount_col and df[5][i]:
                    other['discount'] = df[5][i]

                if date:
                    items[date][description] = other
                else:
                    items[''] = {}
                    items[''][description] = other

        info['items'] = items

    return info



## *******************************


invoices = {}
jobs = {}


def import_workbooks(file_names):
    worksheets = []

    start_time = time.time()
    for fname in file_names:
        print('Loading %s' % fname)
        wb = load_workbook(fname, data_only=True)
        worksheets.extend(wb.worksheets)
    elapsed_string = str(datetime.timedelta(seconds=time.time()-start_time))
    print('workbooks loaded in %s' % elapsed_string)


    start_time = time.time()
    for ws in tqdm(worksheets, total=len(worksheets)):

        # select one invoice sheet from the workbook
        invoices[ws.title] = parse_sheet(ws)

    elapsed_string = str(datetime.timedelta(seconds=time.time()-start_time))
    print('Finished in %s' % elapsed_string)


    with open('invoices.json','w') as out_file:
        out_file.write(json.dumps(invoices, indent=3))


def flatten_dict(d):
    result = []
    for name in invoices:
        if invoices[name]:
            c = invoices[name].copy()
            items = c.pop('items')
            for date in items:
                c['date'] = date
                for description, item in items[date].items():
                    c2 = c.copy()
                    c2['description'] = description
                    c2.update(item)
                    result.append(c2)
    return result




### ************************************


# Read in the original Excel workbooks and create the invoices.json file
import_workbooks(file_names)


# read in invoices.json and convert it to a flat table, which we will call invoice_items.csv
with open('invoices.json','r') as in_file:
    invoices = json.loads(in_file.read())
df = pd.DataFrame(flatten_dict(invoices)).drop_duplicates().dropna()

fields = ['invoice#','title','rate','date','description','hours','discount','subtotal']
df[fields].to_csv('invoice_items.csv',index=False,  encoding='utf-8')


df = pd.read_csv('invoice_items.csv')

# assoicate items with rate-type, room, half/full-day, and discount-type
rate_classes = {'ptm': 'part[-| ]?time', 'ftm':'full[ |-]?time|full member', 'nm':'none?[ |-]member',
         'wn': 'weekend', 'wd':'weekday|wkday'}
# WITH event rate?

for rate in rate_classes:
    df['rate_'+rate] =  df['rate'].str.contains(rate_classes[rate], case=False, na=False)


room_classes = {'broadway':'broadway', 'atrium':'atrium', 'jingletown':'jingletown', 'omi':'gallery|omi',
                     'east_oak':'east', 'west_oak':'west', 'up':'uptown', 'down':'downtown'}
for room in room_classes:
    df[room] =  df['description'].str.contains(room_classes[room], case=False, na=False)


discount_classes = {'fdd':'Full[-| ]?day', 'mrd':'Multi[-| ]?Room', 'pd':'Partnership',
                'fd':'Founder', 'rcd':'Returning[-| ]?client', 'hdd':'Half[-| ]?Day'}
for discount in discount_classes:
    df[discount] =  df['rate'].str.contains(discount_classes[discount], case=False, na=False)

df.to_csv('invoices_flagged.csv',index=False,  encoding='utf-8')

# categories

# Non Member weekend - average $
# Non Member weekday
# part time member weeend
# part tiem member weekday
# full time member weekend
# full time member weekday

# half day = 5.5hrs or less
# full day = 6hrs or more

# On Broadway
# Atrium
# Jingletown Lounge
# OMI Gallery
# Meridian Room
# Uptown
# Downtown
# East Oakland
# West Oakland

# Totals of how often:
# Founder Discount
# Multi Room Discount
# Full Day Discount
# Partner Discount

