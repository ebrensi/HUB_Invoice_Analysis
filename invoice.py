# -*- coding: utf-8 -*-
"""
Created on Mon Aug 17 14:23:20 2015

@author: Efrem
"""

from openpyxl import load_workbook
import pandas as pd
NaN = pd.np.nan
import re
import concurrent.futures as futures
import json
import time
import datetime
import os.path
from collections import OrderedDict
#from tqdm import *

room_classes = {'broadway':'broadway', 'atrium':'atrium', 'jingletown':'jingletown', 'omi':'gallery|omi',
                 'meditation':'meditation', 'kitchen':'kitchen', 'meridian':'meridian', 'east_oak':'east',
                 'west_oak':'west', 'up':'uptown', 'down':'downtown'}

rate_classes = {'ptm': 'part[-| ]?time', 'ftm':'full[ |-]?time|full member', 'nm':'none?[ |-]member',
                    'wkn': 'weekend', 'wkd':'weekday|wkday', 'wth':'with'}

discount_classes = {'fdd':'Full[-| ]?day', 'mrd':'Multi[-| ]?Room', 'pd':'Partnership',
                        'fd':'Founder', 'rcd':'Returning[-| ]?client', 'hdd':'Half[-| ]?Day'}


def flatten(l):
    return [item for sublist in l for item in sublist]

def find(patt):
    def search(x):
        if x:
            return patt.search(unicode(x))
        else:
            return None
    return search

file_names = ['IHO_OnGoing_InvoiceTemplate.xlsx']#,
              # '2015 OnGoing InvoiceTemplate.xlsx']
             #'IHO_OnGoing_QuoteTemplate.xlsx'
#                ]


def parse_sheet(ws):
    info = {}
    pd.set_option('expand_frame_repr', False)
    sname = ws.title

    template_pattern = re.compile('template|quotes', re.IGNORECASE)
    if template_pattern.search(sname):
        return False, False

    # make a dataframe from the current sheet
    df = pd.DataFrame([tuple([cell.value for cell in row]) for row in ws.rows]).dropna(how='all', axis=[0,1])
    df = df.reset_index(drop=True)

    if any(df):
        # # get the invoice number from the sheet content, if it's there
        # # otherwise, try to exctract it from the sheet name
        # patt = re.compile('invoice #?(.*)', re.IGNORECASE)
        # tags = df.dropna(how='all', axis=[0,1]).applymap(find(patt)).dropna(how='all', axis=[0,1])
        # tags_list = flatten(tags.values.tolist())
        # if tags_list and bool(tags_list[0]):
        #     info['invoice#'] = tags_list[0].group(1).strip()
        # else:
        #     match = re.search('(\d+-?\d*) ', sname)
        #     if match:
        #         info['invoice#'] = match.group(1).rstrip('.').strip()
        #     else:
        #         info['invoice#'] = sname.strip()


        # find the cell that contains rate information and parse it
        patt = re.compile('.*(rate:| rate).*', re.IGNORECASE)
        tags = df.dropna(how='all', axis=[0,1]).applymap(find(patt)).dropna(how='all', axis=[0,1])
        tags_list = flatten(tags.values.tolist())
        if tags_list:
            info['rate'] = tags_list[0].group(0).lstrip('RATE:').strip()
        else:
            info['rate'] = ''


        # Determine upper & lower boundaries for the item subtable
        sep = df[df[0].str.contains('^date',case=False, na=False)].index.tolist()
        if sep:
             table_header_row = sep[0]
        sep = df[df[1].str.contains('total',case=False, na=False)].index.tolist()
        if sep:
            last_row = sep[0]
        else:
            last_row = len(df)


        # grab the items sub-table into a DataFrame
        header_row = df.iloc[table_header_row].tolist()
        last_col = next(i for i, j in reversed(list(enumerate(header_row))) if j)
        header = [str(field) for field in header_row[0:last_col+1]]

        subsheet = df.iloc[table_header_row+1:last_row+1, 0:last_col+1]
        header[0] = 'DATE'

        subsheet.columns = header


        if not subsheet['DATE'].iloc[0]:
             subsheet['DATE'].iloc[0] = '?'

        subsheet = subsheet.dropna(how='all',axis=[0,1]).reset_index(drop=True)

        # Fill-in DATE column
        for i in subsheet.index:
            d = subsheet.loc[i, 'DATE']
            if d:
                if isinstance(d, datetime.datetime):
                    subsheet.loc[i, 'DATE'] = str(d.date())
                else:
                    subsheet.loc[i, 'DATE'] = str(d)
            else:
                if i > 0:
                    subsheet.loc[i, 'DATE'] = subsheet.loc[i-1, 'DATE']
                else:
                    subsheet.loc[i, 'DATE'] = "unknown"
        items = subsheet.to_dict("records")

        info['items'] = items

        subsheet['RATE'] = info['rate']
        subsheet['invoice#'] = info['invoice#']
        subsheet['SHEET'] = sname
        # subsheet = subsheet[['SHEET','invoice#']+header+['RATE']]
    return info, subsheet



## *******************************

def xlsx2json(file_names):
    worksheets = []

    start_time = time.time()
    for fname in file_names:
        print('Loading %s' % fname)
        wb = load_workbook(fname, data_only=True)
        worksheets.extend(wb.worksheets)
    elapsed_string = str(datetime.timedelta(seconds=time.time()-start_time))
    print('workbooks loaded in %s' % elapsed_string)

    sheet_names = reversed([ws.title for ws in worksheets])
    invoices = OrderedDict.fromkeys(sheet_names)
    dfs = []
    start_time = time.time()
    for ws in worksheets:
        invoice_dict, invoice_df = parse_sheet(ws)
        if invoice_dict:
            invoices[ws.title] = invoice_dict
            dfs.append(invoice_df)
            print(ws.title)
        else:
            # if nothing was parsed from this invoice then remove it's key from 'invoices'
            invoices.pop(ws.title, None)

#        df = pd.concat(dfs, ignore_index=True)


    elapsed_string = str(datetime.timedelta(seconds=time.time()-start_time))
    print('Finished in %s' % elapsed_string)

    return invoices, dfs

# This function produces a list of dictionaries, each entry one item from a nested invoice dictionary
def flatten_dict(invoices):
    result = []
    for title in invoices:
        if invoices[title]:
            c = invoices[title].copy()
            items = c.pop('items')
            for item in items:
                c2 = c.copy()
                c2.update(item)
                result.append(c2)
    return result




### ************************************

# Read in the original Excel workbooks and create the invoices.json file
if not os.path.isfile('invoices.json'):
    invoices, df = xlsx2json(file_names)
    with open('invoices.json','w') as out_file:
        out_file.write(json.dumps(invoices, indent=3))
else:
    with open('invoices.json','r') as in_file:
        invoices = json.load(in_file)
"""
# Transform json data into a flat table
if not os.path.isfile('invoice_items_flat.csv'):
    df = pd.DataFrame(flatten_dict(invoices)).drop_duplicates().dropna(how='all')

    # sort entries by invoice number
    sort_by_invoice_num = df['title'].str.extract('(\d+)').dropna().astype(int).order().index
    df = df.loc[sort_by_invoice_num]

    fields = ['title','date','description','amount','hours','subtotal','discount','rate']
    df = df[fields]

    # output a multi-index excel file for inspection
    df.set_index(['title','date']).to_excel('invoice_items_flat.xlsx')

    df.to_csv('invoice_items_flat.csv',index=False,  encoding='utf-8')
else:
    df = pd.read_csv('invoice_items_flat.csv', encoding='utf-8')


## clean up numeric columns
# first we take care of implicit full-discount (amounts labeled 'waved', 'comped', 'included')
no_charge_amount = df['amount'].str.contains("waved|comped|included", case=False, na=False)
no_charge_subtot = df['subtotal'].str.contains("waved|comped|included", case=False, na=False)
no_charge = no_charge_amount | no_charge_subtot
df.loc[no_charge,['amount','subtotal']] = 0
df.loc[no_charge, 'discount'] = 1

# We'll need 'hours' field to be numeric too.
#  convert any 'flat fee' indicators to 1 so that the amount identifies with subtotal.
df['hours'].loc[df['hours'].str.contains('flat', case=False, na=False)] = 1

# convert all 'amount' and 'subtotal' values to floats (anything non-numeric becomes NaN)
df[['amount','subtotal','hours']] = df[['amount','subtotal','hours']].convert_objects(convert_numeric=True)

# drop rows where both 'amount' and 'subtotal' are non-numeric
df = df[~df[['amount','subtotal']].isnull().all(axis=1)]

# fill in missing subtotal values.  We need these values for determining income.

# output a multi-index excel file for inspection
df.set_index(['title','date']).to_excel('invoice_items_flat_cleaned.xls')



# # Assoicate items with rate, room, half/full-day, and discount-type
# for rate in rate_classes:
#     df['rate_'+rate] =  df['rate'].str.contains(rate_classes[rate], case=False, na=False)

# for room in room_classes:
#     df[room] =  df['description'].str.contains(room_classes[room], case=False, na=False)


# for discount in discount_classes:
#     df[discount] =  df['rate'].str.contains(discount_classes[discount], case=False, na=False)



# rooms = room_classes.keys()
# grouped = df.groupby(['title','date'])


"""