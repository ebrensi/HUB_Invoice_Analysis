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

service_classes = {'setup/breakdown':'set[-| ]?up', 'staffing':'staff|manager', 'A/V':'A/V|technician',
                    'janitorial':'janitorial|waste|cleaning' }

member_classes = {'part-time': 'part[-| ]?time', 'full-time':'full[ |-]?time|full member',
                    'non-member':'none?[ |-]member', 'org-connect':'Org'}

day_type_classes  =  {'weekend': 'weekend', 'weekday':'weekday|wkday'}

day_duration_classes = {'full-day':'Full[-| ]?day',  'half-day':'Half[-| ]?Day'}

discount_classes = { 'multi-room':'Multi[-| ]?Room', 'founder':'Founder','partnership':'Partnership',
                     'returnng-client':'Returning[-| ]?client'}


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
    pd.set_option('expand_frame_repr', False)
    sname = ws.title.strip()

    template_pattern = re.compile('template|quotes', re.IGNORECASE)
    if template_pattern.search(sname):
        return False

    invoice_num_match = re.match('^(\d+)',sname)

    # exclude invoices with no invoice# or invoice# < 2035
    if (not invoice_num_match) or (int(invoice_num_match.group(1)) < 2035) :
        print("\texcluding '%s'" % (sname))
        return False

    # make a dataframe from the current sheet
    df = pd.DataFrame([tuple([cell.value for cell in row]) for row in ws.rows]).dropna(how='all', axis=[0,1])
    df = df.reset_index(drop=True)

    if any(df):
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
        date_col_name = header[0]
        subsheet.columns = header


        if not subsheet[date_col_name].iloc[0]:
             subsheet[date_col_name].iloc[0] = '?'

        subsheet = subsheet.dropna(how='all',axis=[0,1]).reset_index(drop=True)

        # Fill-in DATE column
        for i in subsheet.index:
            d = subsheet.loc[i, date_col_name]
            if d:
                if isinstance(d, datetime.datetime):
                    subsheet.loc[i, date_col_name] = str(d.date())
                else:
                    subsheet.loc[i, date_col_name] = str(d)
            else:
                if i > 0:
                    subsheet.loc[i, date_col_name] = subsheet.loc[i-1, date_col_name]
                else:
                    subsheet.loc[i, date_col_name] = "unknown"
        items = subsheet.to_dict("records")
        info['items'] = items

    return info



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
    # dfs = []
    start_time = time.time()
    for ws in worksheets:
        invoice_dict = parse_sheet(ws)
        if invoice_dict:
            invoices[ws.title] = invoice_dict
            # dfs.append(invoice_df)
            print(ws.title)
        else:
            # if nothing was parsed from this invoice then remove it's key from 'invoices'
            invoices.pop(ws.title, None)



    elapsed_string = str(datetime.timedelta(seconds=time.time()-start_time))
    print('Finished in %s' % elapsed_string)

    return invoices



# This function produces a list of dictionaries, each entry one item from a nested invoice dictionary
def flatten_dict(invoices):
    result = []
    for title in invoices:
        if invoices[title]:
            c = invoices[title].copy()
            items = c.pop('items')
            c['sheet'] = title
            for item in items:
                c2 = c.copy()
                c2.update(item)
                result.append(c2)
    return result




### ************************************

# Read in the original Excel workbooks and create the invoices.json file
if not os.path.isfile('invoices.json'):
    invoices = xlsx2json(file_names)
    with open('invoices.json','w') as out_file:
        out_file.write(json.dumps(invoices, indent=3))
else:
    with open('invoices.json','r') as in_file:
        invoices = json.load(in_file)

# Transform json data into a flat table
if not os.path.isfile('invoice_items_flat.csv'):
    df = pd.DataFrame(flatten_dict(invoices)).drop_duplicates().dropna(how='all')
    df.to_csv('invoice_items_flat.csv',index=False,  encoding='utf-8')
    df.set_index(['sheet','DATE']).to_excel('invoice_items_flat.xlsx')
else:
    df = pd.read_csv('invoice_items_flat.csv', encoding='utf-8')


# df is a raw flat table.  First we join equivalent columns
df['DATE'].update(df['DATE OF EVENT'])
# df = df.drop('DATE OF EVENT', axis=1)

for col_name in ['ESTIMATED HOURS', 'HOURS']:
    df['HOURS/UNITS'].update(df[col_name])
    # df = df.drop(col_name, axis=1)

for col_name in [' TOTAL', 'ESTIMATE TOTAL', 'ESTIMATED TOTAL']:
    df['TOTAL'].update(df[col_name])
    # df = df.drop(col_name, axis=1)

# df = df.drop('DONATION', axis=1)

df = df[['sheet','DATE','DESCRIPTION','AMOUNT','HOURS/UNITS','SUBTOTAL','DISCOUNT','TOTAL','rate']]


## clean up numeric columns

# first we take care of implicit full-discount (amounts labeled 'waved', 'comped', 'included')
no_charge_indicators = "waved|comped|included|member"
no_charge_amount = df['AMOUNT'].str.contains(no_charge_indicators, case=False, na=False)
no_charge_subtot = df['SUBTOTAL'].str.contains(no_charge_indicators, case=False, na=False)
no_charge_tot = df['TOTAL'].str.contains(no_charge_indicators, case=False, na=False)

no_charge = no_charge_amount | no_charge_subtot | no_charge_tot
df.loc[no_charge,['AMOUNT','SUBTOTAL','TOTAL']] = 0
df.loc[no_charge, 'DISCOUNT'] = 1


#  convert any 'flat fee' indicators to 1 so that the AMOUNT identifies with SUBTOTAL.
df['HOURS/UNITS'].loc[df['HOURS/UNITS'].str.contains('flat', case=False, na=False)] = 1

# convert all 'AMOUNT' and 'SUBTOTAL' values to floats (anything non-numeric becomes NaN)
df[['AMOUNT','SUBTOTAL','HOURS/UNITS','DISCOUNT','TOTAL']] = df[['AMOUNT','SUBTOTAL','HOURS/UNITS','DISCOUNT','TOTAL']].convert_objects(convert_numeric=True)

#  drop rows where 'AMOUNT', 'HOURS/UNITS', SUBTOTAL', 'TOTAL' are all empty or zero
isnull = df[['AMOUNT','SUBTOTAL','HOURS/UNITS','TOTAL']].isnull()
iszero = df[['AMOUNT','SUBTOTAL','HOURS/UNITS','TOTAL']] == 0
is_either = (isnull | iszero).all(axis=1)
df = df[~is_either]

# fill in missing subtotal values.  We need these values for determining income.
# If AMOUNT and HOURS/UNITS are both non-empty then compute SUBTOTAL = AMOUNT * HOURS/UNITS
notnull = df[['AMOUNT','HOURS/UNITS']].notnull().all(axis=1)
df.loc[notnull, 'SUBTOTAL'] = df['AMOUNT'] * df['HOURS/UNITS']


# output a multi-index excel file for inspection
df.set_index(['sheet','DATE']).to_excel('invoice_items_flat_cleaned.xlsx')



##  Classify items into standard categories: 'room', 'service', 'total', 'other'
df['item-type'] = None
df['item'] = None

for room in room_classes:
    this_room_mask = df['DESCRIPTION'].str.contains(room_classes[room], case=False, na=False)
    df.loc[this_room_mask,'item-type'] = 'room'
    df.loc[this_room_mask,'item'] = room

for service in service_classes:
    this_service_mask = df['DESCRIPTION'].str.contains(service_classes[service], case=False, na=False)
    df.loc[this_service_mask,'item-type'] = 'service'
    df.loc[this_service_mask,'item'] = service

this_total_mask = df['DESCRIPTION'].str.contains('total', case=False, na=False)
df.loc[this_total_mask,'item-type'] = 'total'
df.loc[this_total_mask,'item'] = None

other_mask = df['item-type'].isnull()
df.loc[other_mask,'item-type'] = 'other'
df.loc[other_mask,'item'] = df.loc[other_mask,'DESCRIPTION']


## classify RATE info from sheet into discount-type and member-type
df['membership'] = None
for member_type in member_classes:
    member_mask = df['rate'].str.contains(member_classes[member_type], case=False, na=False)
    df.loc[member_mask,'membership'] = member_type

df['day-type'] = None
for day_type in day_type_classes:
    day_type_mask = df['rate'].str.contains(day_type_classes[day_type], case=False, na=False)
    df.loc[day_type_mask,'day-type'] = day_type

df['duration'] = None
for day_duration in day_duration_classes:
    day_duration_mask = df['rate'].str.contains(day_duration_classes[day_duration], case=False, na=False)
    df.loc[day_duration_mask,'duration'] = day_duration




df = df[['sheet','DATE','item-type','item','AMOUNT','HOURS/UNITS','SUBTOTAL','DISCOUNT','TOTAL','rate','membership','day-type']]

df.set_index(['sheet','DATE']).to_excel('invoice_items_prepped.xlsx')



# rooms = room_classes.keys()
# grouped = df.groupby(['title','date'])

