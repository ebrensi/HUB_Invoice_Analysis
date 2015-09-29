
# -*- coding: utf-8 -*-
"""
Created on Mon Aug 17 14:23:20 2015

@author: Efrem
"""

from openpyxl import load_workbook
from StyleFrame import StyleFrame
import pandas as pd
NaN = pd.np.nan
#pd.set_option('expand_frame_repr', False)

import re
import json
import time
import datetime
import os.path
from collections import OrderedDict

# ignore invoices
INVOICE_NUM_CUTOFF = 2035

room_classes = {'broadway':'broadway', 'atrium':'atrium', 'jingletown':'jingle[-| ]?town|mezzanine', 'omi':'gallery|omi',
                 'meditation':'meditation', 'kitchen':'kitchen', 'meridian':'meridian', 'east-oak':'east',
                 'west-oak':'west', 'uptown':'uptown', 'downtown':'downtown', 'courtyard':'courtyard'}

service_classes = {'setup/breakdown':'set[-| ]?up|pre[-| ]?event|post[-| ]?event',
                    'staffing':'staff|manager','A/V':'A/V|technician|sound',
                    'janitorial':'janitorial|waste|cleaning', 'drinks':'coffee|wine', 'compostables':'compost' }

member_classes = {'part-time': 'part[-| ]?time|Part Ttime', 'full-time':'full[ |-]?time|full member',
                    'non-member':'none?[ |-]member', 'org-connect':'Org'}

day_type_classes  =  {'weekend': 'weekend', 'weekday':'weekday|wkday'}

day_duration_classes = {'full-day':'Full[-| ]?day',  'half-day':'Half[-| ]?Day'}

discount_classes = { 'multi-room':'Multi[-| ]?Room', 'multi-day':'Multi[-| ]?day','multi-event':'Multi[-| ]?event', 'founder':'Founder','partnership':'Partner',
                     'returning-client':'Returning[-| ]?client', 'sponsorship':'sponsor', 'reoccuring':'reo?ccuring',
                     'WITH':'WITH|share'}

# 115014 Debby Irving :  Full Member Weekday Rate - 50/50 Rev Share
# 2115 John Chiang : Non Member Weekday Rate - Sharon Cornu Credit Applied
# 2106 LNA_WATER : WITH EVENT
# 2181 BALANCE : Full Time Member Weekday Rate  - WITH event sponsorship
# 2151 B Labs : Non-member weekday; Discount Rate for IHO WITH event
# 2089 Vegan Soul Sunday : WITH event rate
# Non Member Weekday Rate -- Holiday Rate
# 2190 Kerusso Music Group : Non Member Weekday Rate -- Holiday Rate


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

## Extract relevant data from one invoice worksheet and return it as a dict
def parse_sheet(ws):
    info = {}
    pd.set_option('expand_frame_repr', False)
    sname = ws.title.strip()

    template_pattern = re.compile('template|quotes', re.IGNORECASE)
    if template_pattern.search(sname):
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
# Import a workbook of invoice sheets and store relevant data for every invoice as a
#   record in a json dictionary.
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

    start_time = time.time()
    for ws in worksheets:
        invoice_dict = parse_sheet(ws)
        if invoice_dict:
            invoices[ws.title] = invoice_dict
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
            c['invoice'] = title
            for item in items:
                c2 = c.copy()
                c2.update(item)
                result.append(c2)
    return result




### ************************************

# Read in the original Excel workbooks and create the invoices.json file
#  or load 'invoices.json' if it exists
if os.path.isfile('invoices.json'):
    with open('invoices.json','r') as in_file:
        invoices = json.load(in_file)
else:
    invoices = xlsx2json(file_names)
    with open('invoices.json','w') as out_file:
        out_file.write(json.dumps(invoices, indent=3))


# Clean up raw flattened data into the basic columns we want
fname = 'invoice_data'

if os.path.isfile(fname+'.xlsx'):
    df = pd.read_excel(fname+'.xlsx', encoding='utf-8')
else:
    df = pd.DataFrame(flatten_dict(invoices)).drop_duplicates().dropna(how='all')
    # df.to_excel('raw_flat_table.xlsx')

    # exclude invoices with no invoice# or invoice# < INVOICE_NUM_CUTOFF
    invoice_num = df['invoice'].str.extract('(\d+)').str.strip().astype(float)
    exclude_mask = invoice_num.isnull() |  (invoice_num < INVOICE_NUM_CUTOFF)
    # print("Excluding \n%s" % (df['invoice'][exclude_mask].tolist()) )
    df = df[~exclude_mask]


    # Exclude cancellation invoices
    cancellations = df['invoice'].str.contains('cancel', na=False, case=False)
    df = df[~cancellations]



    # Join equivalent columns
    df['DATE'].update(df['DATE OF EVENT'])

    df = df.rename(columns={'HOURS/UNITS': 'HOURS_UNITS'})
    for col_name in ['ESTIMATED HOURS', 'HOURS']:
        df['HOURS_UNITS'].update(df[col_name])

    for col_name in [' TOTAL', 'ESTIMATE TOTAL', 'ESTIMATED TOTAL']:
        df['TOTAL'].update(df[col_name])

    df['DISCOUNT'].update(df['DONATION'])


    # mask = df['OCCURANCE'].notnull()
    # df.loc[mask, 'HOURS_UNITS'] = df.loc[mask, 'HOURS_UNITS'] * df[mask, 'OCCURANCE']

    df = df[['invoice','DATE','DESCRIPTION','AMOUNT','HOURS_UNITS','SUBTOTAL','DISCOUNT','TOTAL','rate']]




    #  drop rows where 'AMOUNT', 'SUBTOTAL', 'DISCOUNT','TOTAL' are all empty or zero
    isnull = df[['AMOUNT','SUBTOTAL','DISCOUNT','TOTAL']].isnull()
    iszero = df[['AMOUNT','SUBTOTAL','DISCOUNT','TOTAL']] == 0
    is_either = (isnull | iszero).all(axis=1)
    df = df[~is_either]




    ##  Parse DESCRIPTION field into 'item_type' and 'item'
    ##  and Classify items into standard categories: 'room', 'service', 'total', 'other'
    df['item_type'] = None
    df['item'] = None

    for room in room_classes:
        this_room_mask = df['DESCRIPTION'].str.contains(room_classes[room], case=False, na=False)
        df.loc[this_room_mask,'item_type'] = 'room'
        df.loc[this_room_mask,'item'] = room

    for service in service_classes:
        this_service_mask = df['DESCRIPTION'].str.contains(service_classes[service], case=False, na=False)
        df.loc[this_service_mask,'item_type'] = 'service'
        df.loc[this_service_mask,'item'] = service

    this_total_mask = df['DESCRIPTION'].str.contains('total', case=False, na=False)
    df.loc[this_total_mask,'item_type'] = 'total'
    df.loc[this_total_mask,'item'] = None

    other_mask = df['item_type'].isnull()
    df.loc[other_mask,'item_type'] = 'other'
    df.loc[other_mask,'item'] = df.loc[other_mask,'DESCRIPTION']



    ## Parse RATE field from sheet into discount-type, day-type, and member-type
    df['membership'] = None
    for member_type in member_classes:
        member_mask = df['rate'].str.contains(member_classes[member_type], case=False, na=False)
        df.loc[member_mask,'membership'] = member_type

    df['day_type'] = None
    for day_type in day_type_classes:
        day_type_mask = df['rate'].str.contains(day_type_classes[day_type], case=False, na=False)
        df.loc[day_type_mask,'day_type'] = day_type

    df['duration'] = None
    for day_duration in day_duration_classes:
        day_duration_mask = df['rate'].str.contains(day_duration_classes[day_duration], case=False, na=False)
        df.loc[day_duration_mask,'duration'] = day_duration

    df['discount_type'] = None
    for discount in discount_classes:
        discount_mask = df['rate'].str.contains(discount_classes[discount], case=False, na=False)
        df.loc[discount_mask,'discount_type'] = discount





    ## Convert numeric field data to numbers and fill-in as much missing information as possible

    # first we set values labeled 'waved', 'comped', 'included' to zero
    no_charge_indicators = "waved|comped|included|member"

    no_charge_amount = df['AMOUNT'].str.contains(no_charge_indicators, case=False, na=False)
    df.loc[no_charge_amount,'AMOUNT'] = 0

    no_charge_subtot = df['SUBTOTAL'].str.contains(no_charge_indicators, case=False, na=False)
    df.loc[no_charge_subtot,'SUBTOTAL'] = 0

    no_charge_tot = df['TOTAL'].str.contains(no_charge_indicators, case=False, na=False)
    df.loc[no_charge_tot,'TOTAL'] = 0

    #  Set DISCOUNT to 1 for any waived fees
    all_no_charge = no_charge_amount | no_charge_subtot | no_charge_tot
    df.loc[all_no_charge, 'DISCOUNT'] = 1

    # If no discount-type is indicated then make an indication
    no_discount_type_indicated = df['discount_type'].isnull()
    df.loc[(all_no_charge & no_discount_type_indicated), 'discount_type'] = 'waived'

    #  convert any 'flat fee' indicators to 1 so that the AMOUNT identifies with SUBTOTAL
    df['HOURS_UNITS'].loc[df['HOURS_UNITS'].str.contains('flat', case=False, na=False)] = 1



    # convert all numeric field values to floats (anything non-numeric becomes NaN)
    NUMERIC_FIELDS = ['AMOUNT','SUBTOTAL','HOURS_UNITS','DISCOUNT','TOTAL']
    df[NUMERIC_FIELDS] = df[NUMERIC_FIELDS].convert_objects(convert_numeric=True)



    # Fill-in missing hours and discounts for room rentals where these values are implied by a multi-room
    #   deal, particularly when DESCRIPTION field contains certain keywords like 'Entire ...'
    multi_room_item_mask =  df['item'].str.contains('entire|level|rentals', na=False, case=False)
    idx_to_drop = []

    for item_idx, item in df[multi_room_item_mask].iterrows():
        associated_rooms = (df['invoice'] == item['invoice']) & (df['item_type'] == 'room') & df['TOTAL'].isnull()

        if associated_rooms.any():

            for field in ['HOURS_UNITS','DISCOUNT']:
                df.loc[associated_rooms, field] = item[field]

            idx_to_drop.append(item_idx)
            # df.loc[mask, 'SUBTOTAL'] = df.loc[mask, 'AMOUNT'] * df.loc[mask, 'HOURS_UNITS']
            # df.loc[mask, 'TOTAL'] = df.loc[mask, 'SUBTOTAL'] * (1 - df.loc[mask, 'DISCOUNT'])

    print('dropping multi-room items %s' % str(idx_to_drop))
    df = df.drop(idx_to_drop)


    # fill in missing subtotal values.  We need these values for determining income.
    # If AMOUNT and HOURS_UNITS are both non-empty then compute SUBTOTAL = AMOUNT * HOURS_UNITS
    notnull = df[['AMOUNT','HOURS_UNITS']].notnull().all(axis=1)
    df.loc[notnull, 'SUBTOTAL'] = df.loc[notnull, 'AMOUNT'] * df.loc[notnull, 'HOURS_UNITS']


    #  Fill-in all missing DISCOUNT entries with zero
    df.loc[df['DISCOUNT'].isnull() & (df['item_type'] != 'total'), 'DISCOUNT'] = 0

    # Fill-in all missing TOTAL for items with non-empty SUBTOTAL
    no_tot = df['TOTAL'].isnull() & df['SUBTOTAL'].notnull() & (df['item_type'] != 'total')
    df.loc[no_tot, 'TOTAL'] = df.loc[no_tot, 'SUBTOTAL'] * (1 - df.loc[no_tot, 'DISCOUNT'])


    # Fill-in SUBTOTAL for items with TOTAL but no SUBTOTAL (taking DISCOUNT into consideration)
    no_subtot = df['TOTAL'].notnull() & df['SUBTOTAL'].isnull() & (df['item_type'] != 'total')
    df.loc[no_subtot, 'SUBTOTAL'] = df['TOTAL'][no_subtot] / (1 - df['DISCOUNT'][no_subtot] )



    # Eliminate items with no SUBTOTAL, DISCOUNT, nor TOTAL
    mask = (df['TOTAL'].isnull() | (df['TOTAL'] == 0) )\
            & (df['SUBTOTAL'].isnull() | (df['SUBTOTAL'] == 0) )\
            & (df['DISCOUNT'].isnull() | (df['DISCOUNT'] == 0) )

    df = df[~mask]




    ## Manually compute SUBTOTAL and TOTAL fields in 'total' items
    tots = [0,0]
    for idx, row in df.iterrows():
        if row['item_type'] == 'total':
            df.loc[idx, ['SUBTOTAL','TOTAL']] = tots
            if tots[1] != 0:
                df.loc[idx, 'DISCOUNT'] = 1 - tots[1]/tots[0]
            tots = [0,0]
        else:
            if not pd.np.isnan(row['SUBTOTAL']):
                tots[0] += row['SUBTOTAL']
            if not pd.np.isnan(row['TOTAL']):
                tots[1] += row['TOTAL']



    df = df[['invoice','DATE','item_type','item','AMOUNT','HOURS_UNITS','SUBTOTAL','DISCOUNT','TOTAL',
                'membership','discount_type','day_type','duration']]

    # # An attempt to set column widths in output Excel spreadsheet
    # excel_writer = StyleFrame.ExcelWriter(fname+'.xlsx')
    # sf = StyleFrame(df)
    # for col in ['invoice', 'DATE']:
    #     width = int(df[col].str.len().max())
    #     sf.set_column_width(col, width)
    # sf.to_excel(excel_writer=excel_writer)
    # excel_writer.save()

    df.to_excel(fname+'.xlsx', index=False)



##   *********** ANALYSIS *******************
columns_to_fill = ['membership','discount_type']
df[columns_to_fill] = df[columns_to_fill].fillna('??')

df_invoices = df.set_index(['invoice','DATE'])

##  Line items for only rooms, not including other associated invoice items
df_rooms_only = df.query('item_type == "room"').drop(['item_type'], axis=1)
df_rooms_only['day'] = (df_rooms_only['HOURS_UNITS'] < 5.5).map({True:'Half', False:'Full'})

grouped_by_room = df_rooms_only.groupby(['item','day','membership','discount_type'])

grouped_by_room['HOURS_UNITS','SUBTOTAL','TOTAL'].sum().to_excel('rooms_only_sum.xlsx')
grouped_by_room['AMOUNT','HOURS_UNITS','SUBTOTAL','DISCOUNT','TOTAL'].mean().to_excel('rooms_only_avg.xlsx')


"""
# The overall length of time for each (invoice,DATE)
event_hours = df.groupby(['invoice','DATE'])['HOURS_UNITS'].max()

tot_fields = ['HOURS_UNITS','SUBTOTAL','DISCOUNT','TOTAL','membership','discount_type','day_type','duration']
df_tots = df_invoices.query('item_type == "total"')[tot_fields]
# df_tots2 = df.query('item_type == "total"').groupby(['invoice','DATE'])[tot_fields]

# df_tots['HOURS_UNITS'] = event_hours

# print df_tots
# df_tots = df_tots['HOURS_UNITS'].update(event_hours)
# print df_tots

# df_tots.mean() is mean income w and w/o discount
"""

