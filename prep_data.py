#! python3
# -*- coding: utf-8 -*-
"""
Created on Mon Aug 17 14:23:20 2015

@author: Efrem
"""

import pandas as pd
NaN = pd.np.nan
import re
import json
import os.path

# ignore invoices
INVOICE_NUM_CUTOFF = 2035

room_classes = {'broadway':'broadway', 'atrium':'atrium', 'jingletown':'jingle[-| ]?town|mezzanine',
                'omi':'gallery|omi','meditation':'meditation', 'kitchen':'kitchen', 'meridian':'meridian',
                'east-oak':'east','west-oak':'west', 'uptown':'uptown', 'downtown':'downtown',
                #'courtyard':'courtyard'
                }

service_classes = {'setup/breakdown':'set[-| ]?up|pre[-| ]?event|post[-| ]?event',
                    'staffing':'staff|manager','A/V':'A/V|technician|sound',
                    'janitorial':'janitorial|waste|cleaning', 'drinks':'coffee|wine', 'compostables':'compost' }

member_classes = {'part-time': 'part[-| ]?time|Part Ttime', 'full-time':'full[ |-]?time|full member',
                    'non-member':'none?[ |-]member', 'org-connect':'Org'}

day_type_classes  =  {'weekend': 'weekend', 'weekday':'weekday|wkday'}

day_duration_classes = {'full-day':'Full[-| ]?day',  'half-day':'Half[-| ]?Day'}

discount_classes = { 'multi-room':'Multi[-| ]?Room', 'multi-day':'Multi[-| ]?day','multi-event':'Multi[-| ]?event|reo?ccuring', 'founder':'Founder',
                     'partnership':'Partner|sposor|WITH|share',
                     'returning-client':'Returning[-| ]?client'}


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

# Read in invoices data from json source
with open('invoices.json','r') as in_file:
    invoices = json.load(in_file)


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
    # df.loc[mask, 'HOURS_UNITS'] = df.loc[mask, 'HOURS_UNITS'].astype(float) * df[mask, 'OCCURANCE'].astype(float)

    df = df[['invoice','invoice_date','DATE','DESCRIPTION','AMOUNT','HOURS_UNITS','SUBTOTAL','DISCOUNT','TOTAL','RATE']]




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
        member_mask = df['RATE'].str.contains(member_classes[member_type], case=False, na=False)
        df.loc[member_mask,'membership'] = member_type

    df['day_type'] = None
    for day_type in day_type_classes:
        day_type_mask = df['RATE'].str.contains(day_type_classes[day_type], case=False, na=False)
        df.loc[day_type_mask,'day_type'] = day_type

    df['duration'] = None
    for day_duration in day_duration_classes:
        day_duration_mask = df['RATE'].str.contains(day_duration_classes[day_duration], case=False, na=False)
        df.loc[day_duration_mask,'duration'] = day_duration

    df['discount_type'] = None
    for discount in discount_classes:
        discount_mask = df['RATE'].str.contains(discount_classes[discount], case=False, na=False)
        df.loc[discount_mask,'discount_type'] = discount





    ## Convert numeric field data to numbers and fill-in as much missing information as possible

    # first we set values labeled 'waved', 'comped', 'included' to zero
    no_charge_indicators = "wai?ved|comped|included|member"
    no_charge_amount = df['AMOUNT'].str.contains(no_charge_indicators, case=False, na=False)
    no_charge_subtot = df['SUBTOTAL'].str.contains(no_charge_indicators, case=False, na=False)
    no_charge_tot = df['TOTAL'].str.contains(no_charge_indicators, case=False, na=False)


    all_no_charge = no_charge_amount | no_charge_subtot | no_charge_tot
    df.loc[all_no_charge, ['AMOUNT','SUBTOTAL','TOTAL'] ] = 0

    #  Set DISCOUNT to 1 for any waived fees
    df.loc[all_no_charge, 'DISCOUNT'] = 1

    # If no discount-type is indicated then make an indication
    no_discount_type_indicated = df['discount_type'].isnull()
    df.loc[(all_no_charge & no_discount_type_indicated), 'discount_type'] = 'waived'

    #  convert any 'flat fee' indicators to 1 so that the AMOUNT identifies with SUBTOTAL
    df['HOURS_UNITS'].loc[df['HOURS_UNITS'].str.contains('flat', case=False, na=False)] = 1



    # convert all numeric field values to floats (anything non-numeric becomes NaN)
    NUMERIC_FIELDS = ['AMOUNT','SUBTOTAL','HOURS_UNITS','DISCOUNT','TOTAL']
    for field in NUMERIC_FIELDS:
        df[field] = pd.to_numeric(df[field], errors='coerce')



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



    df = df[['invoice','invoice_date','DATE','item_type','item','AMOUNT','HOURS_UNITS','SUBTOTAL','DISCOUNT','TOTAL',
                'membership','discount_type','day_type','duration']]

    df.to_excel(fname+'.xlsx', index=False)



##   *********** ANALYSIS *******************
columns_to_fill = ['membership','discount_type']
df[columns_to_fill] = df[columns_to_fill].fillna('??')

df_invoices = df.set_index(['invoice','DATE'])

##  Line items for only rooms, not including other associated invoice items
df_rooms_only = df.query('item_type == "room"').drop(['item_type'], axis=1)
df_rooms_only['day'] = (df_rooms_only['HOURS_UNITS'] < 5.5).map({True:'Half', False:'Full'})

grouped_by_room = df_rooms_only.groupby(['item','day','membership','discount_type'])

## Output Room rental summaries as Excel sheets
# grouped_by_room['HOURS_UNITS','SUBTOTAL','TOTAL'].sum().to_excel('rooms_only_sum.xlsx')
# grouped_by_room['AMOUNT','HOURS_UNITS','SUBTOTAL','DISCOUNT','TOTAL'].mean().to_excel('rooms_only_avg.xlsx')


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

