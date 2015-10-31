#! python3
# -*- coding: utf-8 -*-
"""
Created on Mon Aug 17 14:23:20 2015

@author: Efrem
"""

import pandas as pd
NaN = pd.np.nan

# Read data into a dataframe
df = pd.read_csv('invoice_data.csv')


categorical_data = [
 'item_type','item','membership','discount_type','day_type','day_dur'
]

# for col in categorical_data:
#     df[col] = df[col].astype('category')




##   *********** ANALYSIS *******************
df_invoices = df.set_index(['invoice','DATE'])

##  Line items for only rooms, not including other associated invoice items
df_rooms_only = df.query('item_type == "ROOM"').drop(['item_type'], axis=1)
df_rooms_only['day'] = (df_rooms_only['HOURS_UNITS'] < 5.5).map({True:'Half', False:'Full'})

grouped_by_room = df_rooms_only.groupby(['item','day','membership','discount_type'])

## Output Room rental summaries as Excel sheets
grouped_by_room['HOURS_UNITS','SUBTOTAL','TOTAL'].sum().to_csv('rooms_only_sum.csv')
grouped_by_room['AMOUNT','HOURS_UNITS','SUBTOTAL','DISCOUNT','TOTAL'].mean().to_csv('rooms_only_avg.csv')


table = pd.pivot_table(df_rooms_only,
                        index=['item','day','membership','discount_type'],
                        values=['HOURS_UNITS','SUBTOTAL','TOTAL'],
                        aggfunc=[pd.np.sum, pd.np.mean], fill_value=0 ) 
table.to_csv('pivot.csv')


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

