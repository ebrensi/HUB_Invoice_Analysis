#! python3
# -*- coding: utf-8 -*-
"""
This is the analysis script for the IHO venue pricing project

"""

import pandas as pd
NaN = pd.np.nan


# This is a little function to output an easier-to-read csv file
#  for a multi-indexed DataFrame.  It eliminates duplicated index entries
#  along index columns. 
# The csv file produced is only meant to be viewed by a human.
def to_nice_csv(df, filename):
    x = df.reset_index()
    # mask = df.index.to_series().duplicated()
    # x.loc[mask.values, df.index.names] = ''
    for col_name in df.index.names:
        x.loc[:,col_name] = x.loc[:,col_name].drop_duplicates() 
    
    x.to_csv(filename,index=False)




# Read data into a dataframe
df = pd.read_csv('invoice_data.csv')

df_invoices = df.set_index(['invoice','DATE'])




##  Line items for only rooms, not including other associated invoice items
df_rooms_only = df.query('item_type == "ROOM"').drop(['item_type'], axis=1)



# We'll specify only certain rooms to simplify the output for now
selected_rooms = ['BROADWAY']
df_rooms_only = df_rooms_only[df_rooms_only['item'].isin(selected_rooms)]



df_rooms_only['day'] = (df_rooms_only['HOURS_UNITS'] < 5.5).map({True:'Half', False:'Full'})

grouped_by_room = df_rooms_only.groupby(['item','day','membership','discount_type'])

g1 = grouped_by_room['HOURS_UNITS','SUBTOTAL','TOTAL'].sum()

# Output Room rental summaries
#grouped_by_room['HOURS_UNITS','SUBTOTAL','TOTAL'].sum().to_csv('rooms_only_sum.csv')
to_nice_csv( grouped_by_room['HOURS_UNITS','SUBTOTAL','TOTAL'].sum(), 'rooms_only_sum.csv' )

# grouped_by_room['AMOUNT','HOURS_UNITS','SUBTOTAL','DISCOUNT','TOTAL'].mean().to_csv('rooms_only_avg.csv')
to_nice_csv( grouped_by_room['AMOUNT','HOURS_UNITS','SUBTOTAL','DISCOUNT','TOTAL'].mean(), 'rooms_only_avg.csv')

# Try an alternative method of aggregation
table = pd.pivot_table(df_rooms_only,
                        index=['item','day','membership','discount_type'],
                        values=['SUBTOTAL','TOTAL'],
                        aggfunc=[pd.np.sum, pd.np.mean], fill_value=0 ) 

to_nice_csv( table, 'rooms_only_pivot.csv')



###  Things to add later

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

