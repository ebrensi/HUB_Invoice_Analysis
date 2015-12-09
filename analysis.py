#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from IHO_event_invoice import *

"""
This is the analysis script for the IHO venue pricing project

"""

# define ordering for some data that will be categorcal
categories = {
    member_type: ['NON_MEMBER',
                  'PART_TIME',
                  'FULL_TIME',
                  'FRIEND'],

    day_dur: ['PARTIAL_DAY', 'FULL_DAY'],

    ROOM: [
        'EAST_OAK',
        'WEST_OAK',
        'DOWNTOWN',
        'UPTOWN',
        'MERIDIAN',
        'OMI',
        'JINGLETOWN',
        'ATRIUM',
        'BROADWAY',
        'PATIO',
        'MEDITATION',
        'KITCHEN']
}


# Read line-item data into a dataframe
df = (pd.read_csv(LINE_ITEMS_FNAME + '.csv')
      .set_index(['invoice', 'DATE']))

# invoice summary data
df_inv = (pd.read_csv(INVOICE_SUMMARIES_FNAME + '.csv')
          .set_index(['invoice', 'DATE']))


class_fields = [member_type, discount_type, day_type, day_dur]
df = df.join(df_inv[class_fields])

# Set some data to categorical.
#  This is done mostly to get the ordering right so
# that it is easier to copy and paste to IHO's spreadsheets
for cat in [member_type, day_dur]:
    df[cat] = (df[cat]
               .astype("category",
                       categories=categories[cat],
                       ordered=True))


discounted = df[discount_type] != 'NONE'
df.loc[discounted, discount_type] = 'DISCOUNT'
df.loc[~discounted, discount_type] = 'NO_DISCOUNT'


# Line items for only rooms, not including other associated invoice items
df_rooms_only = (df.query('{} == "{}"'.format(item_type, ROOM))
                 .drop([item_type], axis=1)
                 .rename(columns={'item': ROOM}))

# Set ordering for ROOM names
df_rooms_only[ROOM] = (df_rooms_only[ROOM]
                       .astype("category",
                               categories=categories[ROOM],
                               ordered=True))


df_rooms_only['EFF_RATE'] = (df_rooms_only['TOTAL'] /
                             df_rooms_only['HOURS_UNITS'])


multindex = [day_type,
             day_dur,
             member_type,
             discount_type]

grouped_by_room = (df_rooms_only
                   .groupby([ROOM] + multindex))

room_counts = grouped_by_room[ROOM].count()
room_counts.name = 'count'


room_sums = grouped_by_room['HOURS_UNITS', 'SUBTOTAL', 'TOTAL'].sum()


room_means = grouped_by_room['AMOUNT',
                             'HOURS_UNITS',
                             'SUBTOTAL',
                             'DISCOUNT',
                             'TOTAL',
                             'EFF_RATE'].mean()

# Create room pivot tables
room_pivot = pd.pivot_table(df_rooms_only,
                            index=[ROOM],
                            values=["EFF_RATE"],
                            columns=[day_type, day_dur,
                                     discount_type, member_type],
                            aggfunc=pd.np.mean)


# Now do the same thing for services
df_services_only = (df.query('{} == "{}"'.format(item_type, SERVICE))
                    .drop([item_type], axis=1)
                    .rename(columns={'item': SERVICE}))


df_services_only['HOURS_UNITS'] = df_services_only['HOURS_UNITS'].fillna(1.0)

df_services_only['EFF_RATE'] = (df_services_only['TOTAL'] /
                                df_services_only['HOURS_UNITS'])


grouped_by_service = (df_services_only
                      .groupby([SERVICE] + multindex))

table = pd.pivot_table(df_services_only,
                       index=[SERVICE, day_type, member_type, discount_type],
                       values=['SUBTOTAL', 'TOTAL'],
                       aggfunc=[pd.np.sum, pd.np.mean], fill_value=0)


# Output aggregated results
"""
# Rooms and services only
df_rooms_only.to_excel('rooms_only.xlsx', float_format='%5.2f')
df_services_only.to_excel('services_only.xlsx', float_format='%5.2f')

# Total income for each room
to_nice_csv(pd.concat([room_sums, room_counts], axis=1),
'IHO_pricing_rooms_only_sum.csv')

# Average income for each room
to_nice_csv(pd.concat([room_means, room_counts], axis=1),
'IHO_pricing_rooms_only_avg.csv')


to_nice_csv(room_means[["AMOUNT", "EFF_RATE"]],
'IHO_pricing_effective_room_rates.csv')

to_nice_csv(table, 'IHO_pricing_services_only.csv')
"""
