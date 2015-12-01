#!/usr/bin/env python3

from IHO_event_invoice import *

# -*- coding: utf-8 -*-
"""
This is the analysis script for the IHO venue pricing project

"""


# Read line-item data into a dataframe
df = (pd.read_csv(LINE_ITEMS_FNAME + '.csv')
      .set_index(['invoice', 'DATE']))

# invoice summary data
df_inv = (pd.read_csv(INVOICE_SUMMARIES_FNAME + '.csv')
          .set_index(['invoice', 'DATE']))


class_fields = [member_type, discount_type, day_type, day_dur]
df = df.join(df_inv[class_fields])

discounted = df[discount_type] != 'NONE'
df.loc[discounted, discount_type] = 'DISCOUNT'
df.loc[~discounted, discount_type] = 'NO_DISCOUNT'

# Line items for only rooms, not including other associated invoice items
df_rooms_only = (df.query('{} == "{}"'.format(item_type, ROOM))
                 .drop([item_type], axis=1)
                 .rename(columns={'item': ROOM}))

df_rooms_only['EFF_RATE'] = (df_rooms_only['TOTAL'] /
                             df_rooms_only['HOURS_UNITS'])

# df_rooms_only.to_excel('rooms_only.xls', float_format='%5.2f')

multindex = [day_type,
             day_dur,
             member_type,
             discount_type]

grouped_by_room = (df_rooms_only
                   .groupby([ROOM] + multindex))

room_counts = grouped_by_room[ROOM].count()
room_counts.name = 'count'

# Output Room rental summaries
room_sums = grouped_by_room['HOURS_UNITS', 'SUBTOTAL', 'TOTAL'].sum()

to_nice_csv(pd.concat([room_sums, room_counts], axis=1),
            'IHO_pricing_rooms_only_sum.csv')


room_means = grouped_by_room['AMOUNT',
                             'HOURS_UNITS',
                             'SUBTOTAL',
                             'DISCOUNT',
                             'TOTAL',
                             'EFF_RATE'].mean()


to_nice_csv(pd.concat([room_means, room_counts], axis=1),
            'IHO_pricing_rooms_only_avg.csv')


to_nice_csv(room_means[["AMOUNT", "EFF_RATE"]],
            'IHO_pricing_effective_room_rates.csv')


# Now do the same thing for services
df_services_only = (df.query('{} == "{}"'.format(item_type, SERVICE))
                    .drop([item_type], axis=1)
                    .rename(columns={'item': SERVICE}))

grouped_by_service = (df_services_only
                      .groupby([SERVICE] + multindex))

table = pd.pivot_table(df_services_only,
                       index=[SERVICE, day_type,
                              member_type, discount_type],
                       values=['SUBTOTAL', 'TOTAL'],
                       aggfunc=[pd.np.sum, pd.np.mean], fill_value=0)

to_nice_csv(table, 'IHO_pricing_services_only.csv')
