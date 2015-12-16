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
                  'unknown'
                  ],

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
        # 'PATIO',
        # 'MEDITATION',
        # 'KITCHEN'
    ]
}


# Read line-item data into a dataframe
df = (pd.read_csv(LINE_ITEMS_FNAME + '.csv')
      .set_index(['invoice', 'DATE']))

# Read in invoice summary data
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


# discounted = df[discount_type] != 'NONE'
# df.loc[discounted, discount_type] = 'DISCOUNT'
# df.loc[~discounted, discount_type] = 'NO_DISCOUNT'


service_items = (df[item_type] == 'SERVICE')
df.loc[service_items, 'HOURS_UNITS'].fillna(1.0, inplace=True)

df['EFF_RATE'] = (df['TOTAL'] /
                  df['HOURS_UNITS'])
df.loc[df["DISCOUNT"] == 1, 'EFF_RATE'] = 0


# Line items for only rooms, not including other associated invoice items
df_rooms_only = (df.query('{} == "{}"'.format(item_type, ROOM))
                 .drop([item_type], axis=1)
                 .rename(columns={'item': ROOM}))

# Set ordering for ROOM names
df_rooms_only[ROOM] = (df_rooms_only[ROOM]
                       .astype("category",
                               categories=categories[ROOM],
                               ordered=True))

pivot_rows = []  # possible put discount_type

pivot_columns = [day_type, day_dur, discount_type, member_type]

grouped_by_room = (df_rooms_only
                   .groupby([ROOM] + pivot_rows + pivot_columns))

room_counts = grouped_by_room[ROOM].count()
room_counts.name = 'count'


room_sums = grouped_by_room['HOURS_UNITS', 'SUBTOTAL', 'TOTAL'].sum()
room_sums['count'] = room_counts

room_means = grouped_by_room[
    'AMOUNT',
    'HOURS_UNITS',
    'SUBTOTAL',
    'DISCOUNT',
    'TOTAL',
    'EFF_RATE'].mean()
room_means['count'] = room_counts


# Create room pivot tables
# may or may not want discount type in there
room_rate_pivot = pd.pivot_table(df_rooms_only,
                                 index=pivot_rows + [ROOM],
                                 values=["EFF_RATE"],
                                 columns=pivot_columns,
                                 aggfunc=pd.np.mean)


room_income_pivot = pd.pivot_table(df_rooms_only,
                                   index=pivot_rows + [ROOM],
                                   values=["TOTAL"],
                                   columns=pivot_columns,
                                   aggfunc=pd.np.mean)


# Now do the same thing for services
df_services_only = (df.query('{} == "{}"'.format(item_type, SERVICE))
                    .drop([item_type], axis=1)
                    .rename(columns={'item': SERVICE}))


# ******************  Output aggregated results  ***************************
df_rooms_only.to_nice_csv('rooms_only.csv')

# Total income for each room
to_nice_csv(room_sums, 'IHO_pricing_rooms_only_sum.csv')

# Average income for each room
to_nice_csv(room_means, 'IHO_pricing_rooms_only_avg.csv')

# Effective Rates
to_nice_csv(room_means[["AMOUNT", "EFF_RATE"]],
            'IHO_pricing_effective_room_rates.csv')


# Write everything to one Excel file
writer = pd.ExcelWriter('Rooms&Services.xlsx')
df_rooms_only.to_excel(writer, 'ROOM items', float_format='%5.2f')
df_services_only.to_excel(writer, 'SERVICE items', float_format='%5.2f')
writer.save()
