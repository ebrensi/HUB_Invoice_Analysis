#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon Aug 17 14:23:20 2015

@author: Efrem
"""

from IHO_event_invoice import *
import json
# import re

INVOICE_NUM_CUTOFF = 2035
EXCLUSIONS = {}


item_classes = {
    ROOM: {
        'EAST_OAK': 'east',
        'WEST_OAK': 'west',
        'DOWNTOWN': 'downtown',
        'UPTOWN': 'uptown',
        'MERIDIAN': 'meridian',
        'OMI': 'gallery|omi',
        'JINGLETOWN': 'jingle[-| ]?town|mezzanine',
        'ATRIUM': 'atrium',
        'BROADWAY': 'broadway',
        'PATIO': 'patio',
        'MEDITATION': 'meditation',
        'KITCHEN': 'kitchen',
    },

    SERVICE: {
        'SETUP_RESET': 'set[-| ]?up|pre[-| ]?event|post[-| ]?event',
        'STAFFING': 'staff|manager',
        'A/V': 'A/V|technician|sound',
        'JANITORIAL': 'janitorial|waste|cleaning',
        'DRINKS': 'coffee|wine',
        'COMPOSTABLES': 'compost',
        'SECURITY': 'security'
    },

    ITEM_TOT: {
        None: 'total'
    }
}

# This is so that when we sort by item_type, item-total is last for each invoice
# It must be reversed because we will be sorting by descending order
ITEM_CLASS_SORT_ORDER = [ROOM, SERVICE, OTHER, ITEM_TOT]


# This is how we categorize RATE info
RATE_classes = {
    member_type: {
        'NON_MEMBER': 'none?[ |-]member',
        'PART_TIME': 'part[-| ]?time',
        'FULL_TIME': 'full[ |-]?time|full member',
    },

    day_type: {
        'WEEKEND': 'weekend',
        'WEEKDAY': 'weekday|wkday'
    },

    day_dur: {
        'PARTIAL_DAY': 'Half[-| ]?Day',
        'FULL_DAY': 'Full[-| ]?day',
    },

    discount_type: {
        'MULTI_ROOM': 'Multi[-| ]?Room',
        'MULTI_DAY': 'Multi[-| ]?day',
        'REOCURRING': 'Multi[-| ]?event|reo?ccuring',
        'FOUNDER': 'Founder',
        'FRIEND': 'Partner|sposor|WITH|share|org',
        'RETURNING': 'Returning[-| ]?client'
    }
}


# The data we read in have different fields that correspond to the same thing
#  Here we specify equivalence classes
FIELD_CLASSES = {
    'DATE': ['DATE', 'DATE OF EVENT', 'DATES OF EVENT'],
    'HOURS_UNITS': ['ESTIMATED HOURS', 'HOURS', 'HOURS/UNITS'],
    'TOTAL': ['TOTAL', ' TOTAL', 'ESTIMATE TOTAL', 'ESTIMATED TOTAL'],
    'DISCOUNT': ['DISCOUNT', 'DONATION']
}


# This function produces a list of dictionaries, each entry one item
#  from a nested invoice dictionary
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


# ************************************
# def main(argv):

# Read in invoices data from json source
with open(JSON_DATA_FNAME, 'r') as in_file:
    invoices = json.load(in_file)

df = pd.DataFrame(flatten_dict(invoices)).drop_duplicates().dropna(how='all')

# exclude invoices with no invoice#, cancellations,
# or invoice# < INVOICE_NUM_CUTOFF
invoice_num = df['invoice'].str.extract('(\d+)').str.strip().astype(float)
cancellations = df['invoice'].str.contains('cancel', na=False, case=False)

exclude_mask = (invoice_num.isnull() |
                (invoice_num < INVOICE_NUM_CUTOFF) |
                cancellations)

print("Excluding \n%s" % (df['invoice'][exclude_mask].unique()))

df = df[~exclude_mask]

# Join equivalent columns as specified in FIELD_CLASSES
for field in FIELD_CLASSES:
    data_to_merge = [df[col_name].dropna()
                     for col_name in FIELD_CLASSES[field]]
    df[field] = pd.concat(data_to_merge).reindex_like(df)

df = df[['invoice', 'invoice_date', 'DATE', 'DESCRIPTION', 'AMOUNT',
         'HOURS_UNITS', 'SUBTOTAL', 'DISCOUNT', 'TOTAL', 'RATE']]

#  drop rows where these fields are all empty or zero
isnull = df[['AMOUNT', 'SUBTOTAL', 'DISCOUNT', 'TOTAL']].isnull()
iszero = df[['AMOUNT', 'SUBTOTAL', 'DISCOUNT', 'TOTAL']] == 0
is_either = (isnull | iszero).all(axis=1)
df = df[~is_either]

# Parse DESCRIPTION field entries into item_type and 'item'
#  as defined in the item_classes dictionary
df[item_type] = None
df['item'] = None

for ic in item_classes:
    for item in item_classes[ic]:
        this_item_mask = (
            df['DESCRIPTION'].str
            .contains(item_classes[ic][item],
                      case=False, na=False))

        df.loc[this_item_mask, item_type] = ic
        df.loc[this_item_mask, 'item'] = item

# Remaining items (that were not assigned an item_type)
#   get classified as OTHER
other_mask = df[item_type].isnull()
df.loc[other_mask, item_type] = OTHER
df.loc[other_mask, 'item'] = df.loc[other_mask, 'DESCRIPTION']

# Set item_type as categorical and set sort order
#   as specified by ITEM_CLASS_SORT_ORDER
df.loc[:, item_type] = (df.loc[:, item_type]
                        .astype('category',
                                categories=ITEM_CLASS_SORT_ORDER[::-1],
                                ordered=True))

# Parse RATE field entries into classes
#   as defined in the RATE_classes dictionary
for RATE_class in RATE_classes:
    for rate in RATE_classes[RATE_class]:
        rate_mask = df['RATE'].str.contains(RATE_classes[RATE_class][rate],
                                            case=False,
                                            na=False)

        df.loc[rate_mask, RATE_class] = rate

# Convert numeric field data to numbers
#  and fill-in as much missing information as possible

# first we set values labeled 'waved', 'comped', 'included' to zero
no_charge_indicators = "wai?ved|comped|included|member"
no_charge_amount = df['AMOUNT'].str.contains(no_charge_indicators,
                                             case=False, na=False)
no_charge_subtot = df['SUBTOTAL'].str.contains(no_charge_indicators,
                                               case=False, na=False)

no_charge_tot = df['TOTAL'].str.contains(no_charge_indicators,
                                         case=False, na=False)

all_no_charge = no_charge_amount | no_charge_subtot | no_charge_tot
df.loc[all_no_charge, ['AMOUNT', 'SUBTOTAL', 'TOTAL']] = 0

#  Set DISCOUNT to 1 for any waived fees
df.loc[all_no_charge, 'DISCOUNT'] = 1

#  convert any 'flat fee' indicators to 1
#   so that the AMOUNT identifies with SUBTOTAL
df['HOURS_UNITS'].loc[df['HOURS_UNITS']
                      .str.contains('flat', case=False, na=False)] = 1

# convert all numeric field values to floats
#   (anything non-numeric becomes NaN)
NUMERIC_FIELDS = ['AMOUNT', 'SUBTOTAL', 'HOURS_UNITS', 'DISCOUNT', 'TOTAL']
for field in NUMERIC_FIELDS:
    df[field] = pd.to_numeric(df[field], errors='coerce')

# Fill-in missing hours and discounts for room rentals
#   where these values are implied by a multi-room deal,
#   particularly when DESCRIPTION field contains
#   certain keywords like 'Entire ...'
multi_room_item_mask = df['item'].str.contains('entire|level|rentals',
                                               na=False, case=False)
idx_to_drop = []

for item_idx, item in df[multi_room_item_mask].iterrows():
    associated_rooms = ((df['invoice'] == item['invoice']) &
                        (df[item_type] == ROOM) &
                        df['TOTAL'].isnull())

    if associated_rooms.any():
        for field in ['HOURS_UNITS', 'DISCOUNT']:
            df.loc[associated_rooms, field] = item[field]

        idx_to_drop.append(item_idx)

print('\nDropping multi-room items \n{}'
      .format(df.loc[idx_to_drop,
                     ['invoice', 'invoice_date', 'DESCRIPTION']]))
df = df.drop(idx_to_drop)

# fill in missing subtotal values.
#   We need these values for determining income.
# If AMOUNT and HOURS_UNITS are both non-empty
#  then compute SUBTOTAL = AMOUNT * HOURS_UNITS
notnull = df[['AMOUNT', 'HOURS_UNITS']].notnull().all(axis=1)
df.loc[notnull, 'SUBTOTAL'] = (df.loc[notnull, 'AMOUNT'] *
                               df.loc[notnull, 'HOURS_UNITS'])

not_tot_type = (df[item_type] != ITEM_TOT)

#  Fill-in all missing DISCOUNT entries with zero
df.loc[df['DISCOUNT'].isnull() &
       (df[item_type] != ITEM_TOT), 'DISCOUNT'] = 0

# Fill-in all missing TOTAL for items with non-empty SUBTOTAL
no_tot = df['TOTAL'].isnull() & df['SUBTOTAL'].notnull() & not_tot_type

df.loc[no_tot, 'TOTAL'] = (df.loc[no_tot, 'SUBTOTAL'] *
                           (1 - df.loc[no_tot, 'DISCOUNT']))

# Fill-in SUBTOTAL for items with TOTAL but no SUBTOTAL
#   (taking DISCOUNT into consideration)
no_subtot = df['TOTAL'].notnull() & df['SUBTOTAL'].isnull() & not_tot_type

df.loc[no_subtot, 'SUBTOTAL'] = (df['TOTAL'][no_subtot] /
                                 (1 - df['DISCOUNT'][no_subtot]))

# Eliminate items with no SUBTOTAL, DISCOUNT, nor TOTAL
mask = ((df['TOTAL'].isnull() | (df['TOTAL'] == 0)) &
        (df['SUBTOTAL'].isnull() | (df['SUBTOTAL'] == 0)) &
        (df['DISCOUNT'].isnull() | (df['DISCOUNT'] == 0)))

print('\nDropping items with no SUBTOTAL, DISCOUNT, or TOTAL: \n{}'
      .format(df[['invoice', 'invoice_date', 'DESCRIPTION']][mask]))

df = df[~mask]

# Manually compute SUBTOTAL and TOTAL fields in 'total' items
tots = [0, 0]
for idx, row in df.iterrows():
    if row[item_type] == ITEM_TOT:
        df.loc[idx, ['SUBTOTAL', 'TOTAL']] = tots
        if tots[1] != 0:
            df.loc[idx, 'DISCOUNT'] = 1 - tots[1] / tots[0]
        tots = [0, 0]
    else:
        if not pd.np.isnan(row['SUBTOTAL']):
            tots[0] += row['SUBTOTAL']
        if not pd.np.isnan(row['TOTAL']):
            tots[1] += row['TOTAL']

# Indicate if membership is unknown
df[member_type] = df[member_type].fillna('unknown')

# if a DISCOUNT is nonzero but no discount-type is indicated, give it a type
unknown_discount = df[discount_type].isnull() & df['DISCOUNT']
df.loc[unknown_discount, discount_type] = 'unknown'
df[discount_type] = df[discount_type].fillna('NONE')

# df['DATE'][event_date.isnull()].str.split('&|-')

# Replace missing day_type entries with day type determined by date
event_date = pd.to_datetime(df['DATE'].where(df[day_type].isnull()),
                            errors='coerce')
weekend_ind = event_date.dt.dayofweek >= 5
df.loc[weekend_ind, day_type] = 'WEEKEND'
df.loc[~weekend_ind & event_date.notnull(), day_type] = 'WEEKDAY'


df = df.sort_values(by=['invoice_date', 'invoice', item_type],
                    ascending=False)

line_item_fields = ['invoice',
                    'invoice_date',
                    'DATE',
                    item_type,
                    'item',
                    'AMOUNT',
                    'HOURS_UNITS',
                    'SUBTOTAL',
                    'DISCOUNT',
                    'TOTAL']

df[line_item_fields].to_csv(LINE_ITEMS_FNAME + '.csv',
                            index=False,
                            float_format='%6.2f')


# --- Whole-invoice classifications: ---

# inv_spec is to deal with a little confusion about whether we want to
#  group by invoice (which can contain multiple event dates)
#  or by date within each invoice as well.
inv_spec = ['invoice', 'DATE']

rooms_by_invoice = (df.query("{}=='{}'".format(item_type, ROOM))
                    .groupby(inv_spec, sort=False))

services_by_invoice = (df.query("{}=='{}'".format(item_type, SERVICE))
                       .groupby(inv_spec, sort=False))

setup_by_invoice = (df.query("item=='SETUP_RESET'")
                    .groupby(inv_spec, sort=False))


# Here we determine the overall length of time allotted for each "event"
# max number of hours that a room was rented for each event
event_room_hours = rooms_by_invoice['HOURS_UNITS'].max()

# total amout of setup/reset time for each event
event_setup_hours = setup_by_invoice['HOURS_UNITS'].sum()

# Total time that each event required
total_event_hours = event_room_hours.add(event_setup_hours, fill_value=0)


event_room_income = rooms_by_invoice[['SUBTOTAL', 'TOTAL']].sum()
event_service_income = services_by_invoice[['SUBTOTAL', 'TOTAL']].sum()
event_setup_income = setup_by_invoice[['SUBTOTAL', 'TOTAL']].sum()

# Total income in $ generated by each event
total_event_income = (event_room_income
                      .add(event_service_income, fill_value=0)
                      .add(event_setup_income, fill_value=0))


total_event_income['REDUCTION'] = (1 -
                                   total_event_income['TOTAL'] /
                                   total_event_income['SUBTOTAL'])


# Number of rooms rented for each event
event_num_rooms = rooms_by_invoice['item'].count()


other_fields = ['invoice_date', member_type,
                discount_type, day_type, day_dur]
by_event = (pd.concat([event_num_rooms,
                       total_event_hours,
                       total_event_income,
                       rooms_by_invoice[other_fields].first()], axis=1))

rules = {'item': '# ROOMS', 'HOURS_UNITS': 'HOURS'}
by_event.rename(columns=rules, inplace=True)

# Add a partial/full day specifier column to event summary table
by_event[day_dur] = (by_event['HOURS']
                     .map(lambda x: "FULL_DAY" if x >= 5.5
                          else "PARTIAL_DAY"))


# Add a column with list of rooms used for each event,
#  abbreviated by first letter
def room_list(room_series):
    return ",".join(room_series
                    .drop_duplicates()
                    .str[0]
                    .sort_values()
                    .tolist())

by_event['rooms'] = rooms_by_invoice['item'].agg(room_list)

by_event = by_event.sort_index(ascending=False)

# Save the event summaries table to csv
by_event.to_csv(
    INVOICE_SUMMARIES_FNAME + '.csv', float_format='%6.2f')

# Put tables into MySQL
to_mySQL(df[line_item_fields], 'line_items')
to_mySQL(by_event.reset_index(), 'invoice_summaries')

# if __name__ == "__main__":
#     main(sys.argv)
