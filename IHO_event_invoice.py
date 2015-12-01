#!/usr/bin/env python

# common defs IHO invoice project
import pandas as pd
import sqlalchemy

NaN = pd.np.nan
WORKBOOK_FILES = ['~/Dropbox/Invoices/IHO_OnGoing_InvoiceTemplate.xlsx',
                  '~/Dropbox/Invoices/2015 OnGoing InvoiceTemplate.xlsx']

JSON_DATA_FNAME = 'IHO_event_invoices.json'
LINE_ITEMS_FNAME = 'IHO_event_invoice_line_items'
INVOICE_SUMMARIES_FNAME = 'IHO_event_invoice_summaries'


#  *********** Class and field name definitions *************
# Define item class field names
item_type = 'item_type'
ROOM = 'ROOM'
SERVICE = 'SERVICE'
ITEM_TOT = 'ITEM_TOT'
OTHER = 'OTHER'

# RATE classifier field names
member_type = 'membership_level'
day_type = 'day_type'
day_dur = 'day_dur'
discount_type = 'discount_type'


# This is a little function to output an easier-to-read csv file
#  for a multi-indexed DataFrame.  It eliminates duplicated index entries
#  along index columns.
# The csv file produced is meant to be used for viewing by humans.
def to_nice_csv(df, filename):
    x = df.reset_index()
    cols = df.index.names
    mask = (x[cols] == x[cols].shift())
    x.loc[:, cols] = x[cols].mask(mask, '')

    x.to_csv(filename, index=False, float_format='%5.2f')


# Put DataFrame into MySQL table
def to_mySQL(df, table_name):
    uri = "mysql+pymysql://root:password@localhost/IHO_venue_rentals"
    db = sqlalchemy.create_engine(uri)
    with db.connect() as conn, conn.begin():
        df.to_sql(table_name, conn, index=True, if_exists='replace')
