#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from openpyxl import load_workbook
import pandas as pd
import re
import json
import time
import datetime
from collections import OrderedDict
import concurrent.futures
import sys


WORKBOOK_FILENAMES = ['original_data/IHO_OnGoing_InvoiceTemplate.xlsx',
                      'original_data/2015 OnGoing InvoiceTemplate.xlsx']

MAX_ROW_LENGTH = 10

date_pat = re.compile('(\d{4}-\d{2}-\d{2})')
rate_pat = re.compile('(.*rate:.*|.*rate.*)', re.IGNORECASE)


# Extract relevant data from one invoice worksheet and return it as a dict
def parse_sheet(ws):
    info = OrderedDict.fromkeys(['invoice_date', 'RATE'])
    sname = ws.title.strip()

    template_pattern = re.compile('template|quote', re.IGNORECASE)
    if template_pattern.search(sname):
        return False

    # make a dataframe from the current sheet
    df = pd.DataFrame([tuple([cell.value for cell in row[:MAX_ROW_LENGTH]])
                       for row in ws.rows])

    df = df.dropna(how='all', axis=[0, 1]).reset_index(drop=True)

    if any(df):
        # We will search search DataFrame df by column, from the last column
        #  to the first, since we know that what we're looking for is
        #  on the right side.  we're looking for a date (invoice_date),
        #  and a cell that has text of the form "RATE: XXXX" where XXX
        #  is some words describing the rate charged for this event.
        for col_name in reversed(df.columns):
            col_str = df[col_name].astype(str).str

            if not info['invoice_date']:
                date_cell = col_str.extract(date_pat).dropna()
                if len(date_cell) > 0:
                    info['invoice_date'] = date_cell.iloc[0]

            elif not info['RATE']:
                rate_cell = col_str.extract(rate_pat).dropna()
                if len(rate_cell) > 0:
                    info['RATE'] = (rate_cell.iloc[0]
                                    .replace('RATE:', '').strip())
                    break

        # Determine upper & lower boundaries for the item subtable
        sep = df[df[0].str.contains('^date',
                                    case=False, na=False)].index.tolist()

        if sep:
            table_header_row = sep[0]
        sep = (df[df[1].str.contains('total', case=False, na=False)]
               .index.tolist())
        if sep:
            last_row = sep[0]
        else:
            last_row = len(df)

        # grab the items sub-table into a DataFrame
        header_row = df.iloc[table_header_row].tolist()
        last_col = next(i for i, j
                        in reversed(list(enumerate(header_row)))
                        if j)
        header = [str(field) for field in header_row[0:last_col+1]]

        subsheet = df.iloc[table_header_row+1:last_row+1, 0:last_col+1]
        date_col_name = header[0]
        subsheet.columns = header

        if not subsheet[date_col_name].iloc[0]:
            subsheet[date_col_name].iloc[0] = '?'

        subsheet = subsheet.dropna(how='all',
                                   axis=[0, 1]).reset_index(drop=True)

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
                    subsheet.loc[i, date_col_name] = \
                     subsheet.loc[i-1, date_col_name]
                else:
                    subsheet.loc[i, date_col_name] = "unknown"

        items = subsheet.to_dict(orient="records")
        info['items'] = items

    return info


# *******************************
# Import a workbook of invoice sheets and store relevant data
#   for every invoice as a record in a json dictionary.
def import_workbook(workbook_file_name):

    #   use openpyxl to open workbook
    wb = load_workbook(workbook_file_name, read_only=True, data_only=True)
    worksheets = wb.worksheets

    invoices = {}

    for ws in worksheets:
        invoice_dict = parse_sheet(ws)
        if invoice_dict:
            invoices[ws.title] = invoice_dict
            print(ws.title)

    return invoices


def main(argv):
    start_time = time.time()
    print('Loading workbooks...')

    with concurrent.futures.ProcessPoolExecutor() as executor:
        wb_list = executor.map(import_workbook, WORKBOOK_FILENAMES)

    invoices = {inv_title: inv_dict
                for wb in wb_list
                for inv_title, inv_dict in wb.items()}

    elapsed_string = str(datetime.timedelta(seconds=time.time()-start_time))
    print('workbooks loaded in %s' % elapsed_string)

    with open('invoices.json', 'w') as out_file:
        out_file.write(json.dumps(invoices, indent=3))


if __name__ == "__main__":
    main(sys.argv)
