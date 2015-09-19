# IHO Invoice Analysis
This repo is for Mikayla's project to determine a better room rental pricing scheme.  The current method involves hourly rates tiered by renter type (part-time member,  full-time member, non-member), and various discounts when the price determined using current rates seems too high.  Mikayla wants a more consistent rate scheme that works without resorting to arbitrary fee-waivers and discounts.

In order to do this we look at income from all past room rental invoices (about 1.5 years worth), which give us a sense of what the rooms are worth.

## Data Preparation/cleaning
The orignal invoices are in the form of Excel (.xlsx) sheets combined into two workbooks, that are not included in this repo because they contain peoples' names and contact info.
The majority of the work for this project is in parsing these inconsistently formatted sheets.  Over the last few years, IHO staff (i.e. Mikayla) have modified the invoice template, introducing new fields and moving fields to different columns, etc.  This makes automated parsing not so straightforward.

* Each invoice spreadsheet contains a sub-table of items that fees were charged for, including but not limited to: room, setup/clean-up, IHO staff support, audio technician, etc.  Each of these items has an associated fee: either a flat fee, or a rate and number of hours, and subtotal (rate * hours), and sometimes a discount field given as a percentage.
  - sometimes the amount field contains a string ("comped", "waved", "included") indicating no fee.  In that case we set the amount to the numeric value 0, and set discount to 1 for that item.

* The first step in assembling the invoice data is to put all relevant invoice info into a json dictionary, with sheet-names as keys. this is because accessing Excel files with Python is slow and it's best to parse them as few times as possible.  This dictionary is saved as a json file and without sensitive name and contact info.

## Analysis
Rooms (rentable spaces) are categorized as follows:
 - On Broadway, Atrium, Jingletown Lounge, OMI Gallery, Meridian Room, Uptown, Downtown, East Oakland, West Oakland, kitchen, meditation-room

For each room we are to determine average income from each type of rental from the folowing categories:
 - Non Member weekend
 - Non Member weekday
 - part time member weeend
 - part time member weekday
 - full time member weekend
 - full time member weekday

Rentals are also distinguished by:
 - half day = 5.5hrs or less
 - full day = 6hrs or more


We also want to determine how often discounts are applied.  That is, totals or toal percentage for:
 - Founder Discount
 - Multi Room Discount
 - Full Day Discount
 - Partner Discount


