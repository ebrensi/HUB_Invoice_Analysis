# IHO Invoice Analysis

This repo is for a project at Oakland Impact Hub to determine a better room rental pricing scheme.  The current method involves hourly rates tiered by renter type (part-time member,  full-time member, non-member), and various discounts when the price determined using current rates seems too high.  We want a more consistent rate scheme that works without resorting to arbitrary fee-waivers and discounts.


## Original data
The data we are working with come from invoices for IHO space bookings.  Each invoice is an Excel worksheet with itemized charges for the room being booked, as well as other services.  An example invoice is here: [example_invoice](page.pdf)
* Each invoice spreadsheet contains a sub-table of items that fees were charged for, including but not limited to: room, setup/clean-up, IHO staff support, audio technician, etc.  Each of these items has an associated fee: either a flat fee, or a rate and number of hours, and subtotal (rate * hours), and sometimes a discount field given as a percentage.


Since these invoices contain people's contact information, the original invoices are not included in this repo.  Instead, we include the json file produced by `import_workbooks.py`, which contains all of the invoice info necessary for analysis.

### Data Structure
This json file contains a nested dictionary data structure with sheet-names as keys at the top-level. 

Rooms (rentable spaces) are categorized as follows:
 - On Broadway, Atrium, Jingletown Lounge, OMI Gallery, Meridian Room, Uptown, Downtown, East Oakland, West Oakland, kitchen, meditation-room

For each room we are to determine average income from each type of rental from the folowing categories:
 - Non Member weekend
 - Non Member weekday
 - part time member weeend
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


