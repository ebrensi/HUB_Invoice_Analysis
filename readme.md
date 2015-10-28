# IHO Invoice Analysis

This repo is for a project at Oakland Impact Hub to determine a better room rental pricing scheme.  The current method involves hourly rates tiered by renter type (part-time member,  full-time member, non-member), and various discounts when the price determined using current rates seems too high.  We want a more consistent rate scheme that works without resorting to arbitrary fee-waivers and discounts.


## Original data
The data we are working with come from invoices for IHO space bookings.  Each invoice is an Excel worksheet with itemized charges for the room being booked, as well as for other services.  An example invoice is here: 
![example_invoice](example_invoice.pdf)
* Each invoice spreadsheet contains a sub-table of items that fees were charged for, including but not limited to: room, setup/clean-up, IHO staff support, audio technician, etc.  Each of these items has an associated fee: either a flat fee, or a rate and number of hours, and subtotal (rate * hours), and sometimes a discount field given as a percentage.


Since these invoices contain peoples' contact information, the original invoices are not included in this repo.  Instead, we include the json file produced by `import_workbooks.py`, which contains all of the invoice info necessary for analysis.

### Imported invoice data
[`invoices.json`](invoices.json) contains a nested dictionary data structure with sheet-names as keys at the top-level. 
Each invoice (usually) contains:
* items that IHO charged money for: typically rooms or other services 
* RATE info is contained in one cell of the original invoice, and is based on the type of renter and day-type: 
eg.
  * Non-Member weekend rental
  * Non-Member weekday
  * Part-time member weeend, or weekday rental
  * Full-time member weekend
  * Full-time member weekday rental
  * Discounts:
    - Founder Discount
  	- Multi Room Discount
  	- Full Day Discount
  	- Partner Discount


### Item Classification
[`prep_data.py`](prep_data.py) classifies items into item-type (room, service, or other) and RATE information into rate/discount types.  It also fills as much missing info as possible and computes subtotals.  The result of running `prep_data` is the table `invoice_data`, which contains all invoice items and (sub)totals for each invoice. 


## Analysis
First We want to query general information about which rooms were rented and at what rates, and what discounts were applied, as well as how much income was reduced by each discount.  For example: What was the average income for renting the Broadway room to a part-time member for a full-day (5.5+ hours) on a weekday?

We also would like to make queries about package deals.  For example: what was typically the total income for a rental that included the Atrium?
