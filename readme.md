# IHO Invoice Analysis

This repo is for a project that began in August of 2015 to determine a better room rental pricing scheme for
the Oakland Impact Hub co-working space.  The current method involves hourly rates tiered by renter type (part-time member,  full-time member, non-member), and various discounts when the price determined using current rates seems too high.  We want a more consistent rate scheme that works without resorting to arbitrary fee-waivers and discounts.

## Original data
The data we are working with come from invoices for IHO space bookings.  Each invoice is an Excel worksheet with itemized charges for the room being booked, as well as for other services.  An example invoice is here:  [example_invoice](example_invoice.pdf)

The relevant information from an invoice is contained in two places:
  1. The sub-table of items that fees were charged for, including but not limited to: room, setup/clean-up, IHO staff support, audio technician, etc.  Each of these items has an associated fee: either a flat fee, or a rate and number of hours, and subtotal (rate * hours), and sometimes a discount field given as a percentage.
  2. The RATE information field, contained in the bottom row of the box labeled `Notes:`. This cell contains information about what factors were used to determine rental rate, eg. member type and discount type


Since these invoices contain peoples' contact information, the original invoices are not included in this repo.  Instead, we include the json file produced by `import_workbooks.py`, which contains all of the invoice info necessary for analysis.

### Imported invoice data
[`IHO_event_invoices.json`](IHO_event_invoices.json) contains a nested dictionary data structure with sheet-names as keys at the top-level.
Each invoice record contains:
* items that IHO charged money for: eg. rooms and services
  - AMOUNT, DESCRIPTION, SUBTOTAL, DISCOUNT, TOTAL fields
* Renter and discount info
  - RATE field  


## Classification
The [`prep_data.py`](prep_data.py) script parses invoice information into consistent labels relevant to a whole invoice/event, or to a line-item within an invoice.

### Invoice/event Classification
The file [`IHO_event_invoice_summaries.csv`](IHO_event_invoice_summaries.csv) contains summary information from the line items grouped by invoice, where each row represents an invoice.  Some information is better understood when associated with a whole event rental.  Whole event information is parsed from the invoice RATE field or derived from line-item data.
  
  * The membership level of person/organization (`member_type`) responsible for the transaction is labeled `NON_MEMBER`, `PART_TIME`, `FULL_TIME`, or `unknown`.
  
  * `discount_type`, when given explicitly in the invoice, is classified as `MULTI_ROOM`, `MULTI_DAY`, `REOCURRING` (for an event that happens on several dates), `FOUNDER`, `FRIEND`, or `RETURNING`.
   
  * `day_type` is determined to be `WEEKDAY` or `WEEKEND` based on event-date and/or RATE info.
  
  * `day_dur` is `FULL_DAY` or `PARTIAL_DAY` based on whether the longest room rental plus setup/reset time exceeds 5.5 hours. 

### Item Classification
[`IHO_event_invoice_line_items.csv`](IHO_event_invoice_line_items.csv).  It contains the invoice data classified into item-type (room, service, or other) and RATE information into rate/discount types mentioned above.  We fill-in as much missing info as we can and compute subtotals with and without discount, when a discount is explicitly given as a percentage, or if the fee for an item is given as 'waived' or 'comped'.

#### Line item classification:
We classify line items by text-matching key terms in invoice DESCRIPTION field 
* `item_type` for a line-item is classified as ROOM, SERVICE, or OTHER
  * Items identified as `EAST_OAK`, `WEST_OAK`, `DOWNTOWN`, `UPTOWN`, `MERIDIAN`, `OMI`, `JINGLETOWN`, `ATRIUM`, `BROADWAY`, `PATIO`, `MEDITATION`, or `KITCHEN` are classified as ROOM.
   
  * Items idetified as `SETUP_RESET`, `STAFFING`, `A/V`, `JANITORIAL`,`DRINKS`, `COMPOSTABLES`, or `SECURITY` are classified as SERVICE.
  


## Analysis
First We want to query general information about which rooms were rented and at what rates, and what discounts were applied, as well as how much income was reduced by each discount.  For example: What was the average income for renting the Broadway room to a part-time member for a full-day (5.5+ hours) on a weekday?

We also would like to make queries about package deals.  For example: what was typically the total income for a rental that included the Atrium?



#### Rooms only
Here are tables of averages and totals for income generated by each room, grouped by discount, type of renter, and a few other classifiers.
  * [Averages](IHO_pricing_rooms_only_avg.csv)
  * [Totals](IHO_pricing_rooms_only_sum.csv)

Each invoice indicates room rental rate in the AMOUNT field. Given the actual income generated for each room and the number of hours it was rented, this table summarizes the effective rate for room in practice, itemized by stated discount and rental categories.
  * [Effective Room Rates](IHO_pricing_effective_room_rates.csv)

#### Services only
This table provides similar aggragate information for service line items associated with the venue rentals.
  * [Services](IHO_pricing_services_only.csv)
