# Simple Reseller Reporting System (SRRS)

This is a small collection of scripts for a small business that sells products for other people or businesses (in other words, a reseller), licensed under GPL version 2. The scripts were developed in the context of a second hand store in Sweden that uses iZettle. iZettle is a point of sales system and provider owned by iZettle AB, and we have no affiliation with the company. 

The functionality includes summing particular subsets of the sales within a custom time period, template-based automatic generation of sales reports for every customer, and automatically emailing the sales reports to the customers.


## (Lack of) Design

I would like to point out that this work was done in limited free time and if this was a professional project there are many things I would have done differently (like a proper DB back-end, better help system and online documentation, CLI arguments, better composability and code design, etc.). 

An important design consideration was that the underlying business logic should be fully exposed so that every part done by the scripts could still be done by hand without extensive IT training; basically insurance for the "what-if-I-die-tomorrow" hypothetical, so that my girlfriend would have been able to run her business without me. As a result, some parts that might look amateurish on first sight are actually conscious design choices. 


## Functions

At the moment, there is no user guide for SRRS, you will have to read the source code or run the scripts to figure out how they work. Nevertheless, here is an overview of the available scripts and their purpose. 

### Config

All important configuration parameters can be adjusted in config/srrs_config.py, and the variables should be mostly self-explanatory. All scripts should behave accordingly and deviations are bugs. 

### Templates

The templates for the emails (text file) and the reports (LaTeX file) can be found in the templates sub-directory. 

### Main scripts:

- sum-weekly-commission-sales.py - Sum all sales that match a specific product pattern. 
- generate-reports.py - Generate a week-based sales report for each customer for a specific week. 
- mail-reports.py - Email all reports to the customers. 


### Support scripts:

- sum-own-monthly-sales.py - Sum only own sales. Needed for tax reporting. 
- sum-quarterly-sales.py - Sum all sales of the quarter. Needed for tax reporting. 
- sum-single-sale.py - Sum only the sales of a single product for a specified period. 

- create-new-bookings-file.py - Takes a template sheet in the bookings spreadsheet and copies it 52 times. Run this once a year to create a new bookings file. 
- create-overview-sheet.py - Creates an overview sheet in the bookings file that links into the other sheets to aggregate booking information in one place for an easy overview.
- convert-xls-to-xlsx.sh (***obsolete***) - Convert old Excel file format to new file format. 


## Dependencies

- python3 - Python 3 and its standard library. 
- openpyxl - Python library for reading and writing xlsx files.
- LaTeX - A full LaTeX subsystem is needed to create the customer reports in pdf. Currently texlive with pdflatex is used. Any other LaTeX subsystem should work fine, perhaps with minor tweaks. 


## Limitations

This is neither a very user-friendly system nor is it very robust. A certain number of assumptions must be fulfilled for these scripts to work properly. These assumptions are:

- The file structure of the bookings file and the customer database must have the expected format. They are easy to adapt, but adaptation is necessary if the file structure is not matched.

- The product names are expected to have a specific prefix so that they can be sorted automatically into different categories. At the moment, HXX (where 0 <= X <= 9) specifies a product has an associated rent (H for hylla which is Swedish for shelf), SXX (where 0 <= X <= 9) for sales without a specific rent. Own sales have no corresponding pattern, but probably should have one. Maybe O for own. 

- The sales summation expects the raw data from iZettle, which so far has been stable. If the raw data format from iZettle ever changes, the summation scripts have to be updated.

- On rare occasions file names are expected to follow a specific pattern, for instance the files with the raw sales data downloaded from iZettle should always have the form "sales-$yyyy-$mm.xlsx" and be saved the folder specified by the config file. 


## Personal note

I wrote these scripts originally to automate the workflows in my girlfriends business. They were tailored specifically to her needs, and they contain a few assumptions that are true for a Swedish business, such as different VAT percentages, etc., but anyone with basic programming skills can adapt these to their own needs. 

My girlfriend no longer runs this particular business, so I have no interest in maintaining or developing this software any further. I am open-sourcing it in the hopes it will prove useful for someone at some point. 


## TODO

See the separate TODO file. 


## FAQ

**Wouldn't a proper database back-end be better than spreadsheet files?**

Absolutely, but my assumption was that a spreadsheet is easier to handle for a small business owner than a database, and an important design choice was to keep everything easily accessible to the business owner. 
If you want to add a proper DB back-end instead, that should be pretty simple.

