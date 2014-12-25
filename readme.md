A python2 script that:
Loads an excel spreadsheet,
Compares each product in the loaded spreadsheet to its costco.com listing, using ProductID as the key.  Specifically, the script checks the following values:  ItemNumber, Brand, ProductName, catEntryId (found in source code), categoryId (found in source code).
Lists any discrepancies between the spreadsheet and the costco.com listing.

Dependencies: 

openpyxl 
http://openpyxl.readthedocs.org/

lxml
http://lxml.de/

requests
http://docs.python-requests.org/