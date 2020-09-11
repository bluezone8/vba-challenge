# VBA CHALLENGE
------
### PURPOSE

The purpose of this document is to analyze data related to stock prices and volumes.  
#

### SOFTWARE CHOICE

Excel was the chosen software for anayzing the data for a number of reasons
1. The volume of the structured data is does not exceed the amount that can be processed by a desktop application.  
2.  VBA code is ideal for this scenario as it requires no additional infrastructure in terms of opening external files etc. so time can be better spent on the coding the analysis itself.    
3.  If additional graphical anaysis is desired, it can be easily added with little effort at a later time.

#

### INCLUDED ITEMS
1. EXCEL SPREADSHEET - HWSubmit_Multiple_year_stock_data.xlsm
2. TEXT - FinalSubmissionCode.txt
3. IMAGE - HWSubmit_2014 StockData.jpg
4. IMAGE - HWSubmit_2015 StockData.jpg
5. IMAGE - HWSubmit_2016 StockData.jpg

#

### DISCUSSION
The files included in the repository include:  the analysis workbook , images of the individual worksheets within the workbook after processing by the program, and the complete macro code written for the analysis.

The code has been reasonable commented.  However, one additional comment related to my thinking for the structure of the program is worth noting.  I determined to make the processing for a single worksheet in one subroutine and then call that subroutine from another subroutine which would then process all worksheets in the workbook.  