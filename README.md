This repository contains a .bas file containing a VBA script to go through a dataset and record yearly opening and closing prices for each ticker in the set, and then summarize the biggest movers among all of those tickers. It will repeat for every sheet of data in a workbook.
The repository also contains screenshots of the results for "Multiple_year_stock_data.xlsx", for each sheet in that file.

ASSUMPTIONS/CLARIFICATIONS:

- we are using 1 table and creating 2 more, and I call them:
    (1) raw data - what we start with;
    (2) summary data - summary of each ticker;
    (3) rundown - maxes and min of all tickers

- I wanted to run through the dataset one time, so I saved maxes and min as I go through the rows, and then printed them in the Rundown table after all rows have been processed.

- I made the code as generic as I could, so I wanted to save the last column so that the summary and rundown tables will be placed past the last column of the raw data. This isn't helpful if the raw data has more columns than the sample, because the code still assumes Col 1 (or A) to be Ticker, 3 (or C) to be Open, 6 (F) to be Close, and 7 (G) to be Volume.

- I included the clearData subroutine because I used it a lot while tweaking my code. It makes similar assumptions, and it is not called out in the main subroutine.

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

PSEUDOCODE:
for each sheet:
   sort raw data by date and then by ticker
   set up summary table headers
   initialize summary data (assign variables) from row 2 of raw data
   for each row(3 to end):
       is ticker(current row) different from ticker(saved)?
           yes:
               print saved variables to curr row of summary
                   -conditional formatting for Yearly Change
               check saved vars for rundown min/max
                   (if new max, save value and ticker name)
               go to next row of summary
               calculate variables
               check for min/max for rundown
           no:
               calculate variables
   set up rundown table (row/column titles)
   print maxPctInc and corresponding ticker
   print minPctInc and corresponding ticker
   print maxTotVol and corresponding ticker

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

All code is my original work. I occasionally consulted with Google, which sent me to Microsoft or StackOverflow for advice/reminders about which VBA methods and functions to use, but the code itself is mine. I did learn some things in my class, like finding the last row of data, but I did a little bit of Google research before it was brought up in class.
