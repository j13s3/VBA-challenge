# VBA-challenge
Data Analytics Boot Camp - Module 2 Challenge - VBA Scripting - Jalaj Sharma

For this challenge, a single Excel Visual Basic script has been written with the subroutine "stock_summary".
There are comments in the script for reference.

The script executes logic in the following sequence:
    1. Declare ALL variables and constants
    2. Count the number of sheets in the workbook
    3. For every individual sheet, perform the following tasks as part of a loop
        a. Set up the headers for the summary table
        b. Count the number of rows of data
        c. For every row in the worksheet, do the following as each row is analysed:
            i. Part 1: Isolate unique stock tickers and populate the summary table region
            ii. Part 2: Using the opening price of the year and the closing price of the same year, calculate the change in price and apply conditional formatting (Please note that percentage changes of 0.00% have not been altered)
            iii. Part 3: For every identified unique ticker, accumulate the total stock volume
            iv. With all the three parts completed, create the full summary table
        d. Using the summary table, find the minimum and maximum proce changes and the greatest stock volume traded
        e. Using the FIND function, identify the ticker associated to the minimum/maximum parameters
        d. Perform basic formating on the active sheet
    4. Once all tasks on all sheets have been completed, bring up a message to say that the code executed fully

The following online resources were referred to:
    1. https://excelchamps.com/vba/rows-count/
    2. https://stackoverflow.com/questions/32190029/find-row-number-of-matching-value
    3. https://www.mrexcel.com/board/threads/vba-change-number-of-decimal-places-of-a-percentage.521221/
    4. https://excelchamps.com/vba/autofit/#AutoFit_Entire_Worksheet
