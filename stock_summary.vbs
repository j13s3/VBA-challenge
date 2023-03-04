Sub stock_summary():

    'Declare all variables
    Dim i, j, worksheet_count, summary_row_count As Integer
    Dim row_count As Long
    Dim start_price, end_price, total_stock_vol As Double
    Dim great_inc_range, great_dec_range, great_val_range As Range
    
    'Declare constant variables
    'This allows for the table to be shifted by just adjusting these row/column indexes
    Const cons_summary_column As Integer = 9
    Const cons_summary_row As Integer = 2

    'The first task is to find the number worksheets in a workbook.
    worksheet_count = ActiveWorkbook.Worksheets.Count
    
    'Once the number of sheets is known, a loop can be created to perform
    'the same tasks across the worksheets as a loop
    For i = 1 To worksheet_count
        
        'As the amount of data in each worksheet is not the same, number of
        'filled rows in each worksheet is to be determined
        'Using the worksheet index, it is possible to loop through sheets
        Worksheets(i).Activate
        
        summary_row_count = cons_summary_row 'This is the first row in the summary table where values are put
        
        'Set up the headers for the stock summary data
        Cells(cons_summary_row - 1, cons_summary_column).Value = "Ticker"
        Cells(cons_summary_row - 1, cons_summary_column + 1).Value = "Yearly Change"
        Cells(cons_summary_row - 1, cons_summary_column + 2).Value = "Percent Change"
        Cells(cons_summary_row - 1, cons_summary_column + 3).Value = "Total Stock Volume"
        Cells(cons_summary_row, cons_summary_column + 6).Value = "Greatest % Increase"
        Cells(cons_summary_row + 1, cons_summary_column + 6).Value = "Greatest % Decrease"
        Cells(cons_summary_row + 2, cons_summary_column + 6).Value = "Greatest Total Volume"
        Cells(cons_summary_row - 1, cons_summary_column + 7).Value = "Ticker"
        Cells(cons_summary_row - 1, cons_summary_column + 8).Value = "Value"
        
        'Calculate the number of filled rows (https://excelchamps.com/vba/rows-count/)
        row_count = ActiveSheet.UsedRange.Rows.Count
        
        'At this stage of the program, number of filled rows in the active
        'worksheet is known. The summary table is to be constructed now for
        'the active sheet.
        
        For j = 2 To row_count 'Loop through every row
        
            'Part 1: Unique stock tickers in the data are to be isolated
            'This is done by comparing the current and next cell.
            'The moment a difference in cell values is detected, the ticker
            'can be recorded
            
            'Part 2: Computing the yearly price changes
            'Instead of looking at the first trading day's and last
            'trading day's prices to find the yearly price changes with
            'the assumption that start/end dates are the same for all tickers,
            'the first entry of a particulat ticker should be the opening year
            'price and the last of the same ticker to be the closing year price.
            'It is clearly visible that the data is sequentially logged
            'My reasoning to do this is because sometimes publicly traded companies
            'go into trading halts. If a trading halt happens on the first day of the
            'year for a particular ticker, its start date is different to the others
            'This gets rid of the assumption that the start/end dates are the same for
            'ALL tickers
            
            'Part 3: Calculate total stock volume trade
            'As every row is scanned for a unique ticker code, the total stock volumes
            'can be aggregated. Need to zero the totaliser variable once the next unique
            'ticker is located
                   
            If (j = 2) Then 'For Part 2
                start_price = Cells(j, 3).Value
            End If
            
            If (Cells(j, 1).Value <> Cells(j + 1, 1).Value) Then 'For all Parts
            
                Cells(summary_row_count, cons_summary_column).Value = Cells(j, 1).Value 'For Part 1
                
                end_price = Cells(j, 6).Value 'For Part 2
                Cells(summary_row_count, cons_summary_column + 1).Value = Round(end_price - start_price, 2) 'For Part 2 - Yearly Change
                If (Cells(summary_row_count, cons_summary_column + 1).Value < 0) Then ' For Part 2 - Conditional Formatting
                    Cells(summary_row_count, cons_summary_column + 1).Interior.Color = vbRed
                ElseIf (Cells(summary_row_count, cons_summary_column + 1).Value > 0) Then
                    Cells(summary_row_count, cons_summary_column + 1).Interior.Color = vbGreen
                Else
                    'Do Nothing for Zero Values
                End If
                    
                Cells(summary_row_count, cons_summary_column + 2).Value = (end_price - start_price) / start_price 'For Part 2 - Percent Change
                
                total_stock_vol = total_stock_vol + Cells(j, 7).Value 'For Part 3
                Cells(summary_row_count, cons_summary_column + 3).Value = total_stock_vol
                                
                'At this point of the FOR loop, the next unique ticker has been located
                summary_row_count = summary_row_count + 1 'Prepare for next summary table entry
                start_price = Cells(j + 1, 3).Value
                total_stock_vol = 0
                
            Else
            
                total_stock_vol = total_stock_vol + Cells(j, 7).Value 'For Part 3
                
            End If
            
        Next j
        
        'Part 4: Calculate the greatest parameters - increase, decreases, and total stock volume
        'With the summary table setup, find function and min/max worksheet functions can be used to populate
        'the greatest parameters
        Cells(cons_summary_row, cons_summary_column + 8).Value = WorksheetFunction.Max(Columns(cons_summary_column + 2))
        Cells(cons_summary_row + 1, cons_summary_column + 8).Value = WorksheetFunction.Min(Columns(cons_summary_column + 2))
        Cells(cons_summary_row + 2, cons_summary_column + 8).Value = WorksheetFunction.Max(Columns(cons_summary_column + 3))
        
        'https://stackoverflow.com/questions/32190029/find-row-number-of-matching-value
        'Running into problems with the FIND function when looking for Percent format values.
        'So, best to convert to Test format first then convert to Percent after using FIND
        Columns(cons_summary_column + 2).NumberFormat = "@" 'For Part 2 - Percent Change
        Cells(cons_summary_row, cons_summary_column + 8).NumberFormat = "@" 'For Part 4 - Greatest Percentage Change
        Cells(cons_summary_row + 1, cons_summary_column + 8).NumberFormat = "@"
        Set great_inc_range = Columns(cons_summary_column + 2).Find(Cells(cons_summary_row, cons_summary_column + 8).Value)
        Set great_dec_range = Columns(cons_summary_column + 2).Find(Cells(cons_summary_row + 1, cons_summary_column + 8).Value)
        Set great_val_range = Columns(cons_summary_column + 3).Find(Cells(cons_summary_row + 2, cons_summary_column + 8).Value)
        
        Cells(cons_summary_row, cons_summary_column + 7).Value = Cells(great_inc_range.Row, cons_summary_column).Value
        Cells(cons_summary_row + 1, cons_summary_column + 7).Value = Cells(great_dec_range.Row, cons_summary_column).Value
        Cells(cons_summary_row + 2, cons_summary_column + 7).Value = Cells(great_val_range.Row, cons_summary_column).Value
        
        'Formatting settings to the active sheet
        'https://www.mrexcel.com/board/threads/vba-change-number-of-decimal-places-of-a-percentage.521221/
        Columns(cons_summary_column + 2).NumberFormat = "0.00%" 'For Part 2 - Percent Change
        Cells(cons_summary_row, cons_summary_column + 8).NumberFormat = "0.00%" 'For Part 4 - Greatest Percentage Change
        Cells(cons_summary_row + 1, cons_summary_column + 8).NumberFormat = "0.00%" 'For Part 4 - Greatest Percentage Change
        
        'https://excelchamps.com/vba/autofit/#AutoFit_Entire_Worksheet
        ActiveSheet.Cells.EntireColumn.AutoFit
        ActiveSheet.Cells.EntireRow.AutoFit
        
    Next i
    
    MsgBox ("Stock Summary Compiled")

End Sub