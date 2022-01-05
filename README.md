# Overview of Project
  The purpose of this analysis was to identify which stocks performed positively and negatively using statistics from 2017 and 2018. We will use this analysis in determining which
stocks appear to be worth investing in. This was initially completed using Microsoft Excel's VBA to create a macro to organize and visualize the stock data, but was refactored to maximize efficiency and to provide continuity. 
# Results
To maximize efficiency, we refactored the original code that was written to generate reports for the years 2017 and 2018:
```
Sub AllStocksAnalysisRefactored()

    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet

    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    
    Worksheets("All Stocks Analysis").Activate

    'Initialize array of all tickers
    Dim tickers(12) As String
    
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
    
    'Activate data worksheet
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
    Next i
        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value

        'End If
        
            End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next rowâ€™s ticker doesnâ€™t match, increase the tickerIndex.
        'If  Then
            
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            End If

            '3d Increase the tickerIndex.
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
            
            
        'End If
            End If
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub
```
  Compared to our original code, this refactored code was able to run much faster by creating a ticker index to return cell values from an array. This saved us ~.56 seconds (.69 original vs .13 refactored) when generating the reports for 2017 and 2018. It might not seem much, but an 81% decrease is a significant increase when compiling thousands of different stocks and their performance. 
# 2018 Refactored Timed
![2018Results](/resources/VBA_Challenge_2018.png)
# 2017 Refactored Timed
![2017Results](/resources/VBA_Challenge_2017.png) 

# 2017 Results
  ![2017Ticker](/resources/VBA_Challenge_2017_Tickers.png)
  
  In 2017, we can see that almost all of the stocks, with the exception of TERP, performed healthily with a positive return at the end of the year. While this looks like it is safe to invest in almost any stock, 2018 will show us what stocks remained resilient.
 # 2018 Results
  ![2018Ticker](/resources/VBA_Challenge_2018_Tickers.png)
  
  Only ENPH and RUN were still performing strong at above 81%, meanwhile we could expect returns as low as 62.6% (DQ). Using this information, it would be best to have invested in ENPH and RUN.
  
# What are the advantages or disadvantages of refactoring code?
The advantages of refactoring code allows developers to optimize efficiency, cleanliness, and could make it easier to understand for others. When code is refactored, it is made easier to digest. Disadvantages of refactoring code is the consumption of time. Refactoring code means existing code is being modified, so it is extra work that could be outside of the scope of a project or takes time away from other tasks. 
# How do these pros and cons apply to refactoring the original VBA script?
Refactoring the original VBA script allowed the script to run faster, which the client wanted for continuity as this will be a tool they will use in the future when identifying 
stocks to invest into the later years. We were not on a strict timeline, so the cons did not apply to this project.
