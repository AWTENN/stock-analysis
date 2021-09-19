# Stock-analysis
  ## Overview
- *The purpose of this analysis was to find the Total Daily Volume and Return for each stock listed and refactor the code that did this in the module and make the process time faster.*

## Code
Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis Refactored").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

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
        tickerVolumes(tickerIndex) = 0
        Next i
        

    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value

        
        
     '3b) Check if the current row is the first row with the selected tickerIndex.
         If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            End If

            
    '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
            If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
            
        '3d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
           End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis Refactored").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
        
    Next i
    
    'Formatting
    Worksheets("All Stocks Analysis Refactored").Activate
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

         
  ## Results
   
![VBA_Challenge_2017.png](https://github.com/AWTENN/stock-analysis/blob/main/VBA_Challenge_2017_original.png)
![VBA_Challenge_2017.png](https://github.com/AWTENN/stock-analysis/blob/main/VBA_Challenge_2017.png)
![VBA_Challenge_2018.png](https://github.com/AWTENN/stock-analysis/blob/main/VBA_Challenge_2018_original.png)
![VBA_Challenge_2018.png](https://github.com/AWTENN/stock-analysis/blob/main/VBA_Challenge_2018.png)
![VBA_Challenge_2017.png](https://github.com/AWTENN/stock-analysis/blob/main/VBA_Challenge_2017_Stocks.png)
![VBA_Challenge_2017.png](https://github.com/AWTENN/stock-analysis/blob/main/VBA_Challenge_2018_Stocks.png)
- As seen in the photos above, I would choose the RUN and ENPH stocks to invest in because they are the only two stocks to have positive return rates in 2017 and 2018, with a big increase in 2018 for both Total Daily Volumes of over 300,000 stocks bought.
- The execution times of the original All Stocks Analysis code took around 1000% longer than the refactored code as seen in the pictures above. I believe this was a cause of the if statement in the original code had an “and” statement, and there was a ticker index instead of just using the tickers String statement.

## Summary

- The advantage of refactoring code is the optimization of the time the code takes to produce the information wanted. However, the disadvantage of refactoring is changing the code could be a very long and tedious process, as I learned with this.
- The pro for this VBA Challenge code is it took less time to produce the data we were looking for. The con was that it took me a day or two to refactor the code precisely so that there were no errors, and it took less time than the original code did.
