# Green Stocks Analysis


### Overview of Analysis
For this challenge, we are helping Steve create VBA Macros that will analyze stocks in 'green' companies. We do so by creating loops to analyze large amounts of data very quickly. In this analysis we re-factored our original code, which analyzed green stocks one year at a time. 

### Results
The results of this refactoring showed that the code is able to be analyzed faster than before. 

```
Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
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
    
    For i = 0 To 11
        ticker = tickers(i)
        
    Next i

    '1b) Create three output arrays
    
    Dim tickerVolumes As Long
     
    Dim tickerStaringPrices As Single
    
    Dim tickerEndingPrices As Single
    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
     Worksheets(yearValue).Activate
     tickerVolumes = 0
        
    ''2b) Loop over all the rows in the spreadsheet.
        'this one has to be a j because i has already been used??
    For j = 2 To RowCount
    
        '3a) Increase volume for current ticker
        If Cells(j, 1).Value = tickerIndex Then
            tickerVolumes = tickerVolumes + Cells(j, 8).Value
        End If
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
         If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
           tickerstartingPrices = Cells(j, 6).Value
        End If
            
            
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row‚Äôs ticker doesn‚Äôt match, increase the tickerIndex.
        'If  Then
            If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                tickerEndingPrices = Cells(j, 6).Value
            End If
            

            '3d Increase the tickerIndex.
            If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                    tickerstartingPrices = Cells(j, 6).Value
            End If
            
            If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                tickerEndingPrices = Cells(j, 6).Value
            End If

        'End If
    
    Next j
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For k = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
            Cells(4 + k, 1).Value = tickerIndex
            Cells(4 + k, 2).Value = tickerVolumes
            Cells(4 + k, 3).Value = tickerEndingPrices / tickerstartingPrices - 1

        
    Next k
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.00%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For l = dataRowStart To dataRowEnd
        
        If Cells(l, 3) > 0 Then
            
            Cells(l, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(l, 3).Interior.Color = vbRed
            
        End If
        
    Next l

```

### Summary
#### Advantages and Disadvantages
One advantage to refactoring code is that we only need to make small changes to our code in order to analyze this large dataset. The disadvantages are numerous if you are not organized - commenting throughout the code to explain the purpose of the different sections is immensely helpful. If your code is not organized and annotated, it can be challenging to identify where the code needs to be edited in order to work. 

When refactoring this code it was important to keep organized in order to ensure the code works properly and mistakes can be identified and fixed. It was also faster after it was refactored, meaning that if you were working with an even larger dataset, the amount of time it might save could be significant. 
