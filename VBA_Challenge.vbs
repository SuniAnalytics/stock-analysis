Sub AllStocksAnalysisRefactored()

    Dim startTime As Single
    Dim endTime  As Single
    
    'User prompt - Capture Year for which the output analysis needs to run
    yearValue = InputBox("What year would you like to run the analysis on?")
    
    'Variable to capture start time of the macro run
    startTime = Timer
    
    'Format the output sheet on the "All Stocks Analysis" worksheet
    Worksheets("All Stocks Analysis").Activate
    Range("A1").Value = "All Stocks Analysis (" + yearValue + ")"
   
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

 
    'Initialize an array of all tickers.
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
    
    'Activate the data worksheet based on the year input.
    Worksheets(yearValue).Activate

    'Get the number of rows to loop over from the source data sheet
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row

    '1a) Create a ticker Index
    tickerIndex = 0
    
    '1b) Create three output arrays to capture the data from source sheet and load into analysis sheet
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single

    ''2a) Create a for loop to initialize the tickerVolumes to zero.   
    
    For i = 0 To 11
       tickerVolumes(i) = 0
    Next i
    
    ''2b) Loop over all the rows in the source spreadsheet.
    
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        'Get Volume of the ticket from each looping row and eight column
        
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
    
        '3a) Increase volume for current ticker
        'Get first row's price to store ticker's starting price
        
              
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).value <> tickers(tickerIndex) Then
               
               tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
               
        End If   
   
    
        '3c) check if the current row is the last row with the selected ticker
        'If the next row???s ticker doesn???t match, increase the tickerIndex.
       
         'Get Last row's Closeing price to store ticker's closing price at the year end
         
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).value <> tickers(tickerIndex) Then
        
             tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        
        End If
    
        '3d Increase the tickerIndex if we reach the last row of the ticket. This will move code to next ticker
        
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        
               tickerIndex = tickerIndex + 1
               
        End If
        
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11    
                    
            Sheets("All Stocks Analysis").Activate
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

            'Color the cell green
            Cells(i, 3).Interior.Color = vbGreen

            Else

            'Color the cell red
            Cells(i, 3).Interior.Color = vbRed
        
            End If
    Next i

    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
      
End Sub

