
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
    Dim tickerIndex As Integer
          ' start each ticker index at zero, to restart counting at each row
                tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerStartingPrice(12) As Single
    Dim tickerEndingprice(12) As Single
    Dim tickerVolumes(12) As Long
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For tickerIndex = 0 To 11
        'set tickerVolume to zero and
            tickerVolumes(tickerIndex) = 0
        
    ''2b) Loop over all the rows in the spreadsheet.
          ' Activate the specific worksheet
                Worksheets(yearValue).Activate
          'start at row 2 since row 1 is header
                For j = 2 To RowCount
    
    '3a) Increase volume for current ticker
          'reference the total volume on column H or 8 on spreadsheet
          tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value

    '3b) Check if the current row is the first row with the selected tickerIndex.
            ' need to check if row above = a specific ticker first to determine start point
             'If  Then
            If Cells(j, 1).Value = tickers(tickerIndex) And Cells(j - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrice(tickerIndex) = Cells(j, 6).Value
            'End If
            End If

    '3c) check if the current row is the last row with the selected ticker
          'If the next row’s ticker doesn’t match, increase the tickerIndex.
             If Cells(j, 1).Value = tickers(tickerIndex) And Cells(j + 1, 1).Value <> tickers(tickerIndex) Then
             tickerEndingprice(tickerIndex) = Cells(j, 6).Value
             End If

    '3d Increase the tickerIndex. If next row insn't the same as previos move to next ticker on ticker array
         If Cells(j, 1).Value = tickers(tickerIndex) And Cells(j + 1, 1).Value <> tickers(tickerIndex) Then
         tickerIndex = tickerIndex + 1
         ' End If
           End If
        
        Next j
        
        'Remember to close loop for ticker index!
        
        Next tickerIndex
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
         For j = 0 To 11
        'Need to activate output worksheet
        Worksheets("All Stocks Analysis").Activate
        'Ticker Row staring on row 4 and adding the rows by ticker. Total 12 rows
        Cells(4 + j, 1).Value = tickers(j)
        'Total Daily Volume
        Cells(4 + j, 2).Value = tickerVolumes(j)
        'Return value by divind the Ending/Starting Price
        Cells(4 + j, 3).Value = tickerEndingprice(j) / tickerStartingPrice(j) - 1

         Next j

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