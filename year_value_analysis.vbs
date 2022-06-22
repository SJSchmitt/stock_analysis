Sub yearValueAnalysis()
    Dim startTime As Single
    Dim endTime As Single
    
    yearValue = InputBox("What year would you like to run the analysis on?")
    startTime = Timer
    
     '1) Format output sheet
    Worksheets("All Stocks Analysis").Activate
    Cells(1, 1).Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Returns"
    
    '2) Initialize array of tickers
    Dim tickers(11) As String
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
    
    '3a) initialize variables for starting and ending price
    Dim startingPrice As Double
    Dim endingPrice As Double
    
    '3b) activate data worksheet
    Worksheets(yearValue).Activate
    
    '3c) get number of rows to loop over
    RowCount = Cells(rows.Count, "A").End(xlUp).Row
    
    '4) loop through all tickers
    For i = 0 To 11
    
        ticker = tickers(i)
        totalVolume = 0
        '5) loop through rows in data
        Worksheets(yearValue).Activate
        For j = 2 To RowCount
            '5a) get total daily volume for current ticker
            If Cells(j, 1).Value = ticker Then
                totalVolume = totalVolume + Cells(j, 8).Value
            End If
            '5b) get starting price for current ticker
            If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then
                startingPrice = Cells(j, 6).Value
            End If
            '5c) get ending price for current ticker
            If Cells(j, 1).Value = ticker And Cells(j + 1, 1).Value <> ticker Then
                endingPrice = Cells(j, 6).Value
            End If
            
            
        Next j
        
        '6) output data for current ticker
        Worksheets("All Stocks Analysis").Activate
        Cells(i + 4, 1).Value = tickers(i)
        Cells(i + 4, 2).Value = totalVolume
        Cells(i + 4, 3).Value = (endingPrice / startingPrice) - 1
        
    Next i
     'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    columns("B").AutoFit

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
    MsgBox "This program ran in " & (endTime - startTime) & " seconds for year " & (yearValue)
End Sub
