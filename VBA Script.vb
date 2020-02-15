Sub YearlyChange()

    For Each ws In Worksheets

        'Declare variables
        Dim LRow As Long
        Dim ShortTableRow As Long
        Dim StockVol As Double
        Dim TickerName As String
        Dim InitialPrice As Double
        Dim FinalPrice As Double
        Dim YearlyChange As Double
        Dim YearlyPercChange As Double
    
        'Find out the Last Row
        LRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        'Set Initial values for variables used in the loop
        ShortTableRow = 2
        StockVol = 0
        InitialPrice = Empty
        FinalPrice = Empty

        'Define Headers
        ws.Range("J1").Value = "Ticker"
        ws.Range("K1").Value = "Yearly Change"
        ws.Range("L1").Value = "Percent Change"
        ws.Range("M1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
    
    
        'Start the Loop to go over all the Ticker Values.
        For i = 2 To LRow
        
            'If the next cell is different than the current one
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then
            
                'Set the Ticker Name in the Summarized Table
                TickerName = ws.Cells(i, 1).Value
                ws.Cells(ShortTableRow, 10) = TickerName
            
                'Add the last value to StockValue, post it and reset it
                StockVol = StockVol + ws.Cells(i, 7).Value
                ws.Cells(ShortTableRow, 13).Value = StockVol
            
                'Find out the close rate for the last day of the year
                FinalPrice = ws.Cells(i, 6).Value
            
                'Calculate the Yearly Change and the Yearly % Change
                YearlyChange = FinalPrice - InitialPrice
                
                If InitialPrice = 0 Then
                    YearlyPercChange = 0
                Else
                    YearlyPercChange = (FinalPrice / InitialPrice) - 1
                End If
            
                'Print the value and the conditional format for the Yeary Change
                ws.Cells(ShortTableRow, 11).Value = YearlyChange
                    If YearlyChange < 0 Then
                        ws.Cells(ShortTableRow, 11).Interior.ColorIndex = 3
                    Else
                        ws.Cells(ShortTableRow, 11).Interior.ColorIndex = 10
                    End If
                
                'Print the value and format for the Yeary % Change
                ws.Cells(ShortTableRow, 12).Value = YearlyPercChange
                ws.Cells(ShortTableRow, 12).NumberFormat = "0.00%"
            
                'Increment the Row number to be filled in the Summarized Table
                ShortTableRow = ShortTableRow + 1
                         
                'Reset the Prices for the next Ticker
                InitialPrice = Empty
                FinalPrice = Empty
                StockVol = 0
            
            'If the next cell is = to the one we currently are
            Else
                'Keep Adding the Stock Colume
                StockVol = StockVol + ws.Cells(i, 7).Value
                
                    'Temporary variable for the Open price of the stock
                    If InitialPrice = Empty Then
                        InitialPrice = ws.Cells(i, 3).Value
                    End If
            End If
        Next i
    
    '----------X---------------X---------------X------------------
        'Challenge
    
        'Set Initial values for variables used in the following loops
        GreatPercInc = 0
        GreatPercDec = 1
        GreatVol = 0
    
        'Loop for the Greatest % Increase & Decrease
        For Each Cell In ws.Range("L2:L" & ShortTableRow)
            If Cell > GreatPercInc And Cell < GreatPercDec Then
                GreatPercInc = Cell
                GreatPercIncRow = Cell.Row
                GreatPercDec = Cell
                GreatPercDecRow = Cell.Row
            ElseIf Cell > GreatPercInc Then
                GreatPercInc = Cell
                GreatPercIncRow = Cell.Row
            ElseIf Cell < GreatPercDec Then
                GreatPercDec = Cell
                GreatPercDecRow = Cell.Row
            End If
        Next
        
        'Loop for the Greatest Total Volume
        For Each VolCell In ws.Range("M2:M" & ShortTableRow)
            If VolCell > GreatVol Then
                GreatVol = VolCell
                GreatVolRow = VolCell.Row
            End If
        Next
    
        'Print the values for the Ticker the respective value and the format
        ws.Range("P2") = ws.Cells(GreatPercIncRow, 10)
        ws.Range("Q2") = GreatPercInc
        ws.Range("Q2").NumberFormat = "0.00%"
        
        ws.Range("P3") = ws.Cells(GreatPercDecRow, 10)
        ws.Range("Q3") = GreatPercDec
        ws.Range("Q3").NumberFormat = "0.00%"
    
        ws.Range("P4") = ws.Cells(GreatVolRow, 10)
        ws.Range("Q4") = GreatVol
    
        'Format all columns to Autofit
        ws.Columns("J:Q").AutoFit
         
    Next
End Sub