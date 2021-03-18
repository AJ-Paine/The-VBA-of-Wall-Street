Sub Ticker_Summary()
    For Each ws In Worksheets
        'Clear Cells where summary tables will go
            ws.Range("I:Q").Clear
        
        'Set up summary table including table for bonus
            ws.Range("I1").Value = "Ticker"
            ws.Range("J1").Value = "Yearly Change"
            ws.Range("K1").Value = "Percent Change"
            ws.Range("L1").Value = "Total Stock Volume"
            ws.Range("P1").Value = "Ticker"
            ws.Range("Q1").Value = "Value"
            ws.Range("O2").Value = "Greatest % Increase"
            ws.Range("O3").Value = "Greatest % Decrease"
            ws.Range("O4").Value = "Greatest Total Volume"
            
        ' Define variables and set to inital values    
            Dim lastrow As Long
            lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
            Dim OpenValue As Double
            OpenValue = ws.Cells(2, 3).Value
        
        'Set variables for yearly change, percent change, and holding total volume traded
            Dim PercentChange As Double
            Dim YearlyChange As Double
            Dim CloseValue As Double
            Dim TotalVol As Double
            TotalVol = 0
            
        'Find unique ticker values from column A and add to column I, calculate yearly change & % Change in Columns J & K
            For i = 2 To lastrow
                If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                'Add unique tickers to column I    
                    ws.Cells((ws.Cells(Rows.Count, 9).End(xlUp).Row) + 1, 9).Value = ws.Cells(i, 1).Value

                'Determine Close Value, calculate yearly change and add it to column J
                    CloseValue = ws.Cells(i, 6).Value
                    YearlyChange = CloseValue - OpenValue
                    ws.Cells((ws.Cells(Rows.Count, 10).End(xlUp).Row) + 1, 10).Value = YearlyChange

                    'Defines PercentChange to 0 if Open Value is equal to 0.  Gets around DIV/0 error that stops script.
                    'Calculate and add percent change to Column K
                    If OpenValue = 0 Then
                        PercentChange = 0
                        ws.Cells((ws.Cells(Rows.Count, 11).End(xlUp).Row) + 1, 11).Value = PercentChange
                        Else
                        PercentChange = YearlyChange / OpenValue
                        ws.Cells((ws.Cells(Rows.Count, 11).End(xlUp).Row) + 1, 11).Value = PercentChange
                    End If

                    'Set new Open Value
                    OpenValue = ws.Cells(i + 1, 3)

                    'Calculate Total Volume and add to column L
                    TotalVol = TotalVol + ws.Cells(i, 7)
                    ws.Cells((ws.Cells(Rows.Count, 12).End(xlUp).Row) + 1, 12).Value = TotalVol

                    'Reset Total Volume
                    TotalVol = 0
                Else
                    TotalVol = TotalVol + ws.Cells(i, 7)
                End If
            Next i
        
    

        'Format PercentChange as a percent
        ws.Range("K2:K" & ws.Cells(Rows.Count, 11).End(xlUp).Row).NumberFormat = "0.00%"
        
        'Format Yearly Change > 0 to interior color green and Yearly Change < 0 to interior color red
            For j = 2 To ws.Cells(Rows.Count, 9).End(xlUp).Row
                If ws.Cells(j, 10).Value > 0 Then
                    ws.Cells(j, 10).Interior.ColorIndex = 4
                ElseIf ws.Cells(j, 10).Value < 0 Then
                    ws.Cells(j, 10).Interior.ColorIndex = 3
                End If
            Next j
        
        'Bonus
            Dim maxinc As Double
            Dim maxdec As Double
            Dim maxvol As Double
        
            maxinc = 0
            maxdec = 0
            maxvol = 0
        
        'Find Max Increase
            For i = 2 To ws.Cells(Rows.Count, 11).End(xlUp).Row
                If ws.Cells(i, 11).Value > maxinc Then
                   maxinc = ws.Cells(i, 11).Value
                   ws.Range("Q2").Value = maxinc
                   ws.Range("P2").Value = ws.Cells(i, 9).Value
                End If

        'Find Max Decrease
                If ws.Cells(i, 11).Value < maxdec Then
                    maxdec = ws.Cells(i, 11).Value
                    ws.Range("Q3").Value = maxdec
                    ws.Range("P3").Value = ws.Cells(i, 9).Value
                End If
        
        'Find Max Total Volume
                If ws.Cells(i, 12).Value > maxvol Then
                    maxvol = ws.Cells(i, 12).Value
                    ws.Range("Q4").Value = maxvol
                    ws.Range("P4").Value = ws.Cells(i, 9).Value
                End If
        
            Next i

        'Format bonus area and autofit column width for worksheets
           ws.Range("Q2:Q3").NumberFormat = "0.00%"
           ws.Cells.Columns.AutoFit
           
    Next ws
    
End Sub

