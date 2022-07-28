Sub stocks():
    
    Dim outputrow As Integer
    Dim year_open As Double
    Dim year_close As Double
    Dim year_change As Double
    Dim stock_volume As Double
'------------------------------------------
For Each ws In ThisWorkbook.Worksheets
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    outputrow = 2
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    For i = 2 To lastrow
    
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            'input ticker
            ws.Cells(outputrow, 9).Value = ws.Cells(i, 1).Value
            'yearly change
            year_open = ws.Cells(2, 3).Value
            year_close = ws.Cells(i, 6).Value
            year_change = year_close - year_open
            ws.Cells(outputrow, 10).Value = year_change
            'percent change
            ws.Cells(outputrow, 11).Value = (year_change / year_open) * 100
            'total volume
            stock_volume = stock_volume + ws.Cells(i, 7).Value
            ws.Cells(outputrow, 12).Value = stock_volume
            year_open = ws.Cells(2 + i, 3).Value
            outputrow = outputrow + 1
            
        End If
        If ws.Cells(i, 10).Value > 0 Then
            ws.Cells(i, 10).Interior.Color = vbGreen
        ElseIf ws.Cells(i, 10).Value <= 0 Then
            ws.Cells(i, 10).Interior.Color = vbRed
        End If
    Next i
Next ws
End Sub
