Sub StockTickerChallenge()
    Dim ticker As String
    Dim volume As Double
    Dim year_change As Double
    Dim per_change As Double
    Dim yr_open As Double
    Dim yr_close As Double
    Dim ticker2 As String
    Dim ticker3 As String
    Dim ticker_row As Double
    Dim percent_row As Long
    Dim great_volume As Double
    Dim large_increase As Double
    Dim large_decrease As Double






    For Each ws In Worksheets
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        ws.Cells(2, 14).Value = "Greatest % increase"
        ws.Cells(3, 14).Value = "Greatest % decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"

        ticker_row = 0
        ticker_row = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        volume = 0
        yr_open = ws.Cells(2, 3).Value
        table_row = 2

        For i = 2 To ticker_row
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ticker = ws.Cells(i, 1).Value
                yr_close = ws.Cells(i, 6).Value
                volume = volume + ws.Cells(i, 7).Value
                year_change = yr_close - yr_open
                If yr_open = 0 Then
                    per_change = 0
                Else
                    per_change = year_change / yr_open
                End If
                per_change2 = Format(per_change, "Percent")
                ws.Range("I" & table_row).Value = ticker
                ws.Range("J" & table_row).Value = year_change
                If ws.Range("J" & table_row).Value > 0 Then
                    ws.Cells(table_row, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(table_row, 10).Interior.ColorIndex = 3
                End If
                ws.Range("K" & table_row).Value = per_change2
                ws.Range("L" & table_row).Value = volume
                table_row = table_row + 1
                yr_open = ws.Cells(i + 1, 3).Value
                volume = 0

            Else
                volume = volume + ws.Cells(i, 7).Value
            End If
        Next i

 '======================================================

        large_increase = 0
        large_decrease = 0
        percent_row = ws.Cells(ws.Rows.Count, "K").End(xlUp).Row
        For j = 2 To percent_row
            If ws.Cells(j, 11).Value > large_increase Then
                large_increase = ws.Cells(j, 11).Value
                ticker2 = ws.Cells(j, 9).Value
                large_increase2 = Format(large_increase, "Percent")
            ws.Cells(2, 16).Value = large_increase2
            ws.Cells(2, 15).Value = ticker2
            ElseIf ws.Cells(j, 11).Value < large_decrease Then
                large_decrease = ws.Cells(j, 11).Value
                ticker3 = ws.Cells(j, 9).Value
                large_decrease2 = Format(large_decrease, "Percent")
            ws.Cells(3, 16).Value = large_decrease2
            ws.Cells(3, 15).Value = ticker3
            End If
        Next j

  '======================================================

        great_volume = 0
        For k = 2 To percent_row
            If ws.Cells(k, 12).Value > great_volume Then
                great_volume = ws.Cells(k, 12).Value
                great_ticker = ws.Cells(k, 9).Value
                ws.Cells(4, 16).Value = great_volume
                ws.Cells(4, 15).Value = great_ticker
            End If
        Next k
    Next ws

End Sub
