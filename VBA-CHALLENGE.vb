Sub alphaTest()
Dim ws As Worksheet
Set ws = ActiveSheet
Dim a_endrow As Long
Dim i_row As Long
Dim volume As LongLong
Dim year_change As Double
Dim percent_change As Double
Dim open_price As Double
Dim close_price As Double
Dim k_endrow As Long
Dim max_percent As Double
Dim ticker_range, percent_range, volume_range As Range
Dim sheets_count As Integer
Dim k As Integer
Dim i As Long
Dim j As Long

    'loop for all worksheets
    For Each ws In Worksheets
            ws.Activate
            Debug.Print ws.Name
                        'headers
                        ws.Columns("A:R").AutoFit
                        ws.Range("i1").Value = "Ticker"
                        ws.Range("j1").Value = "Year Change"
                        ws.Range("k1").Value = "Percent Change"
                        ws.Range("l1").Value = "Total Stock Volume"
                        
                        a_endrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
                        i_row = 2
                    'loop for calculations
                    For i = 2 To a_endrow
                        If ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value And ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                            'create unique sticker
                            ws.Cells(i_row, 9).Value = ws.Cells(i, 1).Value
                            'store open price for the year
                            open_price = ws.Cells(i, 3).Value
                        ElseIf ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
                            'calculate volume
                            volume = volume + ws.Cells(i, 7).Value
                        ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                            'store close price
                            close_price = ws.Cells(i, 6).Value
                            'calcute and display year change
                            year_change = close_price - open_price
                            ws.Cells(i_row, 10).Value = year_change
                            'calculate and display percent change
                            percent_change = (year_change / open_price)
                            ws.Cells(i_row, 11).Value = percent_change
                            'add last day volume to total volume and display total volume
                            volume = volume + ws.Cells(i, 7).Value
                            ws.Cells(i_row, 12).Value = volume
                            'move to next row of unique values
                            i_row = i_row + 1
                            'reset volume
                            volume = 0
                        End If
                        Next i
                    
                    
                    'set headers
                    ws.Range("p1").Value = "Ticker"
                    ws.Range("q1").Value = "Value"
                    'define ranges for i, k and l for xlookup function
                    k_endrow = ws.Cells(Rows.Count, 10).End(xlUp).Row
                    Set ticker_range = ws.Range("i2", "i" & k_endrow)
                    Set percent_range = ws.Range("k2", "k" & k_endrow)
                    Set volume_range = ws.Range("l2", "l" & k_endrow)
                    'greatest percentage increase
                    ws.Range("o2").Value = "Greatest % Increase"
                    ws.Range("q2").Value = ws.Application.WorksheetFunction.Max(ws.Range("k2", "k" & k_endrow))
                    ws.Range("q2").NumberFormat = "0.00%"
                    'use xlookup to return tickers for highest percentage increase
                    ws.Range("p2").Value = ws.Application.WorksheetFunction.XLookup(ws.Range("q2"), percent_range, ticker_range)
                    'greatest percentage decrease
                    ws.Range("o3").Value = "Greatest % Decrease"
                    ws.Range("q3").Value = ws.Application.WorksheetFunction.Min(ws.Range("k2", "k" & k_endrow))
                    ws.Range("p3").Value = ws.Application.WorksheetFunction.XLookup(ws.Range("q3"), percent_range, ticker_range)
                    ws.Range("q3").NumberFormat = "0.00%"
                    'greatest total stock volume
                    ws.Range("o4").Value = "Greatest Total Volume"
                    ws.Range("q4").Value = ws.Application.WorksheetFunction.Max(ws.Range("l2", "l" & k_endrow))
                    ws.Range("p4").Value = ws.Application.WorksheetFunction.XLookup(ws.Range("q4"), volume_range, ticker_range)
                    'format year change to currency
                    ws.Range("j2", "j" & k_endrow).NumberFormat = "0.00"
                    'conditional formating for year change aka red if negative green if positive
                    For k = 2 To k_endrow
                        If ws.Cells(k, 10).Value > 0 Then
                            ws.Cells(k, 10).Interior.ColorIndex = 4
                        Else
                            ws.Cells(k, 10).Interior.ColorIndex = 3
                        End If
                    Next k
                    'format percent change column to percent and round to second decimal
                    ws.Range("k2", "k" & k_endrow).NumberFormat = "0.00%"
                    ws.Columns("A:R").AutoFit
                    MsgBox (ws.Name + " is completed.")
    Next
    
    
    MsgBox ("Done")
End Sub
