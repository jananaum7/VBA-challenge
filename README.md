# VBA-challenge
    For Each ws In Worksheets
    RowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row 'always use to count the nuymber of rows\
    j = 2
    For i = 2 To RowCount
    ticker_symbol = ws.Cells(i, 1).Value
    
    Next_ticker_symbol = ws.Cells(i + 1, 1).Value
    
    If ticker_symbol <> Next_ticker_symbol Then
        Total = Total + ws.Cells(i, 7).Value
        ws.Cells(j, 9).Value = ticker_symbol
        ws.Cells(j, 12).Value = Total
        Total = 0
        j = j + 1
    Else
        Total = Total + ws.Cells(i, 7).Value
    End If
    Next i
      Next ws
    'Was given to me by my tutor Mohamed. 
End Sub
