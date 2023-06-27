Sub Tickers()
Dim i As Integer
Dim j As Integer

For i = 2 To 22771
    If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
        Cells(i, 6) = Cells(i, 1).Value
    End If
Next i



End Sub