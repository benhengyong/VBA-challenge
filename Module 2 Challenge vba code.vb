Sub VBA_Challenge()
    Dim ticker(0 To 3000) As String
    ticker(0) = "AAB"
    Dim counter As Integer
    counter = 0

    For i = 2 To 753000
        
        If Cells(i + 1, 1) <> Cells(i, 1) Then
            counter = counter + 1
            ticker(counter) = Cells(i + 1, 1)
            
        End If
    Next i

    For i = 0 To UBound(ticker)
        Cells(i + 2, 9) = ticker(i)
    Next i

End Sub