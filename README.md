If k > 0 Then
    ReDim resultArr(0 To k - 1)
    Dim i As Long
    For i = 0 To k - 1
        resultArr(i) = keep(i)
    Next i
    EntityChunk = Join(resultArr, " ")
Else
    EntityChunk = ""
End If
