
Private Function IsAlphanumeric(ByVal w As String) As Boolean
    Dim hasLetter As Boolean, hasDigit As Boolean
    Dim i As Long, ch As String
    For i = 1 To Len(w)
        ch = Mid$(w, i, 1)
        If ch Like "[A-Za-z]" Then hasLetter = True
        If ch Like "[0-9]" Then hasDigit = True
        If hasLetter And hasDigit Then
            IsAlphanumeric = True
            Exit Function
        End If
    Next i
    IsAlphanumeric = False
End Function
