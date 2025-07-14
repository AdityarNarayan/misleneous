
'--- replace the whole Category function with this version ‑--
'  1 = ALPHA, 2 = SHORTNUM, 3 = LONGNUM, 4 = NONALNUM, 5 = ALPHANUM_ID
Private Function Category(ByVal w As String) As Long
    
    Dim i As Long, ch As String
    Dim hasLetter As Boolean, invalid As Boolean
    
    For i = 1 To Len(w)
        ch = Mid$(w, i, 1)
        Select Case ch
            Case "A" To "Z", "a" To "z": hasLetter = True
            Case "0" To "9"
            Case Else
                If InStr(ALLOWED_SYM, ch) = 0 Then
                    invalid = True: Exit For
                End If
        End Select
    Next i
    
    '--- NEW: detect pattern <letter><digits> --------------
    If Not invalid Then
        If hasLetter And Len(w) >= 2 _
           And Mid$(w, 1, 1) Like "[A-Za-z]" _
           And Mid$(w, 2) Like String(Len(w) - 1, "#") Then
                Category = 5                       'ALPHANUM_ID
                Exit Function
        End If
    End If
    '-------------------------------------------------------
    
    If invalid Then
        Category = 4                               'NONALNUM
    ElseIf Not hasLetter Then
        If Len(w) <= SHORT_NUM_MAXLEN Then
            Category = 2                           'SHORTNUM
        Else
            Category = 3                           'LONGNUM
        End If
    Else
        Category = 1                               'ALPHA
    End If
End Function


