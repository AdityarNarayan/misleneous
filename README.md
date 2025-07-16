
Header
 Header
 Header

 Header
 
 
 
 
If Not invalid Then
    If Len(w) >= 2 Then
        ' Only treat as ALPHANUM_ID if first char is a letter and all others are digits, and ONLY if length = 2 (1 letter + digits)
        If Mid$(w, 1, 1) Like "[A-Za-z]" And _
           Mid$(w, 2) Like String(Len(w) - 1, "#") And _
           Len(w) = 1 + Len(Mid$(w, 2)) Then
            ' Make sure there is only one letter at the start
            Category = 5 'ALPHANUM_ID
            Exit Function
        End If
    End If
End If
