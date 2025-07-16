
Header
 Header
 Header

 Header
 
 
 
 
If Not invalid Then
    ' Only match exactly one letter followed by digits
    If Len(w) >= 2 Then
        If Mid$(w, 1, 1) Like "[A-Za-z]" And _
           Mid$(w, 2) Like String(Len(w) - 1, "#") And _
           Len(Trim$(Mid$(w, 1, 1))) = 1 Then
            Category = 5 'ALPHANUM_ID
            Exit Function
        End If
    End If
End 
