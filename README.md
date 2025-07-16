
Header
 Header
 Header

 Header
 
 
 
 
 If Not invalid Then
        If hasLetter And Len(w) >= 2 Then
            If Mid$(w, 1, 1) Like "[A-Za-z]" And Mid$(w, 2) Like String(Len(w) - 1, "#") Then
                Category = 5 'ALPHANUM_ID
                Exit Function
            End If
        End If
    End If
