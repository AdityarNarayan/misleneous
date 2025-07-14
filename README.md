
Sub Extract_Upto_First_UnwantedWord()

    Dim ws        As Worksheet
    Dim lastRow   As Long
    Dim r         As Long
    Dim txt       As String
    Dim parts()   As String
    Dim outTxt    As String
    Dim i         As Long
    
    Dim reHasLetter As Object   'contains at least one A–Z
    Dim reDigitsOnly As Object  'all digits, no letters
    Dim reNonAlnum As Object    'any non-alphanumeric character
    
    Set ws = ThisWorkbook.Worksheets(1)          'adjust if needed
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Set reHasLetter = CreateObject("VBScript.RegExp")
    reHasLetter.Pattern = "[A-Za-z]"
    
    Set reDigitsOnly = CreateObject("VBScript.RegExp")
    reDigitsOnly.Pattern = "^\d+$"
    
    Set reNonAlnum = CreateObject("VBScript.RegExp")
    reNonAlnum.Pattern = "[^A-Za-z0-9]"
    
    For r = 2 To lastRow                  'assumes header in row 1
        txt = Trim(ws.Cells(r, "A").Value)
        If Len(txt) > 0 Then
            parts = Split(txt, " ")
            outTxt = ""
            
            For i = 0 To UBound(parts)
                
                '1. word with any non-alphanumeric character → STOP
                If reNonAlnum.Test(parts(i)) Then Exit For
                
                '2. word contains at least one letter → KEEP
                If reHasLetter.Test(parts(i)) Then
                    outTxt = outTxt & parts(i) & " "
                    GoTo NextWord
                End If
                
                '3. word is digits-only
                If reDigitsOnly.Test(parts(i)) Then
                    '   Keep only if next word exists AND has a letter
                    If i < UBound(parts) And reHasLetter.Test(parts(i + 1)) Then
                        outTxt = outTxt & parts(i) & " "
                        GoTo NextWord
                    Else
                        Exit For        'numeric word not followed by letters → STOP
                    End If
                End If
                
    NextWord:
            Next i
            
            ws.Cells(r, "B").Value = Trim(outTxt)       'cleaned text
            ws.Cells(r, "C").Value = UCase(Trim(outTxt))'optional uppercase
        End If
    Next r
    
    MsgBox "Extraction complete!", vbInformation

End Sub
