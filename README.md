

'==============================================
' Main entry point – loops through column A,
' writes cleaned result to column B
'==============================================
Sub ExtractEntityName()
    
    Dim ws As Worksheet
    Dim lastRow As Long, r As Long
    Dim original As String, cleaned As String

    Set ws = ThisWorkbook.Worksheets(1)           '<< adjust as needed
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    For r = 2 To lastRow                          'assumes header in row 1
        original = Trim(CStr(ws.Cells(r, "A").Value))
        cleaned = EntityFromString(original)
        ws.Cells(r, "B").Value = cleaned
    Next r
    
    MsgBox "Extraction complete", vbInformation
    
End Sub

'==============================================
' Returns the “entity” portion of a raw string
' Implements the rules:
'   R1  Keep word with at least one letter (ALPHA)
'   R2  Keep SHORTNUM (<=4 digits) *only* if next word is ALPHA
'   R3  Otherwise stop before current word
'==============================================
Private Function EntityFromString(ByVal txt As String) As String
    
    Dim parts() As String
    Dim i As Long, ub As Long
    Dim result As String
    
    '----- Early exit
    If Len(txt) = 0 Then
        EntityFromString = ""
        Exit Function
    End If
    
    parts = Split(txt, " ")
    ub = UBound(parts)
    
    For i = LBound(parts) To ub
        
        Dim w As String
        w = parts(i)
        If Len(w) = 0 Then GoTo NextWord   'skip empty tokens (double spaces)
        
        Select Case WordCategory(w)
        
            Case "ALPHA"
                result = result & w & " "
            
            Case "SHORTNUM"
                If i < ub Then
                    If WordCategory(parts(i + 1)) = "ALPHA" Then
                        result = result & w & " "
                    Else
                        Exit For  'numeric not followed by ALPHA → stop
                    End If
                Else
                    Exit For      'numeric is last word → stop
                End If
                
            Case Else   'LONGNUM or NONALNUM
                Exit For          'stop before current word
                
        End Select
NextWord:
    Next i
    
    EntityFromString = Trim(result)
    
End Function

'==============================================
' Categorises a single word
'   Returns:  "ALPHA", "SHORTNUM", "LONGNUM", "NONALNUM"
'==============================================
Private Function WordCategory(ByVal w As String) As String
    
    Dim reLetter As Object, reDigits As Object, reNonAlnum As Object
    
    Set reLetter = CreateObject("VBScript.RegExp")
    reLetter.Pattern = "[A-Za-z]"
    
    Set reDigits = CreateObject("VBScript.RegExp")
    reDigits.Pattern = "^\d+$"
    
    Set reNonAlnum = CreateObject("VBScript.RegExp")
    reNonAlnum.Pattern = "[^A-Za-z0-9]"
    
    If reNonAlnum.Test(w) Then
        WordCategory = "NONALNUM"
    ElseIf reDigits.Test(w) Then
        If Len(w) <= 4 Then
            WordCategory = "SHORTNUM"
        Else
            WordCategory = "LONGNUM"
        End If
    ElseIf reLetter.Test(w) Then
        WordCategory = "ALPHA"
    Else
        WordCategory = "NONALNUM"   'fallback
    End If
    
End Function
