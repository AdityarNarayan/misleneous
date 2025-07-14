
Option Explicit

'========== Tunables ==========
Private Const SHORT_NUM_MAXLEN As Long = 4                '≤4 digits
Private Const ALLOWED_SYM      As String = "/-&'.()"      'symbols permitted in a name
'  - put the dash (-) either first or last in the string to avoid range behaviour

'========== Public macro ==========
Sub ExtractEntityName_Optimised2()

    Dim ws As Worksheet
    Dim lastRow As Long, r As Long
    Dim rawTxt As String
    
    Set ws = ThisWorkbook.Worksheets(1)                   'adjust as needed
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    For r = 2 To lastRow                                  'assumes header in row 1
        rawTxt = Trim$(CStr(ws.Cells(r, "A").Value))
        ws.Cells(r, "B").Value = EntityChunk(rawTxt)
    Next r
    
    MsgBox "Entity extraction complete", vbInformation
End Sub

'========== Core extractor ==========
Private Function EntityChunk(ByVal txt As String) As String
    
    If Len(txt) = 0 Then Exit Function
    
    Dim tok() As String: tok = Split(txt, " ")
    Dim keep() As String: ReDim keep(UBound(tok))
    Dim i As Long, k As Long, ub As Long: ub = UBound(tok)
    
    For i = 0 To ub
        
        Dim w As String: w = tok(i)
        If Len(w) = 0 Then GoTo nxt                     'skip empty tokens
        
        Select Case Category(w)
            Case 1      'ALPHA
                keep(k) = w: k = k + 1
            
            Case 2      'SHORTNUM (<=4 digits)
                If i < ub And Category(tok(i + 1)) = 1 Then
                    keep(k) = w: k = k + 1
                Else
                    Exit For                            'noise starts
                End If
            
            Case Else    'LONGNUM or NONALNUM
                Exit For
        End Select
nxt:
    Next i
    
    If k > 0 Then
        ReDim Preserve keep(k - 1)
        EntityChunk = Join(keep, " ")
    End If
End Function

'========== Word categoriser ==========
'   1 = ALPHA      (≥1 letter, all chars valid)
'   2 = SHORTNUM   (digits-only, length ≤ SHORT_NUM_MAXLEN)
'   3 = LONGNUM    (digits-only, length > SHORT_NUM_MAXLEN)
'   4 = NONALNUM   (contains an invalid char)
'=====================================
Private Function Category(ByVal w As String) As Long
    
    Dim i As Long, ch As String
    Dim hasLetter As Boolean, invalid As Boolean
    
    For i = 1 To Len(w)
        ch = Mid$(w, i, 1)
        
        Select Case ch
            Case "A" To "Z", "a" To "z"
                hasLetter = True                       'valid, mark letter
            Case "0" To "9"
                'valid digit, nothing to do
            Case Else
                'Allowed symbol?
                If InStr(ALLOWED_SYM, ch) = 0 Then
                    invalid = True
                    Exit For
                End If
        End Select
    Next i
    
    If invalid Then
        Category = 4                                   'NONALNUM
    ElseIf Not hasLetter Then                          'digits-only
        If Len(w) <= SHORT_NUM_MAXLEN Then
            Category = 2                               'SHORTNUM
        Else
            Category = 3                               'LONGNUM
        End If
    Else
        Category = 1                                   'ALPHA
    End If
End Function
