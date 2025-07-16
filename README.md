
Sub ExtractEntityName_WithStoppers()
    InitStoppers                                'load the arrays once

    Dim ws As Worksheet
    Dim lastRow As Long, r As Long, rawTxt As String
    Dim tokens() As String

    Set ws = ThisWorkbook.Worksheets(1)
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    For r = 2 To lastRow
        rawTxt = Trim$(CStr(ws.Cells(r, "A").Value))
        tokens = Split(rawTxt, " ")
        ws.Cells(r, "B").Value = EntityChunk(tokens, UBound(tokens))
    Next r

    MsgBox "Done with stoppers!", vbInformation
End Sub
