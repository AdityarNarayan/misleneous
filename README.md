
Sub ExtractName_OnlyWords()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim fullText As String
    Dim namePart As String
    Dim regex As Object
    Dim matches As Object
    Dim j As Long

    Set ws = ThisWorkbook.Sheets(1) ' Adjust as needed
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "\b[A-Za-z]+\b"  ' Matches words with only letters
    regex.IgnoreCase = True
    regex.Global = True

    For i = 2 To lastRow ' Assuming header in row 1
        fullText = ws.Cells(i, "A").Value
        namePart = ""
        If fullText <> "" Then
            Set matches = regex.Execute(fullText)
            For j = 0 To matches.Count - 1
                namePart = namePart & matches(j).Value & " "
            Next j
            ws.Cells(i, "B").Value = Trim(namePart)
            ws.Cells(i, "C").Value = UCase(Trim(namePart))
        End If
    Next i

    MsgBox "Extraction complete!", vbInformation
End Sub
