Sub ExtractCounterpartyNames()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim fullText As String
    Dim namePart As String
    Dim arrParts() As String

 
Sub ExtractName_RemoveTrailingNumbers()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim fullText As String
    Dim namePart As String
    Dim regex As Object

    Set ws = ThisWorkbook.Sheets(1) ' Adjust as needed
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "\s*\d+\s*$"   ' Matches any number at the end, with optional spaces before/after
    regex.IgnoreCase = True
    regex.Global = False

    For i = 2 To lastRow ' Assuming header in row 1
        fullText = ws.Cells(i, "A").Value
        If fullText <> "" Then
            namePart = regex.Replace(fullText, "") ' Remove trailing number and spaces
            ws.Cells(i, "B").Value = Trim(namePart)
            ws.Cells(i, "C").Value = UCase(Trim(namePart))
        End If
    Next i

    MsgBox "Extraction complete!", vbInformation
End Sub
