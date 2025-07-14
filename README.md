Sub ExtractCounterpartyNames()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim fullText As String
    Dim namePart As String
    Dim arrParts() As String

    Sub ExtractCounterpartyNames()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim fullText As String
    Dim namePart As String
    Dim arrParts() As String
    Dim j As Long

    Set ws = ThisWorkbook.Sheets(1) ' Adjust if needed
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    For i = 2 To lastRow ' Assuming header in row 1
        fullText = Trim(ws.Cells(i, "A").Value)
        If fullText <> "" Then
            arrParts = Split(fullText, " ")
            If UBound(arrParts) > 0 Then
                namePart = ""
                For j = 0 To UBound(arrParts) - 1
                    namePart = namePart & arrParts(j) & " "
                Next j
                namePart = Trim(namePart)
                ws.Cells(i, "B").Value = namePart
                ws.Cells(i, "C").Value = UCase(Trim(namePart))
            Else
                ws.Cells(i, "B").Value = fullText
                ws.Cells(i, "C").Value = UCase(Trim(fullText))
            End If
        End If
    Next i

    MsgBox "Extraction complete!", vbInformation
End Sub
