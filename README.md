Sub ExtractCounterpartyNames()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim fullText As String
    Dim namePart As String
    Dim arrParts() As String

    Set ws = ThisWorkbook.Sheets(1) ' Adjust if needed
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    For i = 2 To lastRow ' Assuming header in row 1
        fullText = Trim(ws.Cells(i, "A").Value)
        If fullText <> "" Then
            arrParts = Split(fullText, " ")
            If UBound(arrParts) > 0 Then
                ' Remove the last part (assumed to be reference number)
                namePart = Join(Application.Index(arrParts, 1, 0), " ", 0, UBound(arrParts) - 1)
                ws.Cells(i, "B").Value = namePart
                ' For grouping: standardize (e.g., uppercase, remove extra spaces)
                ws.Cells(i, "C").Value = UCase(Trim(namePart))
            Else
                ws.Cells(i, "B").Value = fullText
                ws.Cells(i, "C").Value = UCase(Trim(fullText))
            End If
        End If
    Next i

    MsgBox "Extraction complete!", vbInformation
End Sub
