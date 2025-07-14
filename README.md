
Sub ExtractName_KeepAlphanumericAndNumbers()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim fullText As String
    Dim namePart As String
    Dim arrParts() As String
    Dim j As Long
    Dim regex As Object

    Set ws = ThisWorkbook.Sheets(1) ' Adjust as needed
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "^[A-Za-z0-9]+$" ' Only alphanumeric words (letters and/or numbers)
    regex.IgnoreCase = True

    For i = 2 To lastRow ' Assuming header in row 1
        fullText = ws.Cells(i, "A").Value
        namePart = ""
        If fullText <> "" Then
            arrParts = Split(fullText, " ")
            For j = 0 To UBound(arrParts)
                If regex.Test(arrParts(j)) Then
                    namePart = namePart & arrParts(j) & " "
                Else
                    Exit For
                End If
            Next j
            ws.Cells(i, "B").Value = Trim(namePart)
            ws.Cells(i, "C").Value = UCase(Trim(namePart))
        End If
    Next i

    MsgBox "Extraction complete!", vbInformation
End Sub
