
**' 
Sub CopyDataFromFiles()
    Dim wsList As Worksheet, wsDest As Worksheet, wsLog As Worksheet
    Dim lastRow As Long, destRow As Long, logRow As Long
    Dim i As Long
    Dim filePath As String, customerName As String, status As String
    Dim srcWb As Workbook, srcWs As Worksheet
    Dim srcLastRow As Long, srcLastCol As Long
    Dim dataRange As Range, cell As Range
    Dim tempArr As Variant, j As Long

    Set wsList = ThisWorkbook.Sheets(1) 'Assume file paths are in Sheet1
    Set wsDest = ThisWorkbook.Sheets("Sheet1")
    Set wsLog = ThisWorkbook.Sheets("Sheet2")

    lastRow = wsList.Cells(wsList.Rows.Count, "W").End(xlUp).Row
    destRow = wsDest.Cells(wsDest.Rows.Count, 1).End(xlUp).Row + 1
    logRow = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row + 1

    For i = 4 To lastRow
        filePath = wsList.Cells(i, "W").Value
        status = "Success"
        On Error GoTo ErrHandler

        Set srcWb = Workbooks.Open(filePath, ReadOnly:=True)
        Set srcWs = srcWb.Sheets("Source ECS Data")
        customerName = srcWs.Range("C2").Value 'Assuming customer name is in C2

        srcLastRow = srcWs.Cells(srcWs.Rows.Count, 1).End(xlUp).Row
        srcLastCol = srcWs.Cells(2, srcWs.Columns.Count).End(xlToLeft).Column

        Set dataRange = srcWs.Range(srcWs.Cells(2, 1), srcWs.Cells(srcLastRow, srcLastCol))
        tempArr = dataRange.Value

        ' Add customer name as new column to the right of data
        For j = 1 To UBound(tempArr, 1)
            wsDest.Cells(destRow, 1).Resize(1, srcLastCol).Value = Application.Index(tempArr, j, 0)
            wsDest.Cells(destRow, srcLastCol + 1).Value = customerName
            destRow = destRow + 1
        Next j

        srcWb.Close False

        GoTo LogStatus

ErrHandler:
        status = "Error: " & Err.Description
        On Error GoTo 0
        If Not srcWb Is Nothing Then srcWb.Close False

LogStatus:
        wsLog.Cells(logRow, 1).Value = customerName
        wsLog.Cells(logRow, 2).Value = filePath
        wsLog.Cells(logRow, 3).Value = status
        logRow = logRow + 1
        customerName = ""
        status = ""
        Set srcWb = Nothing
        Set srcWs = Nothing
    Next i

    MsgBox "Process completed."
End Sub
**
