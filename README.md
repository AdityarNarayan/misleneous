
Sub CopyDistinctValuesFromMultipleFiles()
    Dim folderPath As String
    Dim fileName As String
    Dim wbSource As Workbook
    Dim wsSource As Worksheet
    Dim wsDest As Worksheet
    Dim lastRow As Long, destLastRow As Long
    Dim dict As Object
    Dim cell As Range
    Dim colNum As Integer
    Dim tempArr As Variant
    Dim i As Long
    Dim maxRowsPerFile As Long
    Dim maxTotalRows As Long

    folderPath = "C:\Your\Folder\Path\" ' <-- Change this to your folder path
    colNum = 1 ' <-- Change this to your column number (1 = A, 2 = B, etc.)
    maxRowsPerFile = 20000
    maxTotalRows = 500000

    Set wsDest = ThisWorkbook.Sheets("Sheet1")
    destLastRow = wsDest.Cells(wsDest.Rows.Count, 1).End(xlUp).Row

    If destLastRow = 1 And wsDest.Cells(1, 1).Value = "" Then destLastRow = 0

    fileName = Dir(folderPath & "*.xls*")
    Do While fileName <> ""
        Set wbSource = Workbooks.Open(folderPath & fileName, ReadOnly:=True)
        On Error Resume Next
        Set wsSource = wbSource.Sheets("Source ECS Data")
        On Error GoTo 0
        If Not wsSource Is Nothing Then
            lastRow = wsSource.Cells(wsSource.Rows.Count, colNum).End(xlUp).Row
            tempArr = wsSource.Range(wsSource.Cells(2, colNum), wsSource.Cells(lastRow, colNum)).Value 'Assume header in row 1

            Set dict = CreateObject("Scripting.Dictionary")
            For i = 1 To UBound(tempArr, 1)
                If Not dict.Exists(tempArr(i, 1)) Then
                    dict.Add tempArr(i, 1), Nothing
                End If
            Next i

            If dict.Count < maxRowsPerFile Then
                ' Check if adding these rows will exceed maxTotalRows
                If destLastRow + dict.Count > maxTotalRows Then
                    MsgBox "Destination sheet reached 500,000 rows. Macro will terminate."
                    wbSource.Close False
                    Exit Sub
                End If

                ' Copy distinct values to destination
                i = 1
                For Each key In dict.Keys
                    wsDest.Cells(destLastRow + i, 1).Value = key
                    i = i + 1
                Next key
                destLastRow = destLastRow + dict.Count
            End If
        End If
        wbSource.Close False
        Set wsSource = Nothing
        fileName = Dir
    Loop

    MsgBox "Process completed."
End Sub
