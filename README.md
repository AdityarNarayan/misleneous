
Sub Extract_Assignment_Names()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "R").End(xlUp).Row
    
    ' Set headers
    ws.Cells(1, "V").Value = "Offshore Analyst"
    ws.Cells(1, "W").Value = "Offshore Quality Control"
    ws.Cells(1, "X").Value = "Offshore Team Lead"
    
    Dim i As Long
    Dim cellValue As String
    Dim assignments As Variant
    Dim offshoreAnalyst As String, offshoreQC As String, offshoreTL As String
    Dim j As Long, temp As String
    
    For i = 2 To lastRow 'Assumes headers in row 1
        cellValue = ws.Cells(i, "R").Value
        assignments = Split(cellValue, ";")
        
        offshoreAnalyst = ""
        offshoreQC = ""
        offshoreTL = ""
        
        For j = LBound(assignments) To UBound(assignments)
            temp = Trim(assignments(j))
            If InStr(1, temp, "Offshore Analyst", vbTextCompare) > 0 Then
                temp = Replace(temp, "Offshore Analyst:", "", , , vbTextCompare)
                temp = Replace(temp, "Offshore Analyst", "", , , vbTextCompare)
                temp = Trim(temp)
                If offshoreAnalyst = "" Then
                    offshoreAnalyst = temp
                Else
                    offshoreAnalyst = offshoreAnalyst & "; " & temp
                End If
            End If
            If InStr(1, temp, "Offshore Quality Control", vbTextCompare) > 0 Or InStr(1, temp, "Offshore QC", vbTextCompare) > 0 Then
                temp = Replace(temp, "Offshore Quality Control:", "", , , vbTextCompare)
                temp = Replace(temp, "Offshore Quality Control", "", , , vbTextCompare)
                temp = Replace(temp, "Offshore QC:", "", , , vbTextCompare)
                temp = Replace(temp, "Offshore QC", "", , , vbTextCompare)
                temp = Trim(temp)
                If offshoreQC = "" Then
                    offshoreQC = temp
                Else
                    offshoreQC = offshoreQC & "; " & temp
                End If
            End If
            If InStr(1, temp, "Offshore Team Lead", vbTextCompare) > 0 Or InStr(1, temp, "Offshore TL", vbTextCompare) > 0 Then
                temp = Replace(temp, "Offshore Team Lead:", "", , , vbTextCompare)
                temp = Replace(temp, "Offshore Team Lead", "", , , vbTextCompare)
                temp = Replace(temp, "Offshore TL:", "", , , vbTextCompare)
                temp = Replace(temp, "Offshore TL", "", , , vbTextCompare)
                temp = Trim(temp)
                If offshoreTL = "" Then
                    offshoreTL = temp
                Else
                    offshoreTL = offshoreTL & "; " & temp
                End If
            End If
        Next j
        
        If offshoreAnalyst = "" Then offshoreAnalyst = "No Assignment"
        If offshoreQC = "" Then offshoreQC = "No Assignment"
        If offshoreTL = "" Then offshoreTL = "No Assignment"
        
        ws.Cells(i, "V").Value = offshoreAnalyst
        ws.Cells(i, "W").Value = offshoreQC
        ws.Cells(i, "X").Value = offshoreTL
    Next i
End Sub

