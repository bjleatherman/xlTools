Function GetLastRowFromColNum(ws As Worksheet, col As Long) As Long
    GetLastRowFromColNum = ws.Cells(ws.Rows.count, col).End(xlUp).row
End Function

Function GetLastColFromRowNum(ws As Worksheet, row As Long) As Long
    GetLastColFromRowNum = ws.Cells(row, ws.Columns.count).End(xlToLeft).Column
End Function

Function GetDataFromColNum(ws As Worksheet, col As Long) As Variant
    With ws
        Dim lastRow As Long: lastRow = GetLastRowFromColNum(ws, col)
        Dim r As Range: Set r = .Range(.Cells(1, col), .Cells(lastRow, col))
        Dim data As Variant: data = r.Value
    End With
    GetDataFromColNum = data
End Function

Function GetDataFromRowNum(ws As Worksheet, row As Long) As Variant
    With ws
        Dim lastCol As Long: lastCol = GetLastColFromRowNum(ws, row)
        Dim r As Range: Set r = .Range(.Cells(row, 1), .Cells(row, lastCol))
        Dim data As Variant: data = r.Value
    End With
    GetDataFromRowNum = data
End Function

Function ColLetToNumber(letter As String) As Integer
    ColLetToNumber = Range(letter & "1").Column
End Function

Function IndexOfMatchInRangeArray(matchValue As Variant, arr As Variant, dimension As Integer) As Long
    Dim i As Long
    For i = LBound(arr, dimension) To UBound(arr, dimension)
        If dimension = 1 Then
            If arr(i, 1) = matchValue Then
                IndexOfMatch = i
                Exit Function
            End If
        Else
            If arr(1, i) = matchValue Then
                IndexOfMatch = i
                Exit Function
            End If
        End If
    Next i
    ' Return -1 if no match is found
    IndexOfMatch = -1
End Function
