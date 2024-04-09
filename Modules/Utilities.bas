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
