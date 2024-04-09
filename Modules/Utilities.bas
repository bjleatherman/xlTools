Function GetLastRowFromColNum(ws As Worksheet, col As Long) As Long
    GetLastRowFromColNum = ws.Cells(ws.Rows.count, col).End(xlUp).Row
End Function

Function GetDataFromColNum(ws As Worksheet, col As Long) As Variant
    With ws
        Dim LastRow As Long: LastRow = GetLastRowFromColNum(ws, col)
        Dim r As Range: Set r = .Range(.Cells(1, col), .Cells(LastRow, col))
        Dim data As Variant: data = r.Value
    End With
    GetDataFromColNum = data
End Function