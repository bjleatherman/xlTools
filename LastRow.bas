Function GetLastRowFromColNum(ws As Worksheet, col As Long) As Long
    GetLastRowFromColNum = ws.Cells(ws.Rows.count, col).End(xlUp).Row
End Function
