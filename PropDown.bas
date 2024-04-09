Sub PropDown()

    Application.ScreenUpdating = False

    Dim ws As Worksheet: Set ws = Application.ActiveSheet
    Dim r As Range: Set r = Selection
    Dim data As Variant: data = r.Value
    Dim LastRow As Integer: LastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    
    Dim wt As Variant
    Dim grade As String
    Dim pipeType As String
    
    Dim wtIndex As Integer: wtIndex = LBound(data, 2)
    Dim gradeIndex As Integer: gradeIndex = LBound(data, 2) + 1
    Dim pipeTypeIndex As Integer: pipeTypeIndex = LBound(data, 2) + 2
    
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    
    For i = LBound(data, 1) To UBound(data, 1)
        If i > LastRow Then Exit For
        If data(i, 1) <> "" Then
            wt = data(i, wtIndex)
            grade = data(i, gradeIndex)
            pipeType = data(i, pipeTypeIndex)
        Else
            data(i, wtIndex) = wt
            data(i, gradeIndex) = grade
            data(i, pipeTypeIndex) = pipeType
        End If
    Next i
    
    Selection.Value = data
    
    r.Copy
    Selection.PasteSpecial Paste:=xlPasteFormats
    
    Application.CutCopyMode = False
    
    ws.Cells(1, 1).Select
    
End Sub
