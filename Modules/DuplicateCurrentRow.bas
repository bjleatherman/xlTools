Sub DuplicateActiveRow()

    Dim currentCell As Range: Set currentCell = ActiveCell
    ActiveCell.Rows("1:1").EntireRow.Select
    Selection.Copy
    Selection.Insert Shift:=xlDown
    Application.CutCopyMode = False
    
    currentCell.Select
    
End Sub
