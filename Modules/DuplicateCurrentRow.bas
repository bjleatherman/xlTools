Sub DuplicateActiveRow()
    ActiveCell.Rows("1:1").EntireRow.Select
    Selection.Copy
    Selection.Insert Shift:=xlDown
    Application.CutCopyMode = False
End Sub
