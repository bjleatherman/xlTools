Sub GetShelbyGwMatch()
    CopySheetFromPersonalWb "Shelby GW Comparison"
End Sub

Sub GetDigSheetTracker()
    CopySheetFromPersonalWb "Dig Options"
    CopySheetFromPersonalWb "Dig Tracker"
End Sub

Sub GetWtReviewer()
    CopySheetFromPersonalWb "JBL Review"
End Sub

Sub CopySheetFromPersonalWb(sheetName As String)
    
    Application.ScreenUpdating = False

    Dim sourceWorkbook As Workbook
    Dim targetWorkbook As Workbook
    Dim sourceSheet As Worksheet

    ' Set the target workbook to the active workbook
    Set targetWorkbook = ActiveWorkbook

    ' Attempt to open PERSONAL.XLSB from the Excel start folder
    On Error Resume Next
    Set sourceWorkbook = Workbooks("PERSONAL.XLSB")
    If sourceWorkbook Is Nothing Then
        Set sourceWorkbook = Workbooks.Open(Environ("APPDATA") & "\Microsoft\Excel\XLSTART\PERSONAL.XLSB")
    End If
    On Error GoTo 0

    ' Check if the sheet exists in the source workbook
    On Error Resume Next
    Set sourceSheet = sourceWorkbook.Sheets(sheetName)
    If Err.Number <> 0 Then
        MsgBox "Sheet '" & sheetName & "' does not exist in PERSONAL.XLSB.", vbExclamation, "Error"
        Exit Sub
    End If
    On Error GoTo 0

    ' Copy the sheet to the active workbook
    sourceSheet.Copy After:=targetWorkbook.Sheets(targetWorkbook.Sheets.count)

    ' Optionally, activate the copied sheet
    targetWorkbook.Sheets(sheetName).Activate

    Application.ScreenUpdating = True

End Sub
