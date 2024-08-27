Sub GetImportFeaturesCsv()
    CopySheetAndSave "JobNum Import Features", "Import Features"
End Sub

Sub GetImportJointNoCsv()
    CopySheetAndSave "JobNum Import Joint No", "Import Joint No"
End Sub

Sub GetImportWeldsCsv()
    CopySheetAndSave "JobNum Import Welds", "Import Welds"
End Sub

Sub GetImportGpsCsv()
    CopySheetAndSave "JobNum GPS Interp", "GPS Interp"
End Sub

Sub GetAddCommentsCsv()
    CopySheetAndSave "JobNum Add Comments", "Add Comments"
End Sub

Sub GetAddFeaturesCsv()
    CopySheetAndSave "JobNum Add Features", "Add Features"
End Sub

Sub GetAddWeldAttributesCsv()
    CopySheetAndSave "JobNum Weld Attributes", "Weld Attributes"
End Sub

Sub CopySheetAndSave(sheetName As String, defaultName As String)
    
    Dim sourceWorkbook As Workbook
    Dim newWorkbook As Workbook
    Dim sourceSheet As Worksheet
    Dim savePath As Variant

    ' Open PERSONAL.XLSB
    On Error Resume Next
    Set sourceWorkbook = Workbooks("PERSONAL.XLSB")
    If sourceWorkbook Is Nothing Then
        Set sourceWorkbook = Workbooks.Open(Environ("APPDATA") & "\Microsoft\Excel\XLSTART\PERSONAL.XLSB")
    End If
    On Error GoTo 0

    ' Specify the sheet name you want to copy
    Set sourceSheet = sourceWorkbook.Sheets(sheetName)

    ' Create a new workbook
    Set newWorkbook = Workbooks.Add

    ' Copy the sheet to the new workbook
    sourceSheet.Copy Before:=newWorkbook.Sheets(1)
    
    ' Delete the empty, default sheet
    Application.DisplayAlerts = False
    newWorkbook.Sheets("Sheet1").Delete
    Application.DisplayAlerts = True

    ' Ask user to specify the filename and path to save the workbook
    savePath = Application.GetSaveAsFilename( _
        InitialFileName:=defaultName & ".csv", _
        FileFilter:="CSV (Comma delimited) (*.csv), *.csv", _
        Title:="Save As")

    ' Check if the user has cancelled the dialog
    If savePath <> False Then
        ' Save the new workbook as a CSV file
        newWorkbook.SaveAs Filename:=savePath, FileFormat:=xlCSV, CreateBackup:=False
    End If

End Sub



