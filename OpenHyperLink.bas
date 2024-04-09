Sub OpenFile()

    FolderColHeader = "File Folder"
    FileColHeader = "Filename"
    
    FolderColIndex = 0
    FileColIndex = 0
    
    currCell = Split(ActiveCell.Address, "$")
    currCol = currCell(1)
    currRow = currCell(2)
    
    Set ws = ActiveSheet
    
    With ws
    
        lastCol = .Cells(1, .Columns.count).End(xlToLeft).Column
        colHeaders = .Range(.Cells(1, 1), .Cells(1, lastCol)).Value
        For i = LBound(colHeaders, 2) To UBound(colHeaders, 2)
            If colHeaders(1, i) = FolderColHeader Then FolderColIndex = i
            If colHeaders(1, i) = FileColHeader Then FileColIndex = i
        Next i
        
        fp = .Cells(currRow, FolderColIndex).Value & .Cells(currRow, FileColIndex).Value
        If Dir(fp) <> "" And fp <> "" Then
            OpenNonExcelFileUsingShellAndRefocusExcel (fp)
        End If
    End With

End Sub

Sub OpenNonExcelFileUsingShellAndRefocusExcel(fp As String)
    
    ' Temporarily change the Excel application's caption to a unique title
    Dim originalCaption As String
    originalCaption = Application.Caption
    Dim uniqueCaption As String
    uniqueCaption = "ExcelApp" & Timer  ' Use the Timer function to ensure uniqueness
    Application.Caption = uniqueCaption
    
    ' Attempt to open the file with the default program associated with its file type
    Shell "explorer.exe """ & fp & """", vbNormalFocus
    
    ' Use a brief pause to allow the shell command to execute
    Application.Wait Now + TimeValue("00:00:02")
    
    ' Attempt to refocus on Excel by activating the window with the unique caption
    On Error Resume Next  ' In case the AppActivate fails
    AppActivate uniqueCaption
    On Error GoTo 0  ' Turn back on regular error handling
    
    ' Restore the original caption
    Application.Caption = originalCaption
End Sub
