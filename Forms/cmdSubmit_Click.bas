Private Sub cmdSubmit_Click()
    
    Dim Data As cPlqMatchData
    
    prevTjlIndex = cmbPrevTjl.ListIndex + 1
    prevWtIndex = cmbPrevWt.ListIndex + 1
    plqSegLenIndex = cmbPlqSegLen.ListIndex + 1
    plqWtIndex = cmbPlqWt.ListIndex + 1
    plqGradeIndex = cmbPlqGrade.ListIndex + 1
    plqTypeIndex = cmdPlqType.ListIndex + 1
    
    Dim formStartCol As String: formStartCol = CStr(txtStartCol.Value)
    Dim formEndCol As String: formEndCol = CStr(txtEndCol.Value)
    
    ' Validate Column letters
    isStartColValid = IsValidColLet(formStartCol)
    isEndColValid = IsValidColLet(formEndCol)
    
    If isStartColValid = False Or isEndColValid = False Then
        MsgBox "Check PLQ start and end column letters"
        Exit Sub
    End If
    
    StartColIndex = ColLetToNumber(formStartCol)
    EndColIndex = ColLetToNumber(formEndCol)
    
    ' Validate Indexes
    If prevTjlIndex = 0 Or prevWtIndex = 0 Or plqSegLenIndex = 0 _
        Or plqWtIndex = 0 Or plqGradeIndex = 0 Or plqTypeIndex = 0 Then
        
        MsgBox "All Dropdowns must have a selection"
        Exit Sub
    End If
    
    pData.PrevTjlColIndex = prevTjlIndex
    pData.PrevWtColIndex = prevWtIndex
    pData.PlqSegLenColIndex = plqSegLenIndex
    pData.PlqWtColIndex = plqWtIndex
    pData.PlqGradeColIndex = plqGradeIndex
    pData.PlqTypeColIndex = plqTypeIndex
    pData.StartColIndex = StartColIndex
    pData.EndColIndex = EndColIndex
    
    Me.Hide
    
End Sub
