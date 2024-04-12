Private Sub cmdSubmit_Click()
    
    Dim Data As cPlqMatchData
    
    prevTjlIndex = cmbPrevTjl.ListIndex + 1
    prevWtIndex = cmbPrevWt.ListIndex + 1
    plqSegLenIndex = cmbPlqSegLen.ListIndex + 1
    plqWtIndex = cmbPlqWt.ListIndex + 1
    plqGradeIndex = cmbPlqGrade.ListIndex + 1
    plqTypeIndex = cmdPlqType.ListIndex + 1
    
    startCol = txtStartCol.Value
    endCol = txtEndCol.Value
    
    ' Validate Column letters
    isStartColValid = IsValidColLet(startCol)
    isEndColValid = IsValidColLet(endCol)
    
    If isStartColValid = False Or isEndColValid Then
        MsgBox "Check PLQ start and end column letters"
        Exit Sub
    End If
    
    startColIndex = ColLetToNumber(startCol)
    endColIndex = ColLetToNumber(endCol)
    
    ' Validate Indexes
    If prevTjlIndex = 0 Or prevWtIndex = 0 Or plqSegLenIndex = 0 _
        Or plqWtIndex = 0 Or plqGradeIndex = 0 Or plqTypeIndex = 0 Then
        
        MsgBox "All Dropdowns must have a selection"
        Exit Sub
    End If
    
    pData.pPrevTjl = prevTjlIndex
    pData.pPrevWt = prevWtIndex
    pData.pPlqSegLen = plqSegLenIndex
    pData.pPlqWt = plqWtIndex
    pData.pPlqGrade = plqGradeIndex
    pData.pPlqType = plqTypeIndex
    pData.pStartCol = startColIndex
    pData.pEndCol = isEndColValid

    Me.Hide

End Sub
