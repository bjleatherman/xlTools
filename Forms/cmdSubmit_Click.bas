Private Function cmdSubmit_Click() As cPlqMatchData
    
    Dim data As cPlqMatchData
    
    prevTjlIndex = cmbPrevTjl.ListIndex + 1
    prevWtIndex = cmbPrevWt.ListIndex + 1
    plqSegLenIndex = cmbPlqSegLen.ListIndex + 1
    plqWtIndex = cmbPlqWt.ListIndex + 1
    plqGradeIndex = cmbPlqGrade.ListIndex + 1
    plqTypeIndex = cmdPlqType.ListIndex + 1
    
    startCol = txtStartCol.Value
    endCol = txtEndCol.Value
    
    'Debug.Print "tjl: " & prevTjlIndex
    'Debug.Print "wt: " & prevWtIndex
    'Debug.Print "seg len: " & plqSegLenIndex
    'Debug.Print "plq wt: " & plqWtIndex
    'Debug.Print "grade: " & plqGradeIndex
    'Debug.Print "type: " & plqTypeIndex
    'Debug.Print "start: " & startCol
    'Debug.Print "end: " & endCol
    
End Function
