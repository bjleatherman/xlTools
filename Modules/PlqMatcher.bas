Sub MatchPlq()

    Application.ScreenUpdating = False

    Dim ws As Worksheet: Set ws = Application.ActiveSheet
    Dim lastRow As Long: lastRow = GetLastRowFromColNum(ws, 1)
    Dim cTjlRow As Long: cTjlRow = 2
    'Dim fisrtInsertCol As Long: firstInsertCol = 4
    'Dim lastInsertCol As Long: lastInsertCol = 10
    Dim targetObj As TjlTarget
    Dim insertRange As Range
    Dim plqTjlStart As Long
    Dim formData As cPlqMatchData
    
    
    'Dim frm As Object
    'Set frm = New UserForm
    'frm.Show
   
    ' show form
    uGetColsForPlqMatch.Show
    
    Set formData = uGetColsForPlqMatch.Data
    
    ' get data from form
    Dim pTjlColIndex As Long: pTjlColIndex = formData.PrevTjlColIndex
    Dim pWtColIndex As Long: pWtColIndex = formData.PrevWtColIndex
    Dim plqTjlColIndex As Long: plqTjlColIndex = formData.PlqSegLenColIndex
    Dim plqWtColIndex As Long: plqWtColIndex = formData.plqWtColIndex
    Dim PlqGradeColIndex As Long: PlqGradeColIndex = formData.PlqGradeColIndex
    Dim PlqTypeColIndex As Long: PlqTypeColIndex = formData.PlqTypeColIndex
    Dim firstInsertCol As Long: firstInsertCol = formData.StartColIndex
    Dim lastInsertCol As Long: lastInsertCol = formData.EndColIndex
    
    Debug.Print TypeName(pWtColIndex)
    Debug.Print pWtColIndex
    
    pWtColIndex = 2
    
    Debug.Print TypeName(pWtColIndex)
    Debug.Print pWtColIndex
    
    Unload uGetColsForPlqMatch
    
    '''''Get Data'''''
    With ws
        
        ' Previous WT
        Dim pWt As Variant: pWt = GetDataFromColNum(ws, pWtColIndex)
        ' Previous TJL
        Dim pTjl As Variant: pTjl = GetDataFromColNum(ws, pTjlColIndex)
        ' PLQ WT
        Dim plqWt As Variant: plqWt = GetDataFromColNum(ws, plqWtColIndex)
        ' PLQ Grade
        Dim plqGrade As Variant: plqGrade = GetDataFromColNum(ws, PlqGradeColIndex)
        ' PLQ Type
        Dim plqType As Variant: plqType = GetDataFromColNum(ws, PlqTypeColIndex)
        ' PLQ TJL
        Dim plqTjl As Variant: plqTjl = GetDataFromColNum(ws, plqTjlColIndex)
    
    '''''Loop through data'''''
    
        ' Find start of data in PLQ TJL
        For i = LBound(plqTjl) + 1 To UBound(plqTjl)
            If plqTjl(i, 1) <> "" Then
                plqTjlStart = i
                cTjlRow = i
                Exit For
            End If
        Next i
    
        ' Find and add insertion points
        For i = plqTjlStart To UBound(plqTjl)
        
            ' If the line is blank, skip it
            If plqTjl(i, 1) = "" Then
                cTjlRow = cTjlRow + 1
            ' Insert the rows
            Else
                ' Find prev Tjl location where the sum matches the seg length
                Set targetObj = GetClosestJoint(CDbl(plqTjl(i, 1)), cTjlRow, pTjl, lastRow)
                
                ' Get Target Row
                targetRow = targetObj.TjlTargetIndex
                
                contextTargetRow = GetTargetRowWithWtContext(pTjl, pWt, cTjlRow, i, plqTjl, plqWt, lastRow)
                
                ' Determine if a new line needs to be inserted
                If targetRow = cTjlRow Then
                    ' No need to insert row; just shift index
                    cTjlRow = cTjlRow + 1
                Else
                    Set insertRange = .Range( _
                            .Cells(cTjlRow + 1, firstInsertCol), _
                            .Cells(targetObj.TjlTargetIndex, lastInsertCol))
                        
                    insertRange.Select
                    insertRange.Insert Shift:=xlShiftDown
                    cTjlRow = targetObj.TjlTargetIndex + 1
                End If
            End If

            Debug.Print i
            DoEvents
        Next i
    End With

End Sub

Function GetTargetRowWithWtContext(pTjl As Variant, pWt As Variant, cRow As Long, plqIndex As Variant, plqTjl As Variant, plqWt As Variant, lastRow As Long) As Long

    Dim minimumCellSearch As Integer: minimumCellSearch = 5
    Dim odoDriftComp As Double: odoDriftComp = 0.01
    Dim cSegLen As Double: cSegLen = plqTjl(plqIndex, 1)
    Dim odoDriftAllowance As Double: odoDriftAllowance = cSegLen * odoDriftComp
    
    Dim forCount As Integer
    Dim revCount As Integer
    
    forCount = SumTjlByDirectionTillMaxSumReached(0, pTjl, cRow, cSegLen, odoDriftAllowance, lastRow)
    revCount = SumTjlByDirectionTillMaxSumReached(1, pTjl, cRow, cSegLen, odoDriftAllowance, lastRow)
    
    'Check if the counts are longer vs the minimum counts
    
    'Check that that both counts stay in bounds

End Function

Function SumTjlByDirectionTillMaxSumReached(direction As Integer, pTjl As Variant, cRow As Long, cSegLen As Double, tolerance As Double, lastRow As Long) As Long

    Dim flag As Boolean: flag = False
    Dim index As Long: index = cRow
    Dim count As Integer: count = 0
    Dim sum As Double: sum = 0
    
    ' If going forwards, start the sum at the current index
    If direction = 0 Then sum = pTjl(cRow, 1)
    
    Do While flag = False
        If index < lastRow And index > 1 Then
            sum = sum + pTjl(index, 1)
            count = count + 1
            If direction = 0 Then
                index = index + 1
            Else
                index = index - 1
            End If
            If Abs(cSegLen - sum) > tolerance Then
                flag = True
            End If
        Else
            flag = True
        End If
    Loop
    
    SumTjlByDirectionTillMaxSumReached = count
    
End Function

Function GetClosestJoint(targetLen As Double, pTjlIndex As Long, pTjl As Variant, lastRow As Long) As TjlTarget

    Dim targetObj As New TjlTarget
    
    Dim underIndex As Long
    Dim underSum As Double
    
    Dim overIndex As Long
    Dim overSum As Double
    
    Dim sum As Double: sum = 0
    Dim i As Integer: i = pTjlIndex
    
    Dim currentTjlVal As Double: currentTjlVal = pTjl(pTjlIndex, 1)
    
    ' If the current tjl is larger than current seg len
    If currentTjlVal > targetLen Then
        targetObj.TjlTargetIndex = pTjlIndex
        targetObj.TjlTargetSum = pTjl(pTjlIndex, 1)
    Else
        ' Sum pTjl until we find the area it changes over
        Do While sum < targetLen
        
            ' If we are at the bottom of the sheet, exit the loop
            If i >= lastRow Then Exit Do
        
            ' Mark the indexes
            underIndex = i - 1
            overIndex = i
            
            ' Adjust the sums
            underSum = sum
            sum = sum + pTjl(i, 1)
            overSum = sum
            
            ' Increment the indexer
            i = i + 1
        Loop
    
        ' Return the index with the closest fum
        If Abs(targetLen - underSum) < Abs(targetLen - overSum) Then
            targetObj.TjlTargetIndex = underIndex
            targetObj.TjlTargetSum = underSum
        Else
            targetObj.TjlTargetIndex = overIndex
            targetObj.TjlTargetSum = overSum
        End If
        
    End If
            
    Set GetClosestJoint = targetObj
    Set targetObj = Nothing

End Function

