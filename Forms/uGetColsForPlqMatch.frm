VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uGetColsForPlqMatch 
   Caption         =   "Select Columns for PLQ Match"
   ClientHeight    =   6240
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10125
   OleObjectBlob   =   "uGetColsForPlqMatch.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uGetColsForPlqMatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbPlqSegLen_Change()

End Sub

Private Sub Label1_Click()

End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim colHeaders As Variant: colHeaders = GetDataFromRowNum(ws, 1)
    'Comboboxes in the UserForm
    Dim cmbs As Variant: cmbs = Array(cmbPrevTjl, cmbPrevWt, cmbPlqSegLen, cmbPlqWt, cmbPlqGrade, cmdPlqType)
    
    ' Clear ComboBoxes
    For i = LBound(cmbs) To UBound(cmbs)
        cmbs(i).Clear
    Next i
    
    i = 0
    For i = LBound(colHeaders, 2) To UBound(colHeaders, 2)
        For j = LBound(cmbs) To UBound(cmbs)
            cmbs(j).AddItem colHeaders(1, i)
        Next j
    Next i
    
End Sub
