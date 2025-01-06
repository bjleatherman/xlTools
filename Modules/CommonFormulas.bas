Sub IndirectAdjSubtraction()
    With ActiveCell
        .Value = "=Round(Indirect(Address(Row(),Column()-1))-Indirect(Address(Row(),Column()+1)),1)"
    End With
End Sub
Sub ReverseOrientation()
    With ActiveCell
        .Value = "=IF(-A2+0.5>0.49930556,(-A2+0.5)-0.5,(-A2+0.5))"
    End With
End Sub

