Sub DeleteRowsEmptyFirstColumn()
    Dim LastRow As Long
    
    With ActiveSheet
        LastRow = .Range("A" & .Rows.Count).End(xlUp).Row
        
        For i = LastRow To 1 Step -1
            If .Cells(i, 1).Value = "" Then
                .Rows(i).Delete
            End If
        Next i
    End With
End Sub
