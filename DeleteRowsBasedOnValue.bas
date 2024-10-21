Sub DeleteRowsBasedOnValue()
    Dim LastRow As Long
    
    With ActiveSheet
        LastRow = .Range("A" & .Rows.Count).End(xlUp).Row
        DeleteValue = InputBox("Rows starting with this value will be deleted", "Value to delete")
        If DeleteValue = "" Then
            Exit Sub
        End If
        
        For i = LastRow To 1 Step -1
            If InStr(1, .Cells(i, 1).Value, DeleteValue, 1) = 1 Then
                .Rows(i).Delete
            End If
        Next i
    End With
End Sub
