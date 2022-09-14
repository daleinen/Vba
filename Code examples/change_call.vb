Private Sub Worksheet_Change(ByVal Target As Range)

    Dim KeyCells As Range
    Set KeyCells = Range("I4:K5")

    If Not Application.Intersect(KeyCells, Range(Target.Address)) Is Nothing Then
            Call MsgBox("Are you sure you would like to change QTY at:  " & Target.Address & "?" & _
            vbCrLf & "Normally these cells have static QTYs", vbExclamation, "Warning!")
    End If

End Sub