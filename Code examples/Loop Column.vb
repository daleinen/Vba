For Each cell In Workbooks("DRIVER FILE").Worksheets(1).Range("B36:B213")
    If Not IsEmpty(cell) Then
        IowaLast = cell.Row
    End If
Next