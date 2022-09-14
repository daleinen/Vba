    
    'function to get last row of data
    Dim r As Range, lastRow As Long
    For Each r In Worksheets("Domestic").UsedRange
        If Not IsEmpty(r) Then
            lastRow = r.Row
        End If
    Next
    'MsgBox lastRow

--------------------------------------------------------------

lastRow = Cells(Rows.Count, 1).End(xlUp).Row
