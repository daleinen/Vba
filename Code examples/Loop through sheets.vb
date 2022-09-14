For Each Sheet In Workbooks("AD.xlsx").Worksheets
    MsgBox Sheet.Name
Next

For i = 1 To 8
    Sheets(i).Activate
    MsgBox Sheets(i).Name
Next