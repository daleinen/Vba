
'deleting row if cell not equal to effort number
Dim j As Long
For j = Cells(Rows.Count, "d").End(xlUp).Row To 1 Step -1
	If Cells(j, "d") <> effort Then Cells(j, "d").EntireRow.Delete xlUp
Next j