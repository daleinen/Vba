Public Sub SaveWorkBook()
'saving out macro sheets to a new workbook

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    Call SaveMasterRoute
    
    ThisWorkbook.Sheets(3).Activate
    Dim strFileName As String
    Dim strJob As String
    strJob = RemoveBefore(CStr(Range("A1").Value), "=")
    
    Call NewWorkBook("TMS SS Import " & UCase(strJob))
    Call ActivateBook("TMS SS")
    strFileName = ActiveWorkbook.Name
    ThisWorkbook.Activate
    Sheets("Saddle St.").Select
    Sheets("Saddle St.").Copy After:=Workbooks(strFileName).Sheets(1)
    Sheets(1).Delete
    Sheets(1).Name = strJob
    Call DeleteSlots
    Workbooks(strFileName).Close SaveChanges:=True
    
    Call NewWorkBook("TMS Insert Import " & UCase(strJob))
    Call ActivateBook("TMS Insert")
    strFileName = ActiveWorkbook.Name
    ThisWorkbook.Activate
    Sheets("Insert").Select
    Sheets("Insert").Copy After:=Workbooks(strFileName).Sheets(1)
    Sheets(1).Delete
    Sheets(1).Name = strJob
    Call DeleteSlots
    Workbooks(strFileName).Close SaveChanges:=True
    
    Call NewWorkBook("TMS Flat Cut Import " & UCase(strJob))
    Call ActivateBook("TMS Flat")
    strFileName = ActiveWorkbook.Name
    ThisWorkbook.Activate
    Sheets("Flat C.").Select
    Sheets("Flat C.").Copy After:=Workbooks(strFileName).Sheets(1)
    Sheets(1).Delete
    Sheets(1).Name = strJob
    Workbooks(strFileName).Close SaveChanges:=True
    
    Call MsgBox("SAVE files complete", vbInformation)
    
    ThisWorkbook.Sheets(1).Activate
    Range("E4").Select
    
End Sub

Public Sub SaveWorkBook()
'saving out macro sheets to a new workbook

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Dim NewBook As Workbook
    Dim strFileName As String, strLasCol As String
    strFileName = Application.GetSaveAsFilename(fileFilter:="Excel Files (*.xlsx), *.xls")
    If strFileName = "False" Then End
    Set NewBook = Workbooks.Add
    NewBook.SaveAs fileName:=strFileName
    Call ActivateBook(RemoveBefore(strFileName, "\"))
    Dim strNF As String: strNF = ActiveWorkbook.Name
    
    ThisWorkbook.Sheets(3).Activate
    Sheets("Distribution").Select
    Cells.Select
    Selection.Copy
    Workbooks(strNF).Activate
    Cells.Select
    ActiveSheet.Paste
    Sheets(1).Name = "Distribution"
    Workbooks(strNF).Close SaveChanges:=True
    
    Call MsgBox("SAVE files complete", vbInformation)
    
    ThisWorkbook.Sheets(1).Activate
    Range("E3").Select
    
End Sub