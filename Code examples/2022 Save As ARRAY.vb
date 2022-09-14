Public Sub Save()

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    Dim NewBook As Workbook
    Dim strFileName As String
    ThisWorkbook.Sheets("MAIN").Activate
    
    strFileName = Application.GetSaveAsFilename(fileFilter:="Excel Files (*.xls), *.xls,Excel Files (*.xlsx), *.xlsx")
    If strFileName = "False" Then End
    
    Set NewBook = Workbooks.Add
    NewBook.SaveAs filename:=strFileName
    Call ActivateBook(RemoveBefore(strFileName, "\"))
    Dim strNF As String: strNF = ActiveWorkbook.Name
    
    ThisWorkbook.Activate
    Sheets(Array("ESRP", "ATT MAIL", "MULTI MAIL", "MULTI WRAP", "CANFGNOTHER", "CSV", "COMBO")).Select
    Sheets(Array("ESRP", "ATT MAIL", "MULTI MAIL", "MULTI WRAP", "CANFGNOTHER", "CSV", "COMBO")).Copy _
    After:=Workbooks(strNF).Sheets(1)
    Call ActivateBook(strNF)
    Sheets(1).Delete
    
    Workbooks(strNF).Close SaveChanges:=True
    Call MsgBox("SAVE files complete", vbInformation)
    
    ThisWorkbook.Sheets(1).Activate
    
End Sub