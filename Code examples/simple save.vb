Public Sub Save()

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    Dim NewBook As Workbook
    Dim strFileName As String
    ThisWorkbook.Sheets("MAIN").Activate
    strFileName = Application.GetSaveAsFilename(fileFilter:="Excel Files (*.xlsx), *.xls")
    If strFileName = "False" Then End
    
    Set NewBook = Workbooks.Add
    NewBook.SaveAs fileName:=strFileName
    Call ActivateBook(RemoveBefore(strFileName, "\"))
    Dim strNF As String: strNF = ActiveWorkbook.Name
    
    ThisWorkbook.Sheets("UPLOAD").Activate
    Cells.Select
    Selection.Copy
    Workbooks(strNF).Activate
    Cells.Select
    ActiveSheet.Paste
    
    Workbooks(strNF).Close SaveChanges:=True
    
    Call MsgBox("SAVE files complete", vbInformation)
    Call Clear
    
    ThisWorkbook.Sheets(1).Activate
    Range("H8").Select
    
End Sub