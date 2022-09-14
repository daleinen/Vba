    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    Dim NewBook As Workbook
    Dim strFileName As String
    
    strFileName = Application.GetSaveAsFilename(fileFilter:="Excel Files (*.xlsx), *.xls")
    If strFileName = "False" Then End
    
    Set NewBook = Workbooks.Add
    NewBook.SaveAs filename:=strFileName
    Call ActivateBook(RemoveBefore(strFileName, "\"))
    Dim strNF As String: strNF = ActiveWorkbook.name
    
    ThisWorkbook.Sheets(2).Activate
    Cells.Select
    Selection.Copy
    Workbooks(strNF).Activate
    Cells.Select
    ActiveSheet.Paste

    Workbooks(strNF).Close SaveChanges:=True
    
    Call MsgBox("SAVE files complete", vbInformation)
    
    ThisWorkbook.Sheets(1).Activate
    Range("H8").Select