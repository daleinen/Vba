    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    Dim NewBook As Workbook
    Dim strFileName As String
    
    strFileName = Application.GetSaveAsFilename(fileFilter:="Excel Files (*.xlsx), *.xls")
    If strFileName = "False" Then End
    
    Set NewBook = Workbooks.Add
    NewBook.SaveAs filename:=strFileName
    Call ActivateBook(RemoveBefore(strFileName, "\"))
    Dim strNF As String: strNF = ActiveWorkbook.Name
    
    ThisWorkbook.Sheets(2).Activate
    Cells.Select
    Selection.Copy
    Workbooks(strNF).Activate
    Cells.Select
    ActiveSheet.Paste
    
    ThisWorkbook.Sheets(3).Activate
    Cells.Select
    Selection.Copy
    Workbooks(strNF).Activate
    Sheets.Add After:=ActiveSheet
    Cells.Select
    ActiveSheet.Paste

    Sheets(1).Name = "Rolls"
    Sheets(2).Name = "Receipts"

    Workbooks(strNF).Close SaveChanges:=True
    
    Call MsgBox("SAVE files complete", vbInformation)
    
    ThisWorkbook.Sheets(1).Activate