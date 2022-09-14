Public Sub SaveAs()

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    Call GetUserVariables
    Dim NewBook As Workbook
    Dim strFileName As String, strN_O As String, strTTL As String, strDefault As String
    ThisWorkbook.Sheets("MAIN").Activate
    strTTL = CStr(Range("F3").value)
    strN_O = CStr(Range("F5").value)
    strDefault = "Upload " & strN_O & " " & strTTL & " " & Format(dteDTE, "mmddyy")
    
    strFileName = Application.GetSaveAsFilename(InitialFileName:=strDefault, fileFilter:="Excel Files (*.xlsx), *.xls")
    If strFileName = "False" Then End
    
    Set NewBook = Workbooks.Add
    NewBook.SaveAs fileName:=strFileName
    Call ActivateBook(RemoveBefore(strFileName, "\"))
    Dim strNF As String: strNF = ActiveWorkbook.Name
    
    ThisWorkbook.Sheets("TEMP").Activate
    Cells.Select
    Selection.Copy
    Workbooks(strNF).Activate
    Cells.Select
    ActiveSheet.Paste
    Sheets(1).Name = strTTL & " " & Format(dteDTE, "mmddyy") & " " & strN_O
    
    If strN_O = "Newspaper" Then
        ThisWorkbook.Sheets("NEWSPAPER STATS").Activate
        Cells.Select
        Selection.Copy
        Workbooks(strNF).Activate
        Sheets.Add After:=ActiveSheet
        Cells.Select
        ActiveSheet.Paste
        Sheets(2).Name = "NEWSPAPER STATS"
    End If

    Workbooks(strNF).Close SaveChanges:=True
    
    Call MsgBox("SAVE files complete", vbInformation)
    Call Clear
    
    ThisWorkbook.Sheets(1).Activate
    Range("H8").Select

End Sub