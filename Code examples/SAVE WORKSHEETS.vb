Public Sub SaveFiles()

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    ThisWorkbook.Sheets(3).Activate
    Dim strFileName As String
    
    Call NewWorkBook("SCF Split")
    Call ActivateBook("SCF Split")
    strFileName = ActiveWorkbook.Name
    ThisWorkbook.Activate
    Sheets(2).Select
    Sheets(2).Copy After:=Workbooks(strFileName).Sheets(1)
    Sheets(1).Delete
    Workbooks(strFileName).Close SaveChanges:=True
    
    Call NewWorkBook("Version Summary")
    Call ActivateBook("Version Summary")
    strFileName = ActiveWorkbook.Name
    ThisWorkbook.Activate
    Sheets("Insert").Select
    Sheets("Insert").Copy After:=Workbooks(strFileName).Sheets(1)
    Sheets(1).Delete
    Workbooks(strFileName).Close SaveChanges:=True
    
    Call MsgBox("SAVE files complete", vbInformation)
    
    ThisWorkbook.Sheets(1).Activate
    Range("H2").Select
    
End Sub