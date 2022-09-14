Attribute VB_Name = "MiscMod"
Option Explicit

Public Sub EndFormat(ByVal strJob As String)
'end format for bind

    ThisWorkbook.Sheets(2).Activate
    Range("A1").Value = "BindJob=" & UCase(strJob)
    ThisWorkbook.Sheets(3).Activate
    Range("A1").Value = "BindJob=" & UCase(strJob)
    ThisWorkbook.Sheets(4).Activate
    Range("A1").Value = "BindJob=" & UCase(strJob)
    ThisWorkbook.Sheets(5).Activate
    'Range("A1").Value = "BindJob=" & UCase(strJob)
    Columns("C:G").Select
    Selection.ColumnWidth = 27.75
    
    ThisWorkbook.Sheets(3).Activate
    Range("A2").Activate
    ActiveCell.FormulaR1C1 = ""
    Range("A2").Select
 
End Sub

Public Function GetFirstRowSS(ByVal lngStart As Long) As Long
'grabing the first row on ss

    Dim r As Range
    For Each r In Range("F" & lngStart + 1 & ":F" & lngStart + 20)
        If Len(r.Value) > 0 Then
            GetFirstRowSS = r.row - 1
            Exit For
        End If
    Next

End Function

Public Sub SaveWorkBook()
'saving out macro sheets to a new workbook

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    Call ShowProgessBar
    Call FractionComplete(0.1)
    
    Call RunSlotDupes
    Call SaveMasterRoute
    
    ThisWorkbook.Sheets(3).Activate
    Dim strFileName As String
    Dim strJob As String
    strJob = RemoveBefore(CStr(Range("A1").Value), "=")
    Call FractionComplete(0.2)
    
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
    Call FractionComplete(0.4)
    
    Call NewWorkBook("TMS Insert Import " & UCase(strJob))
    Call ActivateBook("TMS Insert")
    strFileName = ActiveWorkbook.Name
    ThisWorkbook.Activate
    Sheets("Insert").Select
    Sheets("Insert").Copy After:=Workbooks(strFileName).Sheets(1)
    Sheets(1).Delete
    Sheets(1).Name = strJob
    'Call DeleteSlots
    Workbooks(strFileName).Close SaveChanges:=True
    Call FractionComplete(0.6)
    
    Call NewWorkBook("TMS Flat Cut Import " & UCase(strJob))
    Call ActivateBook("TMS Flat")
    strFileName = ActiveWorkbook.Name
    ThisWorkbook.Activate
    Sheets("Flat C.").Select
    Sheets("Flat C.").Copy After:=Workbooks(strFileName).Sheets(1)
    Sheets(1).Delete
    Sheets(1).Name = strJob
    Workbooks(strFileName).Close SaveChanges:=True
    Call FractionComplete(0.8)
    
    Call NewWorkBook("TMS Offline Mail Import " & UCase(strJob))
    Call ActivateBook("TMS Offline")
    strFileName = ActiveWorkbook.Name
    ThisWorkbook.Activate
    Sheets("Offline M.").Select
    Sheets("Offline M.").Copy After:=Workbooks(strFileName).Sheets(1)
    Sheets(1).Delete
    Sheets(1).Name = strJob
    Workbooks(strFileName).Close SaveChanges:=True
    Call FractionComplete(1)
    Call WasteSeconds(0.75)
    Call EndProgessBar
    
    Call MsgBox("SAVE files complete", vbInformation)
    
    ThisWorkbook.Sheets(1).Activate
    Range("E4").Select
    
End Sub


Public Sub SaveMasterRoute()

    Dim strFileName As String, strMaster As String
    Call ActivateBook("Carrier")
    strFileName = ActiveWorkbook.Name
           
    Call NewWorkBook("TMS Master Carrier Route ")
    Call ActivateBook("Master")
    strMaster = ActiveWorkbook.Name
    
    Workbooks(strFileName).Activate
    Sheets(2).Select
    Sheets(2).Copy After:=Workbooks(strMaster).Sheets(1)
    Sheets(1).Delete
    
    Workbooks(strMaster).Close SaveChanges:=True

End Sub

Public Sub Clear()
'subroutine to clear certain portions of macro
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    ThisWorkbook.Activate
    Sheets(2).Select
    Rows("5:1000").Select
    Selection.ClearContents
    Sheets(3).Select
    Rows("5:1000").Select
    Selection.ClearContents
    Sheets(4).Select
    Rows("5:1000").Select
    Selection.ClearContents
    Sheets(5).Select
    Rows("2:1000").Select
    Selection.ClearContents
    Sheets(1).Activate
    Range("E4").Select
    
    Call CloseWorkbooks

End Sub

Public Sub ShowProgessBar()

    ufProgress.LabelProgress.Width = 0
    ufProgress.Show
    
End Sub

Public Sub FractionComplete(pctdone As Single)

    With ufProgress
        .LabelCaption.Caption = pctdone * 100 & "% Complete"
        .LabelProgress.Width = pctdone * (.FrameProgress.Width)
    End With
    DoEvents
    
End Sub

Public Sub EndProgessBar()
    Unload ufProgress
End Sub

Public Function RemoveSpecial(x As String) As String
' removing certain chars and performing some string formatting

    Dim y As String
    y = x
    
    y = ReplaceString(y, "-", " ")
    y = ReplaceString(y, "/", " ")
    y = ReplaceString(y, "   ", " ")
    y = ReplaceString(y, "  ", " ")
    y = Left(y, 15)
    y = Trim$(y)
    
    RemoveSpecial = y
    
End Function

Public Function ReverseName(x As String) As String
'used to reverse and manipulate strings when needed to compare naming

    Dim a As String
    Dim b As String
    
    If InStr(x, " - ") > 0 Then
        a = RemoveBefore(x, " - ")
        b = RemoveAfter(x, " - ")
        a = Left(a, 8)
        a = RemoveSpecial(a)
        b = RemoveSpecial(b)
        b = Left(b, 15 - Len(a) - 1)
        ReverseName = Trim$(a & " " & b)
    Else
        ReverseName = Trim$(RemoveSpecial(Left(x, 15)))
    End If

End Function

Public Sub DeleteSlots()
'deleting unused slots

    Dim i As Long
    Dim s As String
    Dim cell As Range
    Dim d As Boolean: d = True
    
    For i = 3 To 12
        For Each cell In Range(Cells(5, i), Cells(GetLastRow(1), i))
            If Len(cell.Value) > 0 Then
                d = False
                Exit For
            End If
        Next cell
        If d = True Then s = s & GetColumnLet(i) & ":" & GetColumnLet(i) & ","
        d = True
    Next i
    
    s = Left(s, Len(s) - 1)
    Range(s).Select
    Selection.Delete Shift:=xlToLeft
    Range("A3").Select
    
End Sub

Public Sub SortDemosIN()
    
    ThisWorkbook.Sheets(3).Activate
    Range("A4:B" & GetLastRow(2)).Select
    ActiveWorkbook.Worksheets("Insert").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Insert").Sort.SortFields.Add2 Key:=Range("A5:A" & GetLastRow(2)) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Insert").Sort.SortFields.Add2 Key:=Range("B5:B" & GetLastRow(2)) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortTextAsNumbers
    With ActiveWorkbook.Worksheets("Insert").Sort
        .SetRange Range("A4:B" & GetLastRow(2))
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
End Sub

Public Sub RunNewDemosIN()
Attribute RunNewDemosIN.VB_ProcData.VB_Invoke_Func = "d\n14"

    ThisWorkbook.Sheets(3).Activate
    Dim i As Long
    Dim cell As Range
    Dim temp As Long
    
    For i = 1 To 100
        For Each cell In Range("B5:B" & GetLastRow(2))
             If CLng(cell.Value) = temp Then
                cell.Value = cell.Value + 100
            Else
                temp = CLng(cell.Value)
            End If
             
        Next cell
    Next i
    
End Sub

