VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SuperDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Today As Date
Public Tomorrow As Date
Public Yesterday As Date

Private Sub Class_Initialize()
    
    Today = Date
    Tomorrow = Date + 1
    Yesterday = Date - 1
    
End Sub

Public Function AddSubDays(ByVal numDays As Long, Optional ByVal startDate As Date = "1/1/1970") As Date
    
    If startDate = "1/1/1970" Then startDate = Today
    AddSubDays = startDate + numDays

End Function

Public Function AddSubWorkDays(ByVal numDays As Long, Optional ByVal startDate As Date = "1/1/1970") As Date
    
    If startDate = "1/1/1970" Then startDate = Today
    AddSubWorkDays = CDate(WorksheetFunction.WorkDay(startDate, numDays))

End Function

Public Function AddTimeToDate(ByVal d As Date, ByVal t As Date) As Date

    AddTimeToDate = d + t
    
End Function

Public Function ConvertDay(ByVal d As String) As Byte

    Dim i As Byte
    Dim s As String: s = d
    s = Trim$(s)
    s = UCase(s)
    
    Select Case s
        Case "SUN"
            i = 1
        Case "MON"
            i = 2
        Case "TUE"
            i = 3
        Case "WED"
            i = 4
        Case "THU"
            i = 5
        Case "FRI"
            i = 7
        Case "SAT"
            i = 7
    End Select
    
    ConvertDay = i
            
End Function

Public Function NextNameDate(ByVal iDay As Byte) As Date

    Dim i As Byte
    For i = 1 To 7
        If weekDay(Date + i) = iDay Then
            NextNameDate = CDate(Date + i)
            Exit For
        End If
    Next i

End Function
