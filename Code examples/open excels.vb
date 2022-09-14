Sub OpenInput()
'subroutine to open client PO in input folder

'decalring and setting variables, setting conditions
Dim myPath As String
Dim myFile As String
myPath = "C:\Users\" & Environ("username") & "\Desktop\Bass Pro\input\"
myFile = Dir(myPath)
Application.DisplayAlerts = False
Application.AskToUpdateLinks = False

'use dir to check if file exists
If Dir(myPath) = "" Then
    'if file does not exist display message
    MsgBox "Please make sure PO is in the input folder before pressing OK"
    sFile = Dir(sPath)
End If

'opening and activating PO
Workbooks.Open myPath & myFile
Workbooks(myFile).Activate
    
End Sub

Public Sub OpenInput2()
'subroutine to open the client print plan

Dim myFolder As String
Dim myFile As String
myFolder = "C:\Users\" & Environ("username") & "\Desktop\Bass Pro\input"
myFile = Dir(myFolder & "\*.xlsx")
Application.DisplayAlerts = False
Application.AskToUpdateLinks = False

Do While myFile <> ""
    Workbooks.Open fileName:=myFolder & "\" & myFile
    Workbooks(myFile).Activate
    myFile = Dir
Loop

End Sub