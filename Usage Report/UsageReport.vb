'MODULE:        UsageReport()
'LAST UPDATED:  11/24/21

'OPTIONS:
Option Explicit
Option Compare Text

Public Function SendUsageReport(ByVal inMacroName As String, _
                                Optional ByVal inVersion As String, _
                                Optional ByVal inBody As String) As Boolean
    
    Const TO_ADDRESS As String = "ceautomationteam@quad.com"
    
    Dim appOutlook As Object
    Dim mailItem As Object
    Dim report As String
    
    SendUsageReport = False
    
On Error GoTo Failed
    
    Set appOutlook = CreateObject("Outlook.Application")
    Set mailItem = appOutlook.CreateItem(0)
    
    report = BuildReport(inMacroName, inVersion, inBody)
    
    With mailItem
        .To = TO_ADDRESS
        .Subject = "Automation Usage Report"
        .Body = report
        
        .DeleteAfterSubmit = True
        .Send
        '.Display
        
        SendUsageReport = True
    End With
    
Failed:
    'clean up
    Set mailItem = Nothing
    Set appOutlook = Nothing
    
End Function


Private Function BuildReport(ByVal inMacroName As String, _
                             ByVal inVersion As String, _
                             ByVal inBody As String) As String
    
    Const DIVIDER_TEXT As String = "~~~~~~~~~~~~~~~~~~~~~~"
    
    Const NAME1 As String = "MACRO NAME"
    Const NAME2 As String = "VERSION"
    Const NAME3 As String = "USERNAME"
    Const NAME4 As String = "COMPUTER NAME"
    Const NAME5 As String = "LOCAL TIMESTAMP"
    Const NAME6 As String = "FILE NAME"
    
    Dim val1 As String: val1 = inMacroName
    Dim val2 As String: val2 = inVersion
    Dim val3 As String: val3 = GetUserName()
    Dim val4 As String: val4 = GetComputerName()
    Dim val5 As String: val5 = LCase(Now)
    Dim val6 As String: val6 = ThisWorkbook.Name
    
    Dim header As String
    Dim tags As String
    
    'build header (plain text)
    header = header & NAME1 & ":  " & val1 & vbCr
    header = header & NAME2 & ":  " & val2 & vbCr
    header = header & NAME3 & ":  " & val3 & vbCr
    header = header & NAME4 & ":  " & val4 & vbCr
    header = header & NAME5 & ":  " & val5 & vbCr
    header = header & NAME6 & ":  " & val6 & vbCr
    header = header & DIVIDER_TEXT & vbCr
    
    'build tags (XML)
    tags = DIVIDER_TEXT & vbCr
    tags = tags & AddXMLTags(NAME1, val1) & vbCr
    tags = tags & AddXMLTags(NAME2, val2) & vbCr
    tags = tags & AddXMLTags(NAME3, val3) & vbCr
    tags = tags & AddXMLTags(NAME4, val4) & vbCr
    tags = tags & AddXMLTags(NAME5, val5) & vbCr
    tags = tags & AddXMLTags(NAME6, val6) & vbCr
    
    'combine
    If inBody = "" Then
        BuildReport = header & vbCr & "(No Body Text Included)" & vbCr & vbCr & tags
    Else
        BuildReport = header & vbCr & inBody & vbCr & vbCr & tags
    End If

End Function


Private Function GetUserName() As String
    
On Error GoTo Failed
    
    GetUserName = LCase(Environ$("username"))
    Exit Function
    
Failed:
    GetUserName = "Unknown"
    
End Function


Private Function GetComputerName() As String
    
On Error GoTo Failed
    
    GetComputerName = LCase(Environ$("computername"))
    Exit Function
    
Failed:
    GetComputerName = "Unknown"
    
End Function


Private Function AddXMLTags(ByVal inTag As String, ByVal inVal As String) As String
    
    Dim openTag As String
    Dim closetag As String
    Dim modTag As String
    
    modTag = LCase(inTag)
    modTag = Replace(modTag, " ", "")
    
    openTag = "<" & modTag & ">"
    closetag = "</" & modTag & ">"
    
    AddXMLTags = openTag & inVal & closetag
    
End Function