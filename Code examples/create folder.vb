Private Sub CommandButton1_Click()
    Dim fso As Object
    Dim tdate As Date
    Dim fldrname As String
    Dim fldrpath As String

    tdate = Now()
    Set fso = CreateObject("scripting.filesystemobject")
    fldrname = Format(tdate, "dd-mm-yyyy")
    fldrpath = "C:\Users\username\Desktop\FSO\" & fldrname
    If Not fso.folderexists(fldrpath) Then
        fso.createfolder (fldrpath)
    End If
End Sub