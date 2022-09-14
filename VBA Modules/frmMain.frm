VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMain 
   Caption         =   "<form name>"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7725
   OleObjectBlob   =   "frmMain.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const VERSION_NUM As String = "1.0.1"
Const FORM_NAME As String = "HF PS_PG Utility"

Dim SHIPath As String

Private Sub UserForm_Initialize()
    
    frmMain.Caption = FORM_NAME
    'lblVersion.Caption = "v" & VERSION_NUM
    
End Sub

Private Sub btnProcess_Click()
    
    Call GetOnePlant
    Call GetEventNumber
    'Call GetJobNumber
    
    frmMain.Hide
    Call Main(SHIPath, GetEventNumber, GetJobNumber, GetOnePlant)
    Unload Me
    
End Sub

Private Sub btnBrowseInput1_Click()
    
    SHIPath = FilePicker("E")
    
    'sets text field only if a file was selected
    If Not SHIPath = "" Then
        txtInputPath1.value = RemoveBefore(SHIPath, "\")
    End If
    
End Sub

Private Function GetOnePlant()

    If ToggleButton1.value = True Then
        GetOnePlant = True
    Else
        GetOnePlant = False
    End If

End Function

Private Function GetEventNumber() As String
    
    Dim s As String
    s = CC(TextBox1.value)
    GetEventNumber = s
    
End Function

Private Function GetJobNumber() As String
    
    Dim s As String
    s = CC(TextBox2.value)
    GetJobNumber = s
    
End Function
