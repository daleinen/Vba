VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SuperBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *** Please name instance myBar
Private Const F_COUNT As Single = 7
' *** CHANGE to the F_COUNT const above to number of times function is called throughout macro, + 1

Private iCount As Single
Private Const GWL_STYLE = -16
Private Const WS_CAPTION = &HC00000
Private Const WS_SYSMENU = &H80000
Private Const SC_CLOSE = &HF060

Private Declare PtrSafe Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare PtrSafe Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare PtrSafe Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long

Private Sub Class_Initialize()
    
    iCount = 1
    
End Sub

Public Sub CloseButtonSettings(frm As Object, show As Boolean)

    Dim windowHandle As Long
    Dim menuHandle As Long
    windowHandle = FindWindowA(vbNullString, frm.Caption)
    
    If show = True Then
        menuHandle = GetSystemMenu(windowHandle, 1)
    Else
        menuHandle = GetSystemMenu(windowHandle, 0)
        DeleteMenu menuHandle, SC_CLOSE, 0&
    End If

End Sub

Public Sub EndProgressBar()

    Unload ufProgress
    
End Sub

Public Sub FractionComplete()

    Dim pctdone As Single
    pctdone = iCount / F_COUNT
    pctdone = Round(pctdone, 2)
    iCount = iCount + 1
    
    With ufProgress
        .LabelCaption.Caption = pctdone * 100 & "% Complete"
        .LabelCaption.Font.Bold = True
        .LabelCaption.Font.Size = 10
        .LabelCaption.TextAlign = fmTextAlignCenter
        .LabelProgress.Width = pctdone * (.FrameProgress.Width)
    End With
    DoEvents
    
    If pctdone >= 1 Then Call WasteSeconds(1)
    
End Sub

Public Sub ShowProgessBar()

    ufProgress.LabelProgress.Width = 0
    ufProgress.show
    Call FractionComplete
    
End Sub
