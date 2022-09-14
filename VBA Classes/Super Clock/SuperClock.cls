VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SuperClock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private toTime As Single
Private stTime As Single
Private cmTime As Single
Private p1Time As Single
Private p2Time As Single
Private pTTime As Single

Private Sub Class_Initialize()

    toTime = 0#
    stTime = 0#
    cmTime = 0#
    p1Time = 0#
    p2Time = 0#
    pTTime = 0#
	
End Sub

Public Sub Start()

    stTime = Timer
	
End Sub

Public Sub Pause()

    p1Time = Timer
	
End Sub

Public Sub Proceed()

    p2Time = Timer
    pTTime = pTTime + (p2Time - p1Time)
	
End Sub

Public Function Complete() As Single

    cmTime = Timer
    toTime = (cmTime - stTime) - pTTime
    Complete = toTime
	
End Function



