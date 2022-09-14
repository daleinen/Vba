
Public Sub FractionComplete(pctdone As Single)

    With ufProgress
        .LabelCaption.Caption = pctdone * 100 & "% Complete"
        .LabelProgress.Width = pctdone * (.FrameProgress.Width)
    End With
    DoEvents
    
End Sub

Public Sub ShowProgessBar()

    ufProgress.LabelProgress.Width = 0
    ufProgress.Show
    
End Sub

Public Sub EndProgressBar()

    Unload ufProgress
    
End Sub