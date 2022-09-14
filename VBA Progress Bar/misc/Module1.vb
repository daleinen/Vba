Attribute VB_Name = "Module1"
Option Explicit

Sub DemoMacro()
    '(Step 1) Display your Progress Bar
    ufProgress.LabelProgress.Width = 0
    ufProgress.Show
    FractionComplete (0) '(Step 2)
        '--------------------------------------
        'A portion of your macro goes here
        '--------------------------------------
    FractionComplete (0.25) '(Step 2)
        '--------------------------------------
        'A portion of your macro goes here
        '--------------------------------------
    FractionComplete (0.5) '(Step 2)
        '--------------------------------------
        'A portion of your macro goes here
        '--------------------------------------
    FractionComplete (0.75) '(Step 2)
        '--------------------------------------
        'A portion of your macro goes here
        '--------------------------------------
    FractionComplete (1) '(Step 2)

    '(Step 3) Close the progress bar when you're done
        Unload ufProgress

    End Sub

    Sub FractionComplete(pctdone As Single)
    With ufProgress
        .LabelCaption.Caption = pctdone * 100 & "% Complete"
        .LabelProgress.Width = pctdone * (.FrameProgress.Width)
    End With
    DoEvents
End Sub

Sub LoopThroughRows()

    Dim i As Long, lastrow As Long
    Dim pctdone As Single
    lastrow = Range("B" & Rows.Count).End(xlUp).Row
    
    '(Step 1) Display your Progress Bar
    ufProgress.LabelProgress.Width = 0
    ufProgress.Show
    
    For i = 1 To lastrow
    '(Step 2) Periodically update progress bar
        pctdone = i / lastrow
        With ufProgress
            .LabelCaption.Caption = "Processing Row " & i & " of " & lastrow
            .LabelProgress.Width = pctdone * (.FrameProgress.Width)
        End With
        DoEvents
            '--------------------------------------
            'the rest of your macro goes below here
            '
            '
            '--------------------------------------
    '(Step 3) Close the progress bar when you're done
        If i = lastrow Then Unload ufProgress
    Next i
    
End Sub
