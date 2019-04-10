Attribute VB_Name = "Event_ProgressBar"
Option Explicit
Option Base 1

Sub ShowStatus()
  ProgressForm.Show
End Sub

Sub TimeFrame(pauseTime As Double)
  Start = Timer
  Do
  DoEvents
  Loop Until (Timer - Start) >= pauseTime
End Sub

Sub SampleEventStatusGenerator()
  Dim i As Long
  For i = 0 To 100
    Call UpdateProgressBar(i)
    Call TimeFrame(0.1)
  Next i
End Sub

Sub UpdateProgressBar(lngPercentComplete As Long)
  ProgressForm.Bar.Width = 2 * lngPercentComplete
  ProgressForm.Status.Caption = lngPercentComplete & "% Complete"
End Sub
