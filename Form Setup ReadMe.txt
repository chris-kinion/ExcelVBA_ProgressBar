Form Setup:
	1. Form: Change caption to progress indicator
	2. Frame: Add frame control
		a. Clear caption
		b. Height = 24
		c. Width = 204
	3. Label: Add outside frame
		a. Change caption to 0% Completed
	4. Label: Add inside frame
		a. Change BackColor to Highlight
		b. Clear caption
		c. Height=20
		d. Width=10 (min of 0, max of 200)
		e. Top and Left offset = 0

Form:
Private Sub UserForm_Activate()
  Call SampleEventStatusGenerator
  Unload ProgressForm
End Sub

Private Sub UserForm_Initialize()
  ProgressForm.Status.Caption = "0% Complete"
  ProgressForm.Bar.Width = 10
End Sub

Module:
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
