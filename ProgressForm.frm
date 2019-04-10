VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProgressForm 
   Caption         =   "Progress Indicator"
   ClientHeight    =   1080
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4395
   OleObjectBlob   =   "ProgressForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ProgressForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Activate()
  Call SampleEventStatusGenerator
  Unload ProgressForm
End Sub



Private Sub UserForm_Initialize()
  ProgressForm.Status.Caption = "0% Complete"
  ProgressForm.Bar.Width = 10
End Sub
