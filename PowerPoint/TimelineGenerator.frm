VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TimelineGenerator 
   Caption         =   "Zeitleiste generieren"
   ClientHeight    =   2070
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4110
   OleObjectBlob   =   "TimelineGenerator.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "TimelineGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Generate_Click()
 GenerateTimeline TimelineTop.Value, TimelineHeight.Value, ImageTop.Value, ImageHeight.Value
 Unload Me
End Sub
