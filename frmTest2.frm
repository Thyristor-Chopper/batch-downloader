VERSION 5.00
Begin VB.Form frmTest2 
   Caption         =   "Form2"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows ±âº»°ª
End
Attribute VB_Name = "frmTest2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Skinner As frmSkinnedFrame

Private Sub Form_Load()
    Set Skinner = New frmSkinnedFrame
    Skinner.Init Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Skinner
End Sub
