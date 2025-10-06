VERSION 5.00
Begin VB.Form frmMainContainer 
   BorderStyle     =   0  '없음
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows 기본값
End
Attribute VB_Name = "frmMainContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    frmMain.Show vbModeless
    frmMain.SetFrameTexture
    SetParent frmMain.hWnd, Me.hWnd
    frmMain.Move 0, 0
    Me.Width = frmMain.Width
    Me.Height = frmMain.Height
End Sub
