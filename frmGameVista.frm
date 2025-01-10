VERSION 5.00
Begin VB.Form frmGameVista 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "랜덤 픽셀"
   ClientHeight    =   5400
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6960
   Icon            =   "frmGameVista.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   6960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.PictureBox picCanvas 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '없음
      Height          =   5415
      Left            =   0
      ScaleHeight     =   5415
      ScaleWidth      =   6975
      TabIndex        =   0
      Top             =   0
      Width           =   6975
      Begin VB.Timer timTimer 
         Interval        =   50
         Left            =   1440
         Top             =   1800
      End
   End
End
Attribute VB_Name = "frmGameVista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Randomize
    picCanvas.BackColor = RGB(Int(Rnd * 255), Int(Rnd * 255), Int(Rnd * 255))
End Sub

Private Sub timTimer_Timer()
    Dim x%, y%
    x = Int(Rnd * (picCanvas.Width - 50)) + 25
    y = Int(Rnd * (picCanvas.Height - 50)) + 25
    Me.Caption = "랜덤 픽셀 (" & x & "," & y & ")"
    picCanvas.Circle (x, y), 12, RGB(Int(Rnd * 255), Int(Rnd * 255), Int(Rnd * 255))
End Sub
