VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows ±âº»°ª
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   2520
      Width           =   1455
   End
   Begin VB.PictureBox Picture2 
      Height          =   2295
      Left            =   240
      ScaleHeight     =   2235
      ScaleWidth      =   3915
      TabIndex        =   2
      Top             =   120
      Width           =   3975
      Begin VB.PictureBox Picture1 
         BackColor       =   &H0000FFFF&
         Height          =   1455
         Left            =   840
         ScaleHeight     =   1395
         ScaleWidth      =   2475
         TabIndex        =   4
         Top             =   360
         Width           =   2535
         Begin VB.CommandButton Command3 
            Caption         =   "Command3"
            Height          =   375
            Left            =   1320
            TabIndex        =   5
            Top             =   0
            Width           =   1215
         End
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Skinner As frmSkinnedFrame
Dim Skinner2 As frmSkinnedFrame

Private Sub Command1_Click()
    Me.Caption = Rnd
End Sub

Private Sub Command2_Click()
    frmTest2.Show
End Sub

Private Sub Command4_Click()
    Skinner2.ReloadSkin
End Sub

Private Sub Form_Load()
    Set Skinner = New frmSkinnedFrame
    Skinner.Init Me
    
    Set Skinner2 = New frmSkinnedFrame
    SetWindowLong Picture1.hWnd, GWL_STYLE, GetWindowLong(Picture1.hWnd, GWL_STYLE) Or WS_BORDER Or WS_OVERLAPPED Or WS_CAPTION Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_SYSMENU
    SetWindowText Picture1.hWnd, App.Title
    Picture1.Height = Picture1.Height + 15
    Picture1.Refresh
    Skinner2.Init Picture1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Skinner
    Unload Skinner2
End Sub
