VERSION 5.00
Begin VB.Form Bluemetal 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows 기본값
   Begin VB.PictureBox pbTopMiddle 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  '없음
      Height          =   480
      Left            =   570
      ScaleHeight     =   480
      ScaleWidth      =   2250
      TabIndex        =   0
      Top             =   0
      Width           =   2250
   End
   Begin VB.Image imgBottomMiddle 
      Height          =   60
      Left            =   480
      Picture         =   "frmBluemetal.frx":0000
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Image imgBottomRight 
      Height          =   60
      Left            =   3960
      Picture         =   "frmBluemetal.frx":003E
      Top             =   1560
      Width           =   60
   End
   Begin VB.Image imgBottomLeft 
      Height          =   60
      Left            =   0
      Picture         =   "frmBluemetal.frx":0080
      Top             =   1440
      Width           =   60
   End
   Begin VB.Image imgRight 
      Height          =   930
      Left            =   4545
      Picture         =   "frmBluemetal.frx":00C1
      Stretch         =   -1  'True
      Top             =   480
      Width           =   75
   End
   Begin VB.Image imgLeft 
      Height          =   930
      Left            =   0
      Picture         =   "frmBluemetal.frx":04E3
      Stretch         =   -1  'True
      Top             =   480
      Width           =   75
   End
   Begin VB.Image imgTopMiddle 
      Height          =   480
      Left            =   720
      Picture         =   "frmBluemetal.frx":0905
      Top             =   600
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Image imgTopRight 
      Height          =   480
      Left            =   3360
      Picture         =   "frmBluemetal.frx":0C6B
      Top             =   0
      Width           =   1260
   End
   Begin VB.Image imgTopLeft 
      Height          =   480
      Left            =   0
      Picture         =   "frmBluemetal.frx":1285
      Top             =   0
      Width           =   570
   End
End
Attribute VB_Name = "Bluemetal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim X As Integer
    imgTopRight.Top = 0
    imgTopLeft.Top = 0
    imgTopLeft.Left = 0
    pbTopMiddle.Left = imgTopLeft.Width
    pbTopMiddle.Width = Me.Width - imgTopRight.Width - imgTopLeft.Width
    imgTopRight.Left = Me.Width - imgTopRight.Width
    For X = 0 To Me.Width - imgTopRight.Width - imgTopLeft.Width Step 1035
        pbTopMiddle.PaintPicture imgTopMiddle.Picture, X, 0
    Next X
    imgBottomLeft.Left = 0
    imgBottomLeft.Top = Me.Height - imgBottomLeft.Height
    imgBottomRight.Top = Me.Height - imgBottomRight.Height
    imgBottomMiddle.Top = Me.Height - imgBottomMiddle.Height
    imgBottomMiddle.Left = imgBottomLeft.Width
    imgBottomRight.Left = Me.Width - imgBottomRight.Width
    imgLeft.Left = 0
    imgLeft.Top = imgTopLeft.Height
    imgLeft.Height = Me.Height - imgTopLeft.Height - imgBottomLeft.Height
    imgRight.Left = Me.Width - imgRight.Width
    imgRight.Height = imgLeft.Height
    imgBottomRight.Top = imgBottomLeft.Top
    imgBottomMiddle.Left = imgBottomLeft.Width
    imgBottomMiddle.Width = Me.Width - imgBottomLeft.Width - imgBottomRight.Width
End Sub

Private Sub Form_Resize()
    Exit Sub
    
    Dim Rgn&, Rgn1&, Rgn2&, Rgn3&, Rgn4&, Rgn5&, Rgn6&, Rgn7&, Rgn8&, Rgn9&
    Rgn = CreateRectRgn(0, 0, Me.Width / 15, Me.Height / 15)
    Rgn1 = CreateRectRgn(0, 0, 7, 1)
    Rgn2 = CreateRectRgn(0, 1, 5, 2)
    Rgn3 = CreateRectRgn(0, 2, 3, 3)
    Rgn4 = CreateRectRgn(0, 3, 2, 4)
    Rgn5 = CreateRectRgn(0, 4, 2, 5)
    Rgn6 = CreateRectRgn(0, 5, 1, 6)
    Rgn7 = CreateRectRgn(0, 6, 1, 7)
    CombineRgn Rgn, Rgn, Rgn1, RGN_DIFF
    CombineRgn Rgn, Rgn, Rgn2, RGN_DIFF
    CombineRgn Rgn, Rgn, Rgn3, RGN_DIFF
    CombineRgn Rgn, Rgn, Rgn4, RGN_DIFF
    CombineRgn Rgn, Rgn, Rgn5, RGN_DIFF
    CombineRgn Rgn, Rgn, Rgn6, RGN_DIFF
    CombineRgn Rgn, Rgn, Rgn7, RGN_DIFF
    SetWindowRgn Me.hWnd, Rgn, True
End Sub
