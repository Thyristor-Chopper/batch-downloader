VERSION 5.00
Begin VB.Form Bluemetal 
   Caption         =   "Form1"
   ClientHeight    =   3060
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4650
   BeginProperty Font 
      Name            =   "±¼¸²"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3060
   ScaleWidth      =   4650
   StartUpPosition =   3  'Windows ±âº»°ª
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2520
      TabIndex        =   8
      Top             =   2040
      Width           =   1455
   End
   Begin VB.PictureBox pbRight 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   930
      Left            =   4560
      MousePointer    =   9  'W E Å©±â Á¶Á¤
      ScaleHeight     =   930
      ScaleWidth      =   75
      TabIndex        =   7
      Top             =   480
      Width           =   75
   End
   Begin VB.PictureBox pbBottomRight 
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   60
      Left            =   3720
      MousePointer    =   8  'NW SE Å©±â Á¶Á¤
      Picture         =   "Bluemetal.frx":0000
      ScaleHeight     =   60
      ScaleWidth      =   60
      TabIndex        =   6
      Top             =   1560
      Width           =   60
   End
   Begin VB.PictureBox pbBottomMiddle 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   75
      Left            =   120
      MousePointer    =   7  'N SÅ©±â Á¶Á¤
      ScaleHeight     =   75
      ScaleWidth      =   2055
      TabIndex        =   5
      Top             =   1680
      Width           =   2055
   End
   Begin VB.PictureBox pbBottomLeft 
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   60
      Left            =   0
      MousePointer    =   6  'NE SW Å©±â Á¶Á¤
      Picture         =   "Bluemetal.frx":0042
      ScaleHeight     =   60
      ScaleWidth      =   60
      TabIndex        =   4
      Top             =   1680
      Width           =   60
   End
   Begin VB.PictureBox pbLeft 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   930
      Left            =   0
      MousePointer    =   9  'W E Å©±â Á¶Á¤
      ScaleHeight     =   930
      ScaleWidth      =   75
      TabIndex        =   3
      Top             =   480
      Width           =   75
   End
   Begin VB.PictureBox pbTopRight 
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   3360
      Picture         =   "Bluemetal.frx":0083
      ScaleHeight     =   480
      ScaleWidth      =   1260
      TabIndex        =   2
      Top             =   0
      Width           =   1260
   End
   Begin VB.PictureBox pbTopLeft 
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   0
      Picture         =   "Bluemetal.frx":069D
      ScaleHeight     =   480
      ScaleWidth      =   570
      TabIndex        =   1
      Top             =   0
      Width           =   570
   End
   Begin VB.PictureBox pbTopMiddle 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   570
      ScaleHeight     =   480
      ScaleWidth      =   2250
      TabIndex        =   0
      Top             =   0
      Width           =   2250
      Begin VB.Label lblCaption 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BackStyle       =   0  'Åõ¸í
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   9
         Top             =   150
         Width           =   2295
      End
      Begin VB.Label lblCaptionShadow 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BackStyle       =   0  'Åõ¸í
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   15
         TabIndex        =   10
         Top             =   165
         Width           =   2295
      End
   End
   Begin VB.Image imgTopMiddle 
      Height          =   480
      Left            =   1560
      Picture         =   "Bluemetal.frx":0C18
      Top             =   600
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Image imgBottomMiddle 
      Height          =   60
      Left            =   480
      Picture         =   "Bluemetal.frx":0F7E
      Top             =   1560
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Image imgRight 
      Height          =   930
      Left            =   4320
      Picture         =   "Bluemetal.frx":0FBC
      Top             =   600
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Image imgLeft 
      Height          =   930
      Left            =   1080
      Picture         =   "Bluemetal.frx":13DE
      Top             =   600
      Visible         =   0   'False
      Width           =   75
   End
End
Attribute VB_Name = "Bluemetal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Hook_Bluemetal Me.hWnd
    SetWindowLong pbTopLeft.hWnd, GWL_EXSTYLE, &H80
    SetParent pbTopLeft.hWnd, GetParent(Me.hWnd)
    SetWindowLong pbTopMiddle.hWnd, GWL_EXSTYLE, &H80
    SetParent pbTopMiddle.hWnd, GetParent(Me.hWnd)
    SetWindowLong pbTopRight.hWnd, GWL_EXSTYLE, &H80
    SetParent pbTopRight.hWnd, GetParent(Me.hWnd)
    SetWindowLong pbLeft.hWnd, GWL_EXSTYLE, &H80
    SetParent pbLeft.hWnd, GetParent(Me.hWnd)
    SetWindowLong pbBottomLeft.hWnd, GWL_EXSTYLE, &H80
    SetParent pbBottomLeft.hWnd, GetParent(Me.hWnd)
    SetWindowLong pbBottomMiddle.hWnd, GWL_EXSTYLE, &H80
    SetParent pbBottomMiddle.hWnd, GetParent(Me.hWnd)
    SetWindowLong pbBottomRight.hWnd, GWL_EXSTYLE, &H80
    SetParent pbBottomRight.hWnd, GetParent(Me.hWnd)
    SetWindowLong pbRight.hWnd, GWL_EXSTYLE, &H80
    SetParent pbRight.hWnd, GetParent(Me.hWnd)
    
    Form_Resize
    
    lblCaption.Caption = Me.Caption
    lblCaptionShadow.Caption = Me.Caption
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Dim rc As RECT
    GetWindowRect Me.hWnd, rc
    
    Dim X%, Y%
    
    pbTopMiddle.Width = (rc.Right - rc.Left) * 15 - pbTopLeft.Width - pbTopRight.Width
    For X = 0 To pbTopMiddle.Width Step imgTopMiddle.Width
        pbTopMiddle.PaintPicture imgTopMiddle.Picture, X, 0
    Next X
    
    pbLeft.Height = (rc.Bottom - rc.Top) * 15 - pbTopLeft.Height - pbBottomLeft.Height
    For Y = 0 To pbLeft.Height Step imgLeft.Height
        pbLeft.PaintPicture imgLeft.Picture, 0, Y
    Next Y
    
    pbBottomMiddle.Width = (rc.Right - rc.Left) * 15 - pbBottomLeft.Width - pbBottomRight.Width
    For X = 0 To pbBottomMiddle.Width Step imgBottomMiddle.Width
        pbBottomMiddle.PaintPicture imgBottomMiddle.Picture, X, 0
    Next X
    
    pbRight.Height = (rc.Bottom - rc.Top) * 15 - pbTopRight.Height - pbBottomRight.Height
    For Y = 0 To pbRight.Height Step imgRight.Height
        pbRight.PaintPicture imgRight.Picture, 0, Y
    Next Y
    
    Dim Rgn&, Rgn1&, Rgn2&, Rgn3&, Rgn4&, Rgn5&, Rgn6&, Rgn7&, Rgn8&, Rgn9&
    Rgn = CreateRectRgn(0, 0, rc.Right - rc.Left, rc.Bottom - rc.Top)
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
    
    lblCaption.Width = pbTopMiddle.Width
    lblCaptionShadow.Width = pbTopMiddle.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unhook_Bluemetal Me.hWnd
    SetParent pbTopLeft.hWnd, Me.hWnd
    SetParent pbTopMiddle.hWnd, Me.hWnd
    SetParent pbTopRight.hWnd, Me.hWnd
    Unload Me
End Sub

Private Sub pbBottomLeft_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        ReleaseCapture
        SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTBOTTOMLEFT, 0&
    End If
End Sub

Private Sub pbBottomMiddle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        ReleaseCapture
        SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTBOTTOM, 0&
    End If
End Sub

Private Sub lblCaption_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        ReleaseCapture
        SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    End If
End Sub

Private Sub lblCaptionShadow_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        ReleaseCapture
        SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    End If
End Sub

Private Sub pbBottomRight_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        ReleaseCapture
        SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTBOTTOMRIGHT, 0&
    End If
End Sub

Private Sub pbLeft_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        ReleaseCapture
        SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTLEFT, 0&
    End If
End Sub

Private Sub pbRight_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        ReleaseCapture
        SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTRIGHT, 0&
    End If
End Sub

Private Sub pbTopLeft_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    pbTopMiddle_MouseMove Button, Shift, X, Y
End Sub

Private Sub pbTopMiddle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        ReleaseCapture
        SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    End If
End Sub

Private Sub pbTopRight_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    pbTopMiddle_MouseMove Button, Shift, X, Y
End Sub
