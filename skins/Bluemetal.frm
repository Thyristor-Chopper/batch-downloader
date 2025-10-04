VERSION 5.00
Begin VB.Form Bluemetal 
   Caption         =   "Form1"
   ClientHeight    =   3165
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4650
   BeginProperty Font 
      Name            =   "굴림"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3165
   ScaleWidth      =   4650
   StartUpPosition =   3  'Windows 기본값
   Begin VB.Timer timMinimizeHover 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3600
      Top             =   2160
   End
   Begin VB.Timer timMaximizeHover 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3120
      Top             =   2160
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   2400
      TabIndex        =   17
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Timer timCloseHover 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2640
      Top             =   2160
   End
   Begin VB.PictureBox pbRight 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  '없음
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   930
      Left            =   4560
      MousePointer    =   9  'W E 크기 조정
      ScaleHeight     =   930
      ScaleWidth      =   75
      TabIndex        =   7
      Top             =   480
      Width           =   75
   End
   Begin VB.PictureBox pbBottomRight 
      BorderStyle     =   0  '없음
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   60
      Left            =   3720
      MousePointer    =   8  'NW SE 크기 조정
      Picture         =   "Bluemetal.frx":0000
      ScaleHeight     =   60
      ScaleWidth      =   60
      TabIndex        =   6
      Top             =   1560
      Width           =   60
   End
   Begin VB.PictureBox pbBottomMiddle 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  '없음
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   60
      Left            =   240
      MousePointer    =   7  'N S크기 조정
      ScaleHeight     =   60
      ScaleWidth      =   2055
      TabIndex        =   5
      Top             =   1920
      Width           =   2055
   End
   Begin VB.PictureBox pbBottomLeft 
      BorderStyle     =   0  '없음
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   60
      Left            =   0
      MousePointer    =   6  'NE SW 크기 조정
      Picture         =   "Bluemetal.frx":0042
      ScaleHeight     =   60
      ScaleWidth      =   60
      TabIndex        =   4
      Top             =   1680
      Width           =   60
   End
   Begin VB.PictureBox pbLeft 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  '없음
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   930
      Left            =   0
      MousePointer    =   9  'W E 크기 조정
      ScaleHeight     =   930
      ScaleWidth      =   75
      TabIndex        =   3
      Top             =   480
      Width           =   75
   End
   Begin VB.PictureBox pbTopRight 
      BorderStyle     =   0  '없음
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   3360
      ScaleHeight     =   480
      ScaleWidth      =   1260
      TabIndex        =   2
      Top             =   0
      Width           =   1260
      Begin VB.Image imgMinimizeButton 
         Height          =   315
         Left            =   120
         Top             =   90
         Width           =   315
      End
      Begin VB.Image imgMaximizeButton 
         Height          =   315
         Left            =   480
         Top             =   90
         Width           =   315
      End
      Begin VB.Image imgCloseButton 
         Height          =   315
         Left            =   840
         Top             =   90
         Width           =   315
      End
      Begin VB.Label lblResizeRight 
         BackStyle       =   0  '투명
         Height          =   495
         Left            =   1185
         MousePointer    =   9  'W E 크기 조정
         TabIndex        =   15
         Top             =   75
         Width           =   75
      End
      Begin VB.Label lblResizeTopRight 
         BackStyle       =   0  '투명
         Height          =   75
         Left            =   960
         MousePointer    =   6  'NE SW 크기 조정
         TabIndex        =   14
         Top             =   0
         Width           =   375
      End
      Begin VB.Label lblResizeTop 
         BackStyle       =   0  '투명
         Height          =   75
         Index           =   1
         Left            =   0
         MousePointer    =   7  'N S크기 조정
         TabIndex        =   11
         Top             =   0
         Width           =   960
      End
   End
   Begin VB.PictureBox pbTopLeft 
      BorderStyle     =   0  '없음
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   570
      TabIndex        =   1
      Top             =   0
      Width           =   570
      Begin VB.Image imgControlMenu 
         Height          =   240
         Left            =   120
         Stretch         =   -1  'True
         Top             =   120
         Width           =   240
      End
      Begin VB.Label lblResizeLeft 
         BackStyle       =   0  '투명
         Height          =   390
         Left            =   0
         MousePointer    =   9  'W E 크기 조정
         TabIndex        =   16
         Top             =   75
         Width           =   75
      End
      Begin VB.Label lblResizeTopLeft 
         BackStyle       =   0  '투명
         Height          =   75
         Left            =   0
         MousePointer    =   8  'NW SE 크기 조정
         TabIndex        =   13
         Top             =   0
         Width           =   255
      End
      Begin VB.Label lblResizeTop 
         BackStyle       =   0  '투명
         Height          =   75
         Index           =   2
         Left            =   240
         MousePointer    =   7  'N S크기 조정
         TabIndex        =   12
         Top             =   0
         Width           =   330
      End
   End
   Begin VB.PictureBox pbTopMiddle 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  '없음
      BeginProperty Font 
         Name            =   "돋움"
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
      Begin VB.Label lblResizeTop 
         BackStyle       =   0  '투명
         Height          =   75
         Index           =   0
         Left            =   0
         MousePointer    =   7  'N S크기 조정
         TabIndex        =   10
         Top             =   0
         Width           =   2295
      End
      Begin VB.Label lblCaption 
         Alignment       =   2  '가운데 맞춤
         BackStyle       =   0  '투명
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "굴림"
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
         TabIndex        =   8
         Top             =   150
         Width           =   2295
      End
      Begin VB.Label lblCaptionShadow 
         Alignment       =   2  '가운데 맞춤
         BackStyle       =   0  '투명
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "굴림"
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
         TabIndex        =   9
         Top             =   165
         Width           =   2295
      End
   End
   Begin VB.Image imgMinimize 
      Height          =   315
      Index           =   3
      Left            =   1080
      Picture         =   "Bluemetal.frx":0083
      Top             =   2760
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image imgMinimize 
      Height          =   315
      Index           =   2
      Left            =   720
      Picture         =   "Bluemetal.frx":0539
      Top             =   2760
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgMinimize 
      Height          =   315
      Index           =   0
      Left            =   1440
      Picture         =   "Bluemetal.frx":0A06
      Top             =   2760
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgMinimize 
      Height          =   315
      Index           =   1
      Left            =   360
      Picture         =   "Bluemetal.frx":0ECB
      Top             =   2760
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgMaximize 
      Height          =   315
      Index           =   4
      Left            =   2880
      Picture         =   "Bluemetal.frx":138B
      Top             =   2400
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgMaximize 
      Height          =   315
      Index           =   5
      Left            =   1800
      Picture         =   "Bluemetal.frx":1850
      Top             =   2400
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgMaximize 
      Height          =   315
      Index           =   6
      Left            =   2160
      Picture         =   "Bluemetal.frx":1CFD
      Top             =   2400
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgMaximize 
      Height          =   315
      Index           =   7
      Left            =   2520
      Picture         =   "Bluemetal.frx":21B5
      Top             =   2400
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgTopMiddle 
      Height          =   480
      Index           =   0
      Left            =   840
      Picture         =   "Bluemetal.frx":265C
      Top             =   1320
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Image imgTopLeft 
      Height          =   480
      Index           =   0
      Left            =   240
      Picture         =   "Bluemetal.frx":409E
      Top             =   1320
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Image imgTopRight 
      Height          =   480
      Index           =   0
      Left            =   1920
      Picture         =   "Bluemetal.frx":4F60
      Top             =   1320
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image imgTopRight 
      Height          =   480
      Index           =   1
      Left            =   1920
      Picture         =   "Bluemetal.frx":6F22
      Top             =   720
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image imgTopLeft 
      Height          =   480
      Index           =   1
      Left            =   240
      Picture         =   "Bluemetal.frx":753C
      Top             =   720
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Image imgMaximize 
      Height          =   315
      Index           =   0
      Left            =   1440
      Picture         =   "Bluemetal.frx":7AB7
      Top             =   2400
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgMaximize 
      Height          =   315
      Index           =   3
      Left            =   1080
      Picture         =   "Bluemetal.frx":7F7C
      Top             =   2400
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgMaximize 
      Height          =   315
      Index           =   2
      Left            =   720
      Picture         =   "Bluemetal.frx":8421
      Top             =   2400
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgMaximize 
      Height          =   315
      Index           =   1
      Left            =   360
      Picture         =   "Bluemetal.frx":88D8
      Top             =   2400
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgClose 
      Height          =   315
      Index           =   0
      Left            =   1440
      Picture         =   "Bluemetal.frx":8D84
      Top             =   2040
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgClose 
      Height          =   315
      Index           =   3
      Left            =   1080
      Picture         =   "Bluemetal.frx":9249
      Top             =   2040
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgClose 
      Height          =   315
      Index           =   2
      Left            =   720
      Picture         =   "Bluemetal.frx":96EE
      Top             =   2040
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgClose 
      Height          =   315
      Index           =   1
      Left            =   360
      Picture         =   "Bluemetal.frx":9BA6
      Top             =   2040
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgTopMiddle 
      Height          =   480
      Index           =   1
      Left            =   840
      Picture         =   "Bluemetal.frx":A053
      Top             =   720
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Image imgBottomMiddle 
      Height          =   60
      Left            =   480
      Picture         =   "Bluemetal.frx":A3B9
      Top             =   1560
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Image imgRight 
      Height          =   930
      Left            =   4320
      Picture         =   "Bluemetal.frx":A3F7
      Top             =   600
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Image imgLeft 
      Height          =   930
      Left            =   120
      Picture         =   "Bluemetal.frx":A819
      Top             =   720
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

Public IsWindowActive As Byte

Sub SetSkinTextures()
    If Not Enable Then Exit Sub
    
    Set pbTopLeft.Picture = imgTopLeft(IsWindowActive).Picture
    Set pbTopRight.Picture = imgTopRight(IsWindowActive).Picture
    
    Set imgCloseButton.Picture = imgClose(IsWindowActive).Picture
    
    imgMaximizeButton.Enabled = Me.MaxButton
    SetMaxButtonTexture
    
    imgMinimizeButton.Enabled = Me.MinButton
    If Me.MaxButton Then
        Set imgMinimizeButton.Picture = imgMinimize(IsWindowActive).Picture
    Else
        Set imgMinimizeButton.Picture = imgMinimize(0).Picture
    End If
    
    lblCaptionShadow.Visible = IsWindowActive
    
    imgCloseButton.Top = pbTopRight.Height - imgCloseButton.Height - 3 * Screen.TwipsPerPixelY
    imgMaximizeButton.Top = pbTopRight.Height - imgMaximizeButton.Height - 3 * Screen.TwipsPerPixelY
    imgMinimizeButton.Top = pbTopRight.Height - imgMinimizeButton.Height - 3 * Screen.TwipsPerPixelY
    imgCloseButton.Left = pbTopRight.Width - BorderSize * Screen.TwipsPerPixelX - imgCloseButton.Width - 2 * Screen.TwipsPerPixelX
    imgMaximizeButton.Left = imgCloseButton.Left - 2 * Screen.TwipsPerPixelX - imgMaximizeButton.Width
    imgMinimizeButton.Left = imgMaximizeButton.Left - 2 * Screen.TwipsPerPixelX - imgMinimizeButton.Width
End Sub

Sub SetMaxButtonTexture()
    If Me.MaxButton Then
        Set imgMaximizeButton.Picture = imgMaximize(IsWindowActive - (Me.WindowState = 2) * 4).Picture
    Else
        Set imgMaximizeButton.Picture = imgMaximize(0).Picture
    End If
End Sub

Sub SetSizableSkinTextures()
    If Me.WindowState = 1 Or (Not Enable) Then Exit Sub
    
    Dim rc As RECT
    GetWindowRect Me.hWnd, rc
    
    Dim X%, Y%
    
    pbTopMiddle.Width = (rc.Right - rc.Left) * 15 - pbTopLeft.Width - pbTopRight.Width
    For X = 0 To pbTopMiddle.Width Step imgTopMiddle(IsWindowActive).Width
        pbTopMiddle.PaintPicture imgTopMiddle(IsWindowActive).Picture, X, 0
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
End Sub

Private Sub SetSkin(EnableSkin As Boolean)
    Enable = EnableSkin
    
    If EnableSkin Then
    
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
        
        Set imgControlMenu.Picture = Me.Icon
        SetSkinTextures
        
        SetResizeCursors
    End If
    
    pbTopLeft.Visible = EnableSkin
    pbTopMiddle.Visible = EnableSkin
    pbTopRight.Visible = EnableSkin
    pbLeft.Visible = EnableSkin
    pbRight.Visible = EnableSkin
    pbBottomLeft.Visible = EnableSkin
    pbBottomMiddle.Visible = EnableSkin
    pbBottomRight.Visible = EnableSkin
    
    SetRgn
    
    SetWindowPos Me.hWnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER Or SWP_FRAMECHANGED
End Sub

Private Sub Command1_Click()
    SetSkin Not Enable
End Sub

Private Sub Form_Load()
    IsWindowActive = 1
    TitleHeight = 32
    BorderSize = 5
    Hook_Bluemetal Me.hWnd
    SetSkin False
End Sub

Private Sub SetResizeCursors()
    Dim Active As Byte
    Active = -(Me.WindowState = 0 And Me.BorderStyle = 2)
    lblResizeLeft.MousePointer = 9 * Active
    lblResizeTopLeft.MousePointer = 8 * Active
    lblResizeTop(0).MousePointer = 7 * Active
    lblResizeTop(1).MousePointer = 7 * Active
    lblResizeTop(2).MousePointer = 7 * Active
    lblResizeTopRight.MousePointer = 6 * Active
    lblResizeRight.MousePointer = 9 * Active
    pbLeft.MousePointer = 9 * Active
    pbRight.MousePointer = 9 * Active
    pbBottomLeft.MousePointer = 6 * Active
    pbBottomMiddle.MousePointer = 8 * Active
    pbBottomRight.MousePointer = 8 * Active
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    SetSizableSkinTextures
    
    SetRgn
    
    lblCaption.Width = pbTopMiddle.Width
    lblCaptionShadow.Width = pbTopMiddle.Width
    lblResizeTop(0).Width = pbTopMiddle.Width
End Sub

Sub SetRgn()
    If Not Enable Then SetWindowRgn Me.hWnd, 0&, True: Exit Sub
    Dim rc As RECT
    Dim Rgn&, Rgn1&, Rgn2&, Rgn3&, Rgn4&, Rgn5&, Rgn6&, Rgn7&, Rgn8&, Rgn9&
    Dim Width%, Height%
    
    '창
    GetWindowRect Me.hWnd, rc
    Width = rc.Right - rc.Left: Height = rc.Bottom - rc.Top
    Rgn = CreateRectRgn(0, 0, Width, Height)
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
    DeleteObject Rgn1
    DeleteObject Rgn2
    DeleteObject Rgn3
    DeleteObject Rgn4
    DeleteObject Rgn5
    DeleteObject Rgn6
    DeleteObject Rgn7
    Rgn1 = CreateRectRgn(Width - 7, 0, Width, 1)
    Rgn2 = CreateRectRgn(Width - 5, 1, Width, 2)
    Rgn3 = CreateRectRgn(Width - 3, 2, Width, 3)
    Rgn4 = CreateRectRgn(Width - 2, 3, Width, 4)
    Rgn5 = CreateRectRgn(Width - 2, 4, Width, 5)
    Rgn6 = CreateRectRgn(Width - 1, 5, Width, 6)
    Rgn7 = CreateRectRgn(Width - 1, 6, Width, 7)
    CombineRgn Rgn, Rgn, Rgn1, RGN_DIFF
    CombineRgn Rgn, Rgn, Rgn2, RGN_DIFF
    CombineRgn Rgn, Rgn, Rgn3, RGN_DIFF
    CombineRgn Rgn, Rgn, Rgn4, RGN_DIFF
    CombineRgn Rgn, Rgn, Rgn5, RGN_DIFF
    CombineRgn Rgn, Rgn, Rgn6, RGN_DIFF
    CombineRgn Rgn, Rgn, Rgn7, RGN_DIFF
    DeleteObject Rgn1
    DeleteObject Rgn2
    DeleteObject Rgn3
    DeleteObject Rgn4
    DeleteObject Rgn5
    DeleteObject Rgn6
    DeleteObject Rgn7
    Rgn1 = CreateRectRgn(0, Height - 2, 1, Height - 1)
    Rgn2 = CreateRectRgn(0, Height - 1, 2, Height)
    CombineRgn Rgn, Rgn, Rgn1, RGN_DIFF
    CombineRgn Rgn, Rgn, Rgn2, RGN_DIFF
    DeleteObject Rgn1
    DeleteObject Rgn2
    Rgn1 = CreateRectRgn(Width - 1, Height - 2, Width, Height - 1)
    Rgn2 = CreateRectRgn(Width - 2, Height - 1, Width, Height)
    CombineRgn Rgn, Rgn, Rgn1, RGN_DIFF
    CombineRgn Rgn, Rgn, Rgn2, RGN_DIFF
    DeleteObject Rgn1
    DeleteObject Rgn2
    SetWindowRgn Me.hWnd, Rgn, True
    DeleteObject Rgn
    
    '프레임
    GetWindowRect pbTopLeft.hWnd, rc
    Width = rc.Right - rc.Left: Height = rc.Bottom - rc.Top
    Rgn = CreateRectRgn(0, 0, Width, Height)
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
    DeleteObject Rgn1
    DeleteObject Rgn2
    DeleteObject Rgn3
    DeleteObject Rgn4
    DeleteObject Rgn5
    DeleteObject Rgn6
    DeleteObject Rgn7
    SetWindowRgn pbTopLeft.hWnd, Rgn, True
    DeleteObject Rgn
    
    GetWindowRect pbTopRight.hWnd, rc
    Width = rc.Right - rc.Left: Height = rc.Bottom - rc.Top
    Rgn = CreateRectRgn(0, 0, Width, Height)
    Rgn1 = CreateRectRgn(Width - 7, 0, Width, 1)
    Rgn2 = CreateRectRgn(Width - 5, 1, Width, 2)
    Rgn3 = CreateRectRgn(Width - 3, 2, Width, 3)
    Rgn4 = CreateRectRgn(Width - 2, 3, Width, 4)
    Rgn5 = CreateRectRgn(Width - 2, 4, Width, 5)
    Rgn6 = CreateRectRgn(Width - 1, 5, Width, 6)
    Rgn7 = CreateRectRgn(Width - 1, 6, Width, 7)
    CombineRgn Rgn, Rgn, Rgn1, RGN_DIFF
    CombineRgn Rgn, Rgn, Rgn2, RGN_DIFF
    CombineRgn Rgn, Rgn, Rgn3, RGN_DIFF
    CombineRgn Rgn, Rgn, Rgn4, RGN_DIFF
    CombineRgn Rgn, Rgn, Rgn5, RGN_DIFF
    CombineRgn Rgn, Rgn, Rgn6, RGN_DIFF
    CombineRgn Rgn, Rgn, Rgn7, RGN_DIFF
    DeleteObject Rgn1
    DeleteObject Rgn2
    DeleteObject Rgn3
    DeleteObject Rgn4
    DeleteObject Rgn5
    DeleteObject Rgn6
    DeleteObject Rgn7
    SetWindowRgn pbTopRight.hWnd, Rgn, True
    DeleteObject Rgn
    
    GetWindowRect pbBottomLeft.hWnd, rc
    Width = rc.Right - rc.Left: Height = rc.Bottom - rc.Top
    Rgn = CreateRectRgn(0, 0, Width, Height)
    Rgn1 = CreateRectRgn(0, Height - 2, 1, Height - 1)
    Rgn2 = CreateRectRgn(0, Height - 1, 2, Height)
    CombineRgn Rgn, Rgn, Rgn1, RGN_DIFF
    CombineRgn Rgn, Rgn, Rgn2, RGN_DIFF
    DeleteObject Rgn1
    DeleteObject Rgn2
    SetWindowRgn pbBottomLeft.hWnd, Rgn, True
    DeleteObject Rgn
    
    GetWindowRect pbBottomRight.hWnd, rc
    Width = rc.Right - rc.Left: Height = rc.Bottom - rc.Top
    Rgn = CreateRectRgn(0, 0, Width, Height)
    Rgn1 = CreateRectRgn(Width - 1, Height - 2, Width, Height - 1)
    Rgn2 = CreateRectRgn(Width - 2, Height - 1, Width, Height)
    CombineRgn Rgn, Rgn, Rgn1, RGN_DIFF
    CombineRgn Rgn, Rgn, Rgn2, RGN_DIFF
    DeleteObject Rgn1
    DeleteObject Rgn2
    SetWindowRgn pbBottomRight.hWnd, Rgn, True
    DeleteObject Rgn
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unhook_Bluemetal Me.hWnd
'    SetParent pbTopLeft.hWnd, Me.hWnd
'    SetParent pbTopMiddle.hWnd, Me.hWnd
'    SetParent pbTopRight.hWnd, Me.hWnd
End Sub

Private Sub imgCloseButton_Click()
    Unload Me
End Sub

Private Sub imgCloseButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    Set imgCloseButton.Picture = imgClose(3).Picture
    imgCloseButton.Tag = "down"
End Sub

Private Sub imgCloseButton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imgCloseButton.Tag = "down" Then Exit Sub
    Set imgCloseButton.Picture = imgClose(2).Picture
    timCloseHover.Enabled = True
End Sub

Private Sub imgCloseButton_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    Set imgCloseButton.Picture = imgClose(IsWindowActive).Picture
    imgCloseButton.Tag = ""
End Sub

Private Sub imgControlMenu_DblClick()
    Unload Me
End Sub

Private Sub imgControlMenu_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ShowControlMenu
End Sub

Private Sub ToggleMaximized()
    If Not Me.MaxButton Then Exit Sub
    If Me.WindowState = 2 Then Me.WindowState = 0 Else Me.WindowState = 2
    SetResizeCursors
    SetMaxButtonTexture
End Sub

Private Sub imgMaximizeButton_Click()
    ToggleMaximized
End Sub

Private Sub imgMaximizeButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    Set imgMaximizeButton.Picture = imgMaximize(3 - (Me.WindowState = 2) * 4).Picture
    imgMaximizeButton.Tag = "down"
End Sub

Private Sub imgMaximizeButton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imgMaximizeButton.Tag = "down" Then Exit Sub
    Set imgMaximizeButton.Picture = imgMaximize(2 - (Me.WindowState = 2) * 4).Picture
    timMaximizeHover.Enabled = True
End Sub

Private Sub imgMaximizeButton_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    Set imgMaximizeButton.Picture = imgMaximize(IsWindowActive - (Me.WindowState = 2) * 4).Picture
    imgMaximizeButton.Tag = ""
End Sub

Private Sub imgMinimizeButton_Click()
    Me.WindowState = 1
End Sub

Private Sub imgMinimizeButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    Set imgMinimizeButton.Picture = imgMinimize(3).Picture
    imgMinimizeButton.Tag = "down"
End Sub

Private Sub imgMinimizeButton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imgMaximizeButton.Tag = "down" Then Exit Sub
    Set imgMinimizeButton.Picture = imgMinimize(2).Picture
    timMinimizeHover.Enabled = True
End Sub

Private Sub imgMinimizeButton_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    Set imgMinimizeButton.Picture = imgMinimize(IsWindowActive).Picture
    imgMinimizeButton.Tag = ""
End Sub

Private Sub lblCaption_DblClick()
    ToggleMaximized
End Sub

Private Sub lblCaption_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then ShowControlMenu
End Sub

Private Sub lblCaptionShadow_DblClick()
    ToggleMaximized
End Sub

Private Sub lblCaptionShadow_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then ShowControlMenu
End Sub

Private Sub lblResizeLeft_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        ReleaseCapture
        SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTLEFT, 0&
    End If
End Sub

Private Sub lblResizeRight_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        ReleaseCapture
        SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTRIGHT, 0&
    End If
End Sub

Private Sub lblResizeTop_DblClick(Index As Integer)
    If Me.WindowState = 2 Then ToggleMaximized
End Sub

Private Sub lblResizeTop_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        ReleaseCapture
        SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTTOP, 0&
    End If
End Sub

Private Sub lblResizeTopLeft_DblClick()
    If Me.WindowState = 2 Or Me.MaxButton = False Then Unload Me
End Sub

Private Sub lblResizeTopLeft_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Me.WindowState = 2 Or Me.MaxButton = False Then ShowControlMenu
End Sub

Private Sub lblResizeTopLeft_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        ReleaseCapture
        SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTTOPLEFT, 0&
    End If
End Sub

Private Sub lblResizeTopRight_DblClick()
    If Me.WindowState = 2 Then ToggleMaximized
End Sub

Private Sub lblResizeTopRight_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        ReleaseCapture
        SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTTOPRIGHT, 0&
    End If
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

Private Sub ShowControlMenu()
    Dim hCur As Long
    Dim pt As POINTAPI
    Dim hMenu As Long
    Dim cmd As Long
    
    GetCursorPos pt
    hMenu = GetSystemMenu(Me.hWnd, 0)
    cmd = TrackPopupMenu(hMenu, TPM_LEFTALIGN Or TPM_RETURNCMD, pt.X + 1, pt.Y + 1, 0, Me.hWnd, ByVal 0&)
    If cmd <> 0 Then SendMessage Me.hWnd, WM_SYSCOMMAND, cmd, 0&
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

Private Sub pbTopLeft_DblClick()
    Unload Me
End Sub

Private Sub pbTopLeft_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ShowControlMenu
End Sub

Private Sub pbTopLeft_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    pbTopMiddle_MouseMove Button, Shift, X, Y
End Sub

Private Sub pbTopLeft_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then ShowControlMenu
End Sub

Private Sub pbTopMiddle_DblClick()
    ToggleMaximized
End Sub

Private Sub pbTopMiddle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        ReleaseCapture
        SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    End If
End Sub

Private Sub pbTopMiddle_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then ShowControlMenu
End Sub

Private Sub pbTopRight_DblClick()
    ToggleMaximized
End Sub

Private Sub pbTopRight_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X >= imgCloseButton.Left And X <= imgCloseButton.Left + imgCloseButton.Width And Y >= imgCloseButton.Top And Y <= imgCloseButton.Top + imgCloseButton.Height Then
        Exit Sub
    End If
    pbTopMiddle_MouseMove Button, Shift, X, Y
End Sub

Private Sub pbTopRight_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then ShowControlMenu
End Sub

Private Sub timCloseHover_Timer()
    If (Not IsMouseOn(imgCloseButton)) And imgCloseButton.Tag <> "down" Then
        Set imgCloseButton.Picture = imgClose(IsWindowActive).Picture
        timCloseHover.Enabled = False
    End If
End Sub

Private Function IsMouseOn(imgImage As Image) As Boolean
    Dim pt As POINTAPI
    Dim rectLeft As Single, rectTop As Single, rectRight As Single, rectBottom As Single
    
    GetCursorPos pt
    ScreenToClient pbTopRight.hWnd, pt
    pt.X = pt.X * Screen.TwipsPerPixelX
    pt.Y = pt.Y * Screen.TwipsPerPixelY
    
    rectLeft = imgImage.Left
    rectTop = imgImage.Top
    rectRight = imgImage.Left + imgImage.Width
    rectBottom = imgImage.Top + imgImage.Height
    
    IsMouseOn = (pt.X >= rectLeft And pt.X <= rectRight And pt.Y >= rectTop And pt.Y <= rectBottom)
End Function

Private Sub timMaximizeHover_Timer()
    If (Not IsMouseOn(imgMaximizeButton)) And imgMaximizeButton.Tag <> "down" Then
        Set imgMaximizeButton.Picture = imgMaximize(IsWindowActive - (Me.WindowState = 2) * 4).Picture
        timMaximizeHover.Enabled = False
    End If
End Sub

Private Sub timMinimizeHover_Timer()
    If (Not IsMouseOn(imgMinimizeButton)) And imgMinimizeButton.Tag <> "down" Then
        Set imgMinimizeButton.Picture = imgMinimize(IsWindowActive).Picture
        timMinimizeHover.Enabled = False
    End If
End Sub
