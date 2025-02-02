VERSION 5.00
Begin VB.Form YesNoCancelMsgBox 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "메시지 상자"
   ClientHeight    =   1365
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   28440
   BeginProperty Font 
      Name            =   "굴림"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "YesNoCancelMsgBox.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   28440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin prjDownloadBooster.TygemButton tygOK 
      Height          =   315
      Left            =   7440
      TabIndex        =   10
      Tag             =   "novisibilitychange"
      Top             =   840
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      Caption         =   "확인"
   End
   Begin prjDownloadBooster.TygemButton tygCancel 
      Height          =   315
      Left            =   5880
      TabIndex        =   9
      Tag             =   "novisibilitychange"
      Top             =   840
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      Caption         =   "취소"
   End
   Begin prjDownloadBooster.TygemButton tygNo 
      Height          =   320
      Left            =   4320
      TabIndex        =   8
      Tag             =   "novisibilitychange"
      Top             =   840
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      Caption         =   "아니요"
   End
   Begin prjDownloadBooster.TygemButton tygYes 
      Height          =   315
      Left            =   2760
      TabIndex        =   7
      Tag             =   "novisibilitychange"
      Top             =   840
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      Caption         =   "예"
   End
   Begin VB.Timer timeout 
      Enabled         =   0   'False
      Left            =   360
      Top             =   960
   End
   Begin prjDownloadBooster.CommandButtonW cmdCancel 
      Cancel          =   -1  'True
      Height          =   320
      Left            =   5880
      TabIndex        =   5
      Top             =   840
      Width           =   1455
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "취소"
   End
   Begin prjDownloadBooster.CommandButtonW cmdNo 
      Height          =   320
      Left            =   4320
      TabIndex        =   4
      Top             =   840
      Width           =   1455
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "아니요(&N)"
   End
   Begin prjDownloadBooster.CommandButtonW cmdYes 
      Height          =   320
      Left            =   2760
      TabIndex        =   3
      Top             =   840
      Width           =   1455
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "예(&Y)"
   End
   Begin prjDownloadBooster.OptionButtonW optNo 
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   1320
      Width           =   1575
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "아니요(&N)"
   End
   Begin prjDownloadBooster.OptionButtonW optYes 
      Height          =   255
      Left            =   1080
      TabIndex        =   0
      Top             =   960
      Width           =   1575
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "예(&Y)"
   End
   Begin prjDownloadBooster.CommandButtonW cmdOK 
      Default         =   -1  'True
      Height          =   315
      Left            =   7440
      TabIndex        =   2
      Top             =   840
      Width           =   1455
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "확인"
   End
   Begin VB.Image imgMBIconInfo 
      Height          =   480
      Left            =   240
      Picture         =   "YesNoCancelMsgBox.frx":000C
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgMBIconWarning 
      Height          =   480
      Left            =   240
      Picture         =   "YesNoCancelMsgBox.frx":044E
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgMBIconError 
      Height          =   480
      Left            =   240
      Picture         =   "YesNoCancelMsgBox.frx":0890
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgMBIconQuestion 
      Height          =   480
      Left            =   240
      Picture         =   "YesNoCancelMsgBox.frx":0CD2
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblContent 
      BackColor       =   &H00F8EFE5&
      BackStyle       =   0  '투명
      Caption         =   "내용"
      Height          =   495
      Left            =   960
      TabIndex        =   6
      Top             =   360
      Width           =   27255
   End
End
Attribute VB_Name = "YesNoCancelMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MSET As Boolean
Public Mode As Byte

Private Sub cmdNo_Click()
    Functions.MsgBoxResult = vbNo
    MSET = -1
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If optYes.Visible Then
        If optYes.Value = True Then
            Functions.MsgBoxResult = vbYes
        Else
            Functions.MsgBoxResult = vbNo
        End If
    Else
        Functions.MsgBoxResult = vbOK
    End If
    MSET = -1
    Unload Me
End Sub

Private Sub cmdYes_Click()
    Functions.MsgBoxResult = vbYes
    MSET = -1
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Functions.MsgBoxResult = vbCancel
    MSET = -1
    Unload Me
End Sub

Private Sub Form_Load()
    If GetSetting("DownloadBooster", "Options", "DisableDWMWindow", DefaultDisableDWMWindow) = 1 Then DisableDWMWindow Me.hWnd
    SetFormBackgroundColor Me
    SetFont Me
    SetWindowPos Me.hWnd, IIf(MainFormOnTop, HWND_TOPMOST, HWND_NOTOPMOST), 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    MSET = 0
    If Functions.MsgBoxMode = 2 Then
        Dim SystemMenu As Long
        SystemMenu = GetSystemMenu(Me.hWnd, 0)
        DeleteMenu SystemMenu, SC_CLOSE, MF_BYCOMMAND
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If cmdYes.Visible And cmdNo.Visible And (Not cmdCancel.Visible) And (Not MSET) Then
        Cancel = 1
        Exit Sub
    End If
    If Not MSET Then Functions.MsgBoxResult = vbCancel
    GetSystemMenu Me.hWnd, 1
End Sub

Private Sub timeout_Timer()
    cmdOK_Click
End Sub

Private Sub optNo_Click()
    cmdOK.Enabled = True
End Sub

Private Sub optYes_Click()
    cmdOK.Enabled = True
End Sub

Private Sub tygCancel_Click()
    cmdCancel_Click
End Sub

Private Sub tygNo_Click()
    cmdNo_Click
End Sub

Private Sub tygOK_Click()
    cmdOK_Click
End Sub

Private Sub tygYes_Click()
    cmdYes_Click
End Sub
