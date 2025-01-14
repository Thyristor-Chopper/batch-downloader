VERSION 5.00
Begin VB.Form OKMsgBox 
   BackColor       =   &H00F8EFE5&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "메시지 상자"
   ClientHeight    =   1335
   ClientLeft      =   45
   ClientTop       =   555
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
   Icon            =   "OKMsgBox.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   28440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.Timer timeout 
      Enabled         =   0   'False
      Left            =   480
      Top             =   840
   End
   Begin prjDownloadBooster.CommandButtonW cmdOK 
      Cancel          =   -1  'True
      Default         =   -1  'True
      Height          =   320
      Left            =   2760
      TabIndex        =   0
      Top             =   840
      Width           =   1455
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "확인"
   End
   Begin VB.Image imgMBIconQuestion 
      Height          =   480
      Left            =   240
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgMBIconError 
      Height          =   480
      Left            =   240
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgMBIconWarning 
      Height          =   480
      Left            =   240
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblContent 
      BackStyle       =   0  '투명
      Caption         =   "내용"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   1
      Top             =   360
      Width           =   27255
   End
   Begin VB.Image imgMBIconInfo 
      Height          =   480
      Left            =   240
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "OKMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    If GetSetting("DownloadBooster", "Options", "DisableDWMWindow", DefaultDisableDWMWindow) = 1 Then DisableDWMWindow Me.hWnd
    SetFormBackgroundColor Me
    SetFont Me
    
    imgMBIconQuestion.Picture = YesNoCancelMsgBox.imgMBIconQuestion.Picture
    imgMBIconError.Picture = YesNoCancelMsgBox.imgMBIconError.Picture
    imgMBIconWarning.Picture = YesNoCancelMsgBox.imgMBIconWarning.Picture
    imgMBIconInfo.Picture = YesNoCancelMsgBox.imgMBIconInfo.Picture
End Sub

Private Sub timeout_Timer()
    cmdOK_Click
End Sub
