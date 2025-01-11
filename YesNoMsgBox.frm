VERSION 5.00
Begin VB.Form YesNoMsgBox 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "메시지 상자"
   ClientHeight    =   1335
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
   Icon            =   "YesNoMsgBox.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   28440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.CommandButton cmdCancel 
      Caption         =   "아니요(&N)"
      Default         =   -1  'True
      Height          =   320
      Left            =   4320
      TabIndex        =   2
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "예(&Y)"
      Height          =   320
      Left            =   2760
      TabIndex        =   0
      Top             =   840
      Width           =   1455
   End
   Begin VB.Image imgMBIconQuestion 
      Height          =   480
      Left            =   240
      Picture         =   "YesNoMsgBox.frx":000C
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgMBIconError 
      Height          =   480
      Left            =   240
      Picture         =   "YesNoMsgBox.frx":044E
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgMBIconWarning 
      Height          =   480
      Left            =   240
      Picture         =   "YesNoMsgBox.frx":0890
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblContent 
      BackColor       =   &H00F8EFE5&
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
      Picture         =   "YesNoMsgBox.frx":0CD2
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "YesNoMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ButtonPressed As Boolean

Private Sub cmdCancel_Click()
    Functions.YesNoMsgBoxResult = vbNo
    ButtonPressed = True
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Functions.YesNoMsgBoxResult = vbYes
    ButtonPressed = True
    Unload Me
End Sub

Private Sub Form_Load()
    Dim SystemMenu As Long
    SystemMenu = GetSystemMenu(Me.hWnd, 0)
    DeleteMenu SystemMenu, 6, MF_BYPOSITION
    ButtonPressed = False
    If GetSetting("DownloadBooster", "Options", "DisableDWMWindow", DefaultDisableDWMWindow) = 1 Then DisableDWMWindow Me.hWnd
    SetFormBackgroundColor Me
    SetFont Me
    
    cmdOK.Caption = t("예(&Y)", "&Yes")
    cmdCancel.Caption = t("아니요(&N)", "&No")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not ButtonPressed Then Cancel = 1
End Sub
