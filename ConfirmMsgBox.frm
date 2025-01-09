VERSION 5.00
Begin VB.Form ConfirmMsgBox 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "메시지 상자"
   ClientHeight    =   2040
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
   Icon            =   "ConfirmMsgBox.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   28440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.OptionButton optNo 
      Caption         =   "아니요(&N)"
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   1200
      Width           =   1575
   End
   Begin VB.OptionButton optYes 
      Caption         =   "예(&Y)"
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      Top             =   840
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "취소"
      Height          =   320
      Left            =   4320
      TabIndex        =   2
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "확인"
      Default         =   -1  'True
      Height          =   320
      Left            =   2760
      TabIndex        =   0
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Image imgMBIconQuestion 
      Height          =   480
      Left            =   240
      Picture         =   "ConfirmMsgBox.frx":0442
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgMBIconError 
      Height          =   480
      Left            =   240
      Picture         =   "ConfirmMsgBox.frx":0884
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgMBIconWarning 
      Height          =   480
      Left            =   240
      Picture         =   "ConfirmMsgBox.frx":0CC6
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
      Picture         =   "ConfirmMsgBox.frx":1108
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "ConfirmMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public BeepSnd As Long
Dim isOK As Integer

Private Sub cmdCancel_Click()
    Functions.ConfirmResult = vbCancel
    isOK = 0
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If optYes.Value = True Then
        Functions.ConfirmResult = vbYes
    Else
        Functions.ConfirmResult = vbNo
    End If
    isOK = 1
    Unload Me
End Sub

Private Sub Form_Load()
    If GetSetting("DownloadBooster", "Options", "DisableDWMWindow", 1) = 1 Then DisableDWMWindow Me.hWnd
    cmdOK.Caption = "확인"
    cmdCancel.Caption = "취소"
    isOK = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If isOK <> 1 Then
        Functions.ConfirmResult = vbCancel
        Unload Me
    End If
    Unload Me
End Sub

Private Sub optNo_Click()
    cmdOK.Enabled = True
End Sub

Private Sub optYes_Click()
    cmdOK.Enabled = True
End Sub
