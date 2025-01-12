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
   Begin prjDownloadBooster.CommandButtonW cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "취소"
      Default         =   -1  'True
      Height          =   320
      Left            =   5880
      TabIndex        =   3
      Top             =   840
      Width           =   1455
   End
   Begin prjDownloadBooster.CommandButtonW cmdNo 
      Caption         =   "아니요(&N)"
      Height          =   320
      Left            =   4320
      TabIndex        =   2
      Top             =   840
      Width           =   1455
   End
   Begin prjDownloadBooster.CommandButtonW cmdYes 
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
      Picture         =   "YesNoCancelMsgBox.frx":000C
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgMBIconError 
      Height          =   480
      Left            =   240
      Picture         =   "YesNoCancelMsgBox.frx":044E
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgMBIconWarning 
      Height          =   480
      Left            =   240
      Picture         =   "YesNoCancelMsgBox.frx":0890
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
      TabIndex        =   1
      Top             =   360
      Width           =   27255
   End
   Begin VB.Image imgMBIconInfo 
      Height          =   480
      Left            =   240
      Picture         =   "YesNoCancelMsgBox.frx":0CD2
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "YesNoCancelMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MSET As Boolean

Private Sub cmdNo_Click()
    Functions.YesNoCancelMsgBoxResult = vbNo
    MSET = -1
    Unload Me
End Sub

Private Sub cmdYes_Click()
    Functions.YesNoCancelMsgBoxResult = vbYes
    MSET = -1
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Functions.YesNoCancelMsgBoxResult = vbCancel
    MSET = -1
    Unload Me
End Sub

Private Sub Form_Load()
    If GetSetting("DownloadBooster", "Options", "DisableDWMWindow", DefaultDisableDWMWindow) = 1 Then DisableDWMWindow Me.hWnd
    SetFormBackgroundColor Me
    SetFont Me
    MSET = 0
    
    cmdYes.Caption = t("예(&Y)", "&Yes")
    cmdNo.Caption = t("아니요(&N)", "&No")
    cmdCancel.Caption = t("취소", "Cancel")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not MSET Then Functions.YesNoCancelMsgBoxResult = vbCancel
End Sub
