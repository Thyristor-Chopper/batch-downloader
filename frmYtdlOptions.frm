VERSION 5.00
Begin VB.Form frmYtdlOptions 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "youtube-dl 옵션"
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6030
   BeginProperty Font 
      Name            =   "굴림"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmYtdlOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows 기본값
   Begin prjDownloadBooster.CheckBoxW chkAutoYtdl 
      Height          =   255
      Left            =   600
      TabIndex        =   12
      Top             =   120
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   450
      Caption         =   "youtube-dl 사용 여부 자동 결정(&T)"
   End
   Begin prjDownloadBooster.OptionButtonW optDisableYtdl 
      Height          =   255
      Left            =   600
      TabIndex        =   11
      Top             =   480
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   450
      Value           =   -1  'True
      Caption         =   "youtube-dl 사용 안 함(&D)"
   End
   Begin prjDownloadBooster.OptionButtonW optUseYtdl 
      Height          =   255
      Left            =   600
      TabIndex        =   10
      Top             =   840
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   450
      Caption         =   "youtube-dl 사용(&U)"
   End
   Begin prjDownloadBooster.FrameW fYtdl 
      Height          =   2175
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   3836
      Caption         =   "        "
      Begin prjDownloadBooster.ComboBoxW txtFormat 
         Height          =   300
         Left            =   1200
         TabIndex        =   1
         Top             =   360
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   529
         Text            =   "ComboBoxW1"
      End
      Begin prjDownloadBooster.ComboBoxW cbBitRate 
         Height          =   300
         Left            =   1560
         TabIndex        =   2
         Top             =   1800
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         Style           =   2
         Text            =   "ComboBoxW1"
      End
      Begin prjDownloadBooster.OptionButtonW optCBR 
         Height          =   255
         Left            =   840
         TabIndex        =   3
         Top             =   1800
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
         Caption         =   "&CBR:"
      End
      Begin prjDownloadBooster.ComboBoxW cbAudioFormat 
         Height          =   300
         Left            =   2040
         TabIndex        =   4
         Top             =   1080
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   529
         Style           =   2
         Text            =   "ComboBoxW2"
      End
      Begin prjDownloadBooster.ComboBoxW cbVBR 
         Height          =   300
         Left            =   1560
         TabIndex        =   5
         Top             =   1440
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         Style           =   2
         Text            =   "ComboBoxW1"
      End
      Begin prjDownloadBooster.OptionButtonW optVBR 
         Height          =   255
         Left            =   840
         TabIndex        =   6
         Top             =   1440
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
         Value           =   -1  'True
         Caption         =   "&VBR:"
      End
      Begin prjDownloadBooster.CheckBoxW chkExtractAudio 
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   720
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   450
         Caption         =   "음원만 추출(&E)"
      End
      Begin VB.Label Label3 
         BackStyle       =   0  '투명
         Caption         =   "포맷(&F):"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   390
         Width           =   855
      End
      Begin VB.Label Label4 
         BackStyle       =   0  '투명
         Caption         =   "오디오 형식(&A):"
         Height          =   255
         Left            =   600
         TabIndex        =   8
         Top             =   1125
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmYtdlOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cbAudioFormat_Click()
    optUseYtdl_Click
End Sub

Private Sub chkAutoYtdl_Click()
    optDisableYtdl.Enabled = (chkAutoYtdl.Value = 0)
    optUseYtdl.Enabled = optDisableYtdl.Enabled
    optUseYtdl_Click
End Sub

Private Sub chkExtractAudio_Click()
    optUseYtdl_Click
End Sub

Private Sub optDisableYtdl_Click()
    optUseYtdl_Click
End Sub

Private Sub optUseYtdl_Click()
    Dim ctrl As Control
    On Error Resume Next
    For Each ctrl In fYtdl.ContainedControls
        ctrl.Enabled = (optUseYtdl.Value Or chkAutoYtdl.Value = 1)
    Next ctrl
    
    If optUseYtdl.Value Or chkAutoYtdl.Value = 1 Then
        Label4.Enabled = (chkExtractAudio.Value = 1)
        cbAudioFormat.Enabled = (chkExtractAudio.Value = 1)
        optVBR.Enabled = (chkExtractAudio.Value = 1) And cbAudioFormat.ListIndex = 1
        optCBR.Enabled = (chkExtractAudio.Value = 1) And cbAudioFormat.ListIndex = 1
        cbVBR.Enabled = (chkExtractAudio.Value = 1) And cbAudioFormat.ListIndex = 1
        cbBitRate.Enabled = (chkExtractAudio.Value = 1) And cbAudioFormat.ListIndex = 1
        If chkExtractAudio.Value = 1 And cbAudioFormat.ListIndex = 1 Then
            cbVBR.Enabled = optVBR.Value
            cbBitRate.Enabled = optCBR.Value
        End If
    End If
End Sub

Private Sub optCBR_Click()
    optUseYtdl_Click
End Sub

Private Sub optVBR_Click()
    optUseYtdl_Click
End Sub

Private Sub Form_Load()
    If GetSetting("DownloadBooster", "Options", "DisableDWMWindow", DefaultDisableDWMWindow) = 1 Then DisableDWMWindow Me.hWnd
    SetFormBackgroundColor Me
    SetFont Me
    SetWindowPos Me.hWnd, IIf(MainFormOnTop, HWND_TOPMOST, HWND_NOTOPMOST), 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    
    optUseYtdl.Caption = t(optUseYtdl.Caption, "&Use youtube-dl")
    optUseYtdl.Width = t(optUseYtdl.Width, 1455)
    
    cbAudioFormat.AddItem t("자동", "Auto") & " (M4A/OPUS)"
    cbAudioFormat.AddItem "MP3"
    cbAudioFormat.AddItem "WAV"
    cbAudioFormat.AddItem "FLAC"
    cbAudioFormat.ListIndex = 0
    
    Dim i%
    For i = 0 To 9
         cbVBR.AddItem i
    Next i
    cbVBR.ListIndex = 0
    
    cbBitRate.AddItem "8 kbps"
    cbBitRate.AddItem "16 kbps"
    cbBitRate.AddItem "24 kbps"
    cbBitRate.AddItem "32 kbps"
    cbBitRate.AddItem "40 kbps"
    cbBitRate.AddItem "48 kbps"
    cbBitRate.AddItem "56 kbps"
    cbBitRate.AddItem "64 kbps"
    cbBitRate.AddItem "80 kbps"
    cbBitRate.AddItem "96 kbps"
    cbBitRate.AddItem "112 kbps"
    cbBitRate.AddItem "128 kbps"
    cbBitRate.AddItem "144 kbps"
    cbBitRate.AddItem "160 kbps"
    cbBitRate.AddItem "192 kbps"
    cbBitRate.AddItem "224 kbps"
    cbBitRate.AddItem "256 kbps"
    cbBitRate.AddItem "320 kbps"
    cbBitRate.ListIndex = 14
    
    optUseYtdl_Click
End Sub
