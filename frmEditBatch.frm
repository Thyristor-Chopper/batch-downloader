VERSION 5.00
Begin VB.Form frmEditBatch 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "편집"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5550
   BeginProperty Font 
      Name            =   "굴림"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEditBatch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin prjDownloadBooster.CheckBoxW chkUseYtdl 
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   1440
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   450
      Caption         =   "youtube-dl 사용(&U)"
   End
   Begin prjDownloadBooster.FrameW fYtdl 
      Height          =   2175
      Left            =   120
      TabIndex        =   10
      Top             =   1440
      Visible         =   0   'False
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   3836
      Caption         =   "        "
      Begin prjDownloadBooster.ComboBoxW txtFormat 
         Height          =   300
         Left            =   1200
         TabIndex        =   20
         Top             =   360
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   529
         Text            =   "ComboBoxW1"
      End
      Begin prjDownloadBooster.ComboBoxW cbBitRate 
         Height          =   300
         Left            =   1560
         TabIndex        =   19
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
         TabIndex        =   18
         Top             =   1800
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
         Caption         =   "&CBR:"
      End
      Begin prjDownloadBooster.ComboBoxW cbAudioFormat 
         Height          =   300
         Left            =   2040
         TabIndex        =   17
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
         Top             =   720
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   450
         Caption         =   "음원만 추출(&E)"
      End
      Begin VB.Label Label4 
         BackStyle       =   0  '투명
         Caption         =   "오디오 형식(&A):"
         Height          =   255
         Left            =   600
         TabIndex        =   16
         Top             =   1125
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackStyle       =   0  '투명
         Caption         =   "포맷(&F):"
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   390
         Width           =   855
      End
   End
   Begin prjDownloadBooster.TygemButton tygCancel 
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Top             =   3720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "취소"
      BackColor       =   0
      FontSize        =   0
   End
   Begin prjDownloadBooster.TygemButton tygOK 
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   3720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "확인"
      BackColor       =   0
      FontSize        =   0
   End
   Begin prjDownloadBooster.CommandButtonW cmdOK 
      Default         =   -1  'True
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   3720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "확인"
   End
   Begin prjDownloadBooster.CommandButtonW cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   3720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "취소"
   End
   Begin prjDownloadBooster.TygemButton tygBrowse 
      Height          =   300
      Left            =   3840
      TabIndex        =   4
      Top             =   960
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      Caption         =   "찾아보기..."
      BackColor       =   0
      FontSize        =   0
   End
   Begin prjDownloadBooster.CommandButtonW cmdBrowse 
      Height          =   300
      Left            =   3840
      TabIndex        =   5
      Top             =   960
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      Caption         =   "찾아보기(&B)..."
   End
   Begin prjDownloadBooster.TextBoxW txtFilePath 
      Height          =   300
      Left            =   480
      TabIndex        =   6
      Top             =   960
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   529
   End
   Begin prjDownloadBooster.TextBoxW txtURL 
      Height          =   300
      Left            =   480
      TabIndex        =   7
      Top             =   360
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   529
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      Caption         =   "파일 주소(&A):"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label Label2 
      Caption         =   "저장 경로(&S):"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   720
      Width           =   3495
   End
End
Attribute VB_Name = "frmEditBatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public OriginalURL As String
Public OriginalPath As String

Private Sub cbAudioFormat_Click()
    chkUseYtdl_Click
End Sub

Private Sub chkExtractAudio_Click()
    chkUseYtdl_Click
End Sub

Private Sub chkUseYtdl_Click()
    Dim ctrl As Control
    On Error Resume Next
    For Each ctrl In fYtdl.ContainedControls
        ctrl.Enabled = (chkUseYtdl.Value = 1)
    Next ctrl
    
    If chkUseYtdl.Value = 1 Then
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

Private Sub cmdBrowse_Click()
    Tags.BrowsePresetPath = Trim$(txtFilePath.Text)
    Tags.BrowseTargetForm = 1
    
    If GetSetting("DownloadBooster", "Options", "ForceWin31Dialog", "0") = "1" Then
        Unload frmBrowse
        frmBrowse.Show vbModal, Me
    Else
        Unload frmExplorer
        frmExplorer.Show vbModal, Me
    End If
    
    If FolderExists(txtFilePath.Text) Then
        If Right$(txtFilePath.Text, 1) <> "\" Then txtFilePath.Text = txtFilePath.Text & "\"
        txtFilePath.Text = txtFilePath.Text & Tags.FileNameOnly
    End If
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    txtURL.Text = Trim$(txtURL.Text)
    If Left$(txtURL.Text, 7) <> "http://" And Left$(txtURL.Text, 8) <> "https://" Then
        Alert t("주소가 올바르지 않습니다. 'http://' 또는 'https://'로 시작해야 합니다.", "Invalid address. Must start with 'http://' or 'https://'."), App.Title, Me, 16
        Exit Sub
    End If

    txtFilePath.Text = Trim$(txtFilePath.Text)
    Do While Replace(txtFilePath.Text, "\\", "\") <> txtFilePath.Text
        txtFilePath.Text = Replace(txtFilePath.Text, "\\", "\")
    Loop
    
    If FolderExists(txtFilePath.Text) Then
        If Right$(txtFilePath.Text, 1) <> "\" Then txtFilePath.Text = txtFilePath.Text & "\"
        txtFilePath.Text = txtFilePath.Text & Trim$(Tags.FileNameOnly)
    ElseIf Right$(txtFilePath.Text, 1) = "\" Or (Not FolderExists(GetParentFolderName(txtFilePath.Text))) Then
        Alert t("저장 경로가 존재하지 않습니다. [찾아보기] 기능으로 폴더를 찾아볼 수 있습니다.", "Save path does not exist. Use Broewse to browse folders."), App.Title, Me, 16
        Exit Sub
    End If
    txtFilePath.Text = FilterFilename(txtFilePath.Text, True)
    
    On Error Resume Next
    Dim ParentFolderName As String
    ParentFolderName = GetParentFolderName(txtFilePath.Text)
    If Right$(ParentFolderName, 1) = "\" Then ParentFolderName = Left$(ParentFolderName, Len(ParentFolderName) - 1)
    frmMain.lvBatchFiles.SelectedItem.ListSubItems(2).Text = txtURL.Text
    If frmMain.lvBatchFiles.SelectedItem.ListSubItems(3).Text = t("완료", "Done") Then
        If txtURL.Text <> Trim$(OriginalURL) Then
            frmMain.lvBatchFiles.SelectedItem.ListSubItems(3).Text = t("대기", "Queued")
            frmMain.lvBatchFiles.SelectedItem.Checked = True
            frmMain.lvBatchFiles.SelectedItem.ForeColor = &H80000008
            frmMain.lvBatchFiles.SelectedItem.ListSubItems(1).ForeColor = &H80000008
            frmMain.lvBatchFiles.SelectedItem.ListSubItems(2).ForeColor = &H80000008
            frmMain.lvBatchFiles.SelectedItem.ListSubItems(3).ForeColor = &H80000008
            GoTo changeFilepath
        ElseIf txtFilePath.Text <> Trim$(OriginalPath) Then
            Alert t("다운로드가 이미 완료된 파일의 저장 경로가 수정되지 않았습니다.", "Save path has not been changed because it's already downloaded."), App.Title, Me
        End If
    Else
changeFilepath:
        frmMain.lvBatchFiles.SelectedItem.Text = Replace(txtFilePath.Text, ParentFolderName & "\", "", 1, 1)
        frmMain.lvBatchFiles.SelectedItem.ListSubItems(1).Text = txtFilePath.Text
        If txtFilePath.Text <> Trim$(OriginalPath) Then
            frmMain.lvBatchFiles.SelectedItem.ListSubItems(4).Text = "N"
        End If
    End If
    
    Unload Me
End Sub

Private Sub Form_Load()
    If GetSetting("DownloadBooster", "Options", "DisableDWMWindow", DefaultDisableDWMWindow) = 1 Then DisableDWMWindow Me.hWnd
    SetFormBackgroundColor Me
    SetFont Me
    SetWindowPos Me.hWnd, IIf(MainFormOnTop, HWND_TOPMOST, HWND_NOTOPMOST), 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    
    cmdOK.Caption = t("확인", "OK")
    cmdCancel.Caption = t("취소", "Cancel")
    tygOK.Caption = cmdOK.Caption
    tygCancel.Caption = cmdCancel.Caption
    cmdBrowse.Caption = t(cmdBrowse.Caption, "&Browse...")
    tygBrowse.Caption = t("찾아보기...", "Browse...")
    Label1.Caption = t(Label1.Caption, "File &address:")
    Label2.Caption = t(Label2.Caption, "&Save to:")
    Me.Caption = t(Me.Caption, "Edit")
    chkUseYtdl.Caption = t(chkUseYtdl.Caption, "&Use youtube-dl")
    chkUseYtdl.Width = t(chkUseYtdl.Width, 1455)
    
    On Error Resume Next
    Me.Icon = frmMain.imgEdit.ListImages(1).Picture
    On Error GoTo 0
    
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
    
    chkUseYtdl_Click
End Sub

Private Sub optCBR_Click()
    chkUseYtdl_Click
End Sub

Private Sub optVBR_Click()
    chkUseYtdl_Click
End Sub

Private Sub tygBrowse_Click()
    cmdBrowse_Click
End Sub

Private Sub tygCancel_Click()
    cmdCancel_Click
End Sub

Private Sub tygOK_Click()
    cmdOK_Click
End Sub
