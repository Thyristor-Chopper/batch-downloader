VERSION 5.00
Begin VB.Form frmDownloadOptions 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "다운로드 설정"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12060
   BeginProperty Font 
      Name            =   "굴림"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDownloadOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   12060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.PictureBox pbPanel 
      Height          =   4095
      Index           =   2
      Left            =   6360
      ScaleHeight     =   4035
      ScaleWidth      =   5955
      TabIndex        =   17
      Top             =   600
      Width           =   6015
      Begin prjDownloadBooster.LinkLabel lblDescription 
         Height          =   735
         Left            =   720
         TabIndex        =   18
         Top             =   180
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   1296
         Caption         =   "frmDownloadOptions.frx":000C
         Transparent     =   -1  'True
      End
      Begin prjDownloadBooster.CommandButtonW cmdEditHeaderName 
         Height          =   330
         Left            =   3360
         TabIndex        =   19
         Top             =   3660
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         Enabled         =   0   'False
         Caption         =   "이름변경(&R)"
      End
      Begin prjDownloadBooster.TextBoxW txtEdit 
         Height          =   255
         Left            =   2760
         TabIndex        =   20
         Top             =   960
         Visible         =   0   'False
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   450
         BorderStyle     =   1
      End
      Begin prjDownloadBooster.CommandButtonW cmdDeleteHeader 
         Height          =   330
         Left            =   2040
         TabIndex        =   21
         Top             =   3660
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         Enabled         =   0   'False
         Caption         =   "삭제(&D)"
      End
      Begin prjDownloadBooster.CommandButtonW cmdEditHeaderValue 
         Height          =   330
         Left            =   4680
         TabIndex        =   22
         Top             =   3660
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         Enabled         =   0   'False
         Caption         =   "편집(&E)"
      End
      Begin prjDownloadBooster.CommandButtonW cmdAddHeader 
         Height          =   330
         Left            =   720
         TabIndex        =   23
         Top             =   3660
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         Caption         =   "추가(&A)"
      End
      Begin prjDownloadBooster.ListView lvHeaders 
         Height          =   2655
         Left            =   720
         TabIndex        =   24
         Top             =   960
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   4683
         VisualTheme     =   1
         Icons           =   "imgFiles"
         SmallIcons      =   "imgFiles"
         View            =   3
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HideSelection   =   0   'False
         ShowLabelTips   =   -1  'True
         HighlightColumnHeaders=   -1  'True
         AutoSelectFirstItem=   0   'False
      End
      Begin prjDownloadBooster.ImageList imgFiles 
         Left            =   120
         Top             =   2160
         _ExtentX        =   1005
         _ExtentY        =   1005
         ImageWidth      =   16
         ImageHeight     =   16
         ColorDepth      =   4
         MaskColor       =   16711935
         InitListImages  =   "frmDownloadOptions.frx":00E8
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "frmDownloadOptions.frx":0290
         Top             =   120
         Width           =   480
      End
   End
   Begin VB.PictureBox pbPanel 
      AutoRedraw      =   -1  'True
      Height          =   3135
      Index           =   1
      Left            =   120
      ScaleHeight     =   3075
      ScaleWidth      =   5955
      TabIndex        =   0
      Top             =   600
      Width           =   6015
      Begin prjDownloadBooster.CheckBoxW chkAutoYtdl 
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   120
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   450
         Caption         =   "youtube-dl 사용 여부 자동 결정(&T)"
         Transparent     =   -1  'True
      End
      Begin prjDownloadBooster.OptionButtonW optDisableYtdl 
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   480
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   450
         Value           =   -1  'True
         Caption         =   "youtube-dl 사용 안 함(&D)"
         Transparent     =   -1  'True
      End
      Begin prjDownloadBooster.OptionButtonW optUseYtdl 
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   840
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   450
         Caption         =   "youtube-dl 사용(&U)"
         Transparent     =   -1  'True
      End
      Begin prjDownloadBooster.FrameW fYtdl 
         Height          =   2175
         Left            =   120
         TabIndex        =   4
         Top             =   870
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   3836
         Caption         =   "        "
         Transparent     =   -1  'True
         Begin prjDownloadBooster.ComboBoxW txtFormat 
            Height          =   300
            Left            =   1200
            TabIndex        =   5
            Top             =   360
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   529
         End
         Begin prjDownloadBooster.ComboBoxW cbBitRate 
            Height          =   300
            Left            =   1560
            TabIndex        =   6
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
            TabIndex        =   7
            Top             =   1800
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   450
            Caption         =   "&CBR:"
            Transparent     =   -1  'True
         End
         Begin prjDownloadBooster.ComboBoxW cbAudioFormat 
            Height          =   300
            Left            =   2040
            TabIndex        =   8
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
            TabIndex        =   9
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
            TabIndex        =   10
            Top             =   1440
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   450
            Value           =   -1  'True
            Caption         =   "&VBR:"
            Transparent     =   -1  'True
         End
         Begin prjDownloadBooster.CheckBoxW chkExtractAudio 
            Height          =   255
            Left            =   360
            TabIndex        =   11
            Top             =   720
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   450
            Caption         =   "음원만 추출(&E)"
            Transparent     =   -1  'True
         End
         Begin VB.Label Label4 
            BackStyle       =   0  '투명
            Caption         =   "오디오 형식(&A):"
            Height          =   255
            Left            =   600
            TabIndex        =   13
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
   End
   Begin prjDownloadBooster.TabStrip tsTabStrip 
      Height          =   390
      Left            =   120
      TabIndex        =   14
      Top             =   105
      Width           =   2400
      _ExtentX        =   4233
      _ExtentY        =   688
      InitTabs        =   "frmDownloadOptions.frx":06D2
   End
   Begin prjDownloadBooster.CommandButtonW CancelButton 
      Cancel          =   -1  'True
      Height          =   360
      Left            =   4320
      TabIndex        =   15
      Top             =   120
      Width           =   1320
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "취소"
   End
   Begin prjDownloadBooster.CommandButtonW OKButton 
      Default         =   -1  'True
      Height          =   360
      Left            =   2880
      TabIndex        =   16
      Top             =   120
      Width           =   1320
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "확인"
   End
End
Attribute VB_Name = "frmDownloadOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SelectedListItem As LvwListItem
Dim MouseY As Integer

Private Sub CancelButton_Click()
    Unload Me
End Sub

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

Private Sub OKButton_Click()
    Dim i%
    For i = 1 To SessionHeaders.Count
        SessionHeaders.Remove 1
    Next i
    For i = 1 To SessionHeaderKeys.Count
        SessionHeaderKeys.Remove 1
    Next i
    
    If lvHeaders.ListItems.Count > 0 Then
        Dim RawHeaders$
        RawHeaders = ""
        For i = 1 To lvHeaders.ListItems.Count
            If Trim$(lvHeaders.ListItems(i).Text) <> "" Then
                SessionHeaders.Add CStr(lvHeaders.ListItems(i).ListSubItems(1).Text), CStr(Trim$(lvHeaders.ListItems(i).Text))
                SessionHeaderKeys.Add CStr(Trim$(lvHeaders.ListItems(i).Text))
                RawHeaders = RawHeaders & LCase(Trim$(lvHeaders.ListItems(i).Text)) & ": " & lvHeaders.ListItems(i).ListSubItems(1).Text & vbLf
            End If
        Next i
        If Right$(RawHeaders, 1) = vbLf Then RawHeaders = Left$(RawHeaders, Len(RawHeaders) - 1)
        SessionHeaderCache = btoa(RawHeaders)
    Else
        SessionHeaderCache = ""
    End If
    
    SaveSetting "DownloadBooster", "Options", "AutoDetectYtdlURL", chkAutoYtdl.Value
    
    Unload Me
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
    
    Dim i%
    Dim MaxWidth%, MaxHeight%
    MaxWidth = 15
    MaxHeight = 15
    For i = 1 To pbPanel.Count
        pbPanel(i).Visible = 0
        pbPanel(i).Enabled = 0
        pbPanel(i).Top = 450
        pbPanel(i).Left = 165
        pbPanel(i).BorderStyle = 0
        pbPanel(i).AutoRedraw = True
        If MaxWidth < pbPanel(i).Width Then MaxWidth = pbPanel(i).Width
        If MaxHeight < pbPanel(i).Height Then MaxHeight = pbPanel(i).Height
    Next i
    For i = 1 To pbPanel.Count
        pbPanel(i).Width = MaxWidth
        pbPanel(i).Height = MaxHeight
    Next i
    tsTabStrip.Width = MaxWidth + 105
    tsTabStrip.Height = MaxHeight + 390
    tsTabStrip.Top = 120
    tsTabStrip.Left = 120
    CancelButton.Top = tsTabStrip.Top + tsTabStrip.Height + 60
    OKButton.Top = CancelButton.Top
    CancelButton.Left = tsTabStrip.Left + tsTabStrip.Width - CancelButton.Width
    OKButton.Left = CancelButton.Left - 120 - OKButton.Width
    Me.Height = CancelButton.Top + CancelButton.Height + 540
    Me.Width = tsTabStrip.Width + 240 + 60
    pbPanel(1).Visible = -1
    pbPanel(1).Enabled = -1
    
    For i = 1 To pbPanel.Count
        tsTabStrip.DrawBackground pbPanel(i).hWnd, pbPanel(i).hDC
    Next i
    fYtdl.Refresh
    
    cbAudioFormat.AddItem t("자동", "Auto") & " (M4A/OPUS)"
    cbAudioFormat.AddItem "MP3"
    cbAudioFormat.AddItem "WAV"
    cbAudioFormat.AddItem "FLAC"
    cbAudioFormat.ListIndex = 0
    
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
    
    txtFormat.AddItem t("자동", "Auto")
    txtFormat.ListIndex = 0
    
    chkAutoYtdl.Value = GetSetting("DownloadBooster", "Options", "AutoDetectYtdlURL", 1)
    
    Me.Caption = t(Me.Caption, "Download settings")
    tsTabStrip.Tabs(2).Caption = t("헤더", "Headers")
    
    chkAutoYtdl.Caption = t(chkAutoYtdl.Caption, "Automatically use &youtube-dl for supported links")
    optDisableYtdl.Caption = t(optDisableYtdl.Caption, "Never use youtube-&dl for all links")
    optUseYtdl.Caption = t(optUseYtdl.Caption, "Always &use youtube-dl for all links")
    optUseYtdl.Width = t(optUseYtdl.Width, 2775)
    chkExtractAudio.Caption = t(chkExtractAudio.Caption, "&Extract audio")
    Label4.Caption = t(Label4.Caption, "&Audio format:")
    Label3.Caption = t(Label3.Caption, "&Format:")
    
    OKButton.Caption = t(OKButton.Caption, "OK")
    CancelButton.Caption = t(CancelButton.Caption, "Cancel")
    
    cmdAddHeader.Caption = t(cmdAddHeader.Caption, "&Add")
    cmdDeleteHeader.Caption = t(cmdDeleteHeader.Caption, "&Delete")
    cmdEditHeaderName.Caption = t(cmdEditHeaderName.Caption, "&Rename")
    cmdEditHeaderValue.Caption = t(cmdEditHeaderValue.Caption, "&Edit")
    lblDescription.Caption = t(lblDescription.Caption, "Headers here are only applied in this session. Go to <A>Options</A> to change them permanently.")
    
    lvHeaders.ColumnHeaders.Add , , t("이름", "Name"), 2055
    lvHeaders.ColumnHeaders.Add , , t("값", "Value"), 2775
    lvHeaders.SmallIcons = imgFiles
    
    Dim Header
    For Each Header In SessionHeaderKeys
        lvHeaders.ListItems.Add(, , Header, , 1).ListSubItems.Add , , SessionHeaders(CStr(Header))
    Next Header
    
    optUseYtdl_Click
End Sub

Private Sub tsTabStrip_TabClick(ByVal TabItem As TbsTab)
    Dim i%
    For i = 1 To pbPanel.Count
        If i = TabItem.Index Then
            pbPanel(i).Visible = -1
            pbPanel(i).Enabled = -1
        Else
            pbPanel(i).Visible = 0
            pbPanel(i).Enabled = 0
        End If
    Next i
End Sub

Private Sub cmdAddHeader_Click()
    lvHeaders.SetFocus
    Set lvHeaders.SelectedItem = lvHeaders.ListItems.Add(, , "", , 1)
    lvHeaders.SelectedItem.ListSubItems.Add , , ""
    lvHeaders.StartLabelEdit
End Sub

Private Sub cmdDeleteHeader_Click()
    If Not lvHeaders.SelectedItem Is Nothing Then
        If lvHeaders.SelectedItem.Selected Then
            lvHeaders.ListItems.Remove lvHeaders.SelectedItem.Index
        End If
    End If
End Sub

Private Sub cmdEditHeaderName_Click()
    On Error Resume Next
    lvHeaders.SetFocus
    lvHeaders.StartLabelEdit
End Sub

Private Sub cmdEditHeaderValue_Click()
    On Error GoTo exitsub
    If Not lvHeaders.SelectedItem Is Nothing Then
        Set SelectedListItem = lvHeaders.SelectedItem
        With txtEdit
            .Top = (lvHeaders.Top + MouseY) - Fix((txtEdit.Height) / 2)
            .Left = lvHeaders.Left + lvHeaders.ColumnHeaders(1).Width + 30
            .Width = lvHeaders.ColumnHeaders(2).Width
            .Text = SelectedListItem.ListSubItems(1).Text
            .Visible = True
            .SetFocus
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
        OKButton.Enabled = 0
    End If
exitsub:
End Sub

Private Sub lblDescription_LinkActivate(ByVal Link As LlbLink, ByVal Reason As LlbLinkActivateReasonConstants)
    Load frmOptions
    frmOptions.tsTabStrip.Tabs(2).Selected = -1
    frmOptions.Show vbModal, Me
End Sub

Private Sub txtEdit_LostFocus()
    On Error Resume Next
    SelectedListItem.ListSubItems(1).Text = txtEdit.Text
    txtEdit.Visible = False
    Set SelectedListItem = Nothing
    OKButton.Enabled = -1
End Sub
 
Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Or KeyAscii = 10 Then
        SelectedListItem.ListSubItems(1).Text = txtEdit.Text
        txtEdit.Visible = False
        Set SelectedListItem = Nothing
        OKButton.Enabled = -1
        lvHeaders.SetFocus
    End If
End Sub

Private Sub lvHeaders_AfterLabelEdit(Cancel As Boolean, NewString As String)
    NewString = Trim$(NewString)
    If NewString = "" Then
invalidname:
        Cancel = True
        Alert t("헤더 이름이 잘못되었습니다.", "Invalid header name."), App.Title, Me, 16
        Exit Sub
    End If
    
    Dim i%
    For i = 1 To Len(NewString)
        Select Case Mid$(NewString, i, 1)
            Case "a" To "z", "A" To "Z", "0" To "9", "-", "_"
            Case Else
                GoTo invalidname
        End Select
    Next i
    
    For i = 1 To lvHeaders.ListItems.Count
        If LCase(lvHeaders.ListItems(i).Text) = LCase(NewString) Then
            Cancel = True
            Alert t("해당 이름이 이미 존재합니다.", "Duplicate header name."), App.Title, Me, 16
            Exit Sub
            Exit For
        End If
    Next i
End Sub

Private Sub lvHeaders_ItemDblClick(ByVal Item As LvwListItem, ByVal Button As Integer)
    If Item.Selected Then _
        cmdEditHeaderValue_Click
End Sub

Private Sub lvHeaders_ItemSelect(ByVal Item As LvwListItem, ByVal Selected As Boolean)
    On Error GoTo justdisable
    If Selected Then
        cmdDeleteHeader.Enabled = -1
        cmdEditHeaderName.Enabled = -1
        cmdEditHeaderValue.Enabled = -1
        Exit Sub
    End If
justdisable:
    cmdDeleteHeader.Enabled = 0
    cmdEditHeaderName.Enabled = 0
    cmdEditHeaderValue.Enabled = 0
End Sub

Private Sub lvHeaders_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseY = Y
End Sub

