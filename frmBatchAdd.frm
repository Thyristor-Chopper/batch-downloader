VERSION 5.00
Begin VB.Form frmBatchAdd 
   Caption         =   "일괄 다운로드"
   ClientHeight    =   3795
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5985
   BeginProperty Font 
      Name            =   "굴림"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBatchAdd.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin prjDownloadBooster.TygemButton tygBrowse 
      Height          =   330
      Left            =   4560
      TabIndex        =   8
      Top             =   3360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      Caption         =   "찾아보기..."
      BackColor       =   0
      FontSize        =   0
   End
   Begin prjDownloadBooster.TextBoxW txtSavePath 
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   3405
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   450
   End
   Begin prjDownloadBooster.TygemButton tygCancel 
      Height          =   345
      Left            =   4560
      TabIndex        =   5
      Top             =   510
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   609
   End
   Begin prjDownloadBooster.TygemButton tygOK 
      Height          =   345
      Left            =   4560
      TabIndex        =   4
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   609
   End
   Begin prjDownloadBooster.TextBoxW txtURLs 
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4215
      _ExtentX        =   0
      _ExtentY        =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3
   End
   Begin prjDownloadBooster.CommandButtonW cmdCancel 
      Cancel          =   -1  'True
      Height          =   340
      Left            =   4560
      TabIndex        =   3
      Top             =   510
      Width           =   1335
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "취소"
   End
   Begin prjDownloadBooster.CommandButtonW cmdOK 
      Height          =   340
      Left            =   4560
      TabIndex        =   2
      Top             =   120
      Width           =   1335
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "확인"
   End
   Begin prjDownloadBooster.CommandButtonW cmdBrowse 
      Height          =   330
      Left            =   4560
      TabIndex        =   9
      Top             =   3360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      Caption         =   "찾아보기(&B)..."
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '투명
      Caption         =   "저장 경로(&S):"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3165
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "각 줄에 파일 주소를 입력하십시오(&L)."
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmBatchAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'참고자료
'https://blog.naver.com/wnwlsrb3/220729779017
Dim PrevKeyCode As Integer
Dim Initialized As Boolean

Private Sub cmdBrowse_Click()
    txtSavePath.Text = Trim$(txtSavePath.Text)
    
    Unload frmBrowse
    Unload frmExplorer
    Tags.BrowsePresetPath = txtSavePath.Text
    Tags.BrowseTargetForm = 2
    If GetSetting("DownloadBooster", "Options", "ForceWin31Dialog", "0") = "1" Then
        frmBrowse.Show vbModal, Me
    Else
        frmExplorer.Show vbModal, Me
    End If
    
'    Dim bi As BROWSEINFO
'    Dim pidl As Long
'    Dim Path As String
'    Dim Pos As Integer
'    bi.hOwner = Me.hWnd
'    If FolderExists(txtSavePath.Text) Then
'        bi.pIDLRoot = SHSimpleIDListFromPath(txtSavePath.Text)
'    ElseIf FolderExists(fso.GetParentFolderName(txtSavePath.Text)) Then
'        bi.pIDLRoot = SHSimpleIDListFromPath(fso.GetParentFolderName(txtSavePath.Text))
'    Else
'        bi.pIDLRoot = 0&
'    End If
'    bi.lpszTitle = t("일괄 다운로드할 파일들을 저장할 폴더를 선택하십시오.", "Select the folder to save the downloaded files.")
'    bi.ulFlags = BIF_RETURNONLYFSDIRS Or BIF_DONTGOBELOWDOMAIN Or BIF_STATUSTEXT Or BIF_EDITBOX Or BIF_VALIDATE Or BIF_NEWDIALOGSTYLE
'    pidl = SHBrowseForFolder(bi)
'    Path = Space$(MAX_PATH)
'    If SHGetPathFromIDList(pidl, Path) Then
'        Pos = InStr(Path, Chr$(0))
'        txtSavePath.Text = Left(Path, Pos - 1)
'    End If
'    CoTaskMemFree pidl
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    txtSavePath.Text = Trim$(txtSavePath.Text)
    If Not FolderExists(txtSavePath.Text) Then
        Alert t("저장 경로가 존재하지 않습니다. [찾아보기] 기능으로 폴더를 찾아볼 수 있습니다.", "Save path does not exist. Use Broewse to browse folders."), App.Title, Me, 16
        Exit Sub
    End If
    txtSavePath.Text = FilterFilename(txtSavePath.Text, True)

    Dim URLs() As String
    URLs = Split(txtURLs.Text, vbCrLf)
    For i = 0 To UBound(URLs)
        If Replace(URLs(i), " ", "") <> "" Then
            frmMain.AddBatchURLs URLs(i), txtSavePath.Text
        End If
    Next i
    
    Unload Me
End Sub

Private Sub Form_Activate()
    If Initialized Then Exit Sub
    Initialized = True
    
    txtURLs.SetFocus
End Sub

Private Sub Form_Load()
    If GetSetting("DownloadBooster", "Options", "DisableDWMWindow", DefaultDisableDWMWindow) = 1 Then DisableDWMWindow Me.hWnd
    SetFormBackgroundColor Me
    SetFont Me
    SetWindowPos Me.hWnd, IIf(MainFormOnTop, HWND_TOPMOST, HWND_NOTOPMOST), 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    Initialized = False
    
    Me.Caption = t(Me.Caption, "Batch download")
    cmdOK.Caption = t(cmdOK.Caption, "OK")
    cmdCancel.Caption = t(cmdCancel.Caption, "Cancel")
    tygOK.Caption = cmdOK.Caption
    tygCancel.Caption = cmdCancel.Caption
    Label1.Caption = t(Label1.Caption, "Enter one UR&L per line:")
    Label2.Caption = t(Label2.Caption, "&Save to:")
    cmdBrowse.Caption = t(cmdBrowse.Caption, "&Browse...")
    tygBrowse.Caption = t("찾아보기...", "Browse...")
    
    On Error Resume Next
    Me.Icon = frmMain.Icon
    On Error GoTo 0
    
    SetWindowSizeLimit2 Me.hWnd, 5145 + PaddedBorderWidth * 15 * 2, Screen.Width + 1200, 2310 + PaddedBorderWidth * 15 * 2, Screen.Height + 1200
    On Error Resume Next
    Me.Width = GetSetting("DownloadBooster", "UserData", "BatchURLAddWidth", Me.Width - PaddedBorderWidth * 15 * 2) + PaddedBorderWidth * 15 * 2
    Me.Height = GetSetting("DownloadBooster", "UserData", "BatchURLAddHeight", Me.Height - PaddedBorderWidth * 15 * 2) + PaddedBorderWidth * 15 * 2
End Sub

Private Sub Form_Resize()
    If Me.WindowState = 1 Then Exit Sub
    cmdOK.Left = Me.Width - PaddedBorderWidth * 15 * 2 - 1545
    tygOK.Left = cmdOK.Left
    cmdCancel.Left = cmdOK.Left
    tygCancel.Left = cmdOK.Left
    txtURLs.Width = Me.Width - PaddedBorderWidth * 15 * 2 - 1890
    txtURLs.Height = Me.Height - PaddedBorderWidth * 15 * 2 - 1770
    Label2.Top = Me.Height - PaddedBorderWidth * 15 * 2 - 1140
    txtSavePath.Top = Me.Height - PaddedBorderWidth * 15 * 2 - 900
    cmdBrowse.Left = cmdOK.Left
    tygBrowse.Left = cmdOK.Left
    cmdBrowse.Top = Me.Height - PaddedBorderWidth * 15 * 2 - 945
    tygBrowse.Top = cmdBrowse.Top
    txtSavePath.Width = txtURLs.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Me.WindowState = 0 Then
        SaveSetting "DownloadBooster", "UserData", "BatchURLAddWidth", Me.Width - PaddedBorderWidth * 15 * 2
        SaveSetting "DownloadBooster", "UserData", "BatchURLAddHeight", Me.Height - PaddedBorderWidth * 15 * 2
    End If
    Unhook2 Me.hWnd
End Sub

Private Sub txtURLs_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = 13 Or KeyCode = 10) Then
        If (PrevKeyCode = 13 Or PrevKeyCode = 10) Then
            cmdOK_Click
        End If
        PrevKeyCode = KeyCode
    Else
        PrevKeyCode = 0
    End If
End Sub

Private Sub txtURLs_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    PrevKeyCode = 0
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
