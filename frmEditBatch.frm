VERSION 5.00
Begin VB.Form frmEditBatch 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "편집"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5235
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
   ScaleHeight     =   2475
   ScaleWidth      =   5235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin prjDownloadBooster.TygemButton tygCancel 
      Height          =   375
      Left            =   3720
      TabIndex        =   9
      Top             =   2040
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "취소"
      BackColor       =   0
      FontSize        =   0
   End
   Begin prjDownloadBooster.TygemButton tygOK 
      Height          =   375
      Left            =   2280
      TabIndex        =   8
      Top             =   2040
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
      Left            =   2280
      TabIndex        =   5
      Top             =   2040
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "확인"
   End
   Begin prjDownloadBooster.CommandButtonW cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   3720
      TabIndex        =   6
      Top             =   2040
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "취소"
   End
   Begin prjDownloadBooster.FrameW fInfo 
      Height          =   1815
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   3201
      Caption         =   " 파일 다운로드 정보 "
      Begin prjDownloadBooster.TygemButton tygBrowse 
         Height          =   330
         Left            =   3120
         TabIndex        =   10
         Top             =   1380
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         Caption         =   "찾아보기..."
         BackColor       =   0
         FontSize        =   0
      End
      Begin prjDownloadBooster.CommandButtonW cmdBrowse 
         Height          =   330
         Left            =   3120
         TabIndex        =   4
         Top             =   1380
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         Caption         =   "찾아보기(&B)..."
      End
      Begin prjDownloadBooster.TextBoxW txtFilePath 
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   1080
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   450
      End
      Begin prjDownloadBooster.TextBoxW txtURL 
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   480
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   450
      End
      Begin VB.Label Label2 
         Caption         =   "저장 경로(&S):"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   3495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "파일 주소(&A):"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   2895
      End
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
    fInfo.Caption = t(fInfo.Caption, " File download information ")
    Label1.Caption = t(Label1.Caption, "File &address:")
    Label2.Caption = t(Label2.Caption, "&Save to:")
    Me.Caption = t(Me.Caption, "Edit")
    
    On Error Resume Next
    Me.Icon = frmMain.imgEdit.ListImages(1).Picture
    On Error GoTo 0
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
