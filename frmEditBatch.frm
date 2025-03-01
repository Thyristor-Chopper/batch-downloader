VERSION 5.00
Begin VB.Form frmEditBatch 
   BorderStyle     =   3  'ũ�� ���� ��ȭ ����
   Caption         =   "����"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5685
   BeginProperty Font 
      Name            =   "����"
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
   ScaleHeight     =   3075
   ScaleWidth      =   5685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '������ ���
   Begin prjDownloadBooster.CommandButtonW cmdOK 
      Default         =   -1  'True
      Height          =   330
      Left            =   2760
      TabIndex        =   5
      Top             =   2640
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      Caption         =   "Ȯ��"
   End
   Begin prjDownloadBooster.CommandButtonW cmdCancel 
      Cancel          =   -1  'True
      Height          =   330
      Left            =   4200
      TabIndex        =   6
      Top             =   2640
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      Caption         =   "���"
   End
   Begin prjDownloadBooster.FrameW fInfo 
      Height          =   2415
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   4260
      Caption         =   " ���� �ٿ�ε� ���� "
      Begin prjDownloadBooster.CommandButtonW cmdHeaders 
         Height          =   330
         Left            =   2160
         TabIndex        =   10
         Top             =   1920
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         Caption         =   "���(&H)..."
      End
      Begin prjDownloadBooster.CommandButtonW cmdYtdl 
         Height          =   330
         Left            =   360
         TabIndex        =   9
         Top             =   1920
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         Caption         =   "&youtube-dl..."
      End
      Begin prjDownloadBooster.CommandButtonW cmdBrowse 
         Height          =   330
         Left            =   3720
         TabIndex        =   4
         Top             =   1380
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         Caption         =   "ã�ƺ���(&B)..."
      End
      Begin prjDownloadBooster.TextBoxW txtFilePath 
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   1080
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   450
      End
      Begin prjDownloadBooster.TextBoxW txtURL 
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   480
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   450
      End
      Begin VB.Label Label3 
         BackStyle       =   0  '����
         Caption         =   "�߰� ����(&D):"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1680
         Width           =   3015
      End
      Begin VB.Label Label2 
         Caption         =   "���� ���(&S):"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   3495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '����
         Caption         =   "���� �ּ�(&A):"
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
Public EncodedHeaders As String
Public InitialFileName As String

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

Private Sub cmdHeaders_Click()
    Tags.DownloadOptionsTargetForm = 2
    Dim DecodedHeaders As Collection
    Set DecodedHeaders = DecodeHeaderCache(EncodedHeaders)
    Set frmDownloadOptions.HeaderKeys = DecodedHeaders("keys")
    Set frmDownloadOptions.Headers = DecodedHeaders("values")
#If Not HIDEYTDL Then
    frmDownloadOptions.tsTabStrip.Tabs(2).Selected = True
#End If
    frmDownloadOptions.Show vbModal, Me
End Sub

Private Sub cmdOK_Click()
    txtURL.Text = Trim$(txtURL.Text)
    If Left$(txtURL.Text, 7) <> "http://" And Left$(txtURL.Text, 8) <> "https://" Then
        Alert t("�ּҰ� �ùٸ��� �ʽ��ϴ�. 'http://' �Ǵ� 'https://'�� �����ؾ� �մϴ�.", "Invalid address. Must start with 'http://' or 'https://'."), App.Title, 16
        Exit Sub
    End If

    txtFilePath.Text = Trim$(txtFilePath.Text)
    Do While Replace(txtFilePath.Text, "\\", "\") <> txtFilePath.Text
        txtFilePath.Text = Replace(txtFilePath.Text, "\\", "\")
    Loop
    
    Dim OnlyFolderInPath As Boolean, FileNameModified As Boolean
    OnlyFolderInPath = False
    FileNameModified = True
    If FolderExists(txtFilePath.Text) Then
        If Right$(txtFilePath.Text, 1) <> "\" Then txtFilePath.Text = txtFilePath.Text & "\"
        txtFilePath.Text = txtFilePath.Text & Trim$(Tags.FileNameOnly)
        OnlyFolderInPath = True
    ElseIf Right$(txtFilePath.Text, 1) = "\" Or (Not FolderExists(GetParentFolderName(txtFilePath.Text))) Then
        Alert t("���� ��ΰ� �������� �ʽ��ϴ�. [ã�ƺ���] ������� ������ ã�ƺ� �� �ֽ��ϴ�.", "Save path does not exist. Use Broewse to browse folders."), App.Title, 16
        Exit Sub
    ElseIf LCase(GetFilename(txtFilePath.Text)) = LCase(Tags.FileNameOnly) Then
        FileNameModified = False
    End If
    txtFilePath.Text = FilterFilename(txtFilePath.Text, True)
    
    On Error Resume Next
    Dim ParentFolderName As String
    ParentFolderName = GetParentFolderName(txtFilePath.Text)
    If Right$(ParentFolderName, 1) = "\" Then ParentFolderName = Left$(ParentFolderName, Len(ParentFolderName) - 1)
    frmMain.lvBatchFiles.SelectedItem.ListSubItems(2).Text = txtURL.Text
    If frmMain.lvBatchFiles.SelectedItem.ListSubItems(3).Text = t("�Ϸ�", "Done") Then
        If txtURL.Text <> Trim$(OriginalURL) Then
            frmMain.lvBatchFiles.SelectedItem.ListSubItems(3).Text = t("���", "Queued")
            frmMain.lvBatchFiles.SelectedItem.Checked = True
            frmMain.lvBatchFiles.SelectedItem.ForeColor = &H80000008
            frmMain.lvBatchFiles.SelectedItem.ListSubItems(1).ForeColor = &H80000008
            frmMain.lvBatchFiles.SelectedItem.ListSubItems(2).ForeColor = &H80000008
            frmMain.lvBatchFiles.SelectedItem.ListSubItems(3).ForeColor = &H80000008
            GoTo changeFilepath
        ElseIf txtFilePath.Text <> Trim$(OriginalPath) Then
            Alert t("�ٿ�ε尡 �̹� �Ϸ�� ������ ���� ��ΰ� �������� �ʾҽ��ϴ�.", "Save path has not been changed because it's already downloaded."), App.Title
        End If
    Else
changeFilepath:
        frmMain.lvBatchFiles.SelectedItem.Text = Replace(txtFilePath.Text, ParentFolderName & "\", "", 1, 1)
        frmMain.lvBatchFiles.SelectedItem.ListSubItems(1).Text = txtFilePath.Text
        If txtFilePath.Text <> Trim$(OriginalPath) And (Not OnlyFolderInPath) And FileNameModified Then
            frmMain.lvBatchFiles.SelectedItem.ListSubItems(4).Text = "N"
        End If
    End If
    frmMain.lvBatchFiles.SelectedItem.ListSubItems(5).Text = EncodedHeaders
    
    Unload Me
End Sub

Private Sub cmdYtdl_Click()
    Tags.DownloadOptionsTargetForm = 2
    Dim DecodedHeaders As Collection
    Set DecodedHeaders = DecodeHeaderCache(EncodedHeaders)
    Set frmDownloadOptions.HeaderKeys = DecodedHeaders("keys")
    Set frmDownloadOptions.Headers = DecodedHeaders("values")
    frmDownloadOptions.Show vbModal, Me
End Sub

Private Sub Form_Load()
    If GetSetting("DownloadBooster", "Options", "DisableDWMWindow", DefaultDisableDWMWindow) = 1 Then DisableDWMWindow Me.hWnd
    SetFormBackgroundColor Me
    SetFont Me
    SetWindowPos Me.hWnd, IIf(MainFormOnTop, HWND_TOPMOST, HWND_NOTOPMOST), 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    
    cmdOK.Caption = t("Ȯ��", "OK")
    cmdCancel.Caption = t("���", "Cancel")
    cmdBrowse.Caption = t(cmdBrowse.Caption, "&Browse...")
    fInfo.Caption = t(fInfo.Caption, " File download information ")
    Label1.Caption = t(Label1.Caption, "File &address:")
    Label2.Caption = t(Label2.Caption, "&Save to:")
    Me.Caption = t(Me.Caption, "Edit")
    Label3.Caption = t(Label3.Caption, "A&dditional settings:")
    cmdHeaders.Caption = t(cmdHeaders.Caption, "&Headers...")
    
    EncodedHeaders = frmMain.lvBatchFiles.SelectedItem.ListSubItems(5).Text
    InitialFileName = frmMain.lvBatchFiles.SelectedItem.Text
    Me.Caption = Me.Caption & " - " & InitialFileName
    
    On Error Resume Next
    Me.Icon = frmMain.imgEdit.ListImages(1).Picture
    On Error GoTo 0
    
#If HIDEYTDL Then
    cmdYtdl.Visible = False
    cmdHeaders.Left = cmdYtdl.Left
#End If

End Sub
