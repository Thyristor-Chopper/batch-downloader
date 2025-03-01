VERSION 5.00
Begin VB.Form frmBatchAdd 
   Caption         =   "�ϰ� �ٿ�ε�"
   ClientHeight    =   3795
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5985
   BeginProperty Font 
      Name            =   "����"
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
   StartUpPosition =   1  '������ ���
   Begin prjDownloadBooster.CommandButtonW cmdAdvanced 
      Height          =   340
      Left            =   4560
      TabIndex        =   7
      Top             =   900
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   609
      Caption         =   "���(&V)..."
   End
   Begin prjDownloadBooster.TextBoxW txtSavePath 
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3405
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   450
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
      Caption         =   "���"
   End
   Begin prjDownloadBooster.CommandButtonW cmdOK 
      Height          =   340
      Left            =   4560
      TabIndex        =   2
      Top             =   120
      Width           =   1335
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "Ȯ��"
   End
   Begin prjDownloadBooster.CommandButtonW cmdBrowse 
      Height          =   330
      Left            =   4560
      TabIndex        =   6
      Top             =   3360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      Caption         =   "ã�ƺ���(&B)..."
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '����
      Caption         =   "���� ���(&S):"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3165
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "�� �ٿ� ���� �ּҸ� �Է��Ͻʽÿ�(&L)."
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
'�����ڷ�
'https://blog.naver.com/wnwlsrb3/220729779017
Dim PrevKeyCode As Integer
Dim Initialized As Boolean
Public HeaderCache$

Private Sub cmdAdvanced_Click()
    Tags.DownloadOptionsTargetForm = 1
    Set frmDownloadOptions.HeaderKeys = New Collection
    Set frmDownloadOptions.Headers = New Collection
    frmDownloadOptions.Show vbModal, Me
End Sub

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
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    txtSavePath.Text = Trim$(txtSavePath.Text)
    If Not FolderExists(txtSavePath.Text) Then
        Alert t("���� ��ΰ� �������� �ʽ��ϴ�. [ã�ƺ���] ������� ������ ã�ƺ� �� �ֽ��ϴ�.", "Save path does not exist. Use Broewse to browse folders."), App.Title, 16
        Exit Sub
    End If
    txtSavePath.Text = FilterFilename(txtSavePath.Text, True)

    Dim URLs() As String
    URLs = Split(txtURLs.Text, vbCrLf)
    For i = 0 To UBound(URLs)
        If Replace(URLs(i), " ", "") <> "" Then
            frmMain.AddBatchURLs URLs(i), txtSavePath.Text, HeaderCache
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
    Label1.Caption = t(Label1.Caption, "Enter one UR&L per line:")
    Label2.Caption = t(Label2.Caption, "&Save to:")
    cmdBrowse.Caption = t(cmdBrowse.Caption, "&Browse...")
    tr cmdAdvanced, "Ad&vanced..."
    
    HeaderCache = ""
    
    On Error Resume Next
    Me.Icon = frmMain.Icon
    On Error GoTo 0
    
    Hook_BatchAdd Me.hWnd
    On Error Resume Next
    Me.Width = GetSetting("DownloadBooster", "UserData", "BatchURLAddWidth", Me.Width - PaddedBorderWidth * 15 * 2) + PaddedBorderWidth * 15 * 2
    Me.Height = GetSetting("DownloadBooster", "UserData", "BatchURLAddHeight", Me.Height - PaddedBorderWidth * 15 * 2) + PaddedBorderWidth * 15 * 2
End Sub

Sub Form_Resize()
    If Me.WindowState = 1 Then Exit Sub
    cmdOK.Left = Me.Width - PaddedBorderWidth * 15 * 2 - 1545
    cmdCancel.Left = cmdOK.Left
    cmdAdvanced.Left = cmdOK.Left
    txtURLs.Width = Me.Width - PaddedBorderWidth * 15 * 2 - 1890
    txtURLs.Height = Me.Height - PaddedBorderWidth * 15 * 2 - 1770
    Label2.Top = Me.Height - PaddedBorderWidth * 15 * 2 - 1140
    txtSavePath.Top = Me.Height - PaddedBorderWidth * 15 * 2 - 900
    cmdBrowse.Left = cmdOK.Left
    cmdBrowse.Top = Me.Height - PaddedBorderWidth * 15 * 2 - 945
    txtSavePath.Width = txtURLs.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Me.WindowState = 0 Then
        SaveSetting "DownloadBooster", "UserData", "BatchURLAddWidth", Me.Width - PaddedBorderWidth * 15 * 2
        SaveSetting "DownloadBooster", "UserData", "BatchURLAddHeight", Me.Height - PaddedBorderWidth * 15 * 2
    End If
    Unhook_BatchAdd Me.hWnd
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
