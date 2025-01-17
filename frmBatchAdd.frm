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
   Begin prjDownloadBooster.TygemButton tygBrowse 
      Height          =   330
      Left            =   4560
      TabIndex        =   8
      Top             =   3360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      Caption         =   "ã�ƺ���..."
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
      TabIndex        =   1
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
      TabIndex        =   9
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
      TabIndex        =   6
      Top             =   3165
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "�� �ٿ� ���� �ּҸ� �Է��Ͻʽÿ�(&L)."
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmBatchAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PrevKeyCode As Integer

Private Sub cmdBrowse_Click()
    Unload frmBrowse
    Tags.BrowsePresetPath = Trim$(txtSavePath.Text)
    Tags.BrowseTargetForm = 2
    frmBrowse.Show vbModal, Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    txtSavePath.Text = Trim$(txtSavePath.Text)
    If Not FolderExists(txtSavePath.Text) Then
        Alert t("���� ��ΰ� �������� �ʽ��ϴ�. [ã�ƺ���] ������� ������ ã�ƺ� �� �ֽ��ϴ�.", "Save path does not exist. Use Broewse to browse folders."), App.Title, Me, 16
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

Private Sub Form_Load()
    If GetSetting("DownloadBooster", "Options", "DisableDWMWindow", DefaultDisableDWMWindow) = 1 Then DisableDWMWindow Me.hWnd
    SetFormBackgroundColor Me
    SetFont Me
    
    Me.Caption = t(Me.Caption, "Batch download")
    cmdOK.Caption = t(cmdOK.Caption, "OK")
    cmdCancel.Caption = t(cmdCancel.Caption, "Cancel")
    tygOK.Caption = cmdOK.Caption
    tygCancel.Caption = cmdCancel.Caption
    Label1.Caption = t(Label1.Caption, "Enter one UR&L per line:")
    Label2.Caption = t(Label2.Caption, "&Save to:")
    cmdBrowse.Caption = t(cmdBrowse.Caption, "&Browse...")
    tygBrowse.Caption = t("ã�ƺ���...", "Browse...")
    
    Me.Icon = frmMain.Icon
    
    SetWindowSizeLimit2 Me.hWnd, 5145 + PaddedBorderWidth * 15 * 2, Screen.Width + 1200, 2310 + PaddedBorderWidth * 15 * 2, Screen.Height + 1200
    On Error Resume Next
    Me.Width = GetSetting("DownloadBooster", "UserData", "BatchURLAddWidth", Me.Width - PaddedBorderWidth * 15 * 2) + PaddedBorderWidth * 15 * 2
    Me.Height = GetSetting("DownloadBooster", "UserData", "BatchURLAddHeight", Me.Height - PaddedBorderWidth * 15 * 2) + PaddedBorderWidth * 15 * 2
End Sub

Private Sub Form_Resize()
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
