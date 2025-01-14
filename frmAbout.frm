VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'ũ�� ���� ��ȭ ����
   Caption         =   "MyApp ����"
   ClientHeight    =   4725
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   6750
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "����"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3261.279
   ScaleMode       =   0  '�����
   ScaleWidth      =   6338.599
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '������ ���
   Begin VB.Timer timLicenseLoader 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   5880
      Top             =   120
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "Ȯ��"
      Default         =   -1  'True
      Height          =   345
      Left            =   4845
      TabIndex        =   0
      Top             =   3840
      Width           =   1710
   End
   Begin prjDownloadBooster.LinkLabel lblReadOnline 
      Height          =   255
      Left            =   2520
      TabIndex        =   2
      Tag             =   "nocolorchange"
      Top             =   2040
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      Caption         =   "frmAbout.frx":0742
      Transparent     =   -1  'True
   End
   Begin prjDownloadBooster.TextBoxW txtLicensePlaceholder 
      Height          =   270
      Left            =   360
      TabIndex        =   3
      Top             =   1440
      Width           =   1215
      _ExtentX        =   0
      _ExtentY        =   0
      Locked          =   -1  'True
      ScrollBars      =   2
   End
   Begin prjDownloadBooster.ProgressBar pbLicenseLoadProgress 
      Height          =   255
      Left            =   1680
      Top             =   1320
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   450
      Max             =   812
      Step            =   10
   End
   Begin prjDownloadBooster.TextBoxW txtLicense 
      Height          =   615
      Left            =   1050
      TabIndex        =   4
      Top             =   1440
      Width           =   1695
      _ExtentX        =   0
      _ExtentY        =   0
      Enabled         =   0   'False
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3
   End
   Begin prjDownloadBooster.CommandButtonW cmdSysInfo 
      Height          =   345
      Left            =   4845
      TabIndex        =   5
      Top             =   4275
      Width           =   1710
      _ExtentX        =   3016
      _ExtentY        =   609
      Caption         =   "�ý��� ����(&S)..."
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  '����
      Caption         =   "����"
      Height          =   225
      Left            =   1050
      TabIndex        =   1
      Tag             =   "nocolorchange"
      Top             =   600
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  '����
      Caption         =   "���� ���α׷� ����"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1050
      TabIndex        =   7
      Tag             =   "nocolorchange"
      Top             =   240
      Width           =   3885
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  '����
      Caption         =   "���� ���α׷� ����"
      ForeColor       =   &H00000000&
      Height          =   570
      Left            =   1050
      TabIndex        =   6
      Tag             =   "nocolorchange"
      Top             =   960
      Width           =   4125
   End
   Begin VB.Image picIcon 
      Height          =   480
      Left            =   240
      Top             =   240
      Width           =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  '���� �ܻ�
      Index           =   1
      X1              =   112.686
      X2              =   6237.181
      Y1              =   2515.844
      Y2              =   2515.844
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   100.479
      X2              =   6210.888
      Y1              =   2526.197
      Y2              =   2526.197
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"
Dim LineNum As Integer

Private Sub cmdSysInfo_Click()
  Call StartSysInfo
End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    If GetSetting("DownloadBooster", "Options", "DisableDWMWindow", DefaultDisableDWMWindow) = 1 Then DisableDWMWindow Me.hWnd
    SetFormBackgroundColor Me
    SetFont Me
    
    LineNum = 1
    Me.Caption = t(App.Title & " ����", "About " & App.Title)
    picIcon.Picture = frmMain.Icon
    lblVersion.Caption = t("���� ", "Version ") & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
    lblDescription.Caption = t("�� ���α׷����� Node.js�� ���̳ʸ��� ���ԵǾ� ������," & vbCrLf & "���̼��� ������ ������ �����ϴ�.", "This program includes the binary of Node.js." & vbCrLf & "Check out the license of Node.js below.")
    lblDescription.Width = Me.Width - lblDescription.Left - 180 - 540
    txtLicense.Width = lblDescription.Width
    txtLicense.Height = Line1(0).Y1 - txtLicense.Top - 90 - cmdSysInfo.Height - 90 - pbLicenseLoadProgress.Height - 30 + 120
    txtLicensePlaceholder.Width = txtLicense.Width
    txtLicensePlaceholder.Height = txtLicense.Height
    txtLicensePlaceholder.Top = txtLicense.Top
    txtLicensePlaceholder.Left = txtLicense.Left
    pbLicenseLoadProgress.Width = txtLicense.Width
    pbLicenseLoadProgress.Top = txtLicense.Top + txtLicense.Height + 30
    pbLicenseLoadProgress.Left = txtLicense.Left
    lblReadOnline.Top = txtLicense.Top + txtLicense.Height + 30 + pbLicenseLoadProgress.Height + 60
    lblReadOnline.Left = txtLicense.Left
    cmdSysInfo.Caption = t(cmdSysInfo.Caption, "&System information...")
    lblReadOnline.Caption = t(lblReadOnline.Caption, "<A>[Read online]</A>")
    cmdOK.Caption = t(cmdOK.Caption, "OK")
    
    timLicenseLoader.Enabled = -1
End Sub

Private Sub lblReadOnline_LinkActivate(ByVal Link As LlbLink, ByVal Reason As LlbLinkActivateReasonConstants)
    Shell "cmd /c start """" https://raw.githubusercontent.com/nodejs/node/refs/heads/v0.10/LICENSE"
End Sub

Private Sub picIcon_DblClick()
    If frmMain.AboutEasterEgg And IsKeyPressed(gksKeyboardShift) And IsKeyPressed(gksKeyboardalt) And IsKeyPressed(gksKeyboardctrl) Then
        If WinVer < 5.1 Then
            frmGameWin95.Show 1, Me
        ElseIf WinVer < 6# Then
            frmGameWinXP.Show 1, Me
        ElseIf WinVer < 6.2 Then
            frmGameVista.Show 1, Me
        ElseIf WinVer < 6.4 Then
            frmGameWin95.Show 1, Me
        Else
            frmGame.Show 1, Me
        End If
    End If
End Sub

Sub StartSysInfo()
    On Error GoTo SysInfoErr
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' �ý��� ���� ���α׷��� ��ο� �̸��� ������Ʈ������ ���� �ɴϴ�...
    SysInfoPath = GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, "")
    If SysInfoPath = "" Then
        SysInfoPath = GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, "")
        If SysInfoPath <> "" Then
            ' �˷��� 32��Ʈ ���� ������ ���� ���θ� Ȯ���մϴ�.
            If Dir(SysInfoPath & "\MSINFO32.EXE") <> "" Then
                SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
                
            ' ���� - ������ ã�� �� �����ϴ�...
            Else
                GoTo SysInfoErr
            End If
        ' ���� - ������Ʈ�� �׸��� ã�� �� �����ϴ�...
        Else
            GoTo SysInfoErr
        End If
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    Alert t("������ �ý��� ������ ����� �� �����ϴ�.", "System Information is unavailable."), App.Title, Me, 48
End Sub

Private Sub timLicenseLoader_Timer()
    If LineNum > 812 Then
        timLicenseLoader.Enabled = 0
        pbLicenseLoadProgress.Visible = 0
        txtLicense.Height = txtLicense.Height + pbLicenseLoadProgress.Height + 30
        txtLicense.Enabled = -1
        txtLicensePlaceholder.Visible = 0
        Exit Sub
    End If
    
    On Error GoTo LicenseFail
    txtLicense.Text = txtLicense.Text & LoadResString(LineNum) & vbCrLf
    pbLicenseLoadProgress.Value = LineNum
    txtLicensePlaceholder.Text = t("���̼����� �ҷ����� ��... (", "Loading the license text... (") & Floor(LineNum / 812 * 100) & "%)"
    LineNum = LineNum + 1
    Exit Sub
LicenseFail:
    txtLicense.Text = t("���̼����� �ҷ��� �� �����ϴ�. ���� ��ũ���� Ȯ���� �� �ֽ��ϴ�.", "Unable to load the license. Check this URL:") & vbCrLf & " https://raw.githubusercontent.com/nodejs/node/refs/heads/v0.10/LICENSE"
    timLicenseLoader.Enabled = 0
    pbLicenseLoadProgress.Visible = 0
    txtLicense.Height = txtLicense.Height + pbLicenseLoadProgress.Height + 30
    txtLicense.Enabled = -1
    txtLicensePlaceholder.Visible = 0
End Sub
