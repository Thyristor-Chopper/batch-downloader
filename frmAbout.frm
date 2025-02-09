VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "정보"
   ClientHeight    =   5595
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   7890
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "굴림"
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
   ScaleHeight     =   3861.77
   ScaleMode       =   0  '사용자
   ScaleWidth      =   7409.121
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.PictureBox pbLicenses 
      BorderStyle     =   0  '없음
      Height          =   3255
      Index           =   5
      Left            =   2760
      ScaleHeight     =   3255
      ScaleWidth      =   4815
      TabIndex        =   15
      Top             =   1680
      Width           =   4815
      Begin prjDownloadBooster.TextBoxW txtMisc 
         Height          =   3255
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   5741
         MultiLine       =   -1  'True
         ScrollBars      =   3
      End
   End
   Begin VB.PictureBox pbLicenses 
      BorderStyle     =   0  '없음
      Height          =   3255
      Index           =   4
      Left            =   2760
      ScaleHeight     =   3255
      ScaleWidth      =   4815
      TabIndex        =   19
      Top             =   1680
      Width           =   4815
      Begin prjDownloadBooster.TextBoxW txtPNG 
         Height          =   3255
         Left            =   0
         TabIndex        =   20
         Top             =   0
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   5741
         MultiLine       =   -1  'True
         ScrollBars      =   2
      End
   End
   Begin VB.PictureBox pbLicenses 
      BorderStyle     =   0  '없음
      Height          =   3255
      Index           =   3
      Left            =   2760
      ScaleHeight     =   3255
      ScaleWidth      =   4815
      TabIndex        =   17
      Top             =   1680
      Width           =   4815
      Begin prjDownloadBooster.TextBoxW txtCC 
         Height          =   3255
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   5741
         MultiLine       =   -1  'True
         ScrollBars      =   2
      End
   End
   Begin VB.PictureBox pbLicenses 
      BorderStyle     =   0  '없음
      Height          =   3255
      Index           =   2
      Left            =   2760
      ScaleHeight     =   3255
      ScaleWidth      =   4815
      TabIndex        =   5
      Top             =   1680
      Width           =   4815
      Begin prjDownloadBooster.TextBoxW txtIconv 
         Height          =   3255
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   5741
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2
      End
   End
   Begin VB.PictureBox pbLicenses 
      BorderStyle     =   0  '없음
      Height          =   3255
      Index           =   1
      Left            =   2760
      ScaleHeight     =   3255
      ScaleWidth      =   4815
      TabIndex        =   2
      Top             =   1680
      Width           =   4815
      Begin prjDownloadBooster.LinkLabel lblReadOnline 
         Height          =   255
         Left            =   1440
         TabIndex        =   14
         Tag             =   "nocolorchange"
         Top             =   2280
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         Caption         =   "frmAbout.frx":000C
         Transparent     =   -1  'True
      End
      Begin prjDownloadBooster.TextBoxW txtLicensePlaceholder 
         Height          =   270
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   1215
         _ExtentX        =   0
         _ExtentY        =   0
         Locked          =   -1  'True
         ScrollBars      =   2
      End
      Begin prjDownloadBooster.ProgressBar pbLicenseLoadProgress 
         Height          =   255
         Left            =   0
         Top             =   3000
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   450
         Max             =   812
         Step            =   10
      End
      Begin prjDownloadBooster.TextBoxW txtLicense 
         Height          =   2955
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   5212
         Enabled         =   0   'False
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3
      End
   End
   Begin prjDownloadBooster.ImageList imgItems 
      Left            =   360
      Top             =   4320
      _ExtentX        =   1005
      _ExtentY        =   1005
      ImageWidth      =   32
      ImageHeight     =   32
      ColorDepth      =   8
      MaskColor       =   16711935
      InitListImages  =   "frmAbout.frx":004E
   End
   Begin prjDownloadBooster.FrameW FrameW1 
      Height          =   3615
      Left            =   1080
      TabIndex        =   0
      Top             =   1440
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   6376
      Caption         =   "라이선스(&L)"
      Begin prjDownloadBooster.ListView lvItems 
         Height          =   3255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   5741
         Icons           =   "imgItems"
         Arrange         =   2
         LabelEdit       =   2
         HideSelection   =   0   'False
         ShowInfoTips    =   -1  'True
         ShowLabelTips   =   -1  'True
         ShowColumnTips  =   -1  'True
         SnapToGrid      =   -1  'True
      End
   End
   Begin prjDownloadBooster.TygemButton tygOK 
      Height          =   345
      Left            =   5805
      TabIndex        =   13
      Top             =   5160
      Width           =   1950
      _ExtentX        =   3440
      _ExtentY        =   609
   End
   Begin prjDownloadBooster.TygemButton tygSysInfo 
      Height          =   345
      Left            =   3720
      TabIndex        =   9
      Top             =   5160
      Width           =   1950
      _ExtentX        =   3440
      _ExtentY        =   609
   End
   Begin VB.Timer timLicenseLoader 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   5880
      Top             =   120
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "확인"
      Default         =   -1  'True
      Height          =   345
      Left            =   5805
      TabIndex        =   8
      Top             =   5160
      Width           =   1950
   End
   Begin prjDownloadBooster.CommandButtonW cmdSysInfo 
      Height          =   345
      Left            =   3720
      TabIndex        =   7
      Top             =   5160
      Width           =   1950
      _ExtentX        =   3440
      _ExtentY        =   609
      Caption         =   "시스템 정보(&S)..."
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  '투명
      Caption         =   "버전"
      Height          =   225
      Left            =   1050
      TabIndex        =   10
      Tag             =   "nocolorchange"
      Top             =   600
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  '투명
      Caption         =   "응용 프로그램 제목"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1050
      TabIndex        =   12
      Tag             =   "nocolorchange"
      Top             =   240
      Width           =   3885
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  '투명
      Caption         =   "응용 프로그램 설명"
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   1050
      TabIndex        =   11
      Tag             =   "nocolorchange"
      Top             =   960
      Width           =   6645
   End
   Begin VB.Image picIcon 
      Height          =   480
      Left            =   240
      Picture         =   "frmAbout.frx":11BE
      Top             =   240
      Width           =   480
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

Private Sub Form_Activate()
    'On Error Resume Next
    'lvItems.SetFocus
End Sub

Private Sub Form_Load()
    If GetSetting("DownloadBooster", "Options", "DisableDWMWindow", DefaultDisableDWMWindow) = 1 Then DisableDWMWindow Me.hWnd
    SetFormBackgroundColor Me
    SetFont Me
    SetWindowPos Me.hWnd, IIf(MainFormOnTop, HWND_TOPMOST, HWND_NOTOPMOST), 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    
    LineNum = 1
    Me.Caption = t(App.Title & " 정보", "About " & App.Title)
    'picIcon.Picture = frmMain.Icon
    lblVersion.Caption = t("버전 ", "Version ") & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
    lblDescription.Caption = t("이 프로그램에는 외부 라이브러리가 일부 포함되어 있으며 라이선스 전문은 다음과 같습니다.", "This program includes external libraries. Check out the license of them below.")
    lblDescription.Width = Me.Width - lblDescription.Left - 180 - 540
    'txtLicense.Width = lblDescription.Width
    'txtLicense.Height = Line1(0).Y1 - txtLicense.Top - 90 - cmdSysInfo.Height - 90 - pbLicenseLoadProgress.Height - 30 + 120
    txtLicensePlaceholder.Width = txtLicense.Width
    txtLicensePlaceholder.Height = txtLicense.Height
    txtLicensePlaceholder.Top = txtLicense.Top
    txtLicensePlaceholder.Left = txtLicense.Left
    pbLicenseLoadProgress.Width = txtLicense.Width
    pbLicenseLoadProgress.Top = txtLicense.Top + txtLicense.Height + 30
    pbLicenseLoadProgress.Left = txtLicense.Left
    lblReadOnline.Top = txtLicense.Top + txtLicense.Height + 30 + pbLicenseLoadProgress.Height + 60
    lblReadOnline.Left = txtLicense.Left
    cmdOK.Caption = t(cmdOK.Caption, "OK")
    cmdSysInfo.Caption = t(cmdSysInfo.Caption, "&System information...")
    lblReadOnline.Caption = t(lblReadOnline.Caption, "<A>[Read online]</A>")
    tygOK.Caption = t("확인", "OK")
    tygSysInfo.Caption = t("시스템 정보...", "System information...")
    
    txtIconv.Text = txtIconv.Text & "Copyright (c) 2011 Alexander Shtuchkin" & vbCrLf
    txtIconv.Text = txtIconv.Text & "" & vbCrLf
    txtIconv.Text = txtIconv.Text & "Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the" & _
                                    """Software""), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish," & _
                                    "distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:" & vbCrLf
    txtIconv.Text = txtIconv.Text & "" & vbCrLf
    txtIconv.Text = txtIconv.Text & "The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software." & vbCrLf
    txtIconv.Text = txtIconv.Text & "" & vbCrLf
    txtIconv.Text = txtIconv.Text & "THE SOFTWARE IS PROVIDED ""AS IS"", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF" & _
                                    "MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE" & _
                                    "LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE."
    
    timLicenseLoader.Enabled = -1
    
    lvItems.ListItems.Add , , "Node.js (v0.11.11)", 1
    lvItems.ListItems.Add , , "iconv-lite (v0.6.3)", 2
    lvItems.ListItems.Add , , "Common Controls", 1
    lvItems.ListItems.Add , , "PNG Alpha", 2
    lvItems.ListItems.Add , , t("기타 소스 코드", "Other source codes"), 1
    lvItems.ListItems(1).Selected = True
    
    txtCC.Text = "https://github.com/Kr00l/VBCCR/tree/master/Standard%20EXE%20Version" & vbCrLf & vbCrLf
    txtCC.Text = txtCC.Text & "MIT License" & vbCrLf & vbCrLf
    txtCC.Text = txtCC.Text & "Copyright (c) 2012-present Krool" & vbCrLf & vbCrLf
    txtCC.Text = txtCC.Text & "Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the ""Software""), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:" & vbCrLf & vbCrLf
    txtCC.Text = txtCC.Text & "The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software." & vbCrLf & vbCrLf
    txtCC.Text = txtCC.Text & "THE SOFTWARE IS PROVIDED ""AS IS"", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE."
    
    txtPNG.Text = "https://www.vbforums.com/showthread.php?896878-PNG-with-alpha-channel-into-standard-VB6-image-control" & vbCrLf & vbCrLf
    txtPNG.Text = txtPNG.Text & "Elroy, LaVolpe, Dilettante, Wqweto, Schmidt, & The Trick" & vbCrLf & vbCrLf
    txtPNG.Text = txtPNG.Text & "Any software I (Elroy) post in these forums (VBForums) written by me is provided ""AS IS"" without warranty of any kind, expressed or implied, and permission is hereby granted, free of charge and without restriction, to any person obtaining a copy. To all, peace and happiness." & vbCrLf & vbCrLf
    
    txtMisc.Text = txtMisc.Text & "- https://www.vbforums.com/showthread.php?457171-RESOLVED-How-to-get-Desktop-Path-in-VB" & vbCrLf
    txtMisc.Text = txtMisc.Text & "- https://www.vbforums.com/showthread.php?445574-Reading-shortcut-information" & vbCrLf
    txtMisc.Text = txtMisc.Text & "- https://www.vbforums.com/showthread.php?430704-RESOLVED-Get-drive-size-space" & vbCrLf
    txtMisc.Text = txtMisc.Text & "- https://www.codeguru.com/visual-basic/displaying-the-file-properties-dialog/" & vbCrLf
    txtMisc.Text = txtMisc.Text & "- http://vbcity.com/forums/t/105530.aspx" & vbCrLf
    txtMisc.Text = txtMisc.Text & "- https://www.vbforums.com/showthread.php?696217-How-do-I-load-an-EXE-or-DLL-file-icon" & vbCrLf
    
    FrameW1.Caption = t(FrameW1.Caption, "&License")
End Sub

Private Sub lblReadOnline_LinkActivate(ByVal Link As LlbLink, ByVal Reason As LlbLinkActivateReasonConstants)
    Shell "cmd /c start """" https://raw.githubusercontent.com/nodejs/node/refs/heads/v0.10/LICENSE"
End Sub

Sub StartSysInfo()
    On Error GoTo SysInfoErr
  
    Dim RC As Long
    Dim SysInfoPath As String
    
    ' 시스템 정보 프로그램의 경로와 이름을 레지스트리에서 가져옵니다...
    SysInfoPath = GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, "")
    If SysInfoPath = "" Then
        SysInfoPath = GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, "")
        If SysInfoPath <> "" Then
            ' 알려진 32비트 파일 버전의 존재 여부를 확인합니다.
            If Dir(SysInfoPath & "\MSINFO32.EXE") <> "" Then
                SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
                
            ' 오류 - 파일을 찾을 수 없습니다...
            Else
                GoTo SysInfoErr
            End If
        ' 오류 - 레지스트리 항목을 찾을 수 없습니다...
        Else
            GoTo SysInfoErr
        End If
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    Alert t("지금은 시스템 정보를 사용할 수 없습니다.", "System Information is unavailable."), App.Title, Me, 48
End Sub

Private Sub lvItems_ItemSelect(ByVal Item As LvwListItem, ByVal Selected As Boolean)
    On Error Resume Next
    If Selected = False Then Exit Sub
    
    Dim i%
    For i = pbLicenses.LBound To pbLicenses.UBound
        If i = Item.Index Then
            pbLicenses(i).Visible = -1
        Else
            pbLicenses(i).Visible = 0
        End If
    Next i
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
    txtLicensePlaceholder.Text = t("라이선스를 불러오는 중... (", "Loading the license text... (") & Floor(LineNum / 812 * 100) & "%)"
    LineNum = LineNum + 1
    Exit Sub
LicenseFail:
    txtLicense.Text = t("라이선스를 불러올 수 없습니다. 다음 링크에서 확인할 수 있습니다.", "Unable to load the license. Check this URL:") & vbCrLf & " https://raw.githubusercontent.com/nodejs/node/refs/heads/v0.10/LICENSE"
    timLicenseLoader.Enabled = 0
    pbLicenseLoadProgress.Visible = 0
    txtLicense.Height = txtLicense.Height + pbLicenseLoadProgress.Height + 30
    txtLicense.Enabled = -1
    txtLicensePlaceholder.Visible = 0
End Sub

Private Sub tygOK_Click()
    cmdOK_Click
End Sub

Private Sub tygSysInfo_Click()
    cmdSysInfo_Click
End Sub
