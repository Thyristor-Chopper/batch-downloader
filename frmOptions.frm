VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "스킨 설정"
   ClientHeight    =   5490
   ClientLeft      =   2760
   ClientTop       =   3855
   ClientWidth     =   6750
   BeginProperty Font 
      Name            =   "굴림"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.PictureBox pbPanel 
      AutoRedraw      =   -1  'True
      Height          =   2265
      Index           =   1
      Left            =   1320
      ScaleHeight     =   2205
      ScaleWidth      =   3675
      TabIndex        =   5
      Top             =   960
      Width           =   3735
      Begin prjDownloadBooster.FrameW Frame5 
         Height          =   675
         Left            =   120
         TabIndex        =   34
         Top             =   1080
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   1191
         Caption         =   " 인터페이스 "
         Transparent     =   -1  'True
         Begin VB.ComboBox cbLanguage 
            Height          =   300
            Left            =   1080
            Style           =   2  '드롭다운 목록
            TabIndex        =   36
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '투명
            Caption         =   "언어(&L):"
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   270
            Width           =   975
         End
      End
      Begin prjDownloadBooster.CheckBoxW chkRememberURL 
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   600
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   450
         Caption         =   "파일 주소 기억(&M)"
         Transparent     =   -1  'True
      End
      Begin prjDownloadBooster.CheckBoxW chkNoCleanup 
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   2610
         _ExtentX        =   4604
         _ExtentY        =   450
         Caption         =   "조각 파일 유지(&N)"
         Transparent     =   -1  'True
      End
      Begin prjDownloadBooster.FrameW Frame2 
         Height          =   855
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   3495
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   " 다운로드 설정 "
      End
   End
   Begin VB.PictureBox pbWinXP 
      Height          =   375
      Left            =   4080
      Picture         =   "frmOptions.frx":0442
      ScaleHeight     =   315
      ScaleWidth      =   795
      TabIndex        =   33
      Top             =   6000
      Width           =   855
   End
   Begin VB.PictureBox pbDWM7 
      Height          =   375
      Left            =   3120
      Picture         =   "frmOptions.frx":1783
      ScaleHeight     =   315
      ScaleWidth      =   675
      TabIndex        =   32
      Top             =   6000
      Width           =   735
   End
   Begin VB.PictureBox pbNoDWM 
      Height          =   375
      Left            =   2280
      Picture         =   "frmOptions.frx":2DDB
      ScaleHeight     =   315
      ScaleWidth      =   555
      TabIndex        =   31
      Top             =   6000
      Width           =   615
   End
   Begin VB.PictureBox pbDWM8 
      Height          =   375
      Left            =   1320
      Picture         =   "frmOptions.frx":4046
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   30
      Top             =   6000
      Width           =   495
   End
   Begin VB.PictureBox pbDWM10 
      Height          =   375
      Left            =   720
      Picture         =   "frmOptions.frx":5093
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   29
      Top             =   6000
      Width           =   375
   End
   Begin VB.Timer timLicenseLoader 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   960
      Top             =   5040
   End
   Begin VB.PictureBox pbPanel 
      Height          =   2535
      Index           =   3
      Left            =   3840
      ScaleHeight     =   2475
      ScaleWidth      =   2595
      TabIndex        =   21
      Top             =   1920
      Width           =   2655
      Begin VB.TextBox txtLicensePlaceholder 
         Height          =   270
         Left            =   240
         Locked          =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   27
         Top             =   1320
         Width           =   1215
      End
      Begin prjDownloadBooster.ProgressBar pbLicenseLoadProgress 
         Height          =   255
         Left            =   1560
         Top             =   1200
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   450
         Max             =   812
         Step            =   10
      End
      Begin VB.TextBox txtLicense 
         Enabled         =   0   'False
         Height          =   495
         Left            =   930
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  '양방향
         TabIndex        =   26
         Top             =   1680
         Width           =   1695
      End
      Begin VB.CommandButton cmdSysInfo 
         Caption         =   "시스템 정보(&S)..."
         Height          =   345
         Left            =   120
         TabIndex        =   25
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Image picIcon 
         Height          =   480
         Left            =   120
         Top             =   120
         Width           =   480
      End
      Begin VB.Label lblDescription 
         BackStyle       =   0  '투명
         Caption         =   "응용 프로그램 설명"
         ForeColor       =   &H00000000&
         Height          =   810
         Left            =   930
         TabIndex        =   24
         Top             =   840
         Width           =   4125
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  '투명
         Caption         =   "응용 프로그램 제목"
         ForeColor       =   &H00000000&
         Height          =   480
         Left            =   930
         TabIndex        =   23
         Top             =   120
         Width           =   3885
      End
      Begin VB.Label lblVersion 
         BackStyle       =   0  '투명
         Caption         =   "버전"
         Height          =   225
         Left            =   930
         TabIndex        =   22
         Top             =   480
         Width           =   3885
      End
   End
   Begin VB.PictureBox pbPanel 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  '없음
      Height          =   4425
      Index           =   2
      Left            =   165
      ScaleHeight     =   4425
      ScaleWidth      =   6390
      TabIndex        =   4
      Top             =   450
      Visible         =   0   'False
      Width           =   6390
      Begin VB.PictureBox pbOptionContainer 
         BorderStyle     =   0  '없음
         Height          =   615
         Index           =   2
         Left            =   240
         ScaleHeight     =   615
         ScaleWidth      =   1680
         TabIndex        =   18
         Top             =   1440
         Width           =   1680
         Begin prjDownloadBooster.OptionButtonW optUserFore 
            Height          =   255
            Left            =   0
            TabIndex        =   20
            Top             =   330
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   450
            Caption         =   "사용자 지정(&T):"
            Transparent     =   -1  'True
         End
         Begin prjDownloadBooster.OptionButtonW optSystemFore 
            Height          =   255
            Left            =   0
            TabIndex        =   19
            Top             =   0
            Width           =   1815
            _ExtentX        =   0
            _ExtentY        =   0
            Caption         =   "시스템 색상(&Y)"
            Transparent     =   -1  'True
         End
      End
      Begin VB.PictureBox pbOptionContainer 
         BorderStyle     =   0  '없음
         Height          =   615
         Index           =   1
         Left            =   240
         ScaleHeight     =   615
         ScaleWidth      =   1680
         TabIndex        =   15
         Top             =   360
         Width           =   1680
         Begin prjDownloadBooster.OptionButtonW optSystemColor 
            Height          =   255
            Left            =   0
            TabIndex        =   17
            Top             =   0
            Width           =   1815
            _ExtentX        =   0
            _ExtentY        =   0
            Caption         =   "시스템 색상(&S)"
            Transparent     =   -1  'True
         End
         Begin prjDownloadBooster.OptionButtonW optUserColor 
            Height          =   255
            Left            =   0
            TabIndex        =   16
            Top             =   330
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   450
            Caption         =   "사용자 지정(&U):"
            Transparent     =   -1  'True
         End
      End
      Begin prjDownloadBooster.CheckBoxW chkNoDWMWindow 
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   2520
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   450
         Caption         =   "윈도우 7 모양으로 바꾸기(&I)"
         Transparent     =   -1  'True
      End
      Begin prjDownloadBooster.FrameW Frame3 
         Height          =   1935
         Left            =   120
         TabIndex        =   9
         Top             =   2280
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   3413
         Caption         =   " 창 모양 "
         Begin VB.PictureBox pbPreview 
            Height          =   1215
            Left            =   360
            ScaleHeight     =   1155
            ScaleWidth      =   5475
            TabIndex        =   28
            Top             =   600
            Width           =   5535
         End
      End
      Begin prjDownloadBooster.FrameW Frame1 
         Height          =   975
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   1720
         Caption         =   " 배경색 "
         Begin VB.Label lblSelectColor 
            BackStyle       =   0  '투명
            Height          =   495
            Left            =   1800
            TabIndex        =   12
            Top             =   480
            Width           =   1455
         End
         Begin VB.Shape pgColor 
            BackStyle       =   1  '투명하지 않음
            BorderColor     =   &H00404040&
            FillColor       =   &H00808080&
            FillStyle       =   2  '수평선
            Height          =   375
            Left            =   1800
            Shape           =   4  '둥근 사각형
            Top             =   510
            Width           =   1455
         End
      End
      Begin prjDownloadBooster.FrameW Frame4 
         Height          =   975
         Left            =   120
         TabIndex        =   13
         Top             =   1200
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   1720
         Caption         =   " 글자색 "
         Begin VB.Label lblSelectFore 
            BackStyle       =   0  '투명
            Height          =   495
            Left            =   1800
            TabIndex        =   14
            Top             =   510
            Width           =   1455
         End
         Begin VB.Shape pgFore 
            BackStyle       =   1  '투명하지 않음
            BorderColor     =   &H00404040&
            FillColor       =   &H00808080&
            FillStyle       =   2  '수평선
            Height          =   375
            Left            =   1800
            Shape           =   4  '둥근 사각형
            Top             =   510
            Width           =   1455
         End
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "적용(&A)"
      Enabled         =   0   'False
      Height          =   360
      Left            =   5280
      TabIndex        =   3
      Top             =   5040
      Width           =   1320
   End
   Begin prjDownloadBooster.TabStrip tsTabStrip 
      Height          =   4815
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   8493
      MultiRow        =   0   'False
      TabFixedWidth   =   53
      TabScrollWheel  =   0   'False
      Transparent     =   -1  'True
      InitTabs        =   "frmOptions.frx":5CFF
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "취소"
      Height          =   360
      Left            =   3840
      TabIndex        =   1
      Top             =   5040
      Width           =   1320
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "확인"
      Default         =   -1  'True
      Height          =   360
      Left            =   2400
      TabIndex        =   0
      Top             =   5040
      Width           =   1320
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"
Dim LineNum As Integer
Dim AboutEasterEgg2 As Boolean
Dim Loaded As Boolean

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub cbLanguage_Click()
    If Loaded Then
        MsgBox t("언어를 변경하려면 프로그램을 재시작해야 합니다.", "To change the language you must restart the application."), 64
        cmdApply.Enabled = -1
    End If
End Sub

Private Sub chkNoCleanup_Click()
    If Loaded Then cmdApply.Enabled = -1
End Sub

Private Sub chkNoDWMWindow_Click()
    If Loaded Then cmdApply.Enabled = -1
    
    If WinVer < 6# Then
        pbPreview.Picture = pbWinXP.Picture
        Exit Sub
    End If
    
    If chkNoDWMWindow.Value Then
        pbPreview.Picture = pbNoDWM.Picture
    Else
        If WinVer >= 10# Then
            pbPreview.Picture = pbDWM10.Picture
        ElseIf WinVer >= 6.2 Then
            pbPreview.Picture = pbDWM8.Picture
        Else
            pbPreview.Picture = pbDWM7.Picture
        End If
    End If
End Sub

Private Sub chkRememberURL_Click()
    If Loaded Then cmdApply.Enabled = -1
End Sub

Private Sub cmdApply_Click()
    SaveSetting "DownloadBooster", "Options", "NoCleanup", chkNoCleanup.Value
    If WinVer >= 6.1 Then SaveSetting "DownloadBooster", "Options", "DisableDWMWindow", chkNoDWMWindow.Value
    SaveSetting "DownloadBooster", "Options", "RememberURL", chkRememberURL.Value
    If chkNoDWMWindow.Enabled = True Then
        If chkNoDWMWindow.Value Then
            DisableDWMWindow Me.hWnd
            DisableDWMWindow frmMain.hWnd
        Else
            EnableDWMWindow Me.hWnd
            EnableDWMWindow frmMain.hWnd
        End If
    End If
    If optSystemColor.Value Then
        SaveSetting "DownloadBooster", "Options", "BackColor", "-1"
        pgColor.BackColor = &H8000000F
    ElseIf optUserColor.Value Then
        SaveSetting "DownloadBooster", "Options", "BackColor", CLng(pgColor.BackColor)
    End If
    If optSystemFore.Value Then
        SaveSetting "DownloadBooster", "Options", "ForeColor", "-1"
        pgFore.BackColor = &H80000012
    ElseIf optUserFore.Value Then
        SaveSetting "DownloadBooster", "Options", "ForeColor", CLng(pgFore.BackColor)
    End If
    SetFormBackgroundColor Me
    SetFormBackgroundColor frmMain
    If cbLanguage.ListIndex = 1 Then
        SaveSetting "DownloadBooster", "Options", "Language", 1033
    Else
        SaveSetting "DownloadBooster", "Options", "Language", 1042
    End If
    cmdApply.Enabled = 0
End Sub

Private Sub cmdSysInfo_Click()
    Call StartSysInfo
End Sub

Private Sub Form_Load()
    Loaded = False
    LineNum = 1
    AboutEasterEgg2 = False
    If GetSetting("DownloadBooster", "Options", "DisableDWMWindow", DefaultDisableDWMWindow) = 1 Then DisableDWMWindow Me.hWnd
    SetFormBackgroundColor Me
    SetFont Me
    
    Dim i%
    For i = 1 To pbPanel.Count
        If i <> 1 Then pbPanel(i).Visible = 0
        If i <> 2 Then
            pbPanel(i).Top = pbPanel(2).Top
            pbPanel(i).Left = pbPanel(2).Left
            pbPanel(i).Width = pbPanel(2).Width
            pbPanel(i).Height = pbPanel(2).Height
            pbPanel(i).BorderStyle = 0
            pbPanel(i).AutoRedraw = True
        End If
    Next i
    chkNoCleanup.Value = GetSetting("DownloadBooster", "Options", "NoCleanup", 0)
    chkNoDWMWindow.Value = GetSetting("DownloadBooster", "Options", "DisableDWMWindow", DefaultDisableDWMWindow)
    If (Not IsDWMEnabled()) Then
        chkNoDWMWindow.Enabled = False
        chkNoDWMWindow.Value = 1
    End If
    If WinVer < 6.1 Then chkNoDWMWindow.Value = 0
    chkRememberURL.Value = GetSetting("DownloadBooster", "Options", "RememberURL", 0)
    chkNoDWMWindow.Caption = t(chkNoDWMWindow.Caption, "Use W&indows 7 style")
    If WinVer < 6.2 Then chkNoDWMWindow.Caption = t("Aero 창 사용 안 함(&I)", "Disable Aero w&indow")
    
    Dim clrBackColor As Long
    clrBackColor = GetSetting("DownloadBooster", "Options", "BackColor", DefaultBackColor)
    If clrBackColor < 0 Or clrBackColor > 16777215 Then
        optSystemColor.Value = True
        pgColor.BackColor = &H8000000F
    Else
        optUserColor.Value = True
        pgColor.BackColor = clrBackColor
    End If
    
    Dim clrForeColor As Long
    clrForeColor = GetSetting("DownloadBooster", "Options", "ForeColor", -1)
    If clrForeColor < 0 Or clrForeColor > 16777215 Then
        optSystemFore.Value = True
        pgFore.BackColor = &H80000012
    Else
        optUserFore.Value = True
        pgFore.BackColor = clrForeColor
    End If
    
    cmdApply.Enabled = 0
    'On Error Resume Next
    Dim ctrl As Control
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "PictureBox" Then
            ctrl.AutoRedraw = True
            tsTabStrip.DrawBackground ctrl.hWnd, ctrl.hDC
        End If
    Next ctrl
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "FrameW" Then
            ctrl.Transparent = True
            ctrl.Refresh
        End If
    Next ctrl
    
    cbLanguage.AddItem "한국어"
    cbLanguage.AddItem "English"
    cbLanguage.ListIndex = CInt(GetSetting("DownloadBooster", "Options", "Language", GetUserDefaultLangID()) <> 1042) * -1
    
    picIcon.Picture = frmMain.Icon
    lblVersion.Caption = t("버전 ", "Version ") & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
    lblDescription.Caption = t("일괄 다운로드 및 다운로드 부스팅 프로그램" & vbCrLf & vbCrLf & "이 프로그램에는 Node.js의 바이너리가 포함되어 있으며," & vbCrLf & "라이선스 전문은 다음과 같습니다.", "Batch download booster application." & vbCrLf & vbCrLf & "This program includes the binary of Node.js." & vbCrLf & "Check out the license of Node.js below.")
    lblDescription.Width = pbPanel(1).Width - lblDescription.Left - 180
    txtLicense.Width = lblDescription.Width
    txtLicense.Height = pbPanel(1).Height - txtLicense.Top - 90 - cmdSysInfo.Height - 90 - pbLicenseLoadProgress.Height - 30
    txtLicensePlaceholder.Width = txtLicense.Width
    txtLicensePlaceholder.Height = txtLicense.Height
    txtLicensePlaceholder.Top = txtLicense.Top
    txtLicensePlaceholder.Left = txtLicense.Left
    pbLicenseLoadProgress.Width = txtLicense.Width
    pbLicenseLoadProgress.Top = txtLicense.Top + txtLicense.Height + 30
    pbLicenseLoadProgress.Left = txtLicense.Left
    cmdSysInfo.Top = pbPanel(1).Height - cmdSysInfo.Height - 90
    cmdSysInfo.Left = pbPanel(1).Width - cmdSysInfo.Width - 180
    
    chkNoDWMWindow_Click
    
    tsTabStrip.Tabs(1).Caption = t(tsTabStrip.Tabs(1).Caption, " General ")
    tsTabStrip.Tabs(2).Caption = t(tsTabStrip.Tabs(2).Caption, " Appearance ")
    tsTabStrip.Tabs(3).Caption = t(tsTabStrip.Tabs(3).Caption, " About ")
    Frame1.Caption = t(Frame1.Caption, " Background color ")
    Frame4.Caption = t(Frame4.Caption, " Text color ")
    Frame3.Caption = t(Frame3.Caption, " Window appearance ")
    Frame2.Caption = t(Frame2.Caption, " Download options ")
    Frame5.Caption = t(Frame5.Caption, " Interface ")
    chkNoCleanup.Caption = t(chkNoCleanup.Caption, "Preserve segme&nts")
    chkRememberURL.Caption = t(chkRememberURL.Caption, "Re&member URL")
    optSystemColor.Caption = t(optSystemColor.Caption, "&System color")
    optUserColor.Caption = t(optUserColor.Caption, "C&ustom color:")
    optSystemFore.Caption = t(optSystemFore.Caption, "S&ystem color")
    optUserFore.Caption = t(optUserFore.Caption, "Cus&tom color:")
    cmdSysInfo.Caption = t(cmdSysInfo.Caption, "&System information...")
    Label1.Caption = t(Label1.Caption, "&Language:")
    OKButton.Caption = t(OKButton.Caption, "OK")
    CancelButton.Caption = t(CancelButton.Caption, "Cancel")
    cmdApply.Caption = t(cmdApply.Caption, "&Apply")
    Me.Caption = t(Me.Caption, "Settings")
    
    Loaded = True
End Sub

Private Sub lblSelectColor_Click()
    Dim Color As OLE_COLOR
    Color = ShowColorDialog(Me.hWnd, True, pgColor.BackColor)
    If Color = -1 Then Exit Sub
    pgColor.BackColor = Color
    cmdApply.Enabled = -1
    optUserColor.Value = True
End Sub

Private Sub lblSelectFore_Click()
    Dim Color As OLE_COLOR
    Color = ShowColorDialog(Me.hWnd, True, pgFore.BackColor)
    If Color = -1 Then Exit Sub
    pgFore.BackColor = Color
    cmdApply.Enabled = -1
    optUserFore.Value = True
End Sub

Private Sub OKButton_Click()
    cmdApply_Click
    Unload Me
End Sub

Private Sub optSystemColor_Click()
    If Loaded Then cmdApply.Enabled = -1
End Sub

Private Sub optSystemFore_Click()
    If Loaded Then cmdApply.Enabled = -1
End Sub

Private Sub optUserColor_Click()
    If Loaded Then cmdApply.Enabled = -1
End Sub

Private Sub optUserFore_Click()
    If Loaded Then cmdApply.Enabled = -1
End Sub

Private Sub picIcon_DblClick()
    If frmMain.AboutEasterEgg And AboutEasterEgg2 And IsKeyPressed(gksKeyboardShift) And IsKeyPressed(gksKeyboardalt) And IsKeyPressed(gksKeyboardctrl) Then
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

Private Sub tsTabStrip_TabClick(ByVal TabItem As TbsTab)
    Dim i%
    For i = 1 To pbPanel.Count
        If i = TabItem.Index Then
            pbPanel(i).Visible = -1
        Else
            pbPanel(i).Visible = 0
        End If
    Next i
    
    If TabItem.Index = 3 And txtLicense.Text = "" Then
        timLicenseLoader.Enabled = -1
'        On Error GoTo LicenseFail
'        For i = 1 To 812
'            txtLicense.Text = txtLicense.Text & LoadResString(i) & vbCrLf
'        Next i
'        Exit Sub
'LicenseFail:
'        txtLicense.Text = "라이선스를 불러올 수 없습니다. 다음 링크에서 확인할 수 있습니다." & vbCrLf & " https://raw.githubusercontent.com/nodejs/node/refs/heads/v0.10/LICENSE"
    End If
    
    If IsKeyPressed(gksKeyboardShift) And IsKeyPressed(gksKeyboardalt) And IsKeyPressed(gksKeyboardctrl) Then
        AboutEasterEgg2 = -1
    Else
        AboutEasterEgg2 = 0
    End If
End Sub

Sub StartSysInfo()
    On Error GoTo SysInfoErr
  
    Dim RC As Long
    Dim SysInfoPath As String
    
    ' 시스템 정보 프로그램의 경로와 이름을 레지스트리에서 가져 옵니다...
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
    MsgBox t("지금은 시스템 정보를 사용할 수 없습니다.", "System Information is unavailable."), 48
End Sub
