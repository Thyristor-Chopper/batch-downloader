VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "스킨 설정"
   ClientHeight    =   6570
   ClientLeft      =   2760
   ClientTop       =   3855
   ClientWidth     =   13200
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
   ScaleHeight     =   6570
   ScaleWidth      =   13200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.PictureBox pbPanel 
      AutoRedraw      =   -1  'True
      Height          =   2865
      Index           =   1
      Left            =   6840
      ScaleHeight     =   2805
      ScaleWidth      =   6195
      TabIndex        =   5
      Top             =   2640
      Width           =   6255
      Begin prjDownloadBooster.FrameW Frame5 
         Height          =   675
         Left            =   120
         TabIndex        =   34
         Top             =   2040
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
         Top             =   1080
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   450
         Caption         =   "파일 주소 기억(&M)"
         Transparent     =   -1  'True
      End
      Begin prjDownloadBooster.CheckBoxW chkNoCleanup 
         Height          =   255
         Left            =   3000
         TabIndex        =   8
         Top             =   600
         Width           =   2250
         _ExtentX        =   3969
         _ExtentY        =   450
         Caption         =   "조각 파일 유지(&N)"
         Transparent     =   -1  'True
      End
      Begin prjDownloadBooster.FrameW Frame2 
         Height          =   1815
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   3201
         Caption         =   " 다운로드 설정 "
         Begin VB.ComboBox cbWhenExist 
            Height          =   300
            Left            =   2160
            Style           =   2  '드롭다운 목록
            TabIndex        =   46
            Top             =   1320
            Width           =   1455
         End
         Begin prjDownloadBooster.CheckBoxW chkAutoRetry 
            Height          =   255
            Left            =   2880
            TabIndex        =   44
            Top             =   720
            Width           =   3045
            _ExtentX        =   5371
            _ExtentY        =   450
            Caption         =   "네트워크 오류 시 자동 재시도(&U)"
            Transparent     =   -1  'True
         End
         Begin prjDownloadBooster.CheckBoxW chkAlwaysResume 
            Height          =   255
            Left            =   120
            TabIndex        =   43
            Top             =   720
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   450
            Caption         =   "항상 이어받기(&A)"
            Transparent     =   -1  'True
         End
         Begin prjDownloadBooster.CheckBoxW chkBeepWhenComplete 
            Height          =   255
            Left            =   120
            TabIndex        =   42
            Top             =   480
            Width           =   3255
            _ExtentX        =   0
            _ExtentY        =   0
            Caption         =   "완료 후 신호음 재생(&B)"
            Transparent     =   -1  'True
         End
         Begin prjDownloadBooster.CheckBoxW chkOpenDirWhenComplete 
            Height          =   255
            Left            =   2880
            TabIndex        =   41
            Top             =   240
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   450
            Caption         =   "완료 후 폴더 열기(&P)"
            Transparent     =   -1  'True
         End
         Begin prjDownloadBooster.CheckBoxW chkOpenWhenComplete 
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   240
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   450
            Caption         =   "완료 후 파일 열기(&O)"
            Transparent     =   -1  'True
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '투명
            Caption         =   "중복 파일명 처리(&D):"
            Height          =   255
            Left            =   120
            TabIndex        =   45
            Top             =   1365
            Width           =   1935
         End
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
      Height          =   2415
      Index           =   3
      Left            =   6840
      ScaleHeight     =   2355
      ScaleWidth      =   4395
      TabIndex        =   21
      Top             =   120
      Width           =   4455
      Begin prjDownloadBooster.LinkLabel lblReadOnline 
         Height          =   255
         Left            =   2400
         TabIndex        =   47
         Top             =   1920
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         Caption         =   "frmOptions.frx":5CFF
         Transparent     =   -1  'True
      End
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
         Height          =   615
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
      Begin prjDownloadBooster.FrameW FrameW1 
         Height          =   975
         Left            =   3120
         TabIndex        =   48
         Top             =   2160
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   1720
         Caption         =   " 배경 그림 "
         Transparent     =   -1  'True
         Begin VB.CommandButton cmdChooseBackground 
            Caption         =   "그림 선택(&C)..."
            Height          =   330
            Left            =   360
            TabIndex        =   50
            Top             =   510
            Width           =   1935
         End
         Begin prjDownloadBooster.CheckBoxW chkEnableBackgroundImage 
            Height          =   255
            Left            =   120
            TabIndex        =   49
            Top             =   240
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   450
            Caption         =   "배경 그림 사용(&B)"
            Transparent     =   -1  'True
         End
      End
      Begin prjDownloadBooster.FrameW Frame6 
         Height          =   735
         Left            =   3120
         TabIndex        =   37
         Top             =   3240
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   1296
         Caption         =   " 기타 설정 "
         Transparent     =   -1  'True
         Begin VB.ComboBox cbTabStyle 
            Height          =   300
            Left            =   1200
            Style           =   2  '드롭다운 목록
            TabIndex        =   39
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '투명
            Caption         =   "탭 모양(&E):"
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   285
            Width           =   1080
         End
      End
      Begin VB.PictureBox pbOptionContainer 
         BorderStyle     =   0  '없음
         Height          =   615
         Index           =   2
         Left            =   240
         ScaleHeight     =   615
         ScaleWidth      =   1680
         TabIndex        =   18
         Top             =   3480
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
         Top             =   2400
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
         Top             =   360
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
         Top             =   120
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
         Top             =   2160
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   1720
         Caption         =   " 배경색 "
         Begin VB.Label lblSelectColor 
            BackStyle       =   0  '투명
            Height          =   255
            Left            =   1800
            TabIndex        =   12
            Top             =   240
            Width           =   1455
         End
         Begin VB.Shape pgColor 
            BackStyle       =   1  '투명하지 않음
            BorderColor     =   &H00404040&
            FillColor       =   &H00808080&
            Height          =   255
            Left            =   1800
            Shape           =   4  '둥근 사각형
            Top             =   585
            Width           =   855
         End
      End
      Begin prjDownloadBooster.FrameW Frame4 
         Height          =   975
         Left            =   120
         TabIndex        =   13
         Top             =   3240
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   1720
         Caption         =   " 글자색 "
         Begin VB.Label lblSelectFore 
            BackStyle       =   0  '투명
            Height          =   255
            Left            =   1800
            TabIndex        =   14
            Top             =   240
            Width           =   1455
         End
         Begin VB.Shape pgFore 
            BackStyle       =   1  '투명하지 않음
            BorderColor     =   &H00404040&
            FillColor       =   &H00808080&
            Height          =   255
            Left            =   1800
            Shape           =   4  '둥근 사각형
            Top             =   585
            Width           =   855
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
      InitTabs        =   "frmOptions.frx":5D41
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
Dim ColorChanged As Boolean
Dim TabStyleChanged As Boolean
Public ImageChanged As Boolean

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub cbLanguage_Click()
    If Loaded Then
        Alert t("언어를 변경하려면 프로그램을 재시작해야 합니다.", "To change the language you must restart the application."), App.Title, Me, 64
        cmdApply.Enabled = -1
    End If
End Sub

Private Sub cbTabStyle_Click()
    If Loaded Then
        cmdApply.Enabled = -1
        TabStyleChanged = True
    End If
End Sub

Private Sub cbWhenExist_Click()
    If Loaded Then cmdApply.Enabled = -1
End Sub

Private Sub chkAlwaysResume_Click()
    If Loaded Then cmdApply.Enabled = -1
End Sub

Private Sub chkAutoRetry_Click()
    If Loaded Then cmdApply.Enabled = -1
End Sub

Private Sub chkBeepWhenComplete_Click()
    If Loaded Then cmdApply.Enabled = -1
End Sub

Private Sub chkEnableBackgroundImage_Click()
    If Loaded Then
        cmdApply.Enabled = -1
        ImageChanged = True
    End If
    
    If chkEnableBackgroundImage.Value = 0 Then
        cmdChooseBackground.Enabled = 0
    Else
        cmdChooseBackground.Enabled = -1
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

Private Sub chkOpenDirWhenComplete_Click()
    If Loaded Then cmdApply.Enabled = -1
End Sub

Private Sub chkOpenWhenComplete_Click()
    If Loaded Then cmdApply.Enabled = -1
End Sub

Private Sub chkRememberURL_Click()
    If Loaded Then cmdApply.Enabled = -1
End Sub

Private Sub cmdApply_Click()
    If WinVer >= 6.1 And chkNoDWMWindow.Enabled Then SaveSetting "DownloadBooster", "Options", "DisableDWMWindow", chkNoDWMWindow.Value
    SaveSetting "DownloadBooster", "Options", "NoCleanup", chkNoCleanup.Value
    SaveSetting "DownloadBooster", "Options", "RememberURL", chkRememberURL.Value
    
    SaveSetting "DownloadBooster", "Options", "OpenWhenComplete", chkOpenWhenComplete.Value
    SaveSetting "DownloadBooster", "Options", "OpenFolderWhenComplete", chkOpenDirWhenComplete.Value
    SaveSetting "DownloadBooster", "Options", "PlaySound", chkBeepWhenComplete.Value
    SaveSetting "DownloadBooster", "Options", "ContinueDownload", chkAlwaysResume.Value
    SaveSetting "DownloadBooster", "Options", "AutoRetry", chkAutoRetry.Value
    SaveSetting "DownloadBooster", "Options", "WhenFileExists", cbWhenExist.ListIndex
    frmMain.cbWhenExist.ListIndex = cbWhenExist.ListIndex
    
    frmMain.chkOpenAfterComplete.Value = chkOpenWhenComplete.Value
    frmMain.chkOpenFolder.Value = chkOpenDirWhenComplete.Value
    frmMain.chkPlaySound.Value = chkBeepWhenComplete.Value
    frmMain.chkContinueDownload.Value = chkAlwaysResume.Value
    frmMain.chkAutoRetry.Value = chkAutoRetry.Value
    
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
    If ColorChanged Then
        SetFormBackgroundColor Me
        SetFormBackgroundColor frmMain
    End If
    If cbLanguage.ListIndex = 1 Then
        SaveSetting "DownloadBooster", "Options", "Language", 1033
    Else
        SaveSetting "DownloadBooster", "Options", "Language", 1042
    End If
    SaveSetting "DownloadBooster", "Options", "TabStyle", cbTabStyle.ListIndex
    If TabStyleChanged Then frmMain.SetTabStyle
    If ImageChanged Then
        SaveSetting "DownloadBooster", "Options", "UseBackgroundImage", chkEnableBackgroundImage.Value
        frmMain.SetBackgroundImage
    End If
    
    ColorChanged = False
    TabStyleChanged = False
    cmdApply.Enabled = 0
End Sub

Private Sub cmdChooseBackground_Click()
    frmCustomBackground.Show vbModal, Me
End Sub

Private Sub cmdSysInfo_Click()
    Call StartSysInfo
End Sub

Private Sub Form_Load()
    Loaded = False
    ImageChanged = False
    ColorChanged = False
    LineNum = 1
    AboutEasterEgg2 = False
    TabStyleChanged = False
    If GetSetting("DownloadBooster", "Options", "DisableDWMWindow", DefaultDisableDWMWindow) = 1 Then DisableDWMWindow Me.hWnd
    SetFormBackgroundColor Me
    SetFont Me
    
    Me.Width = 6840
    Me.Height = 5970
    
    lblSelectColor.Top = pgColor.Top
    lblSelectColor.Left = pgColor.Left
    lblSelectColor.Width = pgColor.Width
    lblSelectColor.Height = pgColor.Height
    
    lblSelectFore.Top = pgFore.Top
    lblSelectFore.Left = pgFore.Left
    lblSelectFore.Width = pgFore.Width
    lblSelectFore.Height = pgFore.Height
    
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
    If WinVer < 6# Then
        chkNoDWMWindow.Caption = t("DWM 창 비활성화(&I)", "Disable DWM w&indow")
        chkNoDWMWindow.Value = 1
    ElseIf WinVer < 6.2 Then
        chkNoDWMWindow.Caption = t("Aero 창 사용 안 함(&I)", "Disable Aero w&indow")
    End If
    
    chkEnableBackgroundImage.Value = GetSetting("DownloadBooster", "Options", "UseBackgroundImage", 0)
    If chkEnableBackgroundImage.Value = 0 Then cmdChooseBackground.Enabled = 0
    
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
    
    chkOpenWhenComplete.Value = frmMain.chkOpenAfterComplete.Value
    chkOpenDirWhenComplete.Value = frmMain.chkOpenFolder.Value
    chkBeepWhenComplete.Value = frmMain.chkPlaySound.Value
    chkAlwaysResume.Value = frmMain.chkContinueDownload.Value
    chkAutoRetry.Value = frmMain.chkAutoRetry.Value
    
    cbLanguage.Clear
    cbLanguage.AddItem "한국어"
    cbLanguage.AddItem "English"
    cbLanguage.ListIndex = CInt(GetSetting("DownloadBooster", "Options", "Language", GetUserDefaultLangID()) <> 1042) * -1
    
    cbTabStyle.Clear
    cbTabStyle.AddItem t("단추형 탭", "Buttoned tabs")
    cbTabStyle.AddItem t("납작이 탭", "Flat buttons")
    cbTabStyle.AddItem t("일반 탭", "Normal tabs")
    cbTabStyle.AddItem t("단추", "Push buttons")
    cbTabStyle.AddItem t("라디오", "Radio")
    cbTabStyle.ListIndex = CInt(GetSetting("DownloadBooster", "Options", "TabStyle", 4))
    
    cbWhenExist.Clear
    cbWhenExist.AddItem t("중단", "Abort")
    cbWhenExist.AddItem t("덮어쓰기", "Overwrite")
    cbWhenExist.AddItem t("이름 변경", "Rename")
    cbWhenExist.ListIndex = GetSetting("DownloadBooster", "Options", "WhenFileExists", 0)
    
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
    lblReadOnline.Top = txtLicense.Top + txtLicense.Height + 30 + pbLicenseLoadProgress.Height + 60
    lblReadOnline.Left = txtLicense.Left
    
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
    Frame6.Caption = t(Frame6.Caption, " Other settings ")
    Label2.Caption = t(Label2.Caption, "Tab styl&e:")
    chkOpenWhenComplete.Caption = t(chkOpenWhenComplete.Caption, "&Open file when complete")
    chkOpenDirWhenComplete.Caption = t(chkOpenDirWhenComplete.Caption, "O&pen folder when complete")
    chkBeepWhenComplete.Caption = t(chkBeepWhenComplete.Caption, "&Beep when complete")
    chkAlwaysResume.Caption = t(chkAlwaysResume.Caption, "&Always resume")
    chkAutoRetry.Caption = t(chkAutoRetry.Caption, "A&uto retry on network error")
    Label3.Caption = t(Label3.Caption, "If filename alrea&dy exists:")
    lblReadOnline.Caption = t(lblReadOnline.Caption, "<A>[Read online]</A>")
    FrameW1.Caption = t(FrameW1.Caption, " Background image ")
    chkEnableBackgroundImage.Caption = t(chkEnableBackgroundImage.Caption, "Use &background image")
    cmdChooseBackground.Caption = t(cmdChooseBackground.Caption, "&Choose image...")
    
    Loaded = True
End Sub

Private Sub lblReadOnline_LinkActivate(ByVal Link As LlbLink, ByVal Reason As LlbLinkActivateReasonConstants)
    Shell "cmd /c start """" https://raw.githubusercontent.com/nodejs/node/refs/heads/v0.10/LICENSE"
End Sub

Private Sub lblSelectColor_Click()
    Dim Color As OLE_COLOR
    Color = ShowColorDialog(Me.hWnd, True, pgColor.BackColor)
    If Color = -1 Then Exit Sub
    pgColor.BackColor = Color
    cmdApply.Enabled = -1
    optUserColor.Value = True
    ColorChanged = True
End Sub

Private Sub lblSelectFore_Click()
    Dim Color As OLE_COLOR
    Color = ShowColorDialog(Me.hWnd, True, pgFore.BackColor)
    If Color = -1 Then Exit Sub
    pgFore.BackColor = Color
    cmdApply.Enabled = -1
    optUserFore.Value = True
    ColorChanged = True
End Sub

Private Sub OKButton_Click()
    cmdApply_Click
    Unload Me
End Sub

Private Sub optSystemColor_Click()
    If Loaded Then
        cmdApply.Enabled = -1
        ColorChanged = True
    End If
End Sub

Private Sub optSystemFore_Click()
    If Loaded Then
        cmdApply.Enabled = -1
        ColorChanged = True
    End If
End Sub

Private Sub optUserColor_Click()
    If Loaded Then
        cmdApply.Enabled = -1
        ColorChanged = True
    End If
End Sub

Private Sub optUserFore_Click()
    If Loaded Then
        cmdApply.Enabled = -1
        ColorChanged = True
    End If
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
    Alert t("지금은 시스템 정보를 사용할 수 없습니다.", "System Information is unavailable."), App.Title, Me, 48
End Sub
