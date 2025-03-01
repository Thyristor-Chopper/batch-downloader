VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "옵션"
   ClientHeight    =   12825
   ClientLeft      =   2760
   ClientTop       =   3855
   ClientWidth     =   12975
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
   ScaleHeight     =   12825
   ScaleWidth      =   12975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.PictureBox pbPanel 
      Height          =   4425
      Index           =   4
      Left            =   120
      ScaleHeight     =   4365
      ScaleWidth      =   6315
      TabIndex        =   74
      Top             =   8160
      Width           =   6375
      Begin prjDownloadBooster.FrameW fCompleteSound 
         Height          =   735
         Left            =   720
         TabIndex        =   76
         Top             =   720
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   1296
         BorderStyle     =   0
         Caption         =   "                             "
         Transparent     =   -1  'True
         Begin prjDownloadBooster.CommandButtonW cmdBrowseCompleteSound 
            Height          =   300
            Left            =   4560
            TabIndex        =   79
            Top             =   330
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   529
            ImageList       =   "imgBrowse"
            ImageListAlignment=   4
         End
         Begin prjDownloadBooster.TextBoxW txtCompleteSoundPath 
            Height          =   300
            Left            =   360
            TabIndex        =   78
            Top             =   330
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   529
         End
         Begin prjDownloadBooster.CheckBoxW chkBeepWhenComplete 
            Height          =   255
            Left            =   120
            TabIndex        =   77
            Top             =   0
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   450
            Caption         =   "다운로드 완료(&B)"
            Transparent     =   -1  'True
         End
         Begin prjDownloadBooster.CommandButtonW cmdTestCompleteSound 
            Height          =   300
            Left            =   5160
            TabIndex        =   80
            Top             =   330
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "▶"
         End
      End
      Begin prjDownloadBooster.FrameW fAsterisk 
         Height          =   735
         Left            =   720
         TabIndex        =   81
         Top             =   1440
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   1296
         BorderStyle     =   0
         Caption         =   "                             "
         Transparent     =   -1  'True
         Begin prjDownloadBooster.CommandButtonW cmdBrowseAsterisk 
            Height          =   300
            Left            =   4560
            TabIndex        =   82
            Top             =   330
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   529
            ImageList       =   "imgBrowse"
            ImageListAlignment=   4
            Caption         =   "..."
         End
         Begin prjDownloadBooster.TextBoxW txtAsterisk 
            Height          =   300
            Left            =   360
            TabIndex        =   83
            Top             =   330
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   529
         End
         Begin prjDownloadBooster.CheckBoxW chkAsterisk 
            Height          =   255
            Left            =   120
            TabIndex        =   84
            Top             =   0
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   450
            Caption         =   "일반 메시지(&A)"
            Transparent     =   -1  'True
         End
         Begin prjDownloadBooster.CommandButtonW cmdTestAsterisk 
            Height          =   300
            Left            =   5160
            TabIndex        =   85
            Top             =   330
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "▶"
         End
      End
      Begin prjDownloadBooster.FrameW fExclamation 
         Height          =   735
         Left            =   720
         TabIndex        =   86
         Top             =   2160
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   1296
         BorderStyle     =   0
         Caption         =   "                             "
         Transparent     =   -1  'True
         Begin prjDownloadBooster.CommandButtonW cmdBrowseExclamation 
            Height          =   300
            Left            =   4560
            TabIndex        =   87
            Top             =   330
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   529
            ImageList       =   "imgBrowse"
            ImageListAlignment=   4
         End
         Begin prjDownloadBooster.TextBoxW txtExclamation 
            Height          =   300
            Left            =   360
            TabIndex        =   88
            Top             =   330
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   529
         End
         Begin prjDownloadBooster.CheckBoxW chkExclamation 
            Height          =   255
            Left            =   120
            TabIndex        =   89
            Top             =   0
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   450
            Caption         =   "경고 메시지(&E)"
            Transparent     =   -1  'True
         End
         Begin prjDownloadBooster.CommandButtonW cmdTestExclamation 
            Height          =   300
            Left            =   5160
            TabIndex        =   90
            Top             =   330
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "▶"
         End
      End
      Begin prjDownloadBooster.FrameW fError 
         Height          =   735
         Left            =   720
         TabIndex        =   91
         Top             =   2880
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   1296
         BorderStyle     =   0
         Caption         =   "                             "
         Transparent     =   -1  'True
         Begin prjDownloadBooster.CommandButtonW cmdBrowseError 
            Height          =   300
            Left            =   4560
            TabIndex        =   92
            Top             =   330
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   529
            ImageList       =   "imgBrowse"
            ImageListAlignment=   4
         End
         Begin prjDownloadBooster.TextBoxW txtError 
            Height          =   300
            Left            =   360
            TabIndex        =   93
            Top             =   330
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   529
         End
         Begin prjDownloadBooster.CheckBoxW chkError 
            Height          =   255
            Left            =   120
            TabIndex        =   94
            Top             =   0
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   450
            Caption         =   "오류 메시지(&R)"
            Transparent     =   -1  'True
         End
         Begin prjDownloadBooster.CommandButtonW cmdTestError 
            Height          =   300
            Left            =   5160
            TabIndex        =   95
            Top             =   330
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "▶"
         End
      End
      Begin prjDownloadBooster.FrameW fQuestion 
         Height          =   735
         Left            =   720
         TabIndex        =   96
         Top             =   3600
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   1296
         BorderStyle     =   0
         Caption         =   "                             "
         Transparent     =   -1  'True
         Begin prjDownloadBooster.CommandButtonW cmdBrowseQuestion 
            Height          =   300
            Left            =   4560
            TabIndex        =   97
            Top             =   330
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   529
            ImageList       =   "imgBrowse"
            ImageListAlignment=   4
         End
         Begin prjDownloadBooster.TextBoxW txtQuestion 
            Height          =   300
            Left            =   360
            TabIndex        =   98
            Top             =   330
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   529
         End
         Begin prjDownloadBooster.CheckBoxW chkQuestion 
            Height          =   255
            Left            =   120
            TabIndex        =   99
            Top             =   0
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   450
            Caption         =   "질문(&Q)"
            Transparent     =   -1  'True
         End
         Begin prjDownloadBooster.CommandButtonW cmdTestQuestion 
            Height          =   300
            Left            =   5160
            TabIndex        =   100
            Top             =   330
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "▶"
         End
      End
      Begin VB.Image imgIcon1 
         Height          =   480
         Left            =   120
         Picture         =   "frmOptions.frx":000C
         Stretch         =   -1  'True
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label12 
         BackStyle       =   0  '투명
         Caption         =   "기본값을 사용하려면 필드를 비워두십시오."
         Height          =   255
         Left            =   840
         TabIndex        =   75
         Top             =   240
         Width           =   4815
      End
   End
   Begin prjDownloadBooster.ImageList imgBrowse 
      Left            =   12840
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      ImageWidth      =   16
      ImageHeight     =   16
      ColorDepth      =   8
      MaskColor       =   16711935
      InitListImages  =   "frmOptions.frx":036F
   End
   Begin VB.PictureBox pbPanel 
      AutoRedraw      =   -1  'True
      Height          =   4545
      Index           =   2
      Left            =   6480
      ScaleHeight     =   4485
      ScaleWidth      =   5955
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   5280
      Width           =   6015
      Begin prjDownloadBooster.FrameW FrameW3 
         Height          =   855
         Left            =   120
         TabIndex        =   44
         Top             =   120
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   1508
         Caption         =   " 연결 설정 "
         Begin prjDownloadBooster.CheckBoxW chkIgnore300 
            Height          =   255
            Left            =   3000
            TabIndex        =   45
            Top             =   240
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   450
            Caption         =   "300번대 응답 코드 무시(&I)"
            Transparent     =   -1  'True
         End
         Begin prjDownloadBooster.CheckBoxW chkForceGet 
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   480
            Width           =   2505
            _ExtentX        =   4419
            _ExtentY        =   450
            Caption         =   "파일 검사 시 GET 요청(&Q)"
            Transparent     =   -1  'True
         End
         Begin prjDownloadBooster.CheckBoxW chkNoRedirectCheck 
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   240
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   450
            Caption         =   "리다이렉트 검사 안 함(&R)"
            Transparent     =   -1  'True
         End
      End
      Begin prjDownloadBooster.FrameW fHeaders 
         Height          =   3375
         Left            =   120
         TabIndex        =   48
         Top             =   1080
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   5953
         Caption         =   " 헤더 설정 "
         Transparent     =   -1  'True
         Begin prjDownloadBooster.CommandButtonW cmdEditHeaderName 
            Height          =   330
            Left            =   2760
            TabIndex        =   49
            Top             =   2970
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   582
            Enabled         =   0   'False
            Caption         =   "이름변경(&R)"
         End
         Begin prjDownloadBooster.TextBoxW txtEdit 
            Height          =   255
            Left            =   2640
            TabIndex        =   50
            Top             =   360
            Visible         =   0   'False
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   450
            BorderStyle     =   1
         End
         Begin prjDownloadBooster.CommandButtonW cmdDeleteHeader 
            Height          =   330
            Left            =   1440
            TabIndex        =   51
            Top             =   2970
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   582
            Enabled         =   0   'False
            Caption         =   "삭제(&D)"
         End
         Begin prjDownloadBooster.CommandButtonW cmdEditHeaderValue 
            Height          =   330
            Left            =   4080
            TabIndex        =   52
            Top             =   2970
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   582
            Enabled         =   0   'False
            Caption         =   "편집(&E)"
         End
         Begin prjDownloadBooster.CommandButtonW cmdAddHeader 
            Height          =   330
            Left            =   120
            TabIndex        =   53
            Top             =   2970
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   582
            Caption         =   "추가(&A)"
         End
         Begin prjDownloadBooster.ListView lvHeaders 
            Height          =   2655
            Left            =   120
            TabIndex        =   54
            Top             =   240
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   4683
            VisualTheme     =   1
            View            =   3
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            HideSelection   =   0   'False
            ShowLabelTips   =   -1  'True
            HighlightColumnHeaders=   -1  'True
            AutoSelectFirstItem=   0   'False
         End
      End
   End
   Begin VB.PictureBox pbPanel 
      AutoRedraw      =   -1  'True
      Height          =   4665
      Index           =   1
      Left            =   120
      ScaleHeight     =   4605
      ScaleWidth      =   5595
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   600
      Width           =   5655
      Begin prjDownloadBooster.FrameW Frame5 
         Height          =   2475
         Left            =   120
         TabIndex        =   12
         Top             =   1920
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   4366
         Caption         =   " 인터페이스 "
         Transparent     =   -1  'True
         Begin prjDownloadBooster.CheckBoxW chkAllowDuplicates 
            Height          =   255
            Left            =   120
            TabIndex        =   101
            Top             =   1440
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   450
            Caption         =   "일괄 처리 목록에 중복 항목 허용(&I)"
            Transparent     =   -1  'True
         End
         Begin prjDownloadBooster.CheckBoxW chkDontLoadIcons 
            Height          =   255
            Left            =   120
            TabIndex        =   58
            Top             =   1200
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   450
            Caption         =   "열기 대화 상자에서 같은 파일 아이콘 사용(&M)"
            Transparent     =   -1  'True
         End
         Begin prjDownloadBooster.CheckBoxW chkForceOldDialog 
            Height          =   255
            Left            =   120
            TabIndex        =   57
            Top             =   960
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   450
            Caption         =   "윈도우 3.1 대화 상자 사용(&S)"
            Transparent     =   -1  'True
         End
         Begin prjDownloadBooster.CheckBoxW chkExcludeMergeFromElapsed 
            Height          =   255
            Left            =   120
            TabIndex        =   42
            Top             =   720
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   450
            Caption         =   "경과 시간에서 파일 조각 결합 시간 제외(&E)"
            Transparent     =   -1  'True
         End
         Begin prjDownloadBooster.CheckBoxW chkLazyElapsed 
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   480
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   450
            Caption         =   "첫 바이트 수신 후 경과 시간 계산(&C)"
            Transparent     =   -1  'True
         End
         Begin prjDownloadBooster.CheckBoxW chkAeroWindow 
            Height          =   255
            Left            =   2160
            TabIndex        =   40
            Top             =   240
            Visible         =   0   'False
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   450
            Enabled         =   0   'False
            Caption         =   "유리 창 효과 사용(&R)"
            Transparent     =   -1  'True
         End
         Begin prjDownloadBooster.CheckBoxW chkAlwaysOnTop 
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   240
            Width           =   1920
            _ExtentX        =   3387
            _ExtentY        =   450
            Caption         =   "항상 위에 표시(&W)"
            Transparent     =   -1  'True
         End
         Begin prjDownloadBooster.ComboBoxW cbLanguage 
            Height          =   300
            Left            =   1440
            TabIndex        =   14
            Top             =   1800
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   529
            Style           =   2
         End
         Begin VB.Label Label9 
            BackStyle       =   0  '투명
            Caption         =   "..."
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   2160
            Width           =   4200
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '투명
            Caption         =   "언어(&L):"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Tag             =   "nocolorchange"
            Top             =   1845
            Width           =   975
         End
      End
      Begin prjDownloadBooster.CheckBoxW chkRememberURL 
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   450
         Caption         =   "파일 주소 기억(&M)"
         Transparent     =   -1  'True
      End
      Begin prjDownloadBooster.FrameW Frame2 
         Height          =   1695
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   2990
         Caption         =   " 다운로드 설정 "
         Begin prjDownloadBooster.CheckBoxW chkAutoYtdl 
            Height          =   255
            Left            =   120
            TabIndex        =   59
            Top             =   960
            Width           =   4815
            _ExtentX        =   8493
            _ExtentY        =   450
            Caption         =   "지원되는 링크에서 자동으로 youtube-dl 사용(&Y)"
            Transparent     =   -1  'True
         End
         Begin prjDownloadBooster.ComboBoxW cbWhenExist 
            Height          =   300
            Left            =   2055
            TabIndex        =   21
            Top             =   1320
            Width           =   2040
            _ExtentX        =   3598
            _ExtentY        =   529
            Style           =   2
         End
         Begin prjDownloadBooster.CheckBoxW chkAutoRetry 
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   720
            Width           =   4890
            _ExtentX        =   8625
            _ExtentY        =   450
            Caption         =   "네트워크 오류 시 자동 재시도(&U)"
            Transparent     =   -1  'True
         End
         Begin prjDownloadBooster.CheckBoxW chkAlwaysResume 
            Height          =   255
            Left            =   2640
            TabIndex        =   18
            Top             =   480
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   450
            Caption         =   "항상 이어받기(&A)"
            Transparent     =   -1  'True
         End
         Begin prjDownloadBooster.CheckBoxW chkOpenDirWhenComplete 
            Height          =   255
            Left            =   2640
            TabIndex        =   17
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
            TabIndex        =   16
            Top             =   240
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   450
            Caption         =   "완료 후 파일 열기(&O)"
            Transparent     =   -1  'True
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '투명
            Caption         =   "중복 파일명 처리(&D):"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Tag             =   "nocolorchange"
            Top             =   1365
            Width           =   1935
         End
      End
   End
   Begin VB.PictureBox pbPanel 
      AutoRedraw      =   -1  'True
      Height          =   2895
      Index           =   5
      Left            =   120
      ScaleHeight     =   2835
      ScaleWidth      =   6195
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   5280
      Width           =   6255
      Begin prjDownloadBooster.FrameW FrameW4 
         Height          =   615
         Left            =   120
         TabIndex        =   55
         Top             =   2160
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   1085
         Caption         =   " 고급 다운로드 설정 "
         Transparent     =   -1  'True
         Begin prjDownloadBooster.CheckBoxW chkNoCleanup 
            Height          =   255
            Left            =   120
            TabIndex        =   56
            Top             =   240
            Width           =   2250
            _ExtentX        =   3969
            _ExtentY        =   450
            Caption         =   "조각 파일 유지(&N)"
            Transparent     =   -1  'True
         End
      End
      Begin prjDownloadBooster.FrameW FrameW2 
         Height          =   1935
         Left            =   120
         TabIndex        =   26
         Top             =   120
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   3413
         Caption         =   " 경로 설정 "
         Transparent     =   -1  'True
         Begin prjDownloadBooster.TextBoxW txtYtdlPath 
            Height          =   255
            Left            =   2040
            TabIndex        =   31
            Top             =   1560
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   450
         End
         Begin prjDownloadBooster.TextBoxW txtNodePath 
            Height          =   255
            Left            =   2040
            TabIndex        =   28
            Top             =   840
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   450
         End
         Begin prjDownloadBooster.TextBoxW txtScriptPath 
            Height          =   255
            Left            =   2040
            TabIndex        =   29
            Top             =   1200
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   450
         End
         Begin VB.Image imgIcon2 
            Height          =   480
            Left            =   120
            Picture         =   "frmOptions.frx":0757
            Stretch         =   -1  'True
            Top             =   240
            Width           =   480
         End
         Begin VB.Label Label6 
            BackStyle       =   0  '투명
            Caption         =   "기본값을 사용하려면 필드를 비워두십시오. 이 옵션은 고급 사용자를 위한 것이며 일반적으로 변경할 필요가 없습니다."
            Height          =   480
            Left            =   720
            TabIndex        =   37
            Top             =   300
            Width           =   5175
         End
         Begin VB.Label Label7 
            BackStyle       =   0  '투명
            Caption         =   "&youtube-dl/yt-dlp:"
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   1590
            Width           =   1695
         End
         Begin VB.Label Label5 
            BackStyle       =   0  '투명
            Caption         =   "다운로드 스크립트(&D):"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   1230
            Width           =   1935
         End
         Begin VB.Label Label4 
            BackStyle       =   0  '투명
            Caption         =   "N&ode.js:"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   870
            Width           =   1455
         End
      End
   End
   Begin VB.PictureBox pbPanel 
      AutoRedraw      =   -1  'True
      Height          =   4545
      Index           =   3
      Left            =   6360
      ScaleHeight     =   4485
      ScaleWidth      =   6315
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   600
      Visible         =   0   'False
      Width           =   6375
      Begin prjDownloadBooster.FrameW Frame6 
         Height          =   975
         Left            =   3240
         TabIndex        =   15
         Top             =   3480
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   1720
         Caption         =   " 스킨 "
         Transparent     =   -1  'True
         Begin prjDownloadBooster.ComboBoxW cbSkin 
            Height          =   300
            Left            =   870
            TabIndex        =   36
            Top             =   600
            Width           =   1965
            _ExtentX        =   3466
            _ExtentY        =   529
            Style           =   2
            Text            =   "ComboBoxW1"
         End
         Begin prjDownloadBooster.ComboBoxW cbFrameSkin 
            Height          =   300
            Left            =   870
            TabIndex        =   64
            Top             =   240
            Width           =   1965
            _ExtentX        =   3466
            _ExtentY        =   529
            Style           =   2
            Text            =   "ComboBoxW1"
         End
         Begin VB.Label Label10 
            BackStyle       =   0  '투명
            Caption         =   "창(&W):"
            Height          =   255
            Left            =   120
            TabIndex        =   65
            Top             =   285
            Width           =   735
         End
         Begin VB.Label Label8 
            BackStyle       =   0  '투명
            Caption         =   "단추(&O):"
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   645
            Width           =   735
         End
      End
      Begin prjDownloadBooster.FrameW Frame4 
         Height          =   975
         Left            =   120
         TabIndex        =   10
         Top             =   3480
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   1720
         Caption         =   " 글자색 "
         Begin prjDownloadBooster.OptionButtonW optUserFore 
            Height          =   255
            Left            =   120
            TabIndex        =   62
            Top             =   570
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   450
            Caption         =   "사용자 지정(&T):"
            Transparent     =   -1  'True
         End
         Begin prjDownloadBooster.OptionButtonW optSystemFore 
            Height          =   255
            Left            =   120
            TabIndex        =   63
            Top             =   240
            Width           =   1815
            _ExtentX        =   0
            _ExtentY        =   0
            Caption         =   "시스템 색상(&Y)"
            Transparent     =   -1  'True
         End
         Begin VB.Label lblSelectFore 
            BackStyle       =   0  '투명
            Height          =   255
            Left            =   1800
            TabIndex        =   11
            Top             =   240
            Width           =   1455
         End
         Begin VB.Shape pgFore 
            BackStyle       =   1  '투명하지 않음
            BorderColor     =   &H00404040&
            FillColor       =   &H00808080&
            Height          =   255
            Left            =   2160
            Shape           =   4  '둥근 사각형
            Top             =   585
            Width           =   495
         End
      End
      Begin VB.PictureBox pbOuterPreview 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  '없음
         Height          =   2175
         Left            =   120
         ScaleHeight     =   2175
         ScaleWidth      =   6015
         TabIndex        =   66
         Top             =   120
         Width           =   6015
         Begin VB.PictureBox pbBackground 
            AutoRedraw      =   -1  'True
            Enabled         =   0   'False
            Height          =   1380
            Left            =   600
            ScaleHeight     =   1320
            ScaleWidth      =   3855
            TabIndex        =   68
            Tag             =   "nobgdraw"
            Top             =   360
            Width           =   3915
            Begin prjDownloadBooster.CheckBoxW CheckBoxW1 
               Height          =   255
               Left            =   120
               TabIndex        =   69
               Top             =   990
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   450
               Caption         =   "완료 후 열기"
               Transparent     =   -1  'True
            End
            Begin prjDownloadBooster.TextBoxW TextBoxW1 
               Height          =   255
               Left            =   1080
               TabIndex        =   70
               Top             =   120
               Width           =   2415
               _ExtentX        =   4260
               _ExtentY        =   450
               Text            =   "frmOptions.frx":0ACD
            End
            Begin prjDownloadBooster.FrameW FrameW5 
               Height          =   555
               Left            =   120
               TabIndex        =   71
               Top             =   405
               Width           =   3615
               _ExtentX        =   6376
               _ExtentY        =   979
               Caption         =   " 다운로드 현황 "
               Transparent     =   -1  'True
               Begin prjDownloadBooster.ProgressBar pbSampleClassic 
                  Height          =   225
                  Left            =   120
                  Tag             =   "novisualstylechange"
                  Top             =   240
                  Visible         =   0   'False
                  Width           =   3375
                  _ExtentX        =   5953
                  _ExtentY        =   397
                  VisualStyles    =   0   'False
                  Enabled         =   0   'False
                  Value           =   24
                  Step            =   10
               End
               Begin prjDownloadBooster.ProgressBar pbSample 
                  Height          =   225
                  Left            =   120
                  Tag             =   "novisualstylechange"
                  Top             =   240
                  Width           =   3375
                  _ExtentX        =   5953
                  _ExtentY        =   397
                  Enabled         =   0   'False
                  Value           =   24
                  Step            =   10
                  State           =   3
               End
            End
            Begin prjDownloadBooster.CommandButtonW cmdSample 
               Height          =   285
               Left            =   2160
               TabIndex        =   72
               TabStop         =   0   'False
               Tag             =   "notygchange"
               Top             =   990
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   503
               Caption         =   "다운로드"
               Transparent     =   -1  'True
            End
            Begin VB.Label Label11 
               BackStyle       =   0  '투명
               Caption         =   "파일 주소:"
               Height          =   255
               Left            =   120
               TabIndex        =   73
               Top             =   150
               Width           =   975
            End
            Begin VB.Image imgPreview 
               Height          =   375
               Left            =   3120
               Stretch         =   -1  'True
               Top             =   0
               Visible         =   0   'False
               Width           =   855
            End
         End
         Begin VB.PictureBox pbPreview 
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000001&
            Enabled         =   0   'False
            Height          =   2175
            Left            =   0
            ScaleHeight     =   2115
            ScaleWidth      =   5955
            TabIndex        =   67
            Tag             =   "nobgdraw"
            Top             =   0
            Width           =   6015
            Begin VB.Image imgDesktop 
               Height          =   735
               Left            =   0
               Stretch         =   -1  'True
               Top             =   0
               Width           =   1575
            End
         End
      End
      Begin prjDownloadBooster.FrameW Frame1 
         Height          =   975
         Left            =   120
         TabIndex        =   6
         Top             =   2400
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   1720
         Caption         =   " 배경색 "
         Begin prjDownloadBooster.OptionButtonW optSystemColor 
            Height          =   255
            Left            =   120
            TabIndex        =   60
            Top             =   240
            Width           =   1815
            _ExtentX        =   0
            _ExtentY        =   0
            Caption         =   "시스템 색상(&S)"
            Transparent     =   -1  'True
         End
         Begin prjDownloadBooster.OptionButtonW optUserColor 
            Height          =   255
            Left            =   120
            TabIndex        =   61
            Top             =   570
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   450
            Caption         =   "사용자 지정(&U):"
            Transparent     =   -1  'True
         End
         Begin VB.Label lblSelectColor 
            BackStyle       =   0  '투명
            Height          =   255
            Left            =   1800
            TabIndex        =   9
            Top             =   240
            Width           =   1455
         End
         Begin VB.Shape pgColor 
            BackStyle       =   1  '투명하지 않음
            BorderColor     =   &H00404040&
            FillColor       =   &H00808080&
            Height          =   255
            Left            =   2160
            Shape           =   4  '둥근 사각형
            Top             =   585
            Width           =   495
         End
      End
      Begin prjDownloadBooster.FrameW FrameW1 
         Height          =   975
         Left            =   3240
         TabIndex        =   22
         Top             =   2400
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   1720
         Caption         =   " 배경 그림 "
         Transparent     =   -1  'True
         Begin prjDownloadBooster.ComboBoxW cbImagePosition 
            Height          =   300
            Left            =   960
            TabIndex        =   34
            Top             =   600
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   529
            Style           =   2
            Text            =   "ComboBoxW1"
         End
         Begin prjDownloadBooster.CommandButtonW cmdChooseBackground 
            Height          =   330
            Left            =   2160
            TabIndex        =   24
            Top             =   210
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   582
            ImageList       =   "imgBrowse"
            ImageListAlignment=   4
            Transparent     =   -1  'True
         End
         Begin prjDownloadBooster.CheckBoxW chkEnableBackgroundImage 
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   450
            Caption         =   "배경 그림 사용(&B)"
            Transparent     =   -1  'True
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '투명
            Caption         =   "위치(&P):"
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   645
            Width           =   840
         End
      End
   End
   Begin prjDownloadBooster.CommandButtonW cmdApply 
      Height          =   360
      Left            =   10920
      TabIndex        =   3
      Top             =   120
      Width           =   1320
      _ExtentX        =   0
      _ExtentY        =   0
      Enabled         =   0   'False
      Caption         =   "적용(&A)"
   End
   Begin prjDownloadBooster.TabStrip tsTabStrip 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   661
      MultiRow        =   0   'False
      TabFixedWidth   =   53
      TabScrollWheel  =   0   'False
      Transparent     =   -1  'True
      InitTabs        =   "frmOptions.frx":0B1D
   End
   Begin prjDownloadBooster.CommandButtonW CancelButton 
      Cancel          =   -1  'True
      Height          =   360
      Left            =   9480
      TabIndex        =   1
      Top             =   120
      Width           =   1320
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "취소"
   End
   Begin prjDownloadBooster.CommandButtonW OKButton 
      Default         =   -1  'True
      Height          =   360
      Left            =   8040
      TabIndex        =   0
      Top             =   120
      Width           =   1320
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "확인"
   End
   Begin prjDownloadBooster.ImageList imgFiles 
      Left            =   12240
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      ImageWidth      =   16
      ImageHeight     =   16
      ColorDepth      =   4
      MaskColor       =   16711935
      InitListImages  =   "frmOptions.frx":0CC1
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'참고 자료:
'- https://www.vbforums.com/showthread.php?284592-Listview-StartLabelEdit-second-column-*RESOLVED*

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"
Dim Loaded As Boolean
Dim ColorChanged As Boolean
Public ImageChanged As Boolean
Dim VisualStyleChanged As Boolean
Dim SkinChanged As Boolean
Dim MouseY As Integer, SelectedListItem As LvwListItem

Dim PrevhWnd As Long

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub cbFrameSkin_Click()
    If Loaded Then
        cmdApply.Enabled = -1
        SkinChanged = True
    End If
    
    If (cbFrameSkin.ListCount >= 3 And cbFrameSkin.ListIndex = 2) Or (cbFrameSkin.ListCount < 3 And cbFrameSkin.ListIndex = 1) Then
        SetWindowRgn PrevhWnd, CreateRectRgn(0, 0, Screen.Width / Screen.TwipsPerPixelX + 300, Screen.Height / Screen.TwipsPerPixelY + 300), True
    Else
        SetWindowRgn PrevhWnd, 0&, True
    End If
End Sub

Private Sub cbImagePosition_Click()
    If Loaded Then
        cmdApply.Enabled = -1
    End If
End Sub

Private Sub cbLanguage_Click()
    If Loaded Then
        'Alert t("언어를 변경하려면 프로그램을 재시작해야 합니다.", "To change the language you must restart the application."), App.Title, Me, 64
        cmdApply.Enabled = -1
    End If
End Sub

Private Sub cbSkin_Click()
    cmdSample.VisualStyles = (cbSkin.ListIndex <> 1)
    cmdSample.IsTygemButton = (cbSkin.ListIndex = 2)
    cmdSample.Refresh
    pbSampleClassic.Visible = Not cmdSample.VisualStyles
    Dim ctrl As Control
    On Error Resume Next
    For Each ctrl In Me.Controls
        If ctrl.Container Is pbBackground And ctrl.Name <> "cmdSample" And ctrl.Name <> "pbSample" And ctrl.Name <> "pbSampleClassic" Then
            ctrl.VisualStyles = cmdSample.VisualStyles
        End If
    Next ctrl
    If Loaded Then
        cmdApply.Enabled = -1
        SkinChanged = True
        VisualStyleChanged = True
    End If
End Sub

Private Sub cbWhenExist_Click()
    If Loaded Then
        cmdApply.Enabled = -1
    End If
End Sub

Private Sub chkAllowDuplicates_Click()
    If Loaded Then cmdApply.Enabled = -1
End Sub

Private Sub chkAlwaysOnTop_Click()
    If Loaded Then cmdApply.Enabled = -1
End Sub

Private Sub chkAlwaysResume_Click()
    If Loaded Then cmdApply.Enabled = -1
End Sub

Private Sub chkAsterisk_Click()
    If Loaded Then cmdApply.Enabled = -1
    EnableFrameControls fAsterisk, chkAsterisk, (chkAsterisk.Value = 1)
End Sub

Private Sub chkAutoRetry_Click()
    If Loaded Then cmdApply.Enabled = -1
End Sub

Private Sub chkBeepWhenComplete_Click()
    If Loaded Then cmdApply.Enabled = -1
    EnableFrameControls fCompleteSound, chkBeepWhenComplete, (chkBeepWhenComplete.Value = 1)
End Sub

Private Sub chkDontLoadIcons_Click()
    If Loaded Then
        cmdApply.Enabled = -1
        ImageChanged = True
    End If
End Sub

Private Sub chkEnableBackgroundImage_Click()
    If Loaded Then
        cmdApply.Enabled = -1
        ImageChanged = True
        RedrawPreview
    End If
    
    If chkEnableBackgroundImage.Value = 0 Then
        cmdChooseBackground.Enabled = 0
        imgPreview.Visible = 0
        cmdSample.Refresh
    Else
        cmdChooseBackground.Enabled = -1
        
        On Error Resume Next
        If LCase(Right$(GetSetting("DownloadBooster", "Options", "BackgroundImagePath", ""), 4)) = ".png" Then
            Set imgPreview.Picture = LoadPngIntoPictureWithAlpha(GetSetting("DownloadBooster", "Options", "BackgroundImagePath", ""))
        Else
            imgPreview.Picture = LoadPicture(GetSetting("DownloadBooster", "Options", "BackgroundImagePath", ""))
        End If
        
        imgPreview.Visible = -1
        cmdSample.Refresh
    End If
    RedrawPreview
End Sub

Private Sub chkError_Click()
    If Loaded Then cmdApply.Enabled = -1
    EnableFrameControls fError, chkError, (chkError.Value = 1)
End Sub

Private Sub chkExclamation_Click()
    If Loaded Then cmdApply.Enabled = -1
    EnableFrameControls fExclamation, chkExclamation, (chkExclamation.Value = 1)
End Sub

Private Sub chkExcludeMergeFromElapsed_Click()
    If Loaded Then cmdApply.Enabled = -1
End Sub

Private Sub chkForceGet_Click()
    If Loaded Then cmdApply.Enabled = -1
End Sub

Private Sub chkForceOldDialog_Click()
    If Loaded Then cmdApply.Enabled = -1
End Sub

Private Sub chkIgnore300_Click()
    If Loaded Then cmdApply.Enabled = -1
End Sub

Private Sub chkLazyElapsed_Click()
    If Loaded Then cmdApply.Enabled = -1
End Sub

Private Sub chkNoCleanup_Click()
    If Loaded Then cmdApply.Enabled = -1
End Sub

Private Sub chkNoRedirectCheck_Click()
    If Loaded Then cmdApply.Enabled = -1
End Sub

Private Sub chkOpenDirWhenComplete_Click()
    If Loaded Then cmdApply.Enabled = -1
End Sub

Private Sub chkOpenWhenComplete_Click()
    If Loaded Then cmdApply.Enabled = -1
End Sub

Private Sub chkQuestion_Click()
    If Loaded Then cmdApply.Enabled = -1
    EnableFrameControls fQuestion, chkQuestion, (chkQuestion.Value = 1)
End Sub

Private Sub chkRememberURL_Click()
    If Loaded Then cmdApply.Enabled = -1
End Sub

Private Sub cmdAddHeader_Click()
    If Loaded Then cmdApply.Enabled = -1
    lvHeaders.SetFocus
    Set lvHeaders.SelectedItem = lvHeaders.ListItems.Add(, , "", , 1)
    lvHeaders.SelectedItem.ListSubItems.Add , , ""
    lvHeaders.StartLabelEdit
End Sub

Private Sub cmdApply_Click()
    If WinVer >= 6# And cbFrameSkin.ListCount >= 3 Then
        SaveSetting "DownloadBooster", "Options", "DisableDWMWindow", Abs(CInt(cbFrameSkin.ListIndex = 1))
    End If
    If (cbFrameSkin.ListCount >= 3 And cbFrameSkin.ListIndex = 2) Or (cbFrameSkin.ListCount < 3 And cbFrameSkin.ListIndex = 1) Then
        SaveSetting "DownloadBooster", "Options", "UseClassicThemeFrame", 1
    Else
        SaveSetting "DownloadBooster", "Options", "UseClassicThemeFrame", 0
    End If
    
    Dim i%
    
    SaveSetting "DownloadBooster", "Options", "NoCleanup", chkNoCleanup.Value
    SaveSetting "DownloadBooster", "Options", "RememberURL", chkRememberURL.Value
    SaveSetting "DownloadBooster", "Options", "NoRedirectCheck", chkNoRedirectCheck.Value
    SaveSetting "DownloadBooster", "Options", "ForceGet", chkForceGet.Value
    SaveSetting "DownloadBooster", "Options", "Ignore300", chkIgnore300.Value
    SaveSetting "DownloadBooster", "Options", "LazyElapsed", chkLazyElapsed.Value
    SaveSetting "DownloadBooster", "Options", "ExcludeMergeFromElapsed", chkExcludeMergeFromElapsed.Value
    SaveSetting "DownloadBooster", "Options", "ForceWin31Dialog", chkForceOldDialog.Value
    SaveSetting "DownloadBooster", "Options", "DontLoadIcons", chkDontLoadIcons.Value
    SaveSetting "DownloadBooster", "Options", "AutoDetectYtdlURL", chkAutoYtdl.Value
    SaveSetting "DownloadBooster", "Options", "CompleteSoundPath", Trim$(txtCompleteSoundPath.Text)
    SaveSetting "DownloadBooster", "Options", "AllowDuplicatesInQueue", chkAllowDuplicates.Value
    
    SaveSetting "DownloadBooster", "Options", "EnableAsteriskSound", chkAsterisk.Value
    SaveSetting "DownloadBooster", "Options", "EnableExclamationSound", chkExclamation.Value
    SaveSetting "DownloadBooster", "Options", "EnableErrorSound", chkError.Value
    SaveSetting "DownloadBooster", "Options", "EnableQuestionSound", chkQuestion.Value
    SaveSetting "DownloadBooster", "Options", "AsteriskSound", txtAsterisk.Text
    SaveSetting "DownloadBooster", "Options", "ExclamationSound", txtExclamation.Text
    SaveSetting "DownloadBooster", "Options", "ErrorSound", txtError.Text
    SaveSetting "DownloadBooster", "Options", "QuestionSound", txtQuestion.Text
    
    SaveSetting "DownloadBooster", "Options", "OpenWhenComplete", chkOpenWhenComplete.Value
    SaveSetting "DownloadBooster", "Options", "OpenFolderWhenComplete", chkOpenDirWhenComplete.Value
    SaveSetting "DownloadBooster", "Options", "PlaySound", chkBeepWhenComplete.Value
    SaveSetting "DownloadBooster", "Options", "ContinueDownload", chkAlwaysResume.Value
    SaveSetting "DownloadBooster", "Options", "AutoRetry", chkAutoRetry.Value
    SaveSetting "DownloadBooster", "Options", "WhenFileExists", cbWhenExist.ListIndex
    frmMain.cbWhenExist.ListIndex = cbWhenExist.ListIndex
    
    frmMain.chkOpenAfterComplete.Value = chkOpenWhenComplete.Value
    frmMain.chkOpenFolder.Value = chkOpenDirWhenComplete.Value
    frmMain.chkContinueDownload.Value = chkAlwaysResume.Value
    frmMain.chkAutoRetry.Value = chkAutoRetry.Value
    
    If cbFrameSkin.ListCount >= 3 Then
        If cbFrameSkin.ListIndex = 1 Then
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
        frmMain.pgSettingsBackground.Visible = 0
        frmMain.chkOpenAfterComplete.Tag = ""
        frmMain.chkOpenFolder.Tag = ""
        frmMain.chkContinueDownload.Tag = ""
        frmMain.chkAutoRetry.Tag = ""
        frmMain.chkOpenAfterComplete.Transparent = -1
        frmMain.chkOpenFolder.Transparent = -1
        frmMain.chkContinueDownload.Transparent = -1
        frmMain.chkAutoRetry.Transparent = -1
    ElseIf optUserFore.Value Then
        SaveSetting "DownloadBooster", "Options", "ForeColor", CLng(pgFore.BackColor)
        frmMain.pgSettingsBackground.Visible = -1
        frmMain.chkOpenAfterComplete.Tag = "nobackcolorchange"
        frmMain.chkOpenFolder.Tag = "nobackcolorchange"
        frmMain.chkContinueDownload.Tag = "nobackcolorchange"
        frmMain.chkAutoRetry.Tag = "nobackcolorchange"
        frmMain.chkOpenAfterComplete.Transparent = 0
        frmMain.chkOpenFolder.Transparent = 0
        frmMain.chkContinueDownload.Transparent = 0
        frmMain.chkAutoRetry.Transparent = 0
    End If
    SaveSetting "DownloadBooster", "Options", "DisableVisualStyle", CBool(cbSkin.ListIndex = 1) * (-1)
    SaveSetting "DownloadBooster", "Options", "EnableLiveBadukMemoSkin", CBool(cbSkin.ListIndex = 2) * (-1)
    If ColorChanged Or VisualStyleChanged Or SkinChanged Then
        SetFormBackgroundColor Me, True
        SetFormBackgroundColor frmMain, True
        frmMain.LoadLiveBadukSkin
        RedrawPreview
        cmdChooseBackground.Refresh
        frmMain.pbProgressContainer.Refresh
        frmMain.SetupSplitButtons
    End If
    If VisualStyleChanged Then
        On Error Resume Next
        DrawTabBackground
        cmdChooseBackground.Refresh
        cmdSample.Refresh
        On Error GoTo 0
        frmMain.SetTextColors
    End If
    If cbLanguage.ListIndex = 0 Then
        SaveSetting "DownloadBooster", "Options", "Language", "0"
    ElseIf cbLanguage.ListIndex = 1 Then
        SaveSetting "DownloadBooster", "Options", "Language", 1042
    Else
        SaveSetting "DownloadBooster", "Options", "Language", 1033
    End If
    SaveSetting "DownloadBooster", "Options", "ImagePosition", cbImagePosition.ListIndex
    frmMain.ImagePosition = cbImagePosition.ListIndex
    frmMain.SetBackgroundPosition True
    If ImageChanged Then
        SaveSetting "DownloadBooster", "Options", "UseBackgroundImage", chkEnableBackgroundImage.Value
        frmMain.SetBackgroundImage
    End If
    
    On Error Resume Next
    If GetSetting("DownloadBooster", "Options", "ForeColor", -1) <> -1 Or GetSetting("DownloadBooster", "Options", "UseBackgroundImage", 0) = 1 Then
        For i = frmMain.pgOverlay.LBound To frmMain.pgOverlay.UBound
            frmMain.pgOverlay(i).Visible = -1
            frmMain.lblOverlay(i).Visible = -1
        Next i
        frmMain.optTabDownload2.Transparent = 0
        frmMain.optTabDownload2.BackColor = frmMain.pgOverlay(0).BackColor
        frmMain.optTabDownload2.Refresh
        frmMain.optTabThreads2.Transparent = 0
        frmMain.optTabThreads2.BackColor = frmMain.pgOverlay(0).BackColor
        frmMain.optTabThreads2.Refresh
        frmMain.fTabDownload.Transparent = 0
        frmMain.fTabDownload.BackColor = frmMain.pgOverlay(0).BackColor
        frmMain.fTabDownload.Refresh
        frmMain.fTabThreads.Transparent = 0
        frmMain.fTabThreads.BackColor = frmMain.pgOverlay(0).BackColor
        frmMain.fTabThreads.Refresh
    Else
        For i = frmMain.pgOverlay.LBound To frmMain.pgOverlay.UBound
            frmMain.pgOverlay(i).Visible = 0
            frmMain.lblOverlay(i).Visible = 0
        Next i
        frmMain.optTabDownload2.Transparent = -1
        frmMain.optTabDownload2.Refresh
        frmMain.optTabThreads2.Transparent = -1
        frmMain.optTabThreads2.Refresh
        frmMain.fTabDownload.Transparent = -1
        frmMain.fTabDownload.Refresh
        frmMain.fTabThreads.Transparent = -1
        frmMain.fTabThreads.Refresh
    End If
    On Error GoTo 0
    Dim NoDisable As Boolean
    NoDisable = False
    If Trim$(txtNodePath.Text) <> "" Then
        If FileExists(Trim$(txtNodePath.Text)) Then
            SaveSetting "DownloadBooster", "Options", "NodePath", Trim$(txtNodePath.Text)
        Else
            Alert t("Node.js 경로가 존재하지 않습니다.", "Node.js path does not exist."), App.Title, Me, 16
            NoDisable = True
        End If
    Else
        SaveSetting "DownloadBooster", "Options", "NodePath", ""
    End If
    If Trim$(txtScriptPath.Text) <> "" Then
        If FileExists(Trim$(txtScriptPath.Text)) Then
            SaveSetting "DownloadBooster", "Options", "ScriptPath", Trim$(txtScriptPath.Text)
        Else
            Alert t("다운로드 스크립트 경로가 존재하지 않습니다.", "Download script path does not exist."), App.Title, Me, 16
            NoDisable = True
        End If
    Else
        SaveSetting "DownloadBooster", "Options", "ScriptPath", ""
    End If
    If Trim$(txtYtdlPath.Text) <> "" Then
        If FileExists(Trim$(txtYtdlPath.Text)) Then
            SaveSetting "DownloadBooster", "Options", "YtdlPath", Trim$(txtYtdlPath.Text)
        Else
            Alert t("Youtube-dl 경로가 존재하지 않습니다.", "Youtube-dl path does not exist."), App.Title, Me, 16
            NoDisable = True
        End If
    Else
        SaveSetting "DownloadBooster", "Options", "YtdlPath", ""
    End If
    
    Dim hSysMenu As Long
    Dim MII As MENUITEMINFO
    hSysMenu = GetSystemMenu(frmMain.hWnd, 0)
    MainFormOnTop = (chkAlwaysOnTop.Value = 1)
    SetWindowPos frmMain.hWnd, IIf(MainFormOnTop, HWND_TOPMOST, HWND_NOTOPMOST), 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    SetWindowPos Me.hWnd, IIf(MainFormOnTop, HWND_TOPMOST, HWND_NOTOPMOST), 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    With MII
        .cbSize = Len(MII)
        .fMask = MIIM_STATE
        .fState = MFS_ENABLED Or IIf(MainFormOnTop, MFS_CHECKED, 0)
    End With
    SetMenuItemInfo hSysMenu, 1000, 0, MII
    SaveSetting "DownloadBooster", "Options", "AlwaysOnTop", Abs(CInt(MainFormOnTop))
    
    On Error Resume Next
    SaveSetting "DownloadBooster", "Options\Headers", "_Dummy_", "x" '오류 방지
    DeleteSetting "DownloadBooster", "Options\Headers"
    If lvHeaders.ListItems.Count > 0 Then
        For i = 1 To lvHeaders.ListItems.Count
            If Trim$(lvHeaders.ListItems(i).Text) <> "" Then _
                SaveSetting "DownloadBooster", "Options\Headers", Trim$(lvHeaders.ListItems(i).Text), lvHeaders.ListItems(i).ListSubItems(1).Text
        Next i
    End If
    BuildHeaderCache
    
    RedrawPreview
    ColorChanged = False
    ImageChanged = False
    VisualStyleChanged = False
    SkinChanged = False
    If Not NoDisable Then
        cmdApply.Enabled = 0
    End If
End Sub

Private Sub cmdBrowseAsterisk_Click()
    Tags.BrowseTargetForm = 4
    Tags.BrowsePresetPath = txtAsterisk.Text
    Set Tags.BrowseTargetTextbox = txtAsterisk
    frmExplorer.Show vbModal, Me
End Sub

Private Sub cmdBrowseCompleteSound_Click()
    Tags.BrowseTargetForm = 4
    Tags.BrowsePresetPath = txtCompleteSoundPath.Text
    Set Tags.BrowseTargetTextbox = txtCompleteSoundPath
    frmExplorer.Show vbModal, Me
End Sub

Private Sub cmdBrowseError_Click()
    Tags.BrowseTargetForm = 4
    Tags.BrowsePresetPath = txtError.Text
    Set Tags.BrowseTargetTextbox = txtError
    frmExplorer.Show vbModal, Me
End Sub

Private Sub cmdBrowseExclamation_Click()
    Tags.BrowseTargetForm = 4
    Tags.BrowsePresetPath = txtExclamation.Text
    Set Tags.BrowseTargetTextbox = txtExclamation
    frmExplorer.Show vbModal, Me
End Sub

Private Sub cmdBrowseQuestion_Click()
    Tags.BrowseTargetForm = 4
    Tags.BrowsePresetPath = txtQuestion.Text
    Set Tags.BrowseTargetTextbox = txtQuestion
    frmExplorer.Show vbModal, Me
End Sub

Private Sub cmdChooseBackground_Click()
    If GetSetting("DownloadBooster", "Options", "ForceWin31Dialog", "0") = "1" Then
        frmCustomBackground.Show vbModal, Me
        Exit Sub
    End If
    Tags.BrowseTargetForm = 3
    frmExplorer.Show vbModal, Me
End Sub

Private Sub cmdDeleteHeader_Click()
    If Not lvHeaders.SelectedItem Is Nothing Then
        If lvHeaders.SelectedItem.Selected Then
            lvHeaders.ListItems.Remove lvHeaders.SelectedItem.Index
            If Loaded Then
                cmdApply.Enabled = -1
            End If
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

Private Sub cmdTestAsterisk_Click()
    txtAsterisk.Text = Trim$(txtAsterisk.Text)
    If txtAsterisk.Text = "" Then
        MessageBeep 64
    Else
        PlayWave txtAsterisk.Text
    End If
End Sub

Private Sub cmdTestCompleteSound_Click()
    txtCompleteSoundPath.Text = Trim$(txtCompleteSoundPath.Text)
    If txtCompleteSoundPath.Text = "" Then
        MessageBeep 64
    Else
        PlayWave txtCompleteSoundPath.Text
    End If
End Sub

Private Sub cmdTestError_Click()
    txtError.Text = Trim$(txtError.Text)
    If txtError.Text = "" Then
        MessageBeep 16
    Else
        PlayWave txtError.Text
    End If
End Sub

Private Sub cmdTestExclamation_Click()
    txtExclamation.Text = Trim$(txtExclamation.Text)
    If txtExclamation.Text = "" Then
        MessageBeep 48
    Else
        PlayWave txtExclamation.Text
    End If
End Sub

Private Sub cmdTestQuestion_Click()
    txtQuestion.Text = Trim$(txtQuestion.Text)
    If txtQuestion.Text = "" Then
        MessageBeep 32
    Else
        PlayWave txtQuestion.Text
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DestroyWindow PrevhWnd
    Unhook_Options Me.hWnd
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
    
    If Loaded Then
        cmdApply.Enabled = -1
    End If
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

Private Sub txtAsterisk_Change()
    If Loaded Then cmdApply.Enabled = -1
End Sub

Private Sub txtCompleteSoundPath_Change()
    If Loaded Then cmdApply.Enabled = -1
End Sub

Private Sub txtEdit_LostFocus()
    On Error Resume Next
    SelectedListItem.ListSubItems(1).Text = txtEdit.Text
    txtEdit.Visible = False
    Set SelectedListItem = Nothing
    OKButton.Enabled = -1
    If Loaded Then
        cmdApply.Enabled = -1
    End If
End Sub
 
Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Or KeyAscii = 10 Then
        SelectedListItem.ListSubItems(1).Text = txtEdit.Text
        txtEdit.Visible = False
        Set SelectedListItem = Nothing
        OKButton.Enabled = -1
        If Loaded Then
            cmdApply.Enabled = -1
        End If
        lvHeaders.SetFocus
    End If
End Sub

Private Sub Form_Load()
    VisualStyleChanged = False
    Loaded = False
    ImageChanged = False
    ColorChanged = False
    SkinChanged = False
    If GetSetting("DownloadBooster", "Options", "DisableDWMWindow", DefaultDisableDWMWindow) = 1 Then DisableDWMWindow Me.hWnd
    SetFormBackgroundColor Me
    SetFont Me
    SetWindowPos Me.hWnd, IIf(MainFormOnTop, HWND_TOPMOST, HWND_NOTOPMOST), 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    
    On Error Resume Next
    Me.Icon = frmMain.imgWrench.ListImages(1).Picture
    On Error GoTo 0
    
    lvHeaders.SmallIcons = imgFiles
    
    lblSelectColor.Top = pgColor.Top
    lblSelectColor.Left = pgColor.Left
    lblSelectColor.Width = pgColor.Width
    lblSelectColor.Height = pgColor.Height
    
    lblSelectFore.Top = pgFore.Top
    lblSelectFore.Left = pgFore.Left
    lblSelectFore.Width = pgFore.Width
    lblSelectFore.Height = pgFore.Height
    
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
    cmdApply.Top = tsTabStrip.Top + tsTabStrip.Height + 60
    CancelButton.Top = cmdApply.Top
    OKButton.Top = cmdApply.Top
    cmdApply.Left = tsTabStrip.Left + tsTabStrip.Width - cmdApply.Width
    CancelButton.Left = cmdApply.Left - 120 - CancelButton.Width
    OKButton.Left = CancelButton.Left - 120 - OKButton.Width
    Me.Height = cmdApply.Top + cmdApply.Height + 540
    Me.Width = tsTabStrip.Width + 240 + 60
    pbPanel(1).Visible = -1
    pbPanel(1).Enabled = -1
    
    chkNoCleanup.Value = GetSetting("DownloadBooster", "Options", "NoCleanup", 0)
    chkNoRedirectCheck.Value = GetSetting("DownloadBooster", "Options", "NoRedirectCheck", 0)
    chkForceGet.Value = GetSetting("DownloadBooster", "Options", "ForceGet", 1)
    chkIgnore300.Value = GetSetting("DownloadBooster", "Options", "Ignore300", 0)
    chkAlwaysOnTop.Value = Abs(CInt(MainFormOnTop))
    chkLazyElapsed.Value = GetSetting("DownloadBooster", "Options", "LazyElapsed", 0)
    chkExcludeMergeFromElapsed.Value = GetSetting("DownloadBooster", "Options", "ExcludeMergeFromElapsed", 0)
    chkForceOldDialog.Value = GetSetting("DownloadBooster", "Options", "ForceWin31Dialog", 0)
    chkDontLoadIcons.Value = GetSetting("DownloadBooster", "Options", "DontLoadIcons", 0)
    chkRememberURL.Value = GetSetting("DownloadBooster", "Options", "RememberURL", 1)
    chkAutoYtdl.Value = GetSetting("DownloadBooster", "Options", "AutoDetectYtdlURL", 1)
    txtCompleteSoundPath.Text = Trim$(GetSetting("DownloadBooster", "Options", "CompleteSoundPath", ""))
    chkAllowDuplicates.Value = GetSetting("DownloadBooster", "Options", "AllowDuplicatesInQueue", 0)
    
    pbBackground.BorderStyle = 0
    SetPreviewPosition
    
    imgPreview.Top = 0
    imgPreview.Left = 0
    
    If WinVer < 6.2 And IsDWMEnabled() Then
        cbFrameSkin.AddItem "Windows Aero"
    Else
        cbFrameSkin.AddItem t("시스템 스타일", "System style")
    End If
    If IsDWMEnabled() Then
        If WinVer < 6.2 Or LCase(GetFilename(GetKeyValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\ThemeManager", "DllName", "%SystemRoot%\resources\Themes\Aero\aero.msstyles"))) = "aero.msstyles" Then
            cbFrameSkin.AddItem "Windows " & IIf(WinVer < 6.1, "Vista", "7") & " " & t("베이직", "Basic")
        Else
            cbFrameSkin.AddItem t("시스템", "System") & " (" & t("DWM 없음", "No DWM") & ")"
        End If
    End If
    cbFrameSkin.AddItem t("고전 스타일", "Classic style")
    If GetSetting("DownloadBooster", "Options", "UseClassicThemeFrame", 0) <> 0 Then
        If cbFrameSkin.ListCount >= 3 Then
            cbFrameSkin.ListIndex = 2
        Else
            cbFrameSkin.ListIndex = 1
        End If
    ElseIf GetSetting("DownloadBooster", "Options", "DisableDWMWindow", DefaultDisableDWMWindow) <> 0 And cbFrameSkin.ListCount >= 3 Then
        cbFrameSkin.ListIndex = 1
    Else
        cbFrameSkin.ListIndex = 0
    End If
    
    chkEnableBackgroundImage.Value = GetSetting("DownloadBooster", "Options", "UseBackgroundImage", 0)
    If chkEnableBackgroundImage.Value = 0 Then
        cmdChooseBackground.Enabled = 0
    End If
    
    Dim clrBackColor As Long
    clrBackColor = GetSetting("DownloadBooster", "Options", "BackColor", DefaultBackColor)
    If clrBackColor < 0 Or clrBackColor > 16777215 Then
        optSystemColor.Value = True
        pgColor.BackColor = &H8000000F
    Else
        optUserColor.Value = True
        pgColor.BackColor = clrBackColor
    End If
    pbBackground.BackColor = pgColor.BackColor
    
    Dim clrForeColor As Long
    clrForeColor = GetSetting("DownloadBooster", "Options", "ForeColor", -1)
    If clrForeColor < 0 Or clrForeColor > 16777215 Then
        optSystemFore.Value = True
        pgFore.BackColor = &H80000012
    Else
        optUserFore.Value = True
        pgFore.BackColor = clrForeColor
    End If
    Label11.ForeColor = pgFore.BackColor
    
    cmdApply.Enabled = 0
    
    DrawTabBackground
    
    chkOpenWhenComplete.Value = frmMain.chkOpenAfterComplete.Value
    chkOpenDirWhenComplete.Value = frmMain.chkOpenFolder.Value
    chkBeepWhenComplete.Value = GetSetting("DownloadBooster", "Options", "PlaySound", 1)
    chkAlwaysResume.Value = frmMain.chkContinueDownload.Value
    chkAutoRetry.Value = frmMain.chkAutoRetry.Value
    
    cbSkin.Clear
    cbSkin.AddItem t("시스템 스타일", "System style")
    cbSkin.AddItem t("고전 스타일", "Classic style")
    cbSkin.AddItem t("타이젬바둑 쪽지", "LiveBaduk memo style")
    If CInt(GetSetting("DownloadBooster", "Options", "EnableLiveBadukMemoSkin", 0)) Then
        cbSkin.ListIndex = 2
    ElseIf Abs(CInt(GetSetting("DownloadBooster", "Options", "DisableVisualStyle", 0))) Then
        cbSkin.ListIndex = 1
    Else
        cbSkin.ListIndex = 0
    End If
    
    'chkNoTheming.Value = Abs(CInt(GetSetting("DownloadBooster", "Options", "DisableVisualStyle", 0)))
    cmdSample.VisualStyles = (Not CBool(CInt(GetSetting("DownloadBooster", "Options", "DisableVisualStyle", 0))))
    cmdSample.IsTygemButton = Abs(CInt(GetSetting("DownloadBooster", "Options", "EnableLiveBadukMemoSkin", 0))) * (-1)
    
    cbLanguage.Clear
    cbLanguage.AddItem t("자동", "Auto")
    cbLanguage.AddItem "한국어"
    cbLanguage.AddItem "English"
    Dim LangSet As String
    LangSet = GetSetting("DownloadBooster", "Options", "Language", "0")
    If LangSet = "0" Then
        cbLanguage.ListIndex = 0
    ElseIf LangSet = "1042" Then
        cbLanguage.ListIndex = 1
    Else
        cbLanguage.ListIndex = 2
    End If
    
    cbWhenExist.Clear
    cbWhenExist.AddItem t("다운로드 안 하기", "Don't download")
    cbWhenExist.AddItem t("덮어쓰기", "Overwrite")
    cbWhenExist.AddItem t("이름에 번호 붙이기", "Auto Rename")
    cbWhenExist.ListIndex = GetSetting("DownloadBooster", "Options", "WhenFileExists", 0)
    
    cbImagePosition.Clear
    cbImagePosition.AddItem t("늘이기", "Stretch")
    cbImagePosition.AddItem t("높이에 맞추기", "Fit to height")
    cbImagePosition.AddItem t("너비에 맞추기", "Fit to width")
    cbImagePosition.AddItem t("원본 크기 유지", "True size")
    'cbImagePosition.AddItem t("가운데", "Centered")
    cbImagePosition.ListIndex = GetSetting("DownloadBooster", "Options", "ImagePosition", 1)
    
    txtNodePath.Text = GetSetting("DownloadBooster", "Options", "NodePath", "")
    txtScriptPath.Text = GetSetting("DownloadBooster", "Options", "ScriptPath", "")
    txtYtdlPath.Text = GetSetting("DownloadBooster", "Options", "YtdlPath", "")
    
    On Error Resume Next
    If chkEnableBackgroundImage.Value And GetSetting("DownloadBooster", "Options", "BackgroundImagePath", "") <> "" Then
        If LCase(Right$(GetSetting("DownloadBooster", "Options", "BackgroundImagePath", ""), 4)) = ".png" Then
            Set imgPreview.Picture = LoadPngIntoPictureWithAlpha(GetSetting("DownloadBooster", "Options", "BackgroundImagePath", ""))
        Else
            imgPreview.Picture = LoadPicture(GetSetting("DownloadBooster", "Options", "BackgroundImagePath", ""))
        End If
        
        imgPreview.Visible = -1
    End If
    
    tsTabStrip.Tabs(1).Caption = t(tsTabStrip.Tabs(1).Caption, " General ")
    tsTabStrip.Tabs(2).Caption = t(tsTabStrip.Tabs(2).Caption, " Network ")
    tsTabStrip.Tabs(3).Caption = t(tsTabStrip.Tabs(3).Caption, " Appearance ")
    tsTabStrip.Tabs(4).Caption = t(tsTabStrip.Tabs(4).Caption, " Sound ")
    tsTabStrip.Tabs(5).Caption = t(tsTabStrip.Tabs(5).Caption, " Advanced ")
    Frame1.Caption = t(Frame1.Caption, " Background color ")
    Frame4.Caption = t(Frame4.Caption, " Text color ")
    Label10.Caption = t(Label10.Caption, "&Window:")
    Frame2.Caption = t(Frame2.Caption, " Download options ")
    Frame5.Caption = t(Frame5.Caption, " Interface ")
    chkNoCleanup.Caption = t(chkNoCleanup.Caption, "Preserve segme&nts")
    chkRememberURL.Caption = t(chkRememberURL.Caption, "Re&member URL")
    optSystemColor.Caption = t(optSystemColor.Caption, "&System color")
    optUserColor.Caption = t(optUserColor.Caption, "C&ustom color:")
    optSystemFore.Caption = t(optSystemFore.Caption, "S&ystem color")
    optUserFore.Caption = t(optUserFore.Caption, "Cus&tom color:")
    Label1.Caption = t(Label1.Caption, "&Language:")
    OKButton.Caption = t(OKButton.Caption, "OK")
    CancelButton.Caption = t(CancelButton.Caption, "Cancel")
    cmdApply.Caption = t(cmdApply.Caption, "&Apply")
    Me.Caption = t(Me.Caption, "Options")
    Frame6.Caption = t(Frame6.Caption, " Skin ")
    chkOpenWhenComplete.Caption = t(chkOpenWhenComplete.Caption, "&Open file when complete")
    chkOpenDirWhenComplete.Caption = t(chkOpenDirWhenComplete.Caption, "O&pen folder when complete")
    chkBeepWhenComplete.Caption = t(chkBeepWhenComplete.Caption, "Download &complete")
    chkAlwaysResume.Caption = t(chkAlwaysResume.Caption, "&Always resume")
    chkAutoRetry.Caption = t(chkAutoRetry.Caption, "A&uto retry on network error")
    Label3.Caption = t(Label3.Caption, "If filename alrea&dy exists:")
    FrameW1.Caption = t(FrameW1.Caption, " Background image ")
    chkEnableBackgroundImage.Caption = t(chkEnableBackgroundImage.Caption, "Use &background image")
    'cmdChooseBackground.Caption = t(cmdChooseBackground.Caption, "C&hoose image...")
    'chkNoTheming.Caption = t(chkNoTheming.Caption, "&Use classic theme")
    Label6.Caption = t(Label6.Caption, "Leave the field blank to use defaults. This option is for advanced users and there is no need to change for normal use.")
    FrameW2.Caption = t(FrameW2.Caption, " Directory settings ")
    Label5.Caption = t(Label5.Caption, "&Download script:")
    cmdSample.Caption = t(cmdSample.Caption, "Download")
    Label2.Caption = t(Label2.Caption, "&Position:")
    Label8.Caption = t(Label8.Caption, "Butt&on:")
    fHeaders.Caption = t(fHeaders.Caption, " Header settings ")
    chkNoRedirectCheck.Caption = t(chkNoRedirectCheck.Caption, "Don't check fo&r redirects")
    chkForceGet.Caption = t(chkForceGet.Caption, "Force GET re&quest on file check")
    chkIgnore300.Caption = t(chkIgnore300.Caption, "&Ignore 3XX reponse code")
    Label9.Caption = t("언어를 변경하려면 프로그램을 재시작해야 합니다.", "To change the language you must restart the application.")
    chkAlwaysOnTop.Caption = t(chkAlwaysOnTop.Caption, "Al&ways on top")
    chkAeroWindow.Caption = t(chkAeroWindow.Caption, "Use Ae&ro glass window")
    cmdAddHeader.Caption = t(cmdAddHeader.Caption, "&Add")
    cmdDeleteHeader.Caption = t(cmdDeleteHeader.Caption, "&Delete")
    cmdEditHeaderName.Caption = t(cmdEditHeaderName.Caption, "&Rename")
    cmdEditHeaderValue.Caption = t(cmdEditHeaderValue.Caption, "&Edit")
    chkLazyElapsed.Caption = t(chkLazyElapsed.Caption, "Elapsed time sin&ce first data receive")
    chkExcludeMergeFromElapsed.Caption = t(chkExcludeMergeFromElapsed.Caption, "Exclude merging time from elapsed time")
    FrameW3.Caption = t(FrameW3.Caption, " Network settings ")
    FrameW4.Caption = t(FrameW4.Caption, " Advanced download options ")
    chkForceOldDialog.Caption = t(chkForceOldDialog.Caption, "U&se Windows 3.1 dialogs")
    chkDontLoadIcons.Caption = t(chkDontLoadIcons.Caption, "Use sa&me icons for all files in the open dialog")
    chkAutoYtdl.Caption = t(chkAutoYtdl.Caption, "Automatically use &youtube-dl for supported links")
    Label11.Caption = t(Label11.Caption, "File URL:")
    FrameW5.Caption = t(FrameW5.Caption, " Download status ")
    CheckBoxW1.Caption = t(CheckBoxW1.Caption, "Open when done")
    tr chkAsterisk, "&Asterisk"
    tr chkExclamation, "&Exclamation"
    tr chkError, "E&rror"
    tr chkQuestion, "&Question"
    tr Label12, "Leave the fields blank to use the default sound."
    chkAsterisk.Value = GetSetting("DownloadBooster", "Options", "EnableAsteriskSound", 1)
    chkExclamation.Value = GetSetting("DownloadBooster", "Options", "EnableExclamationSound", 1)
    chkError.Value = GetSetting("DownloadBooster", "Options", "EnableErrorSound", 1)
    chkQuestion.Value = GetSetting("DownloadBooster", "Options", "EnableQuestionSound", 1)
    txtAsterisk.Text = GetSetting("DownloadBooster", "Options", "AsteriskSound", "")
    txtExclamation.Text = GetSetting("DownloadBooster", "Options", "ExclamationSound", "")
    txtError.Text = GetSetting("DownloadBooster", "Options", "ErrorSound", "")
    txtQuestion.Text = GetSetting("DownloadBooster", "Options", "QuestionSound", "")
    tr chkAllowDuplicates, "Allow dupl&icates in queue"
    
    lvHeaders.ColumnHeaders.Add , , t("이름", "Name"), 2055
    lvHeaders.ColumnHeaders.Add , , t("값", "Value"), 2775
    If GetSetting("DownloadBooster", "UserData", "HeaderSettingsInitialized", "0") = "0" Then
        SaveSetting "DownloadBooster", "UserData", "HeaderSettingsInitialized", 1
        SaveSetting "DownloadBooster", "Options\Headers", "User-Agent", "Mozilla/5.0 (Windows NT 5.1; rv:102.0) Gecko/20100101 Firefox/102.0 PaleMoon/33.2"
    End If
    
    Dim Headers() As String
    Headers = GetAllSettings("DownloadBooster", "Options\Headers")
    For i = LBound(Headers) To UBound(Headers)
        lvHeaders.ListItems.Add(, , Headers(i, 0), , 1).ListSubItems.Add , , Headers(i, 1)
    Next i
    
    Hook_Options Me.hWnd
    
    imgDesktop.Width = pbPreview.Width
    imgDesktop.Height = pbPreview.Height
    imgDesktop.Top = 0
    imgDesktop.Left = 0
    
    Dim WallpaperPath$, ActiveDesktopWallpaperPath$
    WallpaperPath = GetKeyValue(HKEY_CURRENT_USER, "Control Panel\Desktop", "Wallpaper")
    If WinVer < 6# Then
        ActiveDesktopWallpaperPath = GetKeyValue(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Desktop\General", "Wallpaper", WallpaperPath)
    Else
        ActiveDesktopWallpaperPath = WallpaperPath
    End If
    
    If Left$(WallpaperPath, 1) = """" And Right$(WallpaperPath, 1) = """" Then WallpaperPath = Mid$(WallpaperPath, 2, Len(WallpaperPath) - 2)
    If Left$(ActiveDesktopWallpaperPath, 1) = """" And Right$(ActiveDesktopWallpaperPath, 1) = """" Then ActiveDesktopWallpaperPath = Mid$(ActiveDesktopWallpaperPath, 2, Len(ActiveDesktopWallpaperPath) - 2)
    
    On Error GoTo activefail
    If Right$(LCase(ActiveDesktopWallpaperPath), 4) = ".png" Then
        Set imgDesktop.Picture = LoadPngIntoPictureWithAlpha(ActiveDesktopWallpaperPath)
    Else
        imgDesktop.Picture = LoadPicture(ActiveDesktopWallpaperPath)
    End If
    GoTo nextcode
    
activefail:
    On Error GoTo nextcode
    If Right$(LCase(WallpaperPath), 4) = ".png" Then
        Set imgDesktop.Picture = LoadPngIntoPictureWithAlpha(WallpaperPath)
    Else
        imgDesktop.Picture = LoadPicture(WallpaperPath)
    End If
    
nextcode:
    cmdSample.ImageList = frmMain.imgDownload
    
#If HIDEYTDL Then
    txtYtdlPath.Visible = False
    chkAutoYtdl.Visible = False
    Label7.Visible = False
    Frame2.Height = Frame2.Height - chkAutoYtdl.Height
    Frame2.Refresh
    Frame5.Top = Frame5.Top - chkAutoYtdl.Height
    Frame5.Refresh
    Label3.Top = Label3.Top - chkAutoYtdl.Height
    cbWhenExist.Top = cbWhenExist.Top - chkAutoYtdl.Height
    FrameW2.Height = FrameW2.Height - txtYtdlPath.Height - 120
    FrameW2.Refresh
    FrameW4.Top = FrameW4.Top - txtYtdlPath.Height - 120
    FrameW4.Refresh
#End If
    
    Loaded = True
End Sub

Sub DrawTabBackground()
    On Error Resume Next
    Dim ctrl As Control
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "PictureBox" And ctrl.Tag <> "nobgdraw" Then
            ctrl.AutoRedraw = True
            tsTabStrip.DrawBackground ctrl.hWnd, ctrl.hDC
        End If
    Next ctrl
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "FrameW" Then
            ctrl.Transparent = True
        End If
        ctrl.Refresh
    Next ctrl
End Sub

Sub SetPreviewPosition()
    If PrevhWnd Then DestroyWindow (PrevhWnd)
    Dim CaptionHeight As Integer
    CaptionHeight = GetSystemMetrics(31)
    Dim Left%, Top%
    Left = 30
    Top = 6
    PrevhWnd = CreateWindowEx(WS_EX_CLIENTEDGE, "STATIC", App.Title, WS_DISABLED Or WS_CHILD Or WS_VISIBLE Or WS_BORDER Or WS_OVERLAPPED Or WS_CAPTION Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_SYSMENU, Left, Top, pbBackground.Width / 15 + SizingBorderWidth * 2, pbBackground.Height / 15 + CaptionHeight + SizingBorderWidth * 2 + 1, pbPreview.hWnd, 0&, App.hInstance, 0&)
    pbBackground.Top = pbPreview.Top + Top * 15 + CaptionHeight * 15 + DialogBorderWidth * 15 + PaddedBorderWidth * 15 + 15 + 30
    pbBackground.Left = pbPreview.Left + Left * 15 + DialogBorderWidth * 15 + PaddedBorderWidth * 15 + 30
    imgPreview.Width = pbBackground.Width
    imgPreview.Height = pbBackground.Height
    RedrawPreview
End Sub

Private Sub lblSelectColor_Click()
    Dim Color As OLE_COLOR
    Color = ShowColorDialog(Me.hWnd, True, pgColor.BackColor)
    If Color = -1 Then Exit Sub
    pgColor.BackColor = Color
    cmdApply.Enabled = -1
    optUserColor.Value = True
    ColorChanged = True
    pbBackground.BackColor = pgColor.BackColor
    cmdSample.Refresh
    RedrawPreview
End Sub

Private Sub lblSelectFore_Click()
    Dim Color As OLE_COLOR
    Color = ShowColorDialog(Me.hWnd, True, pgFore.BackColor, True)
    If Color = -1 Then Exit Sub
    pgFore.BackColor = Color
    cmdApply.Enabled = -1
    optUserFore.Value = True
    ColorChanged = True
    Label11.ForeColor = pgFore.BackColor
End Sub

Private Sub lvHeaders_ColumnFilterChanged(ByVal ColumnHeader As LvwColumnHeader)
    Dim i%
    Dim startIdx As Integer
    startIdx = 0
    On Error Resume Next
    startIdx = lvHeaders.SelectedItem.Index
    If Not lvHeaders.SelectedItem.Selected Then startIdx = 0
    For i = startIdx + 1 To lvHeaders.ListItems.Count
        If i > lvHeaders.ListItems.Count Then Exit For
        If (ColumnHeader.Index = 1 And Replace(lvHeaders.ListItems(i).Text, ColumnHeader.FilterValue, "") <> lvHeaders.ListItems(i).Text) Or _
            (ColumnHeader.Index = 2 And Replace(lvHeaders.ListItems(i).ListSubItems(1).Text, ColumnHeader.FilterValue, "") <> lvHeaders.ListItems(i).Text) Then
            lvHeaders.ListItems(i).Selected = True
            lvHeaders.ListItems(i).EnsureVisible
            Exit For
        End If
    Next i
End Sub

Private Sub lvHeaders_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseY = Y
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
    pbBackground.BackColor = &H8000000F
    cmdSample.Refresh
    RedrawPreview
End Sub

Private Sub optSystemFore_Click()
    If Loaded Then
        cmdApply.Enabled = -1
        ColorChanged = True
    End If
    Label11.ForeColor = &H80000012
End Sub

Private Sub optUserColor_Click()
    If Loaded Then
        cmdApply.Enabled = -1
        ColorChanged = True
    End If
    pbBackground.BackColor = pgColor.BackColor
    cmdSample.Refresh
    RedrawPreview
End Sub

Private Sub optUserFore_Click()
    If Loaded Then
        cmdApply.Enabled = -1
        ColorChanged = True
    End If
    Label11.ForeColor = pgFore.BackColor
End Sub

Private Sub tsTabStrip_TabClick(ByVal TabItem As TbsTab)
    On Error Resume Next
    Dim i%
    For i = 1 To pbPanel.Count
        If i = TabItem.Index Then
            pbPanel(i).Visible = -1
            pbPanel(i).Enabled = -1
            pbPanel(i).SetFocus
        Else
            pbPanel(i).Visible = 0
            pbPanel(i).Enabled = 0
        End If
    Next i
    
    If TabItem.Index = 3 Then
        DoEvents
        RedrawPreview
    End If
End Sub

Sub RedrawPreview()
    DoEvents
    pbBackground.Refresh
    cmdSample.Refresh
    Dim ctrl As Control
    On Error Resume Next
    For Each ctrl In Me.Controls
        If ctrl.Container Is pbBackground Then
            ctrl.Refresh
        End If
    Next ctrl
    FrameW5.Refresh
    CheckBoxW1.Refresh
End Sub

Private Sub txtError_Change()
    If Loaded Then cmdApply.Enabled = -1
End Sub

Private Sub txtExclamation_Change()
    If Loaded Then cmdApply.Enabled = -1
End Sub

Private Sub txtNodePath_Change()
    If Loaded Then cmdApply.Enabled = -1
End Sub

Private Sub txtQuestion_Change()
    If Loaded Then cmdApply.Enabled = -1
End Sub

Private Sub txtScriptPath_Change()
    If Loaded Then cmdApply.Enabled = -1
End Sub

Private Sub txtYtdlPath_Change()
    If Loaded Then cmdApply.Enabled = -1
End Sub
