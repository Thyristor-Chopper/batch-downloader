VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "옵션"
   ClientHeight    =   16470
   ClientLeft      =   2760
   ClientTop       =   3855
   ClientWidth     =   14775
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   16470
   ScaleWidth      =   14775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.PictureBox pbPanel 
      AutoRedraw      =   -1  'True
      Height          =   5055
      Index           =   3
      Left            =   7080
      ScaleHeight     =   4995
      ScaleWidth      =   6675
      TabIndex        =   3
      Top             =   5880
      Width           =   6735
      Begin VB.FileListBox lvBackgroundFiles 
         Height          =   450
         Left            =   5640
         TabIndex        =   136
         TabStop         =   0   'False
         Top             =   960
         Visible         =   0   'False
         Width           =   615
      End
      Begin prjDownloadBooster.FrameW FrameW8 
         Height          =   2295
         Left            =   3480
         TabIndex        =   57
         Top             =   2640
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   4048
         Caption         =   " 배경 무늬(&W) "
         Begin VB.ListBox lvBackgrounds 
            Height          =   960
            Left            =   120
            TabIndex        =   58
            Top             =   240
            Width           =   2895
         End
         Begin VB.ComboBox cbImagePosition 
            Height          =   300
            Left            =   960
            Style           =   2  '드롭다운 목록
            TabIndex        =   61
            Top             =   1680
            Width           =   2055
         End
         Begin prjDownloadBooster.CheckBoxW chkImageCentered 
            Height          =   255
            Left            =   960
            TabIndex        =   62
            Top             =   2010
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   450
            Caption         =   "가운데(&C)"
         End
         Begin prjDownloadBooster.CommandButtonW cmdChooseBackground 
            Height          =   330
            Left            =   960
            TabIndex        =   59
            Top             =   1230
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   582
            Caption         =   "찾아보기(&B)..."
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '투명
            Caption         =   "위치(&P):"
            Height          =   255
            Left            =   120
            TabIndex        =   60
            Top             =   1725
            Width           =   840
         End
      End
      Begin prjDownloadBooster.FrameW FrameW7 
         Height          =   2295
         Left            =   120
         TabIndex        =   53
         Top             =   2640
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   4048
         Caption         =   " 무늬(&P) "
         Begin VB.ListBox lvPatterns 
            Height          =   960
            Left            =   120
            TabIndex        =   54
            Top             =   240
            Width           =   2895
         End
         Begin VB.Label lblFillColorSelect 
            BackStyle       =   0  '투명
            Height          =   255
            Left            =   840
            TabIndex        =   56
            Top             =   1320
            Width           =   615
         End
         Begin VB.Shape pgPatternColor 
            BackColor       =   &H00000000&
            BackStyle       =   1  '투명하지 않음
            BorderColor     =   &H00404040&
            FillColor       =   &H00808080&
            Height          =   255
            Left            =   840
            Shape           =   4  '둥근 사각형
            Top             =   1320
            Width           =   615
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "색(&O):"
            Height          =   180
            Left            =   120
            TabIndex        =   55
            Top             =   1350
            Width           =   525
         End
      End
      Begin VB.PictureBox pbBackgroundPreview 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  '없음
         Height          =   2295
         Left            =   2040
         ScaleHeight     =   2295
         ScaleWidth      =   2775
         TabIndex        =   52
         Top             =   240
         Width           =   2775
         Begin VB.Image Image7 
            Height          =   2190
            Left            =   0
            Picture         =   "frmOptions.frx":000C
            Top             =   0
            Width           =   2655
         End
         Begin VB.Image imgBackgroundPreview 
            Height          =   255
            Left            =   120
            Stretch         =   -1  'True
            Top             =   120
            Width           =   255
         End
         Begin VB.Shape pgPatternPreview 
            BackColor       =   &H8000000F&
            BackStyle       =   1  '투명하지 않음
            BorderStyle     =   0  '투명
            Height          =   255
            Left            =   120
            Top             =   120
            Width           =   255
         End
      End
   End
   Begin VB.PictureBox pbPanel 
      Height          =   4545
      Index           =   5
      Left            =   120
      ScaleHeight     =   4485
      ScaleWidth      =   6555
      TabIndex        =   5
      Top             =   8880
      Width           =   6615
      Begin prjDownloadBooster.FrameW FrameW6 
         Height          =   4335
         Left            =   120
         TabIndex        =   84
         Top             =   120
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   7646
         Caption         =   " 효과음 설정 "
         Begin prjDownloadBooster.FrameW fCompleteSound 
            Height          =   735
            Left            =   720
            TabIndex        =   86
            Top             =   720
            Width           =   5535
            _ExtentX        =   9763
            _ExtentY        =   1296
            BorderStyle     =   0
            Caption         =   "                             "
            Begin prjDownloadBooster.CommandButtonW cmdBrowseCompleteSound 
               Height          =   300
               Left            =   4560
               TabIndex        =   89
               Top             =   330
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   529
               ImageList       =   "imgBrowse"
               ImageListAlignment=   4
            End
            Begin VB.TextBox txtCompleteSoundPath 
               Height          =   300
               Left            =   360
               TabIndex        =   88
               Top             =   330
               Width           =   4095
            End
            Begin prjDownloadBooster.CheckBoxW chkBeepWhenComplete 
               Height          =   255
               Left            =   120
               TabIndex        =   87
               Top             =   0
               Width           =   2535
               _ExtentX        =   4471
               _ExtentY        =   450
               Caption         =   "다운로드 완료(&B)"
            End
            Begin prjDownloadBooster.CommandButtonW cmdTestCompleteSound 
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
         Begin prjDownloadBooster.FrameW fAsterisk 
            Height          =   735
            Left            =   720
            TabIndex        =   91
            Top             =   1440
            Width           =   5535
            _ExtentX        =   9763
            _ExtentY        =   1296
            BorderStyle     =   0
            Caption         =   "                             "
            Begin prjDownloadBooster.CommandButtonW cmdBrowseAsterisk 
               Height          =   300
               Left            =   4560
               TabIndex        =   94
               Top             =   330
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   529
               ImageList       =   "imgBrowse"
               ImageListAlignment=   4
            End
            Begin VB.TextBox txtAsterisk 
               Height          =   300
               Left            =   360
               TabIndex        =   93
               Top             =   330
               Width           =   4095
            End
            Begin prjDownloadBooster.CheckBoxW chkAsterisk 
               Height          =   255
               Left            =   120
               TabIndex        =   92
               Top             =   0
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   450
               Caption         =   "일반 메시지(&A)"
            End
            Begin prjDownloadBooster.CommandButtonW cmdTestAsterisk 
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
         Begin prjDownloadBooster.FrameW fExclamation 
            Height          =   735
            Left            =   720
            TabIndex        =   96
            Top             =   2160
            Width           =   5535
            _ExtentX        =   9763
            _ExtentY        =   1296
            BorderStyle     =   0
            Caption         =   "                             "
            Begin prjDownloadBooster.CommandButtonW cmdBrowseExclamation 
               Height          =   300
               Left            =   4560
               TabIndex        =   99
               Top             =   330
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   529
               ImageList       =   "imgBrowse"
               ImageListAlignment=   4
            End
            Begin VB.TextBox txtExclamation 
               Height          =   300
               Left            =   360
               TabIndex        =   98
               Top             =   330
               Width           =   4095
            End
            Begin prjDownloadBooster.CheckBoxW chkExclamation 
               Height          =   255
               Left            =   120
               TabIndex        =   97
               Top             =   0
               Width           =   3135
               _ExtentX        =   5530
               _ExtentY        =   450
               Caption         =   "경고 메시지(&E)"
            End
            Begin prjDownloadBooster.CommandButtonW cmdTestExclamation 
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
         Begin prjDownloadBooster.FrameW fError 
            Height          =   735
            Left            =   720
            TabIndex        =   101
            Top             =   2880
            Width           =   5535
            _ExtentX        =   9763
            _ExtentY        =   1296
            BorderStyle     =   0
            Caption         =   "                             "
            Begin prjDownloadBooster.CommandButtonW cmdBrowseError 
               Height          =   300
               Left            =   4560
               TabIndex        =   104
               Top             =   330
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   529
               ImageList       =   "imgBrowse"
               ImageListAlignment=   4
            End
            Begin VB.TextBox txtError 
               Height          =   300
               Left            =   360
               TabIndex        =   103
               Top             =   330
               Width           =   4095
            End
            Begin prjDownloadBooster.CheckBoxW chkError 
               Height          =   255
               Left            =   120
               TabIndex        =   102
               Top             =   0
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   450
               Caption         =   "오류 메시지(&R)"
            End
            Begin prjDownloadBooster.CommandButtonW cmdTestError 
               Height          =   300
               Left            =   5160
               TabIndex        =   105
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
            Height          =   675
            Left            =   720
            TabIndex        =   106
            Top             =   3600
            Width           =   5535
            _ExtentX        =   9763
            _ExtentY        =   1191
            BorderStyle     =   0
            Caption         =   "                             "
            Begin prjDownloadBooster.CommandButtonW cmdBrowseQuestion 
               Height          =   300
               Left            =   4560
               TabIndex        =   109
               Top             =   330
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   529
               ImageList       =   "imgBrowse"
               ImageListAlignment=   4
            End
            Begin VB.TextBox txtQuestion 
               Height          =   300
               Left            =   360
               TabIndex        =   108
               Top             =   330
               Width           =   4095
            End
            Begin prjDownloadBooster.CheckBoxW chkQuestion 
               Height          =   255
               Left            =   120
               TabIndex        =   107
               Top             =   0
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   450
               Caption         =   "질문(&Q)"
            End
            Begin prjDownloadBooster.CommandButtonW cmdTestQuestion 
               Height          =   300
               Left            =   5160
               TabIndex        =   110
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
         Begin VB.Label Label12 
            BackStyle       =   0  '투명
            Caption         =   "기본값을 사용하려면 필드를 비워두십시오."
            Height          =   255
            Left            =   840
            TabIndex        =   85
            Top             =   360
            Width           =   4815
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   120
            Picture         =   "frmOptions.frx":0F45
            Top             =   240
            Width           =   480
         End
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
      InitListImages  =   "frmOptions.frx":138F
   End
   Begin VB.PictureBox pbPanel 
      AutoRedraw      =   -1  'True
      Height          =   4905
      Index           =   2
      Left            =   6960
      ScaleHeight     =   4845
      ScaleWidth      =   6555
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   11040
      Width           =   6615
      Begin prjDownloadBooster.FrameW FrameW3 
         Height          =   1215
         Left            =   120
         TabIndex        =   38
         Top             =   120
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   2143
         Caption         =   " 연결 설정 "
         Begin prjDownloadBooster.Slider trRequestInterval 
            Height          =   450
            Left            =   3000
            TabIndex        =   43
            Top             =   720
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   794
            Max             =   7
            Value           =   2
            ShowTip         =   0   'False
            Transparent     =   -1  'True
         End
         Begin prjDownloadBooster.CheckBoxW chkIgnore300 
            Height          =   255
            Left            =   3720
            TabIndex        =   40
            Top             =   240
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   450
            Caption         =   "300번대 응답 코드 무시(&I)"
         End
         Begin prjDownloadBooster.CheckBoxW chkForceGet 
            Height          =   255
            Left            =   840
            TabIndex        =   41
            Top             =   480
            Width           =   2865
            _ExtentX        =   5054
            _ExtentY        =   450
            Caption         =   "파일 검사 시 GET 요청(&Q)"
         End
         Begin prjDownloadBooster.CheckBoxW chkNoRedirectCheck 
            Height          =   255
            Left            =   840
            TabIndex        =   39
            Top             =   240
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   450
            Caption         =   "리다이렉트 검사 안 함(&R)"
         End
         Begin VB.Label lblIntervalDisplay 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "(0.1초)"
            Height          =   180
            Left            =   5070
            TabIndex        =   44
            Top             =   840
            Width           =   570
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "스레드 요청 간격(&N):"
            Height          =   180
            Left            =   1080
            TabIndex        =   42
            Top             =   840
            Width           =   1725
         End
         Begin VB.Image Image4 
            Height          =   480
            Left            =   120
            Picture         =   "frmOptions.frx":1777
            Top             =   240
            Width           =   480
         End
      End
      Begin prjDownloadBooster.FrameW fHeaders 
         Height          =   3375
         Left            =   120
         TabIndex        =   45
         Top             =   1440
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   5953
         Caption         =   " 헤더 설정 "
         Begin prjDownloadBooster.CommandButtonW cmdEditHeaderName 
            Height          =   330
            Left            =   3480
            TabIndex        =   49
            Top             =   2940
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   582
            Enabled         =   0   'False
            Caption         =   "이름변경(&R)"
         End
         Begin VB.TextBox txtEdit 
            Height          =   255
            Left            =   3360
            TabIndex        =   51
            Top             =   2520
            Visible         =   0   'False
            Width           =   2535
         End
         Begin prjDownloadBooster.CommandButtonW cmdDeleteHeader 
            Height          =   330
            Left            =   2160
            TabIndex        =   48
            Top             =   2940
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   582
            Enabled         =   0   'False
            Caption         =   "삭제(&D)"
         End
         Begin prjDownloadBooster.CommandButtonW cmdEditHeaderValue 
            Height          =   330
            Left            =   4800
            TabIndex        =   50
            Top             =   2940
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   582
            Enabled         =   0   'False
            Caption         =   "편집(&E)"
         End
         Begin prjDownloadBooster.CommandButtonW cmdAddHeader 
            Height          =   330
            Left            =   840
            TabIndex        =   47
            Top             =   2940
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   582
            Caption         =   "추가(&A)"
         End
         Begin prjDownloadBooster.ListView lvHeaders 
            Height          =   2145
            Left            =   840
            TabIndex        =   46
            Top             =   720
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   3784
            VisualTheme     =   1
            View            =   3
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            HideSelection   =   0   'False
            ShowLabelTips   =   -1  'True
            HighlightColumnHeaders=   -1  'True
            AutoSelectFirstItem=   0   'False
         End
         Begin VB.Image Image5 
            Height          =   480
            Left            =   120
            Picture         =   "frmOptions.frx":1BB9
            Top             =   240
            Width           =   480
         End
         Begin VB.Label Label17 
            BackStyle       =   0  '투명
            Caption         =   "다운로드 중 서버에 요청할 때 전송할 헤더를 설정합니다. [다운로드 설정]에서 설정한 헤더가 우선적으로 적용됩니다."
            Height          =   495
            Left            =   840
            TabIndex        =   132
            Top             =   240
            Width           =   5175
         End
      End
   End
   Begin VB.PictureBox pbPanel 
      AutoRedraw      =   -1  'True
      Height          =   4905
      Index           =   1
      Left            =   120
      ScaleHeight     =   4845
      ScaleWidth      =   6555
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   600
      Width           =   6615
      Begin prjDownloadBooster.FrameW Frame5 
         Height          =   2715
         Left            =   120
         TabIndex        =   20
         Top             =   2160
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   4789
         Caption         =   " 인터페이스 "
         Begin prjDownloadBooster.OptionButtonW optScreenPerScroll 
            Height          =   255
            Left            =   4200
            TabIndex        =   37
            Top             =   2415
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   450
            Caption         =   "한 화면씩(&R)"
         End
         Begin prjDownloadBooster.OptionButtonW optLinePerScroll 
            Height          =   255
            Left            =   2640
            TabIndex        =   36
            Top             =   2415
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   450
            Caption         =   "한 줄씩(&N)"
         End
         Begin prjDownloadBooster.CheckBoxW chkAllowDuplicates 
            Height          =   255
            Left            =   840
            TabIndex        =   27
            Top             =   1440
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   450
            Caption         =   "일괄 처리 목록에 중복 항목 허용(&I)"
         End
         Begin prjDownloadBooster.CheckBoxW chkDontLoadIcons 
            Height          =   255
            Left            =   840
            TabIndex        =   26
            Top             =   1200
            Width           =   4185
            _ExtentX        =   7382
            _ExtentY        =   450
            Caption         =   "열기 대화 상자에서 같은 파일 아이콘 사용(&M)"
         End
         Begin prjDownloadBooster.CheckBoxW chkForceOldDialog 
            Height          =   255
            Left            =   840
            TabIndex        =   25
            Top             =   960
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   450
            Caption         =   "윈도우 3.1 대화 상자 사용(&S)"
         End
         Begin prjDownloadBooster.CheckBoxW chkExcludeMergeFromElapsed 
            Height          =   255
            Left            =   840
            TabIndex        =   24
            Top             =   720
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   450
            Caption         =   "경과 시간에서 파일 조각 결합 시간 제외(&E)"
         End
         Begin prjDownloadBooster.CheckBoxW chkLazyElapsed 
            Height          =   255
            Left            =   840
            TabIndex        =   23
            Top             =   480
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   450
            Caption         =   "첫 바이트 수신 후 경과 시간 계산(&C)"
         End
         Begin prjDownloadBooster.CheckBoxW chkAeroWindow 
            Height          =   255
            Left            =   2880
            TabIndex        =   22
            Top             =   240
            Visible         =   0   'False
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   450
            Enabled         =   0   'False
            Caption         =   "유리 창 효과 사용(&G)"
         End
         Begin prjDownloadBooster.CheckBoxW chkAlwaysOnTop 
            Height          =   255
            Left            =   840
            TabIndex        =   21
            Top             =   240
            Width           =   1920
            _ExtentX        =   3387
            _ExtentY        =   450
            Caption         =   "항상 위에 표시(&W)"
         End
         Begin VB.ComboBox cbLanguage 
            Height          =   300
            Left            =   2640
            Style           =   2  '드롭다운 목록
            TabIndex        =   29
            Top             =   1710
            Width           =   1455
         End
         Begin VB.ComboBox cbFont 
            Height          =   300
            Left            =   2640
            TabIndex        =   33
            Top             =   2070
            Width           =   1455
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "스레드 스크롤(&H):"
            Height          =   180
            Left            =   1080
            TabIndex        =   35
            Top             =   2460
            Width           =   1470
         End
         Begin VB.Label Label16 
            BackStyle       =   0  '투명
            Caption         =   "(다시 시작 필요)"
            Height          =   255
            Left            =   4200
            TabIndex        =   30
            Top             =   1770
            Width           =   1575
         End
         Begin VB.Label lblFontPreview 
            Alignment       =   2  '가운데 맞춤
            BorderStyle     =   1  '단일 고정
            Height          =   270
            Left            =   4200
            TabIndex        =   34
            Tag             =   "nobgdraw"
            Top             =   2085
            Width           =   1575
         End
         Begin VB.Image Image3 
            Height          =   405
            Left            =   120
            Picture         =   "frmOptions.frx":1FFB
            Top             =   240
            Width           =   435
         End
         Begin VB.Label Label13 
            BackStyle       =   0  '투명
            Caption         =   "글꼴(&F):"
            Height          =   255
            Left            =   1080
            TabIndex        =   32
            Top             =   2115
            Width           =   1335
         End
         Begin VB.Label Label9 
            BackStyle       =   0  '투명
            Caption         =   "..."
            Height          =   255
            Left            =   1080
            TabIndex        =   31
            Top             =   2070
            Visible         =   0   'False
            Width           =   4200
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '투명
            Caption         =   "언어(&L):"
            Height          =   255
            Left            =   1080
            TabIndex        =   28
            Tag             =   "nocolorchange"
            Top             =   1755
            Width           =   975
         End
      End
      Begin prjDownloadBooster.FrameW Frame2 
         Height          =   1935
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   3413
         Caption         =   " 다운로드 설정 "
         Begin prjDownloadBooster.UpDown udMaxThreadCount 
            Height          =   255
            Left            =   3630
            Top             =   1575
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
            BuddyControl    =   "txtMaxThreadCount"
            BuddyProperty   =   "frmOptions.frx":222D
            Min             =   2
            Max             =   91
            Value           =   25
            ThousandsSeparator=   0   'False
         End
         Begin VB.TextBox txtMaxThreadCount 
            Alignment       =   1  '오른쪽 맞춤
            BorderStyle     =   0  '없음
            Height          =   210
            Left            =   3255
            TabIndex        =   18
            Top             =   1605
            Width           =   360
         End
         Begin prjDownloadBooster.CheckBoxW chkAutoYtdl 
            Height          =   255
            Left            =   840
            TabIndex        =   13
            Top             =   960
            Width           =   4815
            _ExtentX        =   8493
            _ExtentY        =   450
            Caption         =   "지원되는 링크에서 자동으로 youtube-dl 사용(&Y)"
         End
         Begin VB.ComboBox cbWhenExist 
            Height          =   300
            Left            =   3210
            Style           =   2  '드롭다운 목록
            TabIndex        =   15
            Top             =   1230
            Width           =   2040
         End
         Begin prjDownloadBooster.CheckBoxW chkAutoRetry 
            Height          =   255
            Left            =   840
            TabIndex        =   12
            Top             =   720
            Width           =   4890
            _ExtentX        =   8625
            _ExtentY        =   450
            Caption         =   "네트워크 오류 시 자동 재시도(&U)"
         End
         Begin prjDownloadBooster.CheckBoxW chkRememberURL 
            Height          =   255
            Left            =   840
            TabIndex        =   10
            Top             =   480
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   450
            Caption         =   "파일 주소 기억(&M)"
         End
         Begin prjDownloadBooster.CheckBoxW chkAlwaysResume 
            Height          =   255
            Left            =   3480
            TabIndex        =   11
            Top             =   480
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   450
            Caption         =   "항상 이어받기(&A)"
         End
         Begin prjDownloadBooster.CheckBoxW chkOpenDirWhenComplete 
            Height          =   255
            Left            =   3480
            TabIndex        =   9
            Top             =   240
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   450
            Caption         =   "완료 후 폴더 열기(&P)"
         End
         Begin prjDownloadBooster.CheckBoxW chkOpenWhenComplete 
            Height          =   255
            Left            =   840
            TabIndex        =   8
            Top             =   240
            Width           =   2520
            _ExtentX        =   4445
            _ExtentY        =   450
            Caption         =   "완료 후 파일 열기(&O)"
         End
         Begin VB.TextBox txtOuterMaxThreadCount 
            Enabled         =   0   'False
            Height          =   285
            Left            =   3210
            Locked          =   -1  'True
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   1560
            Width           =   690
         End
         Begin VB.Image Image2 
            Height          =   480
            Left            =   120
            Picture         =   "frmOptions.frx":225D
            Top             =   240
            Width           =   480
         End
         Begin VB.Label Label15 
            BackStyle       =   0  '투명
            Caption         =   "개 (다시 시작 필요)"
            Height          =   255
            Left            =   3960
            TabIndex        =   19
            Top             =   1605
            Width           =   2055
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "최대 스레드 개수(&X):"
            Height          =   180
            Left            =   1080
            TabIndex        =   16
            Top             =   1605
            Width           =   1710
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "중복 파일명 처리(&D):"
            Height          =   180
            Left            =   1080
            TabIndex        =   14
            Tag             =   "nocolorchange"
            Top             =   1275
            Width           =   1710
         End
      End
   End
   Begin VB.PictureBox pbPanel 
      AutoRedraw      =   -1  'True
      Height          =   2895
      Index           =   6
      Left            =   120
      ScaleHeight     =   2835
      ScaleWidth      =   6555
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   5880
      Width           =   6615
      Begin prjDownloadBooster.FrameW FrameW4 
         Height          =   735
         Left            =   120
         TabIndex        =   119
         Top             =   2040
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   1296
         Caption         =   " 고급 다운로드 설정 "
         Begin prjDownloadBooster.CheckBoxW chkNoCleanup 
            Height          =   255
            Left            =   840
            TabIndex        =   120
            Top             =   240
            Width           =   2250
            _ExtentX        =   3969
            _ExtentY        =   450
            Caption         =   "조각 파일 유지(&N)"
            Transparent     =   -1  'True
         End
         Begin VB.Image Image6 
            Height          =   480
            Left            =   120
            Picture         =   "frmOptions.frx":269F
            Top             =   240
            Width           =   480
         End
      End
      Begin prjDownloadBooster.FrameW FrameW2 
         Height          =   1815
         Left            =   120
         TabIndex        =   111
         Top             =   120
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   3201
         Caption         =   " 경로 설정 "
         Begin VB.TextBox txtYtdlPath 
            Height          =   255
            Left            =   2760
            TabIndex        =   118
            Top             =   1440
            Width           =   3495
         End
         Begin VB.TextBox txtNodePath 
            Height          =   255
            Left            =   2760
            TabIndex        =   114
            Top             =   720
            Width           =   3495
         End
         Begin VB.TextBox txtScriptPath 
            Height          =   255
            Left            =   2760
            TabIndex        =   116
            Top             =   1080
            Width           =   3495
         End
         Begin VB.Image imgIcon2 
            Height          =   480
            Left            =   120
            Picture         =   "frmOptions.frx":2AE1
            Top             =   240
            Width           =   480
         End
         Begin VB.Label Label6 
            BackStyle       =   0  '투명
            Caption         =   "기본값을 사용하려면 필드를 비워두십시오. 아래는 고급 사용자를 위한 것이며 일반적으로 변경할 필요가 없습니다."
            Height          =   480
            Left            =   840
            TabIndex        =   112
            Top             =   240
            Width           =   5415
         End
         Begin VB.Label Label7 
            BackStyle       =   0  '투명
            Caption         =   "&youtube-dl/yt-dlp:"
            Height          =   255
            Left            =   840
            TabIndex        =   117
            Top             =   1470
            Width           =   1695
         End
         Begin VB.Label Label5 
            BackStyle       =   0  '투명
            Caption         =   "다운로드 스크립트(&D):"
            Height          =   255
            Left            =   840
            TabIndex        =   115
            Top             =   1110
            Width           =   1935
         End
         Begin VB.Label Label4 
            BackStyle       =   0  '투명
            Caption         =   "N&ode.js:"
            Height          =   255
            Left            =   840
            TabIndex        =   113
            Top             =   750
            Width           =   1455
         End
      End
   End
   Begin VB.PictureBox pbPanel 
      AutoRedraw      =   -1  'True
      Height          =   5025
      Index           =   4
      Left            =   7560
      ScaleHeight     =   4965
      ScaleWidth      =   6435
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   600
      Visible         =   0   'False
      Width           =   6495
      Begin prjDownloadBooster.FrameW FrameW1 
         Height          =   1230
         Left            =   120
         TabIndex        =   79
         Top             =   2400
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   2170
         Caption         =   " 테마 "
         Begin prjDownloadBooster.CommandButtonW cmdDeleteTheme 
            Height          =   300
            Left            =   960
            TabIndex        =   83
            Top             =   885
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   529
            Caption         =   "삭제(&D)"
         End
         Begin prjDownloadBooster.CommandButtonW cmdSaveTheme 
            Height          =   300
            Left            =   960
            TabIndex        =   82
            Top             =   570
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   529
            Caption         =   "저장(&A)..."
         End
         Begin VB.ComboBox cbTheme 
            Height          =   300
            Left            =   960
            Style           =   2  '드롭다운 목록
            TabIndex        =   81
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "테마(&T):"
            Height          =   180
            Left            =   120
            TabIndex        =   80
            Top             =   285
            Width           =   690
         End
      End
      Begin prjDownloadBooster.FrameW Frame6 
         Height          =   1230
         Left            =   120
         TabIndex        =   73
         Top             =   3720
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   2170
         Caption         =   " 스킨 "
         Begin prjDownloadBooster.CommandButtonW cmdAdvancedSkin 
            Height          =   285
            Left            =   990
            TabIndex        =   78
            Top             =   900
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   503
            Caption         =   "고급(&V)..."
         End
         Begin VB.ComboBox cbSkin 
            Height          =   300
            Left            =   990
            Style           =   2  '드롭다운 목록
            TabIndex        =   77
            Top             =   585
            Width           =   1965
         End
         Begin VB.ComboBox cbFrameSkin 
            Height          =   300
            Left            =   990
            Style           =   2  '드롭다운 목록
            TabIndex        =   75
            Top             =   240
            Width           =   1965
         End
         Begin VB.Label Label10 
            BackStyle       =   0  '투명
            Caption         =   "창(&W):"
            Height          =   255
            Left            =   120
            TabIndex        =   74
            Top             =   285
            Width           =   735
         End
         Begin VB.Label Label8 
            BackStyle       =   0  '투명
            Caption         =   "단추(&O):"
            Height          =   255
            Left            =   120
            TabIndex        =   76
            Top             =   630
            Width           =   735
         End
      End
      Begin prjDownloadBooster.FrameW Frame4 
         Height          =   1230
         Left            =   3360
         TabIndex        =   68
         Top             =   3720
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   2170
         Caption         =   " 글자색 "
         Begin prjDownloadBooster.CheckBoxW chkForeColorMainOnly 
            Height          =   255
            Left            =   360
            TabIndex        =   72
            Top             =   900
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   450
            Caption         =   "메인 창에만 적용(&N)"
         End
         Begin prjDownloadBooster.OptionButtonW optUserFore 
            Height          =   255
            Left            =   120
            TabIndex        =   70
            Top             =   600
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   450
            Caption         =   "사용자 지정(&U)"
         End
         Begin prjDownloadBooster.OptionButtonW optSystemFore 
            Height          =   255
            Left            =   120
            TabIndex        =   69
            Top             =   240
            Width           =   1815
            _ExtentX        =   0
            _ExtentY        =   0
            Caption         =   "시스템 색상(&Y)"
         End
         Begin VB.Label lblSystemWindowText 
            BackStyle       =   0  '투명
            Height          =   255
            Left            =   2040
            TabIndex        =   134
            Top             =   240
            Width           =   615
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H80000008&
            BackStyle       =   1  '투명하지 않음
            BorderColor     =   &H00404040&
            FillColor       =   &H00808080&
            Height          =   255
            Left            =   2040
            Shape           =   4  '둥근 사각형
            Top             =   240
            Width           =   615
         End
         Begin VB.Label lblSelectFore 
            BackStyle       =   0  '투명
            Height          =   255
            Left            =   1800
            TabIndex        =   71
            Top             =   240
            Width           =   1455
         End
         Begin VB.Shape pgFore 
            BackStyle       =   1  '투명하지 않음
            BorderColor     =   &H00404040&
            FillColor       =   &H00808080&
            Height          =   255
            Left            =   2040
            Shape           =   4  '둥근 사각형
            Top             =   600
            Width           =   615
         End
      End
      Begin VB.PictureBox pbOuterPreview 
         AutoRedraw      =   -1  'True
         Height          =   2175
         Left            =   120
         ScaleHeight     =   2115
         ScaleWidth      =   6195
         TabIndex        =   124
         Top             =   120
         Width           =   6255
         Begin prjDownloadBooster.SmallWindow pbBackground 
            Height          =   1860
            Left            =   480
            TabIndex        =   126
            TabStop         =   0   'False
            Tag             =   "nobgdraw"
            Top             =   120
            Width           =   3915
            _ExtentX        =   6906
            _ExtentY        =   3281
            Caption         =   "다운로드 부스터"
            MaximizeBox     =   0   'False
            Begin VB.TextBox txtSampleClassic 
               Height          =   270
               Left            =   1080
               TabIndex        =   135
               Top             =   105
               Visible         =   0   'False
               Width           =   2415
            End
            Begin prjDownloadBooster.CheckBoxW CheckBoxW1 
               Height          =   255
               Left            =   120
               TabIndex        =   127
               Top             =   990
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   450
               Caption         =   "완료 후 열기"
               Transparent     =   -1  'True
            End
            Begin VB.TextBox TextBoxW1 
               Height          =   270
               Left            =   1080
               TabIndex        =   128
               Top             =   105
               Width           =   2415
            End
            Begin prjDownloadBooster.FrameW FrameW5 
               Height          =   555
               Left            =   120
               TabIndex        =   129
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
               TabIndex        =   130
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
               TabIndex        =   131
               Top             =   150
               Width           =   975
            End
            Begin VB.Image imgPreview 
               Height          =   135
               Left            =   0
               Stretch         =   -1  'True
               Top             =   0
               Width           =   135
            End
         End
         Begin VB.PictureBox pbPreview 
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000001&
            BorderStyle     =   0  '없음
            Enabled         =   0   'False
            Height          =   2175
            Left            =   0
            ScaleHeight     =   2175
            ScaleWidth      =   6255
            TabIndex        =   125
            TabStop         =   0   'False
            Tag             =   "nobgdraw"
            Top             =   0
            Width           =   6255
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
         Height          =   1230
         Left            =   3360
         TabIndex        =   63
         Top             =   2400
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   2170
         Caption         =   " 배경색 "
         Begin prjDownloadBooster.CheckBoxW chkBackColorMainOnly 
            Height          =   255
            Left            =   360
            TabIndex        =   67
            Top             =   900
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   450
            Caption         =   "메인 창에만 적용(&O)"
         End
         Begin prjDownloadBooster.OptionButtonW optSystemColor 
            Height          =   255
            Left            =   120
            TabIndex        =   64
            Top             =   240
            Width           =   1815
            _ExtentX        =   0
            _ExtentY        =   0
            Caption         =   "시스템 색상(&S)"
         End
         Begin prjDownloadBooster.OptionButtonW optUserColor 
            Height          =   255
            Left            =   120
            TabIndex        =   65
            Top             =   600
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   450
            Caption         =   "사용자 지정(&C)"
         End
         Begin VB.Label lblSystemBtnFace 
            BackStyle       =   0  '투명
            Height          =   255
            Left            =   2040
            TabIndex        =   133
            Top             =   240
            Width           =   615
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H8000000F&
            BackStyle       =   1  '투명하지 않음
            BorderColor     =   &H00404040&
            FillColor       =   &H00808080&
            Height          =   255
            Left            =   2040
            Shape           =   4  '둥근 사각형
            Top             =   240
            Width           =   615
         End
         Begin VB.Label lblSelectColor 
            BackStyle       =   0  '투명
            Height          =   255
            Left            =   1800
            TabIndex        =   66
            Top             =   240
            Width           =   1455
         End
         Begin VB.Shape pgColor 
            BackStyle       =   1  '투명하지 않음
            BorderColor     =   &H00404040&
            FillColor       =   &H00808080&
            Height          =   255
            Left            =   2040
            Shape           =   4  '둥근 사각형
            Top             =   600
            Width           =   615
         End
      End
   End
   Begin prjDownloadBooster.CommandButtonW cmdApply 
      Height          =   360
      Left            =   10920
      TabIndex        =   123
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
      TabIndex        =   0
      Top             =   120
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   661
      MultiRow        =   0   'False
      TabFixedWidth   =   53
      TabScrollWheel  =   0   'False
      Transparent     =   -1  'True
      InitTabs        =   "frmOptions.frx":2F23
   End
   Begin prjDownloadBooster.CommandButtonW CancelButton 
      Cancel          =   -1  'True
      Height          =   360
      Left            =   9480
      TabIndex        =   122
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
      TabIndex        =   121
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
      InitListImages  =   "frmOptions.frx":310B
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
Public ColorChanged As Boolean
Public ImageChanged As Boolean
Dim VisualStyleChanged As Boolean
Dim SkinChanged As Boolean
Dim FontChanged As Boolean
Dim PatternChanged As Boolean
Dim MouseY As Integer, SelectedListItem As LvwListItem
Dim BackgroundDrawn(6) As Boolean
Dim ScrollChanged As Boolean
Dim IntervalValues(8) As Single
Public ChangedBackgroundPath$

Implements ISubclass

Sub NextTabPage(Optional ByVal Reverse As Boolean = False)
    On Error Resume Next
    If Not Reverse Then
        If tsTabStrip.SelectedItem.Index = tsTabStrip.Tabs.Count Then
            tsTabStrip.Tabs(1).Selected = True
        Else
            tsTabStrip.Tabs(tsTabStrip.SelectedItem.Index + 1).Selected = True
        End If
    Else
        If tsTabStrip.SelectedItem.Index = 1 Then
            tsTabStrip.Tabs(tsTabStrip.Tabs.Count).Selected = True
        Else
            tsTabStrip.Tabs(tsTabStrip.SelectedItem.Index - 1).Selected = True
        End If
    End If
End Sub

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub OnFontChange()
    If FontExists(cbFont.Text) Then
        lblFontPreview.FontName = cbFont.Text
    ElseIf cbFont.Text = t("(기본값)", "(default)") Then
        If t(1, 2) = 2 Then
            lblFontPreview.FontName = "Tahoma"
        Else
            lblFontPreview.FontName = DefaultFont
        End If
    End If
    lblFontPreview.FontSize = IIf(LCase(lblFontPreview.FontName) = "tahoma" Or Left$(lblFontPreview.FontName, 7) = "Tahoma ", 8, 9)
    lblFontPreview.Caption = CInt(lblFontPreview.FontSize) & " pt. " & lblFontPreview.FontName
    lblFontPreview.FontBold = False
    lblFontPreview.FontItalic = False
    If Loaded Then
        cmdApply.Enabled = -1
        FontChanged = True
    End If
End Sub

Private Sub cbFont_Change()
    OnFontChange
End Sub

Private Sub cbFont_Click()
    OnFontChange
End Sub

Private Sub cbFrameSkin_Click()
    If Loaded Then
        cmdApply.Enabled = -1
        SkinChanged = True
    End If
    
    If (cbFrameSkin.ListCount >= 3 And cbFrameSkin.ListIndex = 2) Or (cbFrameSkin.ListCount < 3 And cbFrameSkin.ListIndex = 1) Then
        RemoveVisualStyles pbBackground.hWnd
    ElseIf Loaded Then
        ActivateVisualStyles pbBackground.hWnd
    End If
    pbBackground.Refresh
End Sub

Private Sub cbImagePosition_Click()
    If cbImagePosition.ListIndex > 0 And cbImagePosition.ListIndex < 4 Then chkImageCentered.Enabled = True Else chkImageCentered.Enabled = False
    If Loaded Then
        cmdApply.Enabled = -1
        ImageChanged = True
    End If
End Sub

Private Sub cbLanguage_Click()
    If Loaded Then
        'Alert t("언어를 변경하려면 프로그램을 재시작해야 합니다.", "To change the language you must restart the application."), App.Title, 64
        cmdApply.Enabled = -1
    End If
End Sub

Private Sub cbSkin_Click()
    cmdSample.VisualStyles = (cbSkin.ListIndex <> 1)
    cmdSample.IsTygemButton = (cbSkin.ListIndex = 2)
    cmdSample.Refresh
    pbSampleClassic.Visible = Not cmdSample.VisualStyles
    txtSampleClassic.Visible = (cbSkin.ListIndex = 1)
    Dim ctrl As Control
    On Error Resume Next
    For Each ctrl In Me.Controls
        If ctrl.Container Is pbBackground And ctrl.Name <> "cmdSample" And ctrl.Name <> "pbSample" And ctrl.Name <> "pbSampleClassic" And ctrl.Name <> "txtSampleClassic" Then
            ctrl.VisualStyles = cmdSample.VisualStyles
        End If
    Next ctrl
    If Loaded Then
        cmdApply.Enabled = -1
        SkinChanged = True
        VisualStyleChanged = True
        If cbSkin.ListIndex = 2 And DPI <> 96 Then
            MsgBox t("이 스킨의 일부 요소는 96 DPI(100% 배율)에서만 표시됩니다.", "Some of the elements of this skin only works in 96 DPI (100% size)."), 48
        End If
    End If
    If optUserFore.Value Then
        CheckBoxW1.VisualStyles = False
        FrameW5.VisualStyles = False
        CheckBoxW1.ForeColor = pgFore.BackColor
        FrameW5.ForeColor = pgFore.BackColor
    End If
    cmdAdvancedSkin.Enabled = (cbSkin.ListIndex = 2)
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

Private Sub chkBackColorMainOnly_Click()
    If Loaded Then
        cmdApply.Enabled = -1
        ColorChanged = True
    End If
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

Private Sub chkForeColorMainOnly_Click()
    If Loaded Then
        cmdApply.Enabled = -1
        ColorChanged = True
    End If
End Sub

Private Sub chkIgnore300_Click()
    If Loaded Then cmdApply.Enabled = -1
End Sub

Private Sub chkImageCentered_Click()
    If Loaded Then
        cmdApply.Enabled = -1
        ImageChanged = True
    End If
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

Private Sub cmdAdvancedSkin_Click()
    frmLiveBadukSkinProperties.Show vbModal, Me
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
    SaveSetting "DownloadBooster", "Options", "ScrollOneScreen", IIf(optScreenPerScroll.Value, 1, 0)
    SaveSetting "DownloadBooster", "Options", "BackColorMainOnly", chkBackColorMainOnly.Value
    SaveSetting "DownloadBooster", "Options", "ForeColorMainOnly", chkForeColorMainOnly.Value
    If ScrollChanged Then
        frmMain.ScrollOneScreen = optScreenPerScroll.Value
        frmMain.trThreadCount_Scroll
        frmMain.pbProgressContainer.Top = 0
        frmMain.vsProgressScroll.Value = 0
        frmMain.pbProgressContainer.Refresh
        frmMain.vsProgressScroll.LargeChange = IIf(optScreenPerScroll.Value, 1, 10)
    End If
    If trRequestInterval.Value < 8 Then
        SaveSetting "DownloadBooster", "Options", "ThreadRequestInterval", CInt(IntervalValues(trRequestInterval.Value) * 1000)
        trRequestInterval.Max = 7
    End If
    
    If PatternChanged Then
        SaveSetting "DownloadBooster", "Options", "FormFillStyle", lvPatterns.ListIndex
        SaveSetting "DownloadBooster", "Options", "FormFillColor", pgPatternColor.BackColor
        frmMain.SetPattern
        frmMain.SetBackgroundPosition
    End If
    
    Dim NoDisable As Boolean
    NoDisable = False
    
    If Not IsNumeric(txtMaxThreadCount.Text) Then
maxtrdnotint:
        MsgBox t("최대 쓰레드 개수는 정수여야 합니다.", "Maximum number of threads should be an integer."), 16
        NoDisable = True
        GoTo aftermaxtrdcheck
    ElseIf CStr(CInt(txtMaxThreadCount.Text)) <> txtMaxThreadCount.Text Then
        GoTo maxtrdnotint
    ElseIf CInt(txtMaxThreadCount.Text) > MAX_THREAD_COUNT_CONTROL Or CInt(txtMaxThreadCount.Text) < 2 Then
        MsgBox t("최대 쓰레드 개수는 2개 이상 " & MAX_THREAD_COUNT_CONTROL & "개 이하여야 합니다.", "Maximum number of threads should range in 2-" & MAX_THREAD_COUNT_CONTROL & "."), 16
        NoDisable = True
    Else
        SaveSetting "DownloadBooster", "Options", "MaxThreadCount", CInt(txtMaxThreadCount.Text)
'        If CInt(txtMaxThreadCount.Text) > 50 Then
'            MsgBox t("최대 쓰레드 개수가 너무 클 경우 실제 사용 개수와 관계없이 실행 속도가 느려질 수 있습니다.", "If the maximum number of threads is too high, the application might run slower."), 48
'        End If
    End If
aftermaxtrdcheck:
    
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
    ElseIf optUserColor.Value Then
        SaveSetting "DownloadBooster", "Options", "BackColor", CLng(pgColor.BackColor)
    End If
    If optSystemFore.Value Then
        SaveSetting "DownloadBooster", "Options", "ForeColor", "-1"
    ElseIf optUserFore.Value Then
        SaveSetting "DownloadBooster", "Options", "ForeColor", CLng(pgFore.BackColor)
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
        frmMain.SetTextColors
    End If
    If VisualStyleChanged Then
        On Error Resume Next
        DrawTabBackground True
        cmdChooseBackground.Refresh
        cmdSample.Refresh
        On Error GoTo 0
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
    If ImageChanged Then
        SaveSetting "DownloadBooster", "Options", "BackgroundImageCentered", chkImageCentered.Value
        SaveSetting "DownloadBooster", "Options", "UseBackgroundImage", IIf(lvBackgrounds.ListIndex <> 0, 1, 0)
        SaveSetting "DownloadBooster", "Options", "BackgroundImagePath", ChangedBackgroundPath
        frmMain.SetBackgroundImage
        frmMain.SetBackgroundPosition True
    End If
    
    If Trim$(txtNodePath.Text) <> "" Then
        If FileExists(Trim$(txtNodePath.Text)) Then
            SaveSetting "DownloadBooster", "Options", "NodePath", Trim$(txtNodePath.Text)
        Else
            Alert t("Node.js 경로가 존재하지 않습니다.", "Node.js path does not exist."), App.Title, 16
            NoDisable = True
        End If
    Else
        SaveSetting "DownloadBooster", "Options", "NodePath", ""
    End If
    If Trim$(txtScriptPath.Text) <> "" Then
        If FileExists(Trim$(txtScriptPath.Text)) Then
            SaveSetting "DownloadBooster", "Options", "ScriptPath", Trim$(txtScriptPath.Text)
        Else
            Alert t("다운로드 스크립트 경로가 존재하지 않습니다.", "Download script path does not exist."), App.Title, 16
            NoDisable = True
        End If
    Else
        SaveSetting "DownloadBooster", "Options", "ScriptPath", ""
    End If
    If Trim$(txtYtdlPath.Text) <> "" Then
        If FileExists(Trim$(txtYtdlPath.Text)) Then
            SaveSetting "DownloadBooster", "Options", "YtdlPath", Trim$(txtYtdlPath.Text)
        Else
            Alert t("Youtube-dl 경로가 존재하지 않습니다.", "Youtube-dl path does not exist."), App.Title, 16
            NoDisable = True
        End If
    Else
        SaveSetting "DownloadBooster", "Options", "YtdlPath", ""
    End If
    
    If FontChanged Then
        cbFont.Text = Trim$(cbFont.Text)
        If cbFont.Text <> "" And cbFont.Text <> ("(" & t("기본값", "default") & ")") And (Not FontExists(cbFont.Text)) Then
            MsgBox t("지정한 글꼴이 존재하지 않습니다.", "The specified font does not exist."), vbCritical
            NoDisable = True
        Else
            If cbFont.Text = ("(" & t("기본값", "default") & ")") Then
                SaveSetting "DownloadBooster", "Options", "Font", ""
            Else
                SaveSetting "DownloadBooster", "Options", "Font", cbFont.Text
            End If
            SetFont Me, True
            SetFont frmMain, True
        End If
    End If
    
    If lvBackgrounds.ListIndex <> 0 And GetSetting("DownloadBooster", "Options", "BackgroundImagePath", "") = "" Then
        MsgBox t("배경 그림이 선택되지 않았습니다.", "Background image is not selected."), 48
        SaveSetting "DownloadBooster", "Options", "UseBackgroundImage", "0"
        NoDisable = True
    End If
    
    Dim hSysMenu As Long
    Dim MII As MENUITEMINFO
    hSysMenu = GetSystemMenu(frmMain.hWnd, 0)
    MainFormOnTop = (chkAlwaysOnTop.Value = 1)
    SetWindowPos frmMain.hWnd, IIf(MainFormOnTop, hWnd_TOPMOST, hWnd_NOTOPMOST), 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    SetWindowPos Me.hWnd, IIf(MainFormOnTop, hWnd_TOPMOST, hWnd_NOTOPMOST), 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
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
    
    If optUserFore.Value Then
        CheckBoxW1.VisualStyles = False
        FrameW5.VisualStyles = False
        CheckBoxW1.ForeColor = pgFore.BackColor
        FrameW5.ForeColor = pgFore.BackColor
    End If
    
    RedrawPreview
    ColorChanged = False
    ImageChanged = False
    VisualStyleChanged = False
    SkinChanged = False
    ScrollChanged = False
    FontChanged = False
    PatternChanged = False
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

Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse)
    '
End Property

Private Property Get ISubclass_MsgResponse() As EMsgResponse
    ISubclass_MsgResponse = emrConsume
End Property

Private Function ISubclass_WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error Resume Next
 
    Select Case uMsg
        Case WM_SETTINGCHANGE
            Select Case GetStrFromPtr(lParam)
                Case "WindowMetrics"
                    UpdateBorderWidth
                    SetPreviewPosition
                    DrawTabBackground True
            End Select
        Case WM_THEMECHANGED
            DrawTabBackground True
    End Select
    
    ISubclass_WindowProc = CallOldWindowProc(hWnd, uMsg, wParam, lParam)
End Function

Private Sub lblFillColorSelect_Click()
    Dim Color As OLE_COLOR
    Color = ShowColorDialog(Me.hWnd, True, pgPatternColor.BackColor, True)
    If Color = -1 Then Exit Sub
    pgPatternColor.BackColor = Color
    cmdApply.Enabled = -1
    pgPatternPreview.FillColor = pgPatternColor.BackColor
    PatternChanged = True
End Sub

Private Sub lblSystemBtnFace_Click()
    Shell "rundll32.exe shell32.dll,Control_RunDLL desk.cpl,,2", vbNormalFocus
End Sub

Private Sub lblSystemWindowText_Click()
    Shell "rundll32.exe shell32.dll,Control_RunDLL desk.cpl,,2", vbNormalFocus
End Sub

Private Sub lvBackgrounds_Click()
    On Error GoTo nopicture
    Dim BackgroundPath$
    BackgroundPath = lvBackgroundFiles.Path & IIf(EndsWith(lvBackgroundFiles.Path, "\"), "", "\") & lvBackgrounds.List(lvBackgrounds.ListIndex)
    If lvBackgrounds.ListIndex = 0 Then
nopicture:
        Set imgBackgroundPreview.Picture = Nothing
    ElseIf LCase(Right(BackgroundPath, 4)) = ".png" Then
        Set imgBackgroundPreview.Picture = LoadPngIntoPictureWithAlpha(BackgroundPath)
    Else
        imgBackgroundPreview.Picture = LoadPicture(BackgroundPath)
    End If
    Set imgPreview.Picture = imgBackgroundPreview.Picture
    frmOptions.cmdSample.Refresh
    ChangedBackgroundPath = BackgroundPath
    If Loaded Then
        cmdApply.Enabled = -1
        ImageChanged = True
    End If
End Sub

Private Sub lvHeaders_AfterLabelEdit(Cancel As Boolean, NewString As String)
    NewString = Trim$(NewString)
    If NewString = "" Then
invalidname:
        Cancel = True
        Alert t("헤더 이름이 잘못되었습니다.", "Invalid header name."), App.Title, 16
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
            Alert t("해당 이름이 이미 존재합니다.", "Duplicate header name."), App.Title, 16
            Exit Sub
            Exit For
        End If
    Next i
    
    If Loaded Then cmdApply.Enabled = -1
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

Private Sub lvPatterns_Click()
    pgPatternPreview.FillStyle = lvPatterns.ListIndex + 1
    If Loaded Then
        cmdApply.Enabled = -1
        PatternChanged = True
    End If
End Sub

Private Sub optLinePerScroll_Click()
    If Loaded Then
        cmdApply.Enabled = -1
        ScrollChanged = True
    End If
End Sub

Private Sub optScreenPerScroll_Click()
    If Loaded Then
        cmdApply.Enabled = -1
        ScrollChanged = True
    End If
End Sub

Private Sub trRequestInterval_Change()
    trRequestInterval_Scroll
End Sub

Private Sub trRequestInterval_Scroll()
    If trRequestInterval.Value = 8 Then
        lblIntervalDisplay.Caption = "(" & t("사용자 지정", "Customized") & ")"
    Else
        lblIntervalDisplay.Caption = "(" & IntervalValues(trRequestInterval.Value) & t("초", " second" & IIf(IntervalValues(trRequestInterval.Value) = 1, "", "s")) & ")"
    End If
    If Loaded Then cmdApply.Enabled = -1
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
    ScrollChanged = False
    FontChanged = False
    If GetSetting("DownloadBooster", "Options", "DisableDWMWindow", DefaultDisableDWMWindow) = 1 Then DisableDWMWindow Me.hWnd
    SetFormBackgroundColor Me
    SetFont Me
    SetWindowPos Me.hWnd, IIf(MainFormOnTop, hWnd_TOPMOST, hWnd_NOTOPMOST), 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    
    On Error Resume Next
    Me.Icon = frmMain.imgWrench.ListImages(1).Picture
    On Error GoTo 0
    
    BackgroundDrawn(1) = False
    BackgroundDrawn(2) = False
    BackgroundDrawn(3) = False
    BackgroundDrawn(4) = False
    BackgroundDrawn(5) = False
    BackgroundDrawn(6) = False
    
    IntervalValues(0) = 0.01
    IntervalValues(1) = 0.05
    IntervalValues(2) = 0.1
    IntervalValues(3) = 0.3
    IntervalValues(4) = 0.5
    IntervalValues(5) = 1#
    IntervalValues(6) = 3#
    IntervalValues(7) = 5#
    
    udMaxThreadCount.Max = MAX_THREAD_COUNT_CONTROL
    
    lvHeaders.SmallIcons = imgFiles
    
    lblSelectColor.Top = pgColor.Top
    lblSelectColor.Left = pgColor.Left
    lblSelectColor.Width = pgColor.Width
    lblSelectColor.Height = pgColor.Height
    
    lblSelectFore.Top = pgFore.Top
    lblSelectFore.Left = pgFore.Left
    lblSelectFore.Width = pgFore.Width
    lblSelectFore.Height = pgFore.Height
    
    RemoveVisualStyles txtSampleClassic.hWnd
    
    Dim i%
    Dim MaxWidth%, MaxHeight%
    MaxWidth = 15
    MaxHeight = 15
    For i = 1 To pbPanel.Count
        pbPanel(i).Visible = 0
        pbPanel(i).Enabled = 0
        pbPanel(i).Top = 465
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
    tsTabStrip.Height = MaxHeight + 410
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
    txtMaxThreadCount.Text = GetSetting("DownloadBooster", "Options", "MaxThreadCount", 25)
    udMaxThreadCount.SyncFromBuddy
    optLinePerScroll.Value = True
    optScreenPerScroll.Value = (GetSetting("DownloadBooster", "Options", "ScrollOneScreen", 0) <> 0)
    chkBackColorMainOnly.Value = GetSetting("DownloadBooster", "Options", "BackColorMainOnly", 0)
    chkForeColorMainOnly.Value = GetSetting("DownloadBooster", "Options", "ForeColorMainOnly", 0)
    Select Case CInt(GetSetting("DownloadBooster", "Options", "ThreadRequestInterval", 100))
        Case 10
            trRequestInterval.Value = 0
        Case 50
            trRequestInterval.Value = 1
        Case 100
            trRequestInterval.Value = 2
        Case 300
            trRequestInterval.Value = 3
        Case 500
            trRequestInterval.Value = 4
        Case 1000
            trRequestInterval.Value = 5
        Case 3000
            trRequestInterval.Value = 6
        Case 5000
            trRequestInterval.Value = 7
        Case Else
            trRequestInterval.Max = 8
            trRequestInterval.Value = 8
    End Select
    trRequestInterval_Scroll
    
    pbBackground.Enabled = False
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
    pgPatternPreview.BackColor = pgColor.BackColor
    
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
    cbSkin.AddItem t("라이브 바둑 쪽지", "LiveBaduk memo")
    If CInt(GetSetting("DownloadBooster", "Options", "EnableLiveBadukMemoSkin", 0)) Then
        cbSkin.ListIndex = 2
    ElseIf Abs(CInt(GetSetting("DownloadBooster", "Options", "DisableVisualStyle", 0))) Then
        cbSkin.ListIndex = 1
    Else
        cbSkin.ListIndex = 0
    End If
    
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
    cbWhenExist.AddItem t("건너뛰기", "Skip")
    cbWhenExist.AddItem t("덮어쓰기", "Overwrite")
    cbWhenExist.AddItem t("자동 이름 변경", "Auto Rename")
    cbWhenExist.ListIndex = GetSetting("DownloadBooster", "Options", "WhenFileExists", 0)
    
    pbBackgroundPreview.Enabled = False
    
    lvPatterns.Clear
    lvPatterns.AddItem t("(없음)", "(None)")
    lvPatterns.AddItem t("수평선", "Horizontal lines")
    lvPatterns.AddItem t("수직선", "Vertical lines")
    lvPatterns.AddItem t("하향 대각선", "NW-SE lines")
    lvPatterns.AddItem t("상향 대각선", "NE-SW lines")
    lvPatterns.AddItem t("교차", "Grid")
    lvPatterns.AddItem t("대각선 교차", "45 degrees grid")
    lvPatterns.ListIndex = CInt(GetSetting("DownloadBooster", "Options", "FormFillStyle", 0))
    
    ChangedBackgroundPath = GetSetting("DownloadBooster", "Options", "BackgroundImagePath", "")
    LoadBackgroundList True
    
    pgPatternColor.BackColor = CLng(GetSetting("DownloadBooster", "Options", "FormFillColor", 0))
    pgPatternPreview.FillColor = pgPatternColor.BackColor
    pgPatternPreview.FillStyle = lvPatterns.ListIndex + 1
    pgPatternPreview.Width = pbBackgroundPreview.Width - 240
    pgPatternPreview.Height = pbBackgroundPreview.Height - 240
    imgBackgroundPreview.Width = pbBackgroundPreview.Width - 240
    imgBackgroundPreview.Height = pbBackgroundPreview.Height - 240
    
    cbImagePosition.Clear
    cbImagePosition.AddItem t("늘이기", "Stretch")
    cbImagePosition.AddItem t("높이에 맞춤", "Fit to height")
    cbImagePosition.AddItem t("너비에 맞춤", "Fit to width")
    cbImagePosition.AddItem t("원본 크기 유지", "True size")
    cbImagePosition.AddItem t("바둑판식", "Tile")
    cbImagePosition.ListIndex = GetSetting("DownloadBooster", "Options", "ImagePosition", 1)
    cbImagePosition_Click
    chkImageCentered.Value = GetSetting("DownloadBooster", "Options", "BackgroundImageCentered", 0)
    
    '임시
    FrameW1.Enabled = 0
    Label20.Enabled = 0
    cmdSaveTheme.Enabled = 0
    cmdDeleteTheme.Enabled = 0
    cbTheme.Enabled = 0
    cbTheme.AddItem t("수정된 테마", "Modified theme")
    cbTheme.ListIndex = 0
    
    txtNodePath.Text = GetSetting("DownloadBooster", "Options", "NodePath", "")
    txtScriptPath.Text = GetSetting("DownloadBooster", "Options", "ScriptPath", "")
    txtYtdlPath.Text = GetSetting("DownloadBooster", "Options", "YtdlPath", "")
    
    cbFont.Text = Trim$(GetSetting("DownloadBooster", "Options", "Font", ""))
    If cbFont.Text = "" Then cbFont.Text = ("(" & t("기본값", "default") & ")")
    cbFont.AddItem "(" & t("기본값", "default") & ")"
    If t(1, 2) = 2 Then
        If FontExists("Tahoma") Then cbFont.AddItem "Tahoma"
        If FontExists("Segoe UI") Then cbFont.AddItem "Segoe UI"
    Else
        If FontExists("굴림") Then cbFont.AddItem "굴림"
        If FontExists("돋움") Then cbFont.AddItem "돋움"
        If FontExists("바탕") Then cbFont.AddItem "바탕"
        If FontExists("궁서") Then cbFont.AddItem "궁서"
        If FontExists("맑은 고딕") Then cbFont.AddItem "맑은 고딕"
    End If
    
    On Error Resume Next
    
    Dim clrForeColor As Long
    clrForeColor = GetSetting("DownloadBooster", "Options", "ForeColor", -1)
    If clrForeColor < 0 Or clrForeColor > 16777215 Then
        optSystemFore.Value = True
        pgFore.BackColor = &H80000012
    Else
        optUserFore.Value = True
        pgFore.BackColor = clrForeColor
        CheckBoxW1.VisualStyles = False
        FrameW5.VisualStyles = False
        CheckBoxW1.ForeColor = pgFore.BackColor
        FrameW5.ForeColor = pgFore.BackColor
    End If
    Label11.ForeColor = pgFore.BackColor
    
    tsTabStrip.Tabs(1).Caption = t(tsTabStrip.Tabs(1).Caption, " General ")
    tsTabStrip.Tabs(2).Caption = t(tsTabStrip.Tabs(2).Caption, " Network ")
    tsTabStrip.Tabs(3).Caption = t(tsTabStrip.Tabs(3).Caption, " Background ")
    tsTabStrip.Tabs(4).Caption = t(tsTabStrip.Tabs(4).Caption, " Appearance ")
    tsTabStrip.Tabs(5).Caption = t(tsTabStrip.Tabs(5).Caption, " Sound ")
    tsTabStrip.Tabs(6).Caption = t(tsTabStrip.Tabs(6).Caption, " Advanced ")
    Frame1.Caption = t(Frame1.Caption, " Background color ")
    Frame4.Caption = t(Frame4.Caption, " Text color ")
    Label10.Caption = t(Label10.Caption, "&Window:")
    Frame2.Caption = t(Frame2.Caption, " Download options ")
    Frame5.Caption = t(Frame5.Caption, " Interface ")
    chkNoCleanup.Caption = t(chkNoCleanup.Caption, "Preserve segme&nts")
    chkRememberURL.Caption = t(chkRememberURL.Caption, "Re&member URL")
    optSystemColor.Caption = t(optSystemColor.Caption, "&System color")
    optUserColor.Caption = t(optUserColor.Caption, "&Custom color")
    optSystemFore.Caption = t(optSystemFore.Caption, "S&ystem color")
    optUserFore.Caption = t(optUserFore.Caption, "C&ustom color")
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
    chkAlwaysOnTop.Caption = t(chkAlwaysOnTop.Caption, "Al&ways on top")
    chkAeroWindow.Caption = t(chkAeroWindow.Caption, "Use Aero &glass window")
    cmdAddHeader.Caption = t(cmdAddHeader.Caption, "&Add")
    cmdDeleteHeader.Caption = t(cmdDeleteHeader.Caption, "&Delete")
    cmdEditHeaderName.Caption = t(cmdEditHeaderName.Caption, "&Rename")
    cmdEditHeaderValue.Caption = t(cmdEditHeaderValue.Caption, "&Edit")
    chkLazyElapsed.Caption = t(chkLazyElapsed.Caption, "Elapsed time sin&ce first data receive")
    chkExcludeMergeFromElapsed.Caption = t(chkExcludeMergeFromElapsed.Caption, "Exclude merging time from elapsed time")
    FrameW3.Caption = t(FrameW3.Caption, " Connection settings ")
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
    tr Label13, "&Font:"
    tr Label14, "Ma&x. number of threads:"
    tr Label15, "(restart required)"
    tr Label16, Label15.Caption
    tr FrameW6, " Sound settings "
    tr Label17, "Set the headers when requesting to the server on download. Headers set in Download Options have higher priority."
    tr Label18, "T&hread scroll:"
    tr optLinePerScroll, "Per li&ne"
    tr optScreenPerScroll, "Pe&r screen"
    tr Label19, "Thread request i&nterval:"
    pbBackground.Caption = t("다운로드 부스터", "Download Booster")
    'tr cmdChooseBackground, "B&rowse..."
    tr cmdAdvancedSkin, "Ad&vanced..."
    tr chkBackColorMainOnly, "&Only apply to main window"
    tr chkForeColorMainOnly, "O&nly apply to main window"
    tr FrameW7, " &Patterns "
    tr Label21, "C&olor:"
    tr FrameW8, " &Wallpaper "
    tr cmdChooseBackground, "&Browse..."
    tr chkImageCentered, "&Centered"
    tr FrameW1, " Theme "
    tr Label20, "&Theme:"
    tr cmdSaveTheme, "S&ave..."
    tr cmdDeleteTheme, "&Delete"
    
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
    
    AttachMessage Me, Me.hWnd, WM_SETTINGCHANGE
    AttachMessage Me, Me.hWnd, WM_THEMECHANGED
    
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
    Label14.Top = Label14.Top - chkAutoYtdl.Height
    txtMaxThreadCount.Top = txtMaxThreadCount.Top - chkAutoYtdl.Height
    udMaxThreadCount.Top = udMaxThreadCount.Top - chkAutoYtdl.Height
    Label15.Top = Label15.Top - chkAutoYtdl.Height
    txtOuterMaxThreadCount.Top = txtOuterMaxThreadCount.Top - chkAutoYtdl.Height
#End If
    
    Loaded = True
End Sub

Sub LoadBackgroundList(Optional ByVal OnLoad As Boolean = False)
    Dim BackgroundImagePath$
    BackgroundImagePath = ChangedBackgroundPath
    lvBackgrounds.Clear
    lvBackgrounds.AddItem t("(없음)", "(None)")
    lvBackgrounds.ListIndex = 0
    Dim BackgroundImageEnabled As Boolean
    BackgroundImageEnabled = (GetSetting("DownloadBooster", "Options", "UseBackgroundImage", 0) <> 0)
    If FileExists(BackgroundImagePath) Then
        Dim li&
        lvBackgroundFiles.Path = GetParentFolderName(BackgroundImagePath)
        For li = 0 To lvBackgroundFiles.ListCount - 1
            lvBackgrounds.AddItem lvBackgroundFiles.List(li)
            If (Not OnLoad) Or BackgroundImageEnabled Then
                If LCase(GetFilename(BackgroundImagePath)) = LCase(lvBackgroundFiles.List(li)) Then
                    lvBackgrounds.ListIndex = li + 1
                End If
            End If
        Next li
    End If
    lvBackgrounds_Click
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 9 And IsKeyPressed(gksKeyboardctrl) Then
        NextTabPage IsKeyPressed(gksKeyboardShift)
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DetachMessage Me, Me.hWnd, WM_SETTINGCHANGE
    DetachMessage Me, Me.hWnd, WM_THEMECHANGED
End Sub

Sub DrawTabBackground(Optional Force As Boolean = False)
    On Error Resume Next
    Dim ctrl As Control
    Dim CtrlTypeName As String
    
    If Force Then
        BackgroundDrawn(1) = False
        BackgroundDrawn(2) = False
        BackgroundDrawn(3) = False
        BackgroundDrawn(4) = False
        BackgroundDrawn(5) = False
        BackgroundDrawn(6) = False
    ElseIf BackgroundDrawn(tsTabStrip.SelectedItem.Index) Then
        Exit Sub
    End If
    
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "PictureBox" And ctrl.Tag <> "nobgdraw" Then
            If ctrl.Name = "pbPanel" Or (ctrl.Container.Name = "pbPanel" And ctrl.Container.Index = tsTabStrip.SelectedItem.Index) Then
                ctrl.AutoRedraw = True
                tsTabStrip.DrawBackground ctrl.hWnd, ctrl.hDC
            End If
        End If
    Next ctrl
    For Each ctrl In Me.Controls
        If ctrl.Container.Name = "pbPanel" And ctrl.Container.Index = tsTabStrip.SelectedItem.Index Then
            CtrlTypeName = TypeName(ctrl)
            If CtrlTypeName = "FrameW" Or CtrlTypeName = "CheckBoxW" Or CtrlTypeName = "OptionButtonW" Or CtrlTypeName = "CommandButtonW" Or CtrlTypeName = "LinkLabel" Then
                ctrl.Transparent = True
                ctrl.Refresh
            ElseIf CtrlTypeName = "Label" And ctrl.Tag <> "nobgdraw" Then
                ctrl.BackStyle = 0
            ElseIf CtrlTypeName = "PictureBox" Then
                ctrl.Refresh
            End If
        End If
    Next ctrl
    
    BackgroundDrawn(tsTabStrip.SelectedItem.Index) = True
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
    pgPatternPreview.BackColor = pgColor.BackColor
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
    CheckBoxW1.VisualStyles = False
    FrameW5.VisualStyles = False
    CheckBoxW1.ForeColor = pgFore.BackColor
    FrameW5.ForeColor = pgFore.BackColor
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
    If cmdApply.Enabled Then cmdApply_Click
    Unload Me
End Sub

Private Sub optSystemColor_Click()
    If Loaded Then
        cmdApply.Enabled = -1
        ColorChanged = True
    End If
    pbBackground.BackColor = &H8000000F
    pgPatternPreview.BackColor = pbBackground.BackColor
    cmdSample.Refresh
    RedrawPreview
End Sub

Private Sub optSystemFore_Click()
    If Loaded Then
        cmdApply.Enabled = -1
        ColorChanged = True
    End If
    Label11.ForeColor = &H80000012
    CheckBoxW1.VisualStyles = (cbSkin.ListIndex <> 1)
    FrameW5.VisualStyles = (cbSkin.ListIndex <> 1)
    CheckBoxW1.ForeColor = &H80000012
    FrameW5.ForeColor = &H80000012
End Sub

Private Sub optUserColor_Click()
    If Loaded Then
        cmdApply.Enabled = -1
        ColorChanged = True
    End If
    pbBackground.BackColor = pgColor.BackColor
    pgPatternPreview.BackColor = pbBackground.BackColor
    cmdSample.Refresh
    RedrawPreview
End Sub

Private Sub optUserFore_Click()
    If Loaded Then
        cmdApply.Enabled = -1
        ColorChanged = True
    End If
    Label11.ForeColor = pgFore.BackColor
    CheckBoxW1.VisualStyles = False
    FrameW5.VisualStyles = False
    CheckBoxW1.ForeColor = pgFore.BackColor
    FrameW5.ForeColor = pgFore.BackColor
End Sub

Private Sub tsTabStrip_TabClick(ByVal TabItem As TbsTab)
    On Error Resume Next
    DrawTabBackground
    
    Dim i%
    For i = 1 To pbPanel.Count
        If i = TabItem.Index Then
            pbPanel(i).Visible = -1
            pbPanel(i).Enabled = -1
        Else
            pbPanel(i).Visible = 0
            pbPanel(i).Enabled = 0
        End If
    Next i
    
    If TabItem.Index = 4 Then
        DoEvents
        RedrawPreview
    End If
End Sub

Sub SetPreviewPosition()
    Dim Left%, Top%, Width%, Height%
    Left = 30
    Top = 6
    Width = 3915
    Height = 1380
    pbBackground.Top = pbPreview.Top + Top * 15 + 15 + 30
    pbBackground.Left = pbPreview.Left + Left * 15
    imgPreview.Width = Width
    imgPreview.Height = Height
    pbBackground.Width = Width + PaddedBorderWidth * 15 + DialogBorderWidth * 30
    pbBackground.Height = Height + PaddedBorderWidth * 15 + DialogBorderWidth * 30 + CaptionHeight * 15
    RedrawPreview
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

Private Sub txtMaxThreadCount_Change()
    If IsNumeric(txtMaxThreadCount.Text) Then udMaxThreadCount.Value = CInt(txtMaxThreadCount.Text)
    If Loaded Then cmdApply.Enabled = -1
End Sub

Private Sub txtMaxThreadCount_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not IsNumeric(txtMaxThreadCount.Text) Then Exit Sub
    If KeyCode = 38 And txtMaxThreadCount.Text < MAX_THREAD_COUNT_CONTROL Then
        txtMaxThreadCount.Text = CDbl(txtMaxThreadCount.Text) + 1
    ElseIf KeyCode = 40 And txtMaxThreadCount.Text > 2 Then
        txtMaxThreadCount.Text = CDbl(txtMaxThreadCount.Text) - 1
    End If
End Sub

Private Sub txtMaxThreadCount_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 8 And (KeyAscii < 48 Or KeyAscii > 57) Then KeyAscii = 0
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
