VERSION 5.00
Begin VB.Form frmOptions 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "옵션"
   ClientHeight    =   10650
   ClientLeft      =   2760
   ClientTop       =   3855
   ClientWidth     =   14805
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
   ScaleHeight     =   10650
   ScaleWidth      =   14805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.PictureBox pbPanel 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  '없음
      Enabled         =   0   'False
      Height          =   4425
      Index           =   4
      Left            =   7200
      ScaleHeight     =   4425
      ScaleWidth      =   6975
      TabIndex        =   4
      Top             =   5040
      Visible         =   0   'False
      Width           =   6975
      Begin prjDownloadBooster.FrameW FrameW6 
         Height          =   975
         Left            =   120
         TabIndex        =   72
         Top             =   120
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   1720
         Caption         =   "다운로드 알림"
         Transparent     =   -1  'True
         Begin prjDownloadBooster.FrameW fCompleteSound 
            Height          =   705
            Left            =   720
            TabIndex        =   73
            Top             =   240
            Width           =   5535
            _ExtentX        =   9763
            _ExtentY        =   1244
            BorderStyle     =   0
            Caption         =   "                             "
            Transparent     =   -1  'True
            Begin prjDownloadBooster.CommandButtonW cmdBrowseCompleteSound 
               Height          =   300
               Left            =   4560
               TabIndex        =   76
               Top             =   330
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   529
               ImageList       =   "imgBrowse"
               ImageListAlignment=   4
               Transparent     =   -1  'True
            End
            Begin VB.TextBox txtCompleteSoundPath 
               Height          =   300
               Left            =   360
               TabIndex        =   75
               Top             =   330
               Width           =   4095
            End
            Begin prjDownloadBooster.CheckBoxW chkBeepWhenComplete 
               Height          =   255
               Left            =   120
               TabIndex        =   74
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
               TabIndex        =   77
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
               Transparent     =   -1  'True
            End
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   120
            Picture         =   "frmOptions.frx":000C
            Top             =   120
            Width           =   480
         End
      End
      Begin prjDownloadBooster.FrameW FrameW9 
         Height          =   3135
         Left            =   120
         TabIndex        =   117
         Top             =   1200
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   5530
         Caption         =   "메시지 상자"
         Transparent     =   -1  'True
         Begin prjDownloadBooster.FrameW fAsterisk 
            Height          =   735
            Left            =   720
            TabIndex        =   78
            Top             =   240
            Width           =   5535
            _ExtentX        =   9763
            _ExtentY        =   1296
            BorderStyle     =   0
            Caption         =   "                             "
            Transparent     =   -1  'True
            Begin prjDownloadBooster.CommandButtonW cmdBrowseAsterisk 
               Height          =   300
               Left            =   4560
               TabIndex        =   81
               Top             =   330
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   529
               ImageList       =   "imgBrowse"
               ImageListAlignment=   4
               Transparent     =   -1  'True
            End
            Begin VB.TextBox txtAsterisk 
               Height          =   300
               Left            =   360
               TabIndex        =   80
               Top             =   330
               Width           =   4095
            End
            Begin prjDownloadBooster.CheckBoxW chkAsterisk 
               Height          =   255
               Left            =   120
               TabIndex        =   79
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
               TabIndex        =   82
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
               Transparent     =   -1  'True
            End
         End
         Begin prjDownloadBooster.FrameW fExclamation 
            Height          =   735
            Left            =   720
            TabIndex        =   83
            Top             =   960
            Width           =   5535
            _ExtentX        =   9763
            _ExtentY        =   1296
            BorderStyle     =   0
            Caption         =   "                             "
            Transparent     =   -1  'True
            Begin prjDownloadBooster.CommandButtonW cmdBrowseExclamation 
               Height          =   300
               Left            =   4560
               TabIndex        =   86
               Top             =   330
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   529
               ImageList       =   "imgBrowse"
               ImageListAlignment=   4
               Transparent     =   -1  'True
            End
            Begin VB.TextBox txtExclamation 
               Height          =   300
               Left            =   360
               TabIndex        =   85
               Top             =   330
               Width           =   4095
            End
            Begin prjDownloadBooster.CheckBoxW chkExclamation 
               Height          =   255
               Left            =   120
               TabIndex        =   84
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
               TabIndex        =   87
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
               Transparent     =   -1  'True
            End
         End
         Begin prjDownloadBooster.FrameW fError 
            Height          =   735
            Left            =   720
            TabIndex        =   88
            Top             =   1680
            Width           =   5535
            _ExtentX        =   9763
            _ExtentY        =   1296
            BorderStyle     =   0
            Caption         =   "                             "
            Transparent     =   -1  'True
            Begin prjDownloadBooster.CommandButtonW cmdBrowseError 
               Height          =   300
               Left            =   4560
               TabIndex        =   91
               Top             =   330
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   529
               ImageList       =   "imgBrowse"
               ImageListAlignment=   4
               Transparent     =   -1  'True
            End
            Begin VB.TextBox txtError 
               Height          =   300
               Left            =   360
               TabIndex        =   90
               Top             =   330
               Width           =   4095
            End
            Begin prjDownloadBooster.CheckBoxW chkError 
               Height          =   255
               Left            =   120
               TabIndex        =   89
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
               TabIndex        =   92
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
               Transparent     =   -1  'True
            End
         End
         Begin prjDownloadBooster.FrameW fQuestion 
            Height          =   675
            Left            =   720
            TabIndex        =   93
            Top             =   2400
            Width           =   5535
            _ExtentX        =   9763
            _ExtentY        =   1191
            BorderStyle     =   0
            Caption         =   "                             "
            Transparent     =   -1  'True
            Begin prjDownloadBooster.CommandButtonW cmdBrowseQuestion 
               Height          =   300
               Left            =   4560
               TabIndex        =   96
               Top             =   330
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   529
               ImageList       =   "imgBrowse"
               ImageListAlignment=   4
               Transparent     =   -1  'True
            End
            Begin VB.TextBox txtQuestion 
               Height          =   300
               Left            =   360
               TabIndex        =   95
               Top             =   330
               Width           =   4095
            End
            Begin prjDownloadBooster.CheckBoxW chkQuestion 
               Height          =   255
               Left            =   120
               TabIndex        =   94
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
               TabIndex        =   97
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
               Transparent     =   -1  'True
            End
         End
         Begin VB.Image Image8 
            Height          =   480
            Left            =   120
            Picture         =   "frmOptions.frx":0456
            Top             =   240
            Width           =   480
         End
      End
   End
   Begin prjDownloadBooster.ImageList imgBrowse 
      Left            =   11760
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      ImageWidth      =   16
      ImageHeight     =   16
      ColorDepth      =   8
      MaskColor       =   16711935
      InitListImages  =   "frmOptions.frx":0898
   End
   Begin VB.PictureBox pbPanel 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  '없음
      Enabled         =   0   'False
      Height          =   2385
      Index           =   2
      Left            =   7200
      ScaleHeight     =   2385
      ScaleWidth      =   6975
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   600
      Visible         =   0   'False
      Width           =   6975
      Begin prjDownloadBooster.FrameW FrameW3 
         Height          =   1215
         Left            =   120
         TabIndex        =   31
         Top             =   120
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   2143
         Caption         =   "서버 연결"
         Transparent     =   -1  'True
         Begin prjDownloadBooster.Slider trRequestInterval 
            Height          =   450
            Left            =   3000
            TabIndex        =   36
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
            TabIndex        =   33
            Top             =   240
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   450
            Caption         =   "300번대 응답 코드 무시(&I)"
            Transparent     =   -1  'True
         End
         Begin prjDownloadBooster.CheckBoxW chkForceGet 
            Height          =   255
            Left            =   840
            TabIndex        =   34
            Top             =   480
            Width           =   2865
            _ExtentX        =   5054
            _ExtentY        =   450
            Caption         =   "파일 검사 시 GET 요청(&Q)"
            Transparent     =   -1  'True
         End
         Begin prjDownloadBooster.CheckBoxW chkNoRedirectCheck 
            Height          =   255
            Left            =   840
            TabIndex        =   32
            Top             =   240
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   450
            Caption         =   "리다이렉트 검사 안 함(&R)"
            Transparent     =   -1  'True
         End
         Begin VB.Label lblIntervalDisplay 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "(0.1초)"
            Height          =   180
            Left            =   5070
            TabIndex        =   37
            Top             =   840
            Width           =   570
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "스레드 요청 간격(&N):"
            Height          =   180
            Left            =   1080
            TabIndex        =   35
            Top             =   840
            Width           =   1725
         End
         Begin VB.Image Image4 
            Height          =   480
            Left            =   120
            Picture         =   "frmOptions.frx":0C80
            Top             =   240
            Width           =   480
         End
      End
      Begin prjDownloadBooster.FrameW fHeaders 
         Height          =   795
         Left            =   120
         TabIndex        =   38
         Top             =   1440
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   1402
         Caption         =   "사용자 지정 헤더"
         Transparent     =   -1  'True
         Begin prjDownloadBooster.CommandButtonW cmdHeaders 
            Height          =   330
            Left            =   5040
            TabIndex        =   120
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   582
            Caption         =   "헤더(&H)..."
         End
         Begin VB.Image Image5 
            Height          =   480
            Left            =   120
            Picture         =   "frmOptions.frx":10C2
            Top             =   240
            Width           =   480
         End
         Begin VB.Label Label17 
            BackStyle       =   0  '투명
            Caption         =   "다운로드 중 서버에 요청할 때 전송할 헤더를 설정합니다."
            Height          =   495
            Left            =   840
            TabIndex        =   115
            Top             =   240
            Width           =   4095
         End
      End
   End
   Begin VB.PictureBox pbPanel 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  '없음
      Height          =   4185
      Index           =   1
      Left            =   120
      ScaleHeight     =   4185
      ScaleWidth      =   6975
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   600
      Width           =   6975
      Begin prjDownloadBooster.FrameW Frame5 
         Height          =   1905
         Left            =   120
         TabIndex        =   19
         Top             =   2160
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   3360
         Caption         =   "인터페이스"
         Transparent     =   -1  'True
         Begin VB.OptionButton optScreenPerScroll 
            Caption         =   "한 화면씩(&R)"
            Height          =   255
            Left            =   4200
            TabIndex        =   30
            Top             =   1575
            Width           =   1500
         End
         Begin VB.OptionButton optLinePerScroll 
            Caption         =   "한 줄씩(&N)"
            Height          =   255
            Left            =   2640
            TabIndex        =   29
            Top             =   1575
            Width           =   1395
         End
         Begin prjDownloadBooster.CheckBoxW chkAllowDuplicates 
            Height          =   255
            Left            =   840
            TabIndex        =   24
            Top             =   960
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   450
            Caption         =   "일괄 처리 목록에 중복 항목 허용(&I)"
            Transparent     =   -1  'True
         End
         Begin prjDownloadBooster.CheckBoxW chkForceOldDialog 
            Height          =   255
            Left            =   840
            TabIndex        =   23
            Top             =   720
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   450
            Caption         =   "윈도우 3.1 대화 상자 사용(&S)"
            Transparent     =   -1  'True
         End
         Begin prjDownloadBooster.CheckBoxW chkLazyElapsed 
            Height          =   255
            Left            =   840
            TabIndex        =   22
            Top             =   480
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   450
            Caption         =   "첫 바이트 수신 후 경과 시간 계산(&C)"
            Transparent     =   -1  'True
         End
         Begin prjDownloadBooster.CheckBoxW chkAeroWindow 
            Height          =   255
            Left            =   2880
            TabIndex        =   21
            Top             =   240
            Visible         =   0   'False
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   450
            Enabled         =   0   'False
            Caption         =   "유리 창 효과 사용(&G)"
            Transparent     =   -1  'True
         End
         Begin prjDownloadBooster.CheckBoxW chkAlwaysOnTop 
            Height          =   255
            Left            =   840
            TabIndex        =   20
            Top             =   240
            Width           =   1920
            _ExtentX        =   3387
            _ExtentY        =   450
            Caption         =   "항상 위에 표시(&W)"
            Transparent     =   -1  'True
         End
         Begin VB.ComboBox cbLanguage 
            Height          =   300
            Left            =   2640
            Style           =   2  '드롭다운 목록
            TabIndex        =   26
            Top             =   1230
            Width           =   1455
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "스레드 스크롤(&H):"
            Height          =   180
            Left            =   1080
            TabIndex        =   28
            Top             =   1620
            Width           =   1470
         End
         Begin VB.Label Label16 
            BackStyle       =   0  '투명
            Caption         =   "(다시 시작 필요)"
            Height          =   255
            Left            =   4200
            TabIndex        =   27
            Top             =   1290
            Width           =   1575
         End
         Begin VB.Image Image3 
            Height          =   405
            Left            =   120
            Picture         =   "frmOptions.frx":1504
            Top             =   240
            Width           =   435
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '투명
            Caption         =   "언어(&L):"
            Height          =   255
            Left            =   1080
            TabIndex        =   25
            Tag             =   "nocolorchange"
            Top             =   1275
            Width           =   975
         End
      End
      Begin prjDownloadBooster.FrameW Frame2 
         Height          =   1935
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   3413
         Caption         =   "다운로드 설정"
         Transparent     =   -1  'True
         Begin prjDownloadBooster.CheckBoxW chkUseServerModified 
            Height          =   255
            Left            =   3480
            TabIndex        =   12
            Top             =   720
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   450
            Caption         =   "서버의 수정한 날짜 반영(&D)"
            Transparent     =   -1  'True
         End
         Begin prjDownloadBooster.SpinBox txtMaxThreadCount 
            Height          =   270
            Left            =   3210
            TabIndex        =   17
            Top             =   1560
            Width           =   720
            _ExtentX        =   0
            _ExtentY        =   0
            Min             =   2
            Max             =   655
            Value           =   25
            AllowOnlyNumbers=   -1  'True
            TextAlignment   =   1
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
            Transparent     =   -1  'True
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
            TabIndex        =   11
            Top             =   720
            Width           =   2610
            _ExtentX        =   4604
            _ExtentY        =   450
            Caption         =   "오류 시 자동 재시도(&U)"
            Transparent     =   -1  'True
         End
         Begin prjDownloadBooster.CheckBoxW chkRememberURL 
            Height          =   255
            Left            =   840
            TabIndex        =   9
            Top             =   480
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   450
            Caption         =   "파일 주소 기억(&M)"
            Transparent     =   -1  'True
         End
         Begin prjDownloadBooster.CheckBoxW chkAlwaysResume 
            Height          =   255
            Left            =   3480
            TabIndex        =   10
            Top             =   480
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   450
            Caption         =   "항상 이어받기(&A)"
            Transparent     =   -1  'True
         End
         Begin prjDownloadBooster.CheckBoxW chkOpenDirWhenComplete 
            Height          =   255
            Left            =   3480
            TabIndex        =   8
            Top             =   240
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   450
            Caption         =   "완료 후 폴더 열기(&P)"
            Transparent     =   -1  'True
         End
         Begin prjDownloadBooster.CheckBoxW chkOpenWhenComplete 
            Height          =   255
            Left            =   840
            TabIndex        =   7
            Top             =   240
            Width           =   2520
            _ExtentX        =   4445
            _ExtentY        =   450
            Caption         =   "완료 후 파일 열기(&O)"
            Transparent     =   -1  'True
         End
         Begin VB.Image Image2 
            Height          =   480
            Left            =   120
            Picture         =   "frmOptions.frx":1736
            Top             =   240
            Width           =   480
         End
         Begin VB.Label Label15 
            BackStyle       =   0  '투명
            Caption         =   "개 (다시 시작 필요)"
            Height          =   255
            Left            =   3960
            TabIndex        =   18
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
      BorderStyle     =   0  '없음
      Enabled         =   0   'False
      Height          =   1695
      Index           =   5
      Left            =   7200
      ScaleHeight     =   1695
      ScaleWidth      =   6975
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3120
      Visible         =   0   'False
      Width           =   6975
      Begin prjDownloadBooster.FrameW FrameW2 
         Height          =   1455
         Left            =   120
         TabIndex        =   98
         Top             =   120
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   2566
         Caption         =   "프로그램 경로"
         Transparent     =   -1  'True
         Begin VB.TextBox txtYtdlPath 
            Height          =   270
            Left            =   2520
            TabIndex        =   103
            Top             =   1080
            Width           =   4095
         End
         Begin VB.TextBox txtNodePath 
            Height          =   270
            Left            =   2520
            TabIndex        =   101
            Top             =   720
            Width           =   4095
         End
         Begin VB.Image imgIcon2 
            Height          =   480
            Left            =   120
            Picture         =   "frmOptions.frx":1B78
            Top             =   240
            Width           =   480
         End
         Begin VB.Label Label6 
            BackStyle       =   0  '투명
            Caption         =   "기본값을 사용하려면 필드를 비워두십시오. 아래는 고급 사용자를 위한 것이며 일반적으로 변경할 필요가 없습니다."
            Height          =   480
            Left            =   840
            TabIndex        =   99
            Top             =   240
            Width           =   5775
         End
         Begin VB.Label Label7 
            BackStyle       =   0  '투명
            Caption         =   "&youtube-dl/yt-dlp:"
            Height          =   255
            Left            =   840
            TabIndex        =   102
            Top             =   1110
            Width           =   1695
         End
         Begin VB.Label Label4 
            BackStyle       =   0  '투명
            Caption         =   "N&ode.js:"
            Height          =   255
            Left            =   840
            TabIndex        =   100
            Top             =   750
            Width           =   1455
         End
      End
   End
   Begin VB.PictureBox pbPanel 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  '없음
      Enabled         =   0   'False
      Height          =   5385
      Index           =   3
      Left            =   120
      ScaleHeight     =   5385
      ScaleWidth      =   6975
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   5040
      Visible         =   0   'False
      Width           =   6975
      Begin prjDownloadBooster.FrameW FrameW1 
         Height          =   1275
         Left            =   120
         TabIndex        =   39
         Top             =   2280
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   2249
         Caption         =   "테마"
         Transparent     =   -1  'True
         Begin prjDownloadBooster.CommandButtonW cmdDeleteTheme 
            Height          =   300
            Left            =   960
            TabIndex        =   43
            Top             =   885
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   529
            Caption         =   "삭제(&D)"
            Transparent     =   -1  'True
         End
         Begin prjDownloadBooster.CommandButtonW cmdSaveTheme 
            Height          =   300
            Left            =   960
            TabIndex        =   42
            Top             =   570
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   529
            Caption         =   "저장(&A)..."
            Transparent     =   -1  'True
         End
         Begin VB.ComboBox cbTheme 
            Height          =   300
            Left            =   960
            Style           =   2  '드롭다운 목록
            TabIndex        =   41
            Top             =   240
            Width           =   2115
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "테마(&T):"
            Height          =   180
            Left            =   120
            TabIndex        =   40
            Top             =   285
            Width           =   690
         End
      End
      Begin prjDownloadBooster.FrameW Frame6 
         Height          =   1635
         Left            =   120
         TabIndex        =   44
         Top             =   3660
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   2884
         Caption         =   "스킨"
         Transparent     =   -1  'True
         Begin VB.ComboBox cbProgressSkin 
            Height          =   300
            Left            =   870
            Style           =   2  '드롭다운 목록
            TabIndex        =   122
            Top             =   1230
            Width           =   1665
         End
         Begin VB.ComboBox cbFont 
            Height          =   300
            Left            =   870
            TabIndex        =   51
            Top             =   900
            Width           =   2205
         End
         Begin prjDownloadBooster.CommandButtonW cmdAdvancedSkin 
            Height          =   300
            Left            =   2580
            TabIndex        =   49
            Top             =   570
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   529
            ImageListAlignment=   4
            Transparent     =   -1  'True
         End
         Begin VB.ComboBox cbSkin 
            Height          =   300
            Left            =   870
            Style           =   2  '드롭다운 목록
            TabIndex        =   48
            Top             =   570
            Width           =   1665
         End
         Begin VB.ComboBox cbFrameSkin 
            Height          =   300
            Left            =   870
            Style           =   2  '드롭다운 목록
            TabIndex        =   46
            Top             =   240
            Width           =   2205
         End
         Begin prjDownloadBooster.CommandButtonW cmdAdvancedProgressSkin 
            Height          =   300
            Left            =   2580
            TabIndex        =   123
            Top             =   1230
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   529
            ImageListAlignment=   4
            Transparent     =   -1  'True
         End
         Begin VB.Label Label5 
            BackStyle       =   0  '투명
            Caption         =   "현황(&R):"
            Height          =   255
            Left            =   120
            TabIndex        =   124
            Top             =   1275
            Width           =   855
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "글꼴(&F):"
            Height          =   180
            Left            =   120
            TabIndex        =   50
            Top             =   945
            Width           =   675
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "창(&W):"
            Height          =   180
            Left            =   120
            TabIndex        =   45
            Top             =   285
            Width           =   540
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "단추(&O):"
            Height          =   180
            Left            =   120
            TabIndex        =   47
            Top             =   615
            Width           =   705
         End
      End
      Begin prjDownloadBooster.FrameW Frame4 
         Height          =   1635
         Left            =   3600
         TabIndex        =   63
         Top             =   3660
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   2884
         Caption         =   "배경"
         Transparent     =   -1  'True
         Begin prjDownloadBooster.CheckBoxW chkCenter 
            Height          =   255
            Left            =   960
            TabIndex        =   121
            Top             =   1260
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   450
            Caption         =   "가운데(&C)"
            Transparent     =   -1  'True
         End
         Begin VB.ComboBox cbImagePosition 
            Height          =   300
            Left            =   960
            Style           =   2  '드롭다운 목록
            TabIndex        =   71
            Top             =   900
            Width           =   2055
         End
         Begin VB.ComboBox lvBackgrounds 
            Height          =   300
            Left            =   960
            Style           =   2  '드롭다운 목록
            TabIndex        =   68
            Top             =   570
            Width           =   1515
         End
         Begin VB.ComboBox lvPatterns 
            Height          =   300
            Left            =   960
            Style           =   2  '드롭다운 목록
            TabIndex        =   65
            Top             =   240
            Width           =   1515
         End
         Begin VB.FileListBox lvBackgroundFiles 
            Height          =   450
            Left            =   -480
            Pattern         =   "*.JPG;*.JPEG;*.JPE;*.JFIF;*.GIF;*.BMP;*.DIB;*.PNG;*.WMF;*.EMF;*.ICO;*.CUR;*.TIF;*.TIFF;*.RLE"
            TabIndex        =   119
            TabStop         =   0   'False
            Top             =   240
            Visible         =   0   'False
            Width           =   615
         End
         Begin prjDownloadBooster.CommandButtonW cmdChooseBackground 
            Height          =   300
            Left            =   2520
            TabIndex        =   69
            Top             =   570
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   529
            ImageList       =   "imgBrowse"
            ImageListAlignment=   4
         End
         Begin VB.Label Label22 
            BackStyle       =   0  '투명
            Caption         =   "배경(&A):"
            Height          =   255
            Left            =   120
            TabIndex        =   67
            Top             =   630
            Width           =   855
         End
         Begin VB.Label Label21 
            BackStyle       =   0  '투명
            Caption         =   "무늬(&P):"
            Height          =   255
            Left            =   120
            TabIndex        =   64
            Top             =   300
            Width           =   855
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '투명
            Caption         =   "위치(&S):"
            Height          =   255
            Left            =   120
            TabIndex        =   70
            Top             =   945
            Width           =   840
         End
         Begin VB.Label lblFillColorSelect 
            BackStyle       =   0  '투명
            Height          =   255
            Left            =   2520
            TabIndex        =   66
            Top             =   255
            Width           =   495
         End
         Begin VB.Shape pgPatternColor 
            BackColor       =   &H00000000&
            BackStyle       =   1  '투명하지 않음
            BorderColor     =   &H00404040&
            FillColor       =   &H00808080&
            Height          =   255
            Left            =   2520
            Shape           =   4  '둥근 사각형
            Top             =   255
            Width           =   495
         End
      End
      Begin VB.PictureBox pbOuterPreview 
         AutoRedraw      =   -1  'True
         Height          =   2055
         Left            =   120
         ScaleHeight     =   1995
         ScaleWidth      =   6675
         TabIndex        =   107
         TabStop         =   0   'False
         Top             =   120
         Width           =   6735
         Begin VB.PictureBox pbBackground 
            Height          =   1320
            Left            =   360
            ScaleHeight     =   1260
            ScaleWidth      =   3735
            TabIndex        =   109
            TabStop         =   0   'False
            Tag             =   "nobgdraw"
            Top             =   60
            Width           =   3795
            Begin VB.TextBox txtSampleClassic 
               Height          =   270
               Left            =   1020
               TabIndex        =   116
               Top             =   60
               Visible         =   0   'False
               Width           =   2415
            End
            Begin prjDownloadBooster.CheckBoxW CheckBoxW1 
               Height          =   255
               Left            =   60
               TabIndex        =   110
               Top             =   960
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   450
               Caption         =   "완료 후 열기"
               Transparent     =   -1  'True
            End
            Begin VB.TextBox TextBoxW1 
               Height          =   270
               Left            =   1020
               TabIndex        =   111
               Top             =   60
               Width           =   2415
            End
            Begin prjDownloadBooster.FrameW FrameW5 
               Height          =   555
               Left            =   60
               TabIndex        =   112
               Top             =   375
               Width           =   3615
               _ExtentX        =   6376
               _ExtentY        =   979
               Caption         =   "다운로드 현황"
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
               Left            =   2100
               TabIndex        =   113
               TabStop         =   0   'False
               Tag             =   "notygchange"
               Top             =   960
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
               Left            =   60
               TabIndex        =   114
               Top             =   120
               Width           =   975
            End
            Begin VB.Image imgPreview 
               Height          =   135
               Left            =   0
               Stretch         =   -1  'True
               Top             =   0
               Width           =   135
            End
            Begin VB.Shape pgPatternPreview 
               BackColor       =   &H8000000F&
               BackStyle       =   1  '투명하지 않음
               BorderStyle     =   0  '투명
               Height          =   255
               Left            =   0
               Top             =   0
               Width           =   255
            End
         End
         Begin VB.PictureBox pbPreview 
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000001&
            BorderStyle     =   0  '없음
            Enabled         =   0   'False
            Height          =   2055
            Left            =   0
            ScaleHeight     =   2055
            ScaleWidth      =   6735
            TabIndex        =   108
            TabStop         =   0   'False
            Tag             =   "nobgdraw"
            Top             =   0
            Width           =   6735
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
         Height          =   1275
         Left            =   3600
         TabIndex        =   52
         Top             =   2280
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   2249
         Caption         =   "색상"
         Transparent     =   -1  'True
         Begin prjDownloadBooster.FrameW FrameW10 
            Height          =   510
            Left            =   120
            TabIndex        =   118
            Top             =   480
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   900
            BorderStyle     =   0
            Caption         =   "FrameW10"
            Transparent     =   -1  'True
            Begin VB.OptionButton optSystemColor 
               Caption         =   "시스템(&S)"
               Height          =   255
               Left            =   120
               TabIndex        =   54
               Top             =   0
               Width           =   1215
            End
            Begin VB.OptionButton optUserColor 
               Height          =   255
               Left            =   120
               TabIndex        =   55
               Top             =   240
               Width           =   255
            End
            Begin VB.Label lblSelectColor 
               BackStyle       =   0  '투명
               Height          =   255
               Left            =   375
               TabIndex        =   56
               Top             =   255
               Width           =   600
            End
            Begin VB.Shape pgColor 
               BackStyle       =   1  '투명하지 않음
               BorderColor     =   &H00404040&
               FillColor       =   &H00808080&
               Height          =   255
               Left            =   375
               Shape           =   4  '둥근 사각형
               Top             =   255
               Width           =   615
            End
         End
         Begin prjDownloadBooster.CheckBoxW chkBackColorMainOnly 
            Height          =   255
            Left            =   240
            TabIndex        =   57
            Top             =   990
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   450
            Caption         =   "메인 창만(&O)"
            Transparent     =   -1  'True
         End
         Begin prjDownloadBooster.CheckBoxW chkForeColorMainOnly 
            Height          =   255
            Left            =   1680
            TabIndex        =   62
            Top             =   990
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   450
            Enabled         =   0   'False
            Caption         =   "메인 창만(&N)"
            Transparent     =   -1  'True
         End
         Begin VB.OptionButton optUserFore 
            Height          =   255
            Left            =   1680
            TabIndex        =   60
            Top             =   720
            Width           =   255
         End
         Begin VB.OptionButton optSystemFore 
            Caption         =   "시스템(&Y)"
            Height          =   255
            Left            =   1680
            TabIndex        =   59
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label lblSelectFore 
            BackStyle       =   0  '투명
            Height          =   255
            Left            =   1935
            TabIndex        =   61
            Top             =   735
            Width           =   615
         End
         Begin VB.Label Label12 
            BackStyle       =   0  '투명
            Caption         =   "글자(&E):"
            Height          =   255
            Left            =   1560
            TabIndex        =   58
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label9 
            BackStyle       =   0  '투명
            Caption         =   "배경(&B):"
            Height          =   255
            Left            =   120
            TabIndex        =   53
            Top             =   240
            Width           =   1575
         End
         Begin VB.Shape pgFore 
            BackStyle       =   1  '투명하지 않음
            BorderColor     =   &H00404040&
            FillColor       =   &H00808080&
            Height          =   255
            Left            =   1935
            Shape           =   4  '둥근 사각형
            Top             =   735
            Width           =   615
         End
      End
   End
   Begin prjDownloadBooster.CommandButtonW cmdApply 
      Height          =   360
      Left            =   9840
      TabIndex        =   106
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
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   661
      TabFixedWidth   =   53
      TabMinWidth     =   48
      TabScrollWheel  =   0   'False
      Transparent     =   -1  'True
      InitTabs        =   "frmOptions.frx":1FBA
   End
   Begin prjDownloadBooster.CommandButtonW CancelButton 
      Cancel          =   -1  'True
      Height          =   360
      Left            =   8400
      TabIndex        =   105
      Top             =   120
      Width           =   1320
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "취소"
   End
   Begin prjDownloadBooster.CommandButtonW OKButton 
      Default         =   -1  'True
      Height          =   360
      Left            =   6960
      TabIndex        =   104
      Top             =   120
      Width           =   1320
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "확인"
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'참고 자료:
'- https://www.vbforums.com/showthread.php?284592-Listview-StartLabelEdit-second-column-*RESOLVED*

Dim Loaded As Boolean
Public ColorChanged As Boolean
Public ImageChanged As Boolean
Public VisualStyleChanged As Boolean
Dim SkinChanged As Boolean
Public ProgressSkinChanged As Boolean
Public FontChanged As Boolean
Dim PatternChanged As Boolean
Dim ScrollChanged As Boolean
Dim IntervalValues(7) As Single
Public ChangedBackgroundPath$
Dim PreviewControls(4) As Control
Dim DoLoadTheme As Boolean

Public RoundClassicButtons As Byte, DisableVisualStyle As Byte
Public LiveBadukMemoSkinShadowColor&, LiveBadukMemoSkinFrameColor&, LiveBadukMemoSkinFrameType$, LiveBadukMemoSkinTextColor&, LiveBadukMemoSkinEnableShadow As Byte, LiveBadukMemoSkinEnableTextColor As Byte, LiveBadukMemoSkinEnableBorder As Byte, LiveBadukMemoSkinFrameBackgroundType$, LiveBadukMemoSkinFrameBackgroundColor&, LiveBadukMemoSkinContentTextColor&, LiveBadukMemoSkinFrameTexture$, LiveBadukMemoSkinFrameBackground$, LiveBadukMemoSkinLabelFontSize As Integer, LiveBadukMemoSkinLabelFontBold As Byte, LiveBadukMemoSkinEnableLabelFontSize As Byte

Implements IBSSubclass

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub OnFontChange()
    On Error Resume Next
    
    Dim FontName$, FontSize%
    If FontExists(cbFont.Text) Then
        FontName = cbFont.Text
    ElseIf cbFont.Text = t("(기본값)", "(default)") Then
        If t(1, 2) = 2 Then
            FontName = "Tahoma"
        Else
            FontName = DefaultFont
        End If
    End If
    If LCase(FontName) = "tahoma" Or Left$(FontName, 7) = "Tahoma " Then FontSize = 8 Else FontSize = 9
    
    Dim i As Byte
    For i = LBound(PreviewControls) To UBound(PreviewControls)
        PreviewControls(i).Font.Name = FontName
        PreviewControls(i).Font.Size = FontSize
        PreviewControls(i).Font.Bold = False
        PreviewControls(i).Font.Italic = False
    Next i
    
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
    If Loaded Then
        cmdApply.Enabled = -1
        ImageChanged = True
    End If
End Sub

Private Sub cbLanguage_Click()
    If Loaded Then cmdApply.Enabled = -1
End Sub

Private Sub cbProgressSkin_Click()
    cmdAdvancedProgressSkin.Enabled = (cbProgressSkin.ListIndex <> 0)
    If Loaded Then
        ProgressSkinChanged = True
        cmdApply.Enabled = -1
    End If
End Sub

Private Sub cbSkin_Click()
    cmdSample.IsTygemButton = (cbSkin.ListIndex > 0)
    If cmdSample.IsTygemButton Then cmdSample.GetTygemButton().Skin = cbSkin.ListIndex
    cmdSample.Refresh
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
'        If cbSkin.ListIndex = 2 And DPI <> 96 Then
'            MsgBox t("이 스킨의 일부 요소는 96 DPI(100% 배율)에서만 표시됩니다.", "Some of the elements of this skin only works in 96 DPI (100% size)."), 48
'        End If
    End If
    If optUserFore.Value Then
        CheckBoxW1.VisualStyles = False
        FrameW5.VisualStyles = False
        CheckBoxW1.ForeColor = pgFore.BackColor
        FrameW5.ForeColor = pgFore.BackColor
    End If
    'cmdAdvancedSkin.Enabled = (cbSkin.ListIndex = 0)
    cmdSample.RoundButton = (GetSetting("DownloadBooster", "Options", "RoundClassicButtons", 0) <> 0)
End Sub

Private Sub LoadTheme(Optional ByVal ThemeName As String = "")
    Dim Section$
    If ThemeName = "" Then Section = "Options" Else Section = "Options\Themes\" & ThemeName
    On Error Resume Next
    
    txtCompleteSoundPath.Text = Trim$(GetSetting("DownloadBooster", Section, "CompleteSoundPath", ""))
    
    If GetSetting("DownloadBooster", Section, "UseClassicThemeFrame", 0) <> 0 Then
        If cbFrameSkin.ListCount >= 3 Then
            cbFrameSkin.ListIndex = 2
        Else
            cbFrameSkin.ListIndex = 1
        End If
    ElseIf GetSetting("DownloadBooster", Section, "DisableDWMWindow", DefaultDisableDWMWindow) <> 0 And cbFrameSkin.ListCount >= 3 Then
        cbFrameSkin.ListIndex = 1
    Else
        cbFrameSkin.ListIndex = 0
    End If
    
    Dim clrBackColor As Long
    clrBackColor = GetSetting("DownloadBooster", Section, "BackColor", DefaultBackColor)
    If clrBackColor < 0 Or clrBackColor > 16777215 Then
        optSystemColor.Value = True
        pgColor.BackColor = &H8000000F
    Else
        optUserColor.Value = True
        pgColor.BackColor = clrBackColor
    End If
    pbBackground.BackColor = pgColor.BackColor
    pgPatternPreview.BackColor = pgColor.BackColor
    
    chkBeepWhenComplete.Value = GetSetting("DownloadBooster", Section, "PlaySound", 1)
    
    cbSkin.ListIndex = CurrentButtonSkin
    cbProgressSkin.ListIndex = GetSetting("DownloadBooster", Section, "ProgressFrameSkin", 1)
    
    cmdSample.RoundButton = RoundClassicButtons
    cmdSample.VisualStyles = (DisableVisualStyle = 0)
    cmdSample.IsTygemButton = cbSkin.ListIndex > 0
    
    lvPatterns.ListIndex = CInt(GetSetting("DownloadBooster", Section, "FormFillStyle", 0))
    
    ChangedBackgroundPath = GetSetting("DownloadBooster", Section, "BackgroundImagePath", "")
    LoadBackgroundList (ThemeName = "")
    
    pgPatternColor.BackColor = CLng(GetSetting("DownloadBooster", Section, "FormFillColor", 0))
    pgPatternPreview.FillColor = pgPatternColor.BackColor
    pgPatternPreview.FillStyle = lvPatterns.ListIndex + 1
    
    cbImagePosition.ListIndex = GetSetting("DownloadBooster", Section, "ImagePosition", 1)
    cbImagePosition_Click
    
    cbFont.Text = Trim$(GetSetting("DownloadBooster", Section, "Font", ""))
    If cbFont.Text = "" Then cbFont.Text = ("(" & t("기본값", "default") & ")")
    
    Dim clrForeColor As Long
    clrForeColor = GetSetting("DownloadBooster", Section, "ForeColor", -1)
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
End Sub

Private Sub cbTheme_Click()
    cmdDeleteTheme.Enabled = (cbTheme.ListIndex > 0)
    If cbTheme.ListIndex = 0 Then Exit Sub
    If Not Loaded Then Exit Sub
    
    On Error Resume Next
    
    Dim ThemeName$
    ThemeName = cbTheme.List(cbTheme.ListIndex)
    
    If DoLoadTheme Then
        VisualStyleChanged = True
        ImageChanged = True
        ColorChanged = True
        SkinChanged = True
        FontChanged = True
    
        DisableVisualStyle = GetSetting("DownloadBooster", "Options\Themes\" & ThemeName, "DisableVisualStyle", 0)
        RoundClassicButtons = GetSetting("DownloadBooster", "Options\Themes\" & ThemeName, "RoundClassicButtons", 0)
                
        LiveBadukMemoSkinShadowColor = CLng(GetSetting("DownloadBooster", "Options\Themes\" & ThemeName, "LiveBadukMemoSkinShadowColor", 16777215))
        LiveBadukMemoSkinFrameColor = CLng(GetSetting("DownloadBooster", "Options\Themes\" & ThemeName, "LiveBadukMemoSkinFrameColor", 11194780))
        LiveBadukMemoSkinFrameType = GetSetting("DownloadBooster", "Options\Themes\" & ThemeName, "LiveBadukMemoSkinFrameType", "solidcolor")
        LiveBadukMemoSkinTextColor = CLng(GetSetting("DownloadBooster", "Options\Themes\" & ThemeName, "LiveBadukMemoSkinTextColor", 2902560))
        LiveBadukMemoSkinEnableShadow = CByte(GetSetting("DownloadBooster", "Options\Themes\" & ThemeName, "LiveBadukMemoSkinEnableShadow", 1))
        LiveBadukMemoSkinEnableTextColor = CByte(GetSetting("DownloadBooster", "Options\Themes\" & ThemeName, "LiveBadukMemoSkinEnableTextColor", 1))
        LiveBadukMemoSkinEnableBorder = CByte(GetSetting("DownloadBooster", "Options\Themes\" & ThemeName, "LiveBadukMemoSkinEnableBorder", 1))
        LiveBadukMemoSkinFrameBackgroundType = GetSetting("DownloadBooster", "Options\Themes\" & ThemeName, "LiveBadukMemoSkinFrameBackgroundType", "transparent")
        LiveBadukMemoSkinFrameBackgroundColor = CLng(GetSetting("DownloadBooster", "Options\Themes\" & ThemeName, "LiveBadukMemoSkinFrameBackgroundColor", 16777215))
        LiveBadukMemoSkinContentTextColor = CLng(GetSetting("DownloadBooster", "Options\Themes\" & ThemeName, "LiveBadukMemoSkinContentTextColor", 0))
        LiveBadukMemoSkinFrameTexture = GetSetting("DownloadBooster", "Options\Themes\" & ThemeName, "LiveBadukMemoSkinFrameTexture", "")
        LiveBadukMemoSkinFrameBackground = GetSetting("DownloadBooster", "Options\Themes\" & ThemeName, "LiveBadukMemoSkinFrameBackground", "")
        LiveBadukMemoSkinLabelFontSize = CInt(GetSetting("DownloadBooster", "Options\Themes\" & ThemeName, "LiveBadukMemoSkinLabelFontSize", 10))
        LiveBadukMemoSkinLabelFontBold = CByte(GetSetting("DownloadBooster", "Options\Themes\" & ThemeName, "LiveBadukMemoSkinLabelFontBold", 1))
        LiveBadukMemoSkinEnableLabelFontSize = CByte(GetSetting("DownloadBooster", "Options\Themes\" & ThemeName, "LiveBadukMemoSkinEnableLabelFontSize", 1))
        
        LoadTheme ThemeName
    End If
    
    RedrawPreview
    
    cmdApply.Enabled = True
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

Private Sub chkCenter_Click()
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

Private Sub chkLazyElapsed_Click()
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

Private Sub chkUseServerModified_Click()
    If Loaded Then cmdApply.Enabled = -1
End Sub

Private Sub cmdAdvancedProgressSkin_Click()
    frmLiveBadukSkinProperties.Show vbModal, Me
End Sub

Private Sub cmdAdvancedSkin_Click()
    frmSystemSkinProperties.Show vbModal, Me
End Sub

Private Sub cmdApply_Click()
    If cbFrameSkin.ListCount >= 3 Then
        SaveSetting "DownloadBooster", "Options", "DisableDWMWindow", -(cbFrameSkin.ListIndex = 1)
    End If
    SaveSetting "DownloadBooster", "Options", "UseClassicThemeFrame", -((cbFrameSkin.ListCount >= 3 And cbFrameSkin.ListIndex = 2) Or (cbFrameSkin.ListCount < 3 And cbFrameSkin.ListIndex = 1))
    SaveSetting "DownloadBooster", "Options", "RememberURL", chkRememberURL.Value
    SaveSetting "DownloadBooster", "Options", "NoRedirectCheck", chkNoRedirectCheck.Value
    SaveSetting "DownloadBooster", "Options", "ForceGet", chkForceGet.Value
    SaveSetting "DownloadBooster", "Options", "Ignore300", chkIgnore300.Value
    SaveSetting "DownloadBooster", "Options", "LazyElapsed", chkLazyElapsed.Value
    SaveSetting "DownloadBooster", "Options", "ForceWin31Dialog", chkForceOldDialog.Value
    SaveSetting "DownloadBooster", "Options", "AutoDetectYtdlURL", chkAutoYtdl.Value
    SaveSetting "DownloadBooster", "Options", "CompleteSoundPath", Trim$(txtCompleteSoundPath.Text)
    SaveSetting "DownloadBooster", "Options", "AllowDuplicatesInQueue", chkAllowDuplicates.Value
    SaveSetting "DownloadBooster", "Options", "ScrollOneScreen", -optScreenPerScroll.Value
    SaveSetting "DownloadBooster", "Options", "BackColorMainOnly", chkBackColorMainOnly.Value
    SaveSetting "DownloadBooster", "Options", "ForeColorMainOnly", chkForeColorMainOnly.Value
    SaveSetting "DownloadBooster", "Options", "UseServerModifiedDate", chkUseServerModified.Value
    If ScrollChanged Then
        frmMain.ScrollOneScreen = optScreenPerScroll.Value
        frmMain.trThreadCount_Scroll
        frmMain.pbProgressContainer.Top = 0
        frmMain.vsProgressScroll.Value = 0
        frmMain.pbProgressContainer.Refresh
        Dim LargeChange As Byte
        If optScreenPerScroll.Value Then LargeChange = 1 Else LargeChange = 10
        frmMain.vsProgressScroll.LargeChange = LargeChange
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
    
    On Error GoTo maxtrdnotint
    If Not IsNumeric(txtMaxThreadCount.Value) Then
maxtrdnotint:
        MsgBox t("최대 쓰레드 개수는 정수여야 합니다.", "Maximum number of threads should be an integer."), 16
        NoDisable = True
        GoTo aftermaxtrdcheck
    ElseIf txtMaxThreadCount.Value > MAX_THREAD_COUNT_CONTROL Or txtMaxThreadCount.Text < 2& Then
        MsgBox t("최대 쓰레드 개수는 2개 이상 " & MAX_THREAD_COUNT_CONTROL & "개 이하여야 합니다.", "Maximum number of threads should range in 2-" & MAX_THREAD_COUNT_CONTROL & "."), 16
        NoDisable = True
    Else
        SaveSetting "DownloadBooster", "Options", "MaxThreadCount", txtMaxThreadCount.Value
        If txtMaxThreadCount.Value = 655 Then txtMaxThreadCount.Value = 655
'        If CInt(txtMaxThreadCount.Text) > 50 Then
'            MsgBox t("최대 쓰레드 개수가 너무 클 경우 실제 사용 개수와 관계없이 실행 속도가 느려질 수 있습니다.", "If the maximum number of threads is too high, the application might run slower."), 48
'        End If
    End If
aftermaxtrdcheck:
    On Error GoTo 0
    
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
    If ColorChanged Or VisualStyleChanged Or SkinChanged Then
        SaveSetting "DownloadBooster", "Options", "DisableVisualStyle", DisableVisualStyle
        SaveSetting "DownloadBooster", "Options", "RoundClassicButtons", RoundClassicButtons
        
        SaveSetting "DownloadBooster", "Options", "LiveBadukMemoSkinShadowColor", LiveBadukMemoSkinShadowColor
        SaveSetting "DownloadBooster", "Options", "LiveBadukMemoSkinFrameColor", LiveBadukMemoSkinFrameColor
        SaveSetting "DownloadBooster", "Options", "LiveBadukMemoSkinFrameType", LiveBadukMemoSkinFrameType
        SaveSetting "DownloadBooster", "Options", "LiveBadukMemoSkinTextColor", LiveBadukMemoSkinTextColor
        SaveSetting "DownloadBooster", "Options", "LiveBadukMemoSkinEnableShadow", LiveBadukMemoSkinEnableShadow
        SaveSetting "DownloadBooster", "Options", "LiveBadukMemoSkinEnableTextColor", LiveBadukMemoSkinEnableTextColor
        SaveSetting "DownloadBooster", "Options", "LiveBadukMemoSkinEnableBorder", LiveBadukMemoSkinEnableBorder
        SaveSetting "DownloadBooster", "Options", "LiveBadukMemoSkinFrameBackgroundType", LiveBadukMemoSkinFrameBackgroundType
        SaveSetting "DownloadBooster", "Options", "LiveBadukMemoSkinFrameBackgroundColor", LiveBadukMemoSkinFrameBackgroundColor
        SaveSetting "DownloadBooster", "Options", "LiveBadukMemoSkinContentTextColor", LiveBadukMemoSkinContentTextColor
        SaveSetting "DownloadBooster", "Options", "LiveBadukMemoSkinFrameTexture", LiveBadukMemoSkinFrameTexture
        SaveSetting "DownloadBooster", "Options", "LiveBadukMemoSkinFrameBackground", LiveBadukMemoSkinFrameBackground
        SaveSetting "DownloadBooster", "Options", "LiveBadukMemoSkinLabelFontSize", LiveBadukMemoSkinLabelFontSize
        SaveSetting "DownloadBooster", "Options", "LiveBadukMemoSkinLabelFontBold", LiveBadukMemoSkinLabelFontBold
        SaveSetting "DownloadBooster", "Options", "LiveBadukMemoSkinEnableLabelFontSize", LiveBadukMemoSkinEnableLabelFontSize
    
        SaveSetting "DownloadBooster", "Options", "ButtonSkin", cbSkin.ListIndex
        CurrentButtonSkin = cbSkin.ListIndex
        
        SetFormBackgroundColor Me, True
        SetFormBackgroundColor frmMain, True
        RedrawPreview
        cmdChooseBackground.Refresh
        frmMain.pbProgressContainer.Refresh
        frmMain.SetTextColors
    End If
    If ProgressSkinChanged Then
        SaveSetting "DownloadBooster", "Options", "ProgressFrameSkin", cbProgressSkin.ListIndex
        frmMain.LoadLiveBadukSkin
        frmMain.SetupSplitButtons
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
        SaveSetting "DownloadBooster", "Options", "UseBackgroundImage", -(lvBackgrounds.ListIndex <> 0)
        SaveSetting "DownloadBooster", "Options", "BackgroundImagePath", ChangedBackgroundPath
        frmMain.SetBackgroundImage
        frmMain.SetBackgroundPosition True
    End If
    
    If LenB(Trim$(txtNodePath.Text)) Then
        If FileExists(Trim$(txtNodePath.Text)) Then
            SaveSetting "DownloadBooster", "Options", "NodePath", Trim$(txtNodePath.Text)
        Else
            MsgBox t("Node.js 경로가 존재하지 않습니다.", "Node.js path does not exist."), 16
            NoDisable = True
        End If
    Else
        SaveSetting "DownloadBooster", "Options", "NodePath", ""
    End If
    If LenB(Trim$(txtYtdlPath.Text)) Then
        If FileExists(Trim$(txtYtdlPath.Text)) Then
            SaveSetting "DownloadBooster", "Options", "YtdlPath", Trim$(txtYtdlPath.Text)
        Else
            MsgBox t("Youtube-dl 경로가 존재하지 않습니다.", "Youtube-dl path does not exist."), 16
            NoDisable = True
        End If
    Else
        SaveSetting "DownloadBooster", "Options", "YtdlPath", ""
    End If
    
    If FontChanged Then
        cbFont.Text = Trim$(cbFont.Text)
        If LenB(cbFont.Text) And cbFont.Text <> ("(" & t("기본값", "default") & ")") And (Not FontExists(cbFont.Text)) Then
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
        InitPropertySheetDimensions Me, tsTabStrip, pbPanel, OKButton, CancelButton, cmdApply
    End If
    
    If lvBackgrounds.ListIndex <> 0 And LenB(GetSetting("DownloadBooster", "Options", "BackgroundImagePath", "")) = 0 Then
        MsgBox t("배경 그림이 선택되지 않았습니다.", "Background image is not selected."), 48
        SaveSetting "DownloadBooster", "Options", "UseBackgroundImage", "0"
        NoDisable = True
    End If
    
    Dim hSysMenu As Long
    Dim MII As MENUITEMINFO
    hSysMenu = GetSystemMenu(frmMain.hWnd, 0)
    MainFormOnTop = (chkAlwaysOnTop.Value = 1)
    Dim InsertAfter As Long
    If MainFormOnTop Then InsertAfter = hWnd_TOPMOST Else InsertAfter = hWnd_NOTOPMOST
    SetWindowPos frmMain.hWnd, InsertAfter, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    SetWindowPos Me.hWnd, InsertAfter, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    With MII
        .cbSize = Len(MII)
        .fMask = MIIM_STATE
        .fState = MFS_ENABLED
        If MainFormOnTop Then .fState = .fState Or MFS_CHECKED
    End With
    SetMenuItemInfo hSysMenu, 1000, 0, MII
    SaveSetting "DownloadBooster", "Options", "AlwaysOnTop", -MainFormOnTop
    
    If optUserFore.Value Then
        CheckBoxW1.VisualStyles = False
        FrameW5.VisualStyles = False
        CheckBoxW1.ForeColor = pgFore.BackColor
        FrameW5.ForeColor = pgFore.BackColor
    End If
    
    Dim SaveThemeName As String
    If cbTheme.ListIndex = 0 Then SaveThemeName = "" Else SaveThemeName = cbTheme.List(cbTheme.ListIndex)
    SaveSetting "DownloadBooster", "Options", "Theme", SaveThemeName
    
    RedrawPreview
    ResetChanged
    If Not NoDisable Then cmdApply.Enabled = 0
End Sub

Private Sub ResetChanged()
    ColorChanged = False
    ImageChanged = False
    VisualStyleChanged = False
    SkinChanged = False
    ProgressSkinChanged = False
    ScrollChanged = False
    FontChanged = False
    PatternChanged = False
End Sub

Private Sub cmdBrowseAsterisk_Click()
    Set Tags.BrowseTargetTextbox = txtAsterisk
    ShowFileDialog 4, Tags.BrowseTargetTextbox.Text, True
End Sub

Private Sub cmdBrowseCompleteSound_Click()
    Set Tags.BrowseTargetTextbox = txtCompleteSoundPath
    ShowFileDialog 4, Tags.BrowseTargetTextbox.Text, True
End Sub

Private Sub cmdBrowseError_Click()
    Set Tags.BrowseTargetTextbox = txtError
    ShowFileDialog 4, Tags.BrowseTargetTextbox.Text, True
End Sub

Private Sub cmdBrowseExclamation_Click()
    Set Tags.BrowseTargetTextbox = txtExclamation
    ShowFileDialog 4, Tags.BrowseTargetTextbox.Text, True
End Sub

Private Sub cmdBrowseQuestion_Click()
    Set Tags.BrowseTargetTextbox = txtQuestion
    ShowFileDialog 4, Tags.BrowseTargetTextbox.Text, True
End Sub

Private Sub cmdChooseBackground_Click()
    ShowFileDialog 3, GetSetting("DownloadBooster", "Options", "BackgroundImagePath", ""), ForceNewDialog:=True
End Sub

Private Sub cmdDeleteTheme_Click()
    If cbTheme.ListIndex = 0 Then Exit Sub
    On Error Resume Next
    If MsgBox(t("선택한 테마를 삭제하시겠습니까?", "Delete the selected theme?"), vbQuestion + vbYesNo) = vbYes Then
        DeleteSetting "DownloadBooster", "Options\Themes\" & cbTheme.List(cbTheme.ListIndex)
        cbTheme.RemoveItem cbTheme.ListIndex
        cbTheme.ListIndex = 0
    End If
End Sub

Private Sub cmdHeaders_Click()
    Tags.DownloadOptionsTargetForm = 3
    Load frmDownloadOptions
#If HIDEYTDL Then
#Else
    frmDownloadOptions.RemoveYtdlTab
#End If
    frmDownloadOptions.Show vbModal, Me
End Sub

Private Sub cmdSaveTheme_Click()
    Dim ThemeName$, DefThemeName$
    If cbTheme.ListIndex = 0 Then DefThemeName = "" Else DefThemeName = cbTheme.List(cbTheme.ListIndex)
    ThemeName = InputBoxEx(t("테마 이름을 입력하십시오.", "Choose your theme name."), t("테마 저장", "Save theme"), DefThemeName)
    If ThemeName = "" Then
        Exit Sub
    ElseIf InStr(ThemeName, "\") Then
        MsgBox t("테마 이름에 허용되지 않은 문자가 포함되어 있습니다.", "Theme name contains invalid characters."), 16
        Exit Sub
    ElseIf ThemeName = "수정된 테마" Or LCase(ThemeName) = "modified theme" Then
        MsgBox t("테마 이름이 올바르지 않습니다.", "Theme name is invalid."), 16
        Exit Sub
    End If
    
    On Error Resume Next
    DeleteSetting "DownloadBooster", "Options\Themes\" & ThemeName
    
    cbFont.Text = Trim$(cbFont.Text)
    If Not (cbFont.Text <> "" And cbFont.Text <> ("(" & t("기본값", "default") & ")") And (Not FontExists(cbFont.Text))) Then
        If cbFont.Text = ("(" & t("기본값", "default") & ")") Then
            SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "Font", ""
        Else
            SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "Font", cbFont.Text
        End If
    End If
    
    If cbFrameSkin.ListCount >= 3 Then
        SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "DisableDWMWindow", -(cbFrameSkin.ListIndex = 1)
    End If
    SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "UseClassicThemeFrame", -((cbFrameSkin.ListCount >= 3 And cbFrameSkin.ListIndex = 2) Or (cbFrameSkin.ListCount < 3 And cbFrameSkin.ListIndex = 1))
    
    SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "CompleteSoundPath", Trim$(txtCompleteSoundPath.Text)
    SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "BackColorMainOnly", chkBackColorMainOnly.Value
    SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "ForeColorMainOnly", chkForeColorMainOnly.Value
    
    SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "FormFillStyle", lvPatterns.ListIndex
    SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "FormFillColor", pgPatternColor.BackColor
    
    SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "EnableAsteriskSound", chkAsterisk.Value
    SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "EnableExclamationSound", chkExclamation.Value
    SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "EnableErrorSound", chkError.Value
    SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "EnableQuestionSound", chkQuestion.Value
    SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "AsteriskSound", txtAsterisk.Text
    SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "ExclamationSound", txtExclamation.Text
    SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "ErrorSound", txtError.Text
    SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "QuestionSound", txtQuestion.Text
    
    SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "PlaySound", chkBeepWhenComplete.Value
    
    If optSystemColor.Value Then
        SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "BackColor", "-1"
    ElseIf optUserColor.Value Then
        SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "BackColor", CLng(pgColor.BackColor)
    End If
    If optSystemFore.Value Then
        SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "ForeColor", "-1"
    ElseIf optUserFore.Value Then
        SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "ForeColor", CLng(pgFore.BackColor)
    End If
    SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "ButtonSkin", cbSkin.ListIndex
    SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "ProgressFrameSkin", cbProgressSkin.ListIndex
    SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "ImagePosition", cbImagePosition.ListIndex
    SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "UseBackgroundImage", -(lvBackgrounds.ListIndex <> 0)
    SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "BackgroundImagePath", ChangedBackgroundPath
    
    If lvBackgrounds.ListIndex <> 0 And LenB(GetSetting("DownloadBooster", "Options", "BackgroundImagePath", "")) = 0 Then
        SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "UseBackgroundImage", "0"
    End If
    
    SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "DisableVisualStyle", DisableVisualStyle
    SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "RoundClassicButtons", RoundClassicButtons
    
    SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "LiveBadukMemoSkinShadowColor", LiveBadukMemoSkinShadowColor
    SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "LiveBadukMemoSkinFrameColor", LiveBadukMemoSkinFrameColor
    SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "LiveBadukMemoSkinFrameType", LiveBadukMemoSkinFrameType
    SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "LiveBadukMemoSkinTextColor", LiveBadukMemoSkinTextColor
    SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "LiveBadukMemoSkinEnableShadow", LiveBadukMemoSkinEnableShadow
    SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "LiveBadukMemoSkinEnableTextColor", LiveBadukMemoSkinEnableTextColor
    SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "LiveBadukMemoSkinEnableBorder", LiveBadukMemoSkinEnableBorder
    SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "LiveBadukMemoSkinFrameBackgroundType", LiveBadukMemoSkinFrameBackgroundType
    SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "LiveBadukMemoSkinFrameBackgroundColor", LiveBadukMemoSkinFrameBackgroundColor
    SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "LiveBadukMemoSkinContentTextColor", LiveBadukMemoSkinContentTextColor
    SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "LiveBadukMemoSkinFrameTexture", LiveBadukMemoSkinFrameTexture
    SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "LiveBadukMemoSkinFrameBackground", LiveBadukMemoSkinFrameBackground
    SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "LiveBadukMemoSkinLabelFontSize", LiveBadukMemoSkinLabelFontSize
    SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "LiveBadukMemoSkinLabelFontBold", LiveBadukMemoSkinLabelFontBold
    SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "LiveBadukMemoSkinEnableLabelFontSize", LiveBadukMemoSkinEnableLabelFontSize
    
    Dim i%, ThemeFound As Boolean
    ThemeFound = False
    For i = 1 To cbTheme.ListCount - 1
        If cbTheme.List(i) = ThemeName Then
            ThemeFound = True
            cbTheme.ListIndex = i
        End If
    Next i
    If Not ThemeFound Then
        AddItemToComboBox cbTheme, ThemeName
        cbTheme.ListIndex = cbTheme.ListCount - 1
    End If
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
    IBSSubclass_UnsubclassIt
End Sub

Private Function IBSSubclass_MsgResponse(ByVal hWnd As Long, ByVal uMsg As Long) As EMsgResponse
    IBSSubclass_MsgResponse = emrConsume
End Function

Private Sub IBSSubclass_UnsubclassIt()
    DetachMessage Me, Me.hWnd, WM_SETTINGCHANGE
    DetachMessage Me, Me.hWnd, WM_THEMECHANGED
End Sub

Private Function IBSSubclass_WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByRef wParam As Long, ByRef lParam As Long, ByRef bConsume As Boolean) As Long
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
    
    IBSSubclass_WindowProc = CallOldWindowProc(hWnd, uMsg, wParam, lParam)
End Function

Private Sub lblFillColorSelect_Click()
    Dim Color As OLE_COLOR
    Color = ShowColorDialog(Me.hWnd, True, pgPatternColor.BackColor, True)
    If Color = -1 Then Exit Sub
    pgPatternColor.BackColor = Color
    cmdApply.Enabled = -1
    pgPatternPreview.FillColor = pgPatternColor.BackColor
    PatternChanged = True
    RedrawPreview
End Sub

Private Sub lvBackgrounds_Click()
    On Error GoTo nopicture
    Dim BackgroundPath$
    BackgroundPath = lvBackgroundFiles.Path & IIf(EndsWith(lvBackgroundFiles.Path, "\"), "", "\") & lvBackgrounds.List(lvBackgrounds.ListIndex)
    If lvBackgrounds.ListIndex = 0 Then
nopicture:
        Set imgPreview.Picture = Nothing
    Else
        Set imgPreview.Picture = LoadPictureFromFile(BackgroundPath)
    End If
    frmOptions.cmdSample.Refresh
    ChangedBackgroundPath = BackgroundPath
    If Loaded Then
        cmdApply.Enabled = -1
        ImageChanged = True
        RedrawPreview
    End If
End Sub

Private Sub lvPatterns_Click()
    pgPatternPreview.FillStyle = lvPatterns.ListIndex + 1
    If Loaded Then
        cmdApply.Enabled = -1
        PatternChanged = True
        RedrawPreview
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

Private Sub Form_Load()
    InitForm Me
    
#If HIDEYTDL Then
    If LaunchFromMemory Then tsTabStrip.Tabs.Remove 5
#End If

    Set PreviewControls(0) = Label11
    Set PreviewControls(1) = TextBoxW1
    Set PreviewControls(2) = FrameW5
    Set PreviewControls(3) = CheckBoxW1
    Set PreviewControls(4) = cmdSample
    
    IntervalValues(0) = 0.01
    IntervalValues(1) = 0.05
    IntervalValues(2) = 0.1
    IntervalValues(3) = 0.3
    IntervalValues(4) = 0.5
    IntervalValues(5) = 1#
    IntervalValues(6) = 3#
    IntervalValues(7) = 5#
    
    Set cmdAdvancedSkin.ImageList = frmMain.imgWrench
    Set cmdAdvancedProgressSkin.ImageList = frmMain.imgWrench
    
    RemoveVisualStyles txtSampleClassic.hWnd
    
    InitPropertySheetDimensions Me, tsTabStrip, pbPanel, OKButton, CancelButton, cmdApply
    
    AddItemToComboBox cbFont, "(" & t("기본값", "default") & ")"
    If t(1, 2) = 2 Then
        If FontExists("Tahoma") Then AddItemToComboBox cbFont, "Tahoma"
        If FontExists("Segoe UI") Then AddItemToComboBox cbFont, "Segoe UI"
    Else
        If FontExists("굴림") Then AddItemToComboBox cbFont, "굴림"
        If FontExists("돋움") Then AddItemToComboBox cbFont, "돋움"
        If FontExists("바탕") Then AddItemToComboBox cbFont, "바탕"
        If FontExists("궁서") Then AddItemToComboBox cbFont, "궁서"
        If FontExists("맑은 고딕") Then AddItemToComboBox cbFont, "맑은 고딕"
    End If
    
    If WinVer < 6.2 And IsDWMEnabled() Then
        AddItemToComboBox cbFrameSkin, "Windows Aero"
    Else
        AddItemToComboBox cbFrameSkin, t("시스템 스타일", "System style")
    End If
    If IsDWMEnabled() Then
        If WinVer < 6.2 Or LCase(GetFilename(GetKeyValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\ThemeManager", "DllName", "%SystemRoot%\resources\Themes\Aero\aero.msstyles"))) = "aero.msstyles" Then
            AddItemToComboBox cbFrameSkin, "Windows " & IIf(WinVer < 6.1, "Vista", "7") & " " & t("베이직", "Basic")
        Else
            AddItemToComboBox cbFrameSkin, t("시스템 (DWM 없음)", "System (No DWM)")
        End If
    End If
    AddItemToComboBox cbFrameSkin, t("고전 스타일", "Classic style")
    
    pbBackground.Enabled = False
    SetPreviewPosition
    
    imgPreview.Top = 0
    imgPreview.Left = 0
    
    DrawTabBackground
    
    AddItemToComboBox cbSkin, t("시스템 스타일", "System style")
    AddItemToComboBox cbSkin, t("라이브바둑 쪽지", "LiveBaduk memo")
    AddItemToComboBox cbSkin, t("디스크키퍼 2011", "Diskeeper 2011")
    AddItemToComboBox cbSkin, t("광택 파랑", "Glossy blue")
    AddItemToComboBox cbSkin, t("광택 초록", "Glossy green")
    
    AddItemToComboBox cbProgressSkin, t("(없음)", "(None)")
    AddItemToComboBox cbProgressSkin, t("기본 스킨", "Default skin")
    AddItemToComboBox cbProgressSkin, t("라이브바둑 쪽지", "LiveBaduk memo")
    
    AddItemToComboBox cbLanguage, t("자동", "Auto")
    AddItemToComboBox cbLanguage, "한국어"
    AddItemToComboBox cbLanguage, "English"
    
    AddItemToComboBox cbWhenExist, t("건너뛰기", "Skip")
    AddItemToComboBox cbWhenExist, t("덮어쓰기", "Overwrite")
    AddItemToComboBox cbWhenExist, t("자동 이름 변경", "Auto Rename")
    
    AddItemToComboBox lvPatterns, t("(없음)", "(None)")
    AddItemToComboBox lvPatterns, t("수평선", "Horizontal lines")
    AddItemToComboBox lvPatterns, t("수직선", "Vertical lines")
    AddItemToComboBox lvPatterns, t("하향 대각선", "NW-SE lines")
    AddItemToComboBox lvPatterns, t("상향 대각선", "NE-SW lines")
    AddItemToComboBox lvPatterns, t("교차", "Grid")
    AddItemToComboBox lvPatterns, t("대각선 교차", "45 degrees grid")
    
    AddItemToComboBox cbImagePosition, t("늘이기", "Stretch")
    AddItemToComboBox cbImagePosition, t("높이에 맞춤", "Fit to height")
    AddItemToComboBox cbImagePosition, t("너비에 맞춤", "Fit to width")
    AddItemToComboBox cbImagePosition, t("원본 크기 유지", "True size")
    AddItemToComboBox cbImagePosition, t("높이 맞춤(가운데)", "Fit to height (centered)")
    AddItemToComboBox cbImagePosition, t("너비 맞춤(가운데)", "Fit to width (centered)")
    AddItemToComboBox cbImagePosition, t("가운데", "True size (centered)")
    AddItemToComboBox cbImagePosition, t("바둑판식", "Tile")
    
    LoadSettings
    
    tr chkCenter, "&Centered"
    tr tsTabStrip.Tabs(1), "General"
    tr tsTabStrip.Tabs(2), "Connection"
    tr tsTabStrip.Tabs(3), "Appearance"
    tr tsTabStrip.Tabs(4), "Sounds"
#If HIDEYTDL Then
    If Not LaunchFromMemory Then tr tsTabStrip.Tabs(5), "Advanced"
#End If
    tr Label5, "P&rogress:"
    tr Frame1, "Colors"
    tr Frame4, "Wallpaper"
    tr Label10, "&Window:"
    tr Frame2, "Download settings"
    tr Frame5, "Interface"
    tr chkRememberURL, "Re&member URL"
    tr optSystemColor, "&System"
    tr optSystemFore, "S&ystem"
    tr Label1, "&Language:"
    tr OKButton, "OK"
    tr CancelButton, "Cancel"
    tr cmdApply, "&Apply"
    tr Me, "Options"
    tr Frame6, "Skin"
    tr chkOpenWhenComplete, "&Open file when complete"
    tr chkOpenDirWhenComplete, "O&pen folder when complete"
    tr chkBeepWhenComplete, "Download &complete"
    tr chkAlwaysResume, "&Always resume"
    tr chkAutoRetry, "A&uto retry on error"
    tr Label3, "If filename alrea&dy exists:"
    tr Label6, "Leave the field blank to use defaults. This option is for advanced users and there is no need to change for normal use."
    tr FrameW2, "Paths"
    tr cmdSample, "Download"
    tr Label2, "Po&sition:"
    tr Label8, "Butt&on:"
    tr fHeaders, "Headers"
    tr chkNoRedirectCheck, "Don't check fo&r redirects"
    tr chkForceGet, "Force GET re&quest on file check"
    tr chkIgnore300, "&Ignore 3XX reponse code"
    tr chkAlwaysOnTop, "Al&ways on top"
    tr chkAeroWindow, "Use Aero &glass window"
    tr chkLazyElapsed, "Elapsed time sin&ce first data receive"
    tr FrameW3, "Server connection"
    tr chkForceOldDialog, "U&se Windows 3.1 dialogs"
    tr chkAutoYtdl, "Automatically use &youtube-dl for supported links"
    tr Label11, "File URL:"
    tr FrameW5, "Download status"
    tr CheckBoxW1, "Open when done"
    tr chkAsterisk, "&Asterisk"
    tr chkExclamation, "&Exclamation"
    tr chkError, "E&rror"
    tr chkQuestion, "&Question"
    tr chkAllowDuplicates, "Allow dupl&icates in queue"
    tr Label13, "&Font:"
    tr Label14, "Ma&x. number of threads:"
    tr Label15, "(restart required)"
    tr Label16, Label15.Caption
    tr FrameW6, "Download notifications"
    tr FrameW9, "Message boxes"
    tr Label17, "Set the headers when requesting to the server on download."
    tr Label18, "T&hread scroll:"
    tr optLinePerScroll, "Per li&ne"
    tr optScreenPerScroll, "Pe&r screen"
    tr Label19, "Thread request i&nterval:"
    'tr cmdAdvancedSkin, "Ad&vanced..."
    tr chkBackColorMainOnly, "Main &only"
    tr chkForeColorMainOnly, "Mai&n only"
    tr Label21, "&Patterns:"
    tr Label22, "W&allpaper:"
    tr FrameW1, "Theme"
    tr Label20, "&Theme:"
    tr cmdSaveTheme, "S&ave..."
    tr cmdDeleteTheme, "&Delete"
    tr Label9, "&Background:"
    tr Label12, "T&ext:"
    tr chkUseServerModified, "Use server's mo&dified date"
    tr cmdHeaders, "&Headers..."
    
    If LaunchFromMemory Then
        Label4.Enabled = 0
        txtNodePath.Enabled = 0
    End If
    
    AttachMessage Me, Me.hWnd, WM_SETTINGCHANGE
    AttachMessage Me, Me.hWnd, WM_THEMECHANGED
    
    imgDesktop.Width = pbPreview.Width
    imgDesktop.Height = pbPreview.Height
    imgDesktop.Top = 0
    imgDesktop.Left = 0
    
'    Dim WallpaperPath$, ActiveDesktopWallpaperPath$
'    WallpaperPath = GetKeyValue(HKEY_CURRENT_USER, "Control Panel\Desktop", "Wallpaper")
'    If WinVer < 6# Then
'        ActiveDesktopWallpaperPath = GetKeyValue(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Desktop\General", "Wallpaper", WallpaperPath)
'    Else
'        ActiveDesktopWallpaperPath = WallpaperPath
'    End If
'
'    If Left$(WallpaperPath, 1) = """" And Right$(WallpaperPath, 1) = """" Then WallpaperPath = Mid$(WallpaperPath, 2, Len(WallpaperPath) - 2)
'    If Left$(ActiveDesktopWallpaperPath, 1) = """" And Right$(ActiveDesktopWallpaperPath, 1) = """" Then ActiveDesktopWallpaperPath = Mid$(ActiveDesktopWallpaperPath, 2, Len(ActiveDesktopWallpaperPath) - 2)
'
'    On Error GoTo activefail
'    If Right$(LCase(ActiveDesktopWallpaperPath), 4) = ".png" Then
'        Set imgDesktop.Picture = LoadPngFromFile(ActiveDesktopWallpaperPath)
'    Else
'        imgDesktop.Picture = LoadPicture(ActiveDesktopWallpaperPath)
'    End If
'    GoTo nextcode
'
'activefail:
'    On Error GoTo nextcode
'    If Right$(LCase(WallpaperPath), 4) = ".png" Then
'        Set imgDesktop.Picture = LoadPngFromFile(WallpaperPath)
'    Else
'        imgDesktop.Picture = LoadPicture(WallpaperPath)
'    End If
'
'nextcode:
    Set cmdSample.ImageList = frmMain.imgDownload
    
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
    Label14.Top = Label14.Top - chkAutoYtdl.Height
    txtMaxThreadCount.Top = txtMaxThreadCount.Top - chkAutoYtdl.Height
    Label15.Top = Label15.Top - chkAutoYtdl.Height
#End If
    
    Loaded = True
End Sub

Private Sub LoadSettings()
    'SetWindowPos Me.hWnd, IIf(MainFormOnTop, hWnd_TOPMOST, hWnd_NOTOPMOST), 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    On Error Resume Next
    
    chkNoRedirectCheck.Value = GetSetting("DownloadBooster", "Options", "NoRedirectCheck", 0)
    chkForceGet.Value = GetSetting("DownloadBooster", "Options", "ForceGet", 1)
    chkIgnore300.Value = GetSetting("DownloadBooster", "Options", "Ignore300", 0)
    chkAlwaysOnTop.Value = -(MainFormOnTop)
    chkLazyElapsed.Value = GetSetting("DownloadBooster", "Options", "LazyElapsed", 0)
    chkForceOldDialog.Value = GetSetting("DownloadBooster", "Options", "ForceWin31Dialog", 0)
    chkRememberURL.Value = GetSetting("DownloadBooster", "Options", "RememberURL", 1)
    chkAutoYtdl.Value = GetSetting("DownloadBooster", "Options", "AutoDetectYtdlURL", 1)
    chkAllowDuplicates.Value = GetSetting("DownloadBooster", "Options", "AllowDuplicatesInQueue", 0)
    txtMaxThreadCount.Value = GetSetting("DownloadBooster", "Options", "MaxThreadCount", 25)
    optLinePerScroll.Value = True
    optScreenPerScroll.Value = (GetSetting("DownloadBooster", "Options", "ScrollOneScreen", 0) <> 0)
    chkBackColorMainOnly.Value = GetSetting("DownloadBooster", "Options", "BackColorMainOnly", 0)
    chkForeColorMainOnly.Value = GetSetting("DownloadBooster", "Options", "ForeColorMainOnly", 0)
    chkUseServerModified.Value = GetSetting("DownloadBooster", "Options", "UseServerModifiedDate", 1)
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
    
    DisableVisualStyle = CByte(GetSetting("DownloadBooster", "Options", "DisableVisualStyle", 0))
    RoundClassicButtons = CByte(GetSetting("DownloadBooster", "Options", "RoundClassicButtons", 0))
    
    LiveBadukMemoSkinShadowColor = CLng(GetSetting("DownloadBooster", "Options", "LiveBadukMemoSkinShadowColor", 16777215))
    LiveBadukMemoSkinFrameColor = CLng(GetSetting("DownloadBooster", "Options", "LiveBadukMemoSkinFrameColor", 11194780))
    LiveBadukMemoSkinFrameType = LCase(GetSetting("DownloadBooster", "Options", "LiveBadukMemoSkinFrameType", "solidcolor"))
    LiveBadukMemoSkinTextColor = CLng(GetSetting("DownloadBooster", "Options", "LiveBadukMemoSkinTextColor", 2902560))
    LiveBadukMemoSkinEnableShadow = CByte(GetSetting("DownloadBooster", "Options", "LiveBadukMemoSkinEnableShadow", 1))
    LiveBadukMemoSkinEnableTextColor = CByte(GetSetting("DownloadBooster", "Options", "LiveBadukMemoSkinEnableTextColor", 1))
    LiveBadukMemoSkinEnableBorder = CByte(GetSetting("DownloadBooster", "Options", "LiveBadukMemoSkinEnableBorder", 1))
    LiveBadukMemoSkinFrameBackgroundType = LCase(GetSetting("DownloadBooster", "Options", "LiveBadukMemoSkinFrameBackgroundType", "transparent"))
    LiveBadukMemoSkinFrameBackgroundColor = CLng(GetSetting("DownloadBooster", "Options", "LiveBadukMemoSkinFrameBackgroundColor", 16777215))
    LiveBadukMemoSkinContentTextColor = CLng(GetSetting("DownloadBooster", "Options", "LiveBadukMemoSkinContentTextColor", 0))
    LiveBadukMemoSkinFrameTexture = GetSetting("DownloadBooster", "Options", "LiveBadukMemoSkinFrameTexture", "")
    LiveBadukMemoSkinFrameBackground = GetSetting("DownloadBooster", "Options", "LiveBadukMemoSkinFrameBackground", "")
    LiveBadukMemoSkinLabelFontSize = CInt(GetSetting("DownloadBooster", "Options", "LiveBadukMemoSkinLabelFontSize", 10))
    LiveBadukMemoSkinLabelFontBold = CByte(GetSetting("DownloadBooster", "Options", "LiveBadukMemoSkinLabelFontBold", 1))
    LiveBadukMemoSkinEnableLabelFontSize = CByte(GetSetting("DownloadBooster", "Options", "LiveBadukMemoSkinEnableLabelFontSize", 1))
    
    LoadTheme
    
    chkOpenWhenComplete.Value = frmMain.chkOpenAfterComplete.Value
    chkOpenDirWhenComplete.Value = frmMain.chkOpenFolder.Value
    chkAlwaysResume.Value = frmMain.chkContinueDownload.Value
    chkAutoRetry.Value = frmMain.chkAutoRetry.Value
    
    Dim LangSet As String
    LangSet = GetSetting("DownloadBooster", "Options", "Language", "0")
    If LangSet = "0" Then
        cbLanguage.ListIndex = 0
    ElseIf LangSet = "1042" Then
        cbLanguage.ListIndex = 1
    Else
        cbLanguage.ListIndex = 2
    End If
    cbWhenExist.ListIndex = GetSetting("DownloadBooster", "Options", "WhenFileExists", 0)
    pgPatternPreview.Width = pbBackground.Width
    pgPatternPreview.Height = pbBackground.Height
    
    DoLoadTheme = False
    
    ClearComboBox cbTheme
    AddItemToComboBox cbTheme, t("수정된 테마", "Modified theme")
    cbTheme.ListIndex = 0
    On Error Resume Next
    Dim ThemeList() As String
    ThemeList = GetSubkeys(HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\DownloadBooster\Options\Themes")
    Dim CurrentTheme$
    CurrentTheme = GetSetting("DownloadBooster", "Options", "Theme", "")
    For i = LBound(ThemeList) To UBound(ThemeList)
        AddItemToComboBox cbTheme, ThemeList(i)
        If ThemeList(i) = CurrentTheme Then cbTheme.ListIndex = cbTheme.ListCount - 1
    Next i
    
    DoLoadTheme = True
    
    txtNodePath.Text = GetSetting("DownloadBooster", "Options", "NodePath", "")
    txtYtdlPath.Text = GetSetting("DownloadBooster", "Options", "YtdlPath", "")
    
    chkAsterisk.Value = GetSetting("DownloadBooster", "Options", "EnableAsteriskSound", 1)
    chkExclamation.Value = GetSetting("DownloadBooster", "Options", "EnableExclamationSound", 1)
    chkError.Value = GetSetting("DownloadBooster", "Options", "EnableErrorSound", 1)
    chkQuestion.Value = GetSetting("DownloadBooster", "Options", "EnableQuestionSound", 1)
    txtAsterisk.Text = GetSetting("DownloadBooster", "Options", "AsteriskSound", "")
    txtExclamation.Text = GetSetting("DownloadBooster", "Options", "ExclamationSound", "")
    txtError.Text = GetSetting("DownloadBooster", "Options", "ErrorSound", "")
    txtQuestion.Text = GetSetting("DownloadBooster", "Options", "QuestionSound", "")
    
    tsTabStrip.Tabs(1).Selected = True
    ResetChanged
    cmdApply.Enabled = False
End Sub

Sub LoadBackgroundList(Optional ByVal OnLoad As Boolean = False)
    Dim BackgroundImagePath$
    Dim BackgroundImageEnabled As Boolean
    BackgroundImagePath = ChangedBackgroundPath
    ClearComboBox lvBackgrounds
    AddItemToComboBox lvBackgrounds, t("(없음)", "(None)")
    lvBackgrounds.ListIndex = 0
    BackgroundImageEnabled = (GetSetting("DownloadBooster", "Options", "UseBackgroundImage", 0) <> 0)
    If FileExists(BackgroundImagePath) Then
        Dim li&
        lvBackgroundFiles.Path = GetParentFolderName(BackgroundImagePath)
        For li = 0 To lvBackgroundFiles.ListCount - 1
            AddItemToComboBox lvBackgrounds, lvBackgroundFiles.List(li)
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
        NextTabPage tsTabStrip, IsKeyPressed(gksKeyboardShift)
    End If
End Sub

Private Sub DrawTabBackground(Optional Force As Boolean = False)
    On Error Resume Next
    Dim ctrl As Control
    Dim i As Byte
    
    For i = pbPanel.LBound To pbPanel.UBound
        tsTabStrip.DrawBackground pbPanel(i).hWnd, pbPanel(i).hDC
    Next i
    For Each ctrl In Me.Controls
        If TypeOf ctrl Is FrameW Or TypeOf ctrl Is CheckBoxW Or TypeOf ctrl Is CommandButtonW Then
            ctrl.Refresh
        ElseIf TypeOf ctrl Is Slider Then
            ctrl.Refresh
            ctrl.VisualStyles = Not ctrl.VisualStyles
            ctrl.VisualStyles = Not ctrl.VisualStyles
        End If
    Next ctrl
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
    chkBackColorMainOnly.Enabled = False
End Sub

Private Sub optSystemFore_Click()
    If Loaded Then
        cmdApply.Enabled = -1
        ColorChanged = True
    End If
    Label11.ForeColor = &H80000012
    CheckBoxW1.VisualStyles = (DisableVisualStyle = 0)
    FrameW5.VisualStyles = (DisableVisualStyle = 0)
    CheckBoxW1.ForeColor = &H80000012
    FrameW5.ForeColor = &H80000012
    chkForeColorMainOnly.Enabled = False
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
    chkBackColorMainOnly.Enabled = True
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
    chkForeColorMainOnly.Enabled = True
End Sub

Private Sub tsTabStrip_TabClick(TabItem As TbsTab)
    On Error Resume Next
    
    Static i As Byte, Show As Boolean
    For i = 1 To pbPanel.Count
        Show = (i = TabItem.Index)
        pbPanel(i).Visible = Show
        pbPanel(i).Enabled = Show
    Next i
    
    If TabItem.Index = 3 Then
        DoEvents
        RedrawPreview
    End If
End Sub

Private Sub SetPreviewPosition()
    Dim Left%, Top%, Width%, Height%
    Left = 20
    Top = 2
    Width = 3795
    Height = 1320
    pbBackground.BorderStyle = 0
    SetWindowLong pbBackground.hWnd, GWL_STYLE, GetWindowLong(pbBackground.hWnd, GWL_STYLE) Or WS_BORDER Or WS_OVERLAPPED Or WS_CAPTION Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_SYSMENU
    SetWindowText pbBackground.hWnd, App.Title
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
    If Loaded Then cmdApply.Enabled = -1
End Sub

Private Sub txtNodePath_Change()
    If Loaded Then cmdApply.Enabled = -1
End Sub

Private Sub txtQuestion_Change()
    If Loaded Then cmdApply.Enabled = -1
End Sub

Private Sub txtYtdlPath_Change()
    If Loaded Then cmdApply.Enabled = -1
End Sub
