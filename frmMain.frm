VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "다운로드 부스터"
   ClientHeight    =   7740
   ClientLeft      =   60
   ClientTop       =   480
   ClientWidth     =   11115
   BeginProperty Font 
      Name            =   "굴림"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "frmMain"
   MaxButton       =   0   'False
   ScaleHeight     =   7740
   ScaleWidth      =   11115
   StartUpPosition =   3  'Windows 기본값
   Begin prjDownloadBooster.TygemButton tygStop 
      Height          =   330
      Left            =   7320
      TabIndex        =   133
      Top             =   4755
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      Enabled         =   0   'False
      Caption         =   "중지"
   End
   Begin prjDownloadBooster.TygemButton tygGo 
      Height          =   330
      Left            =   7320
      TabIndex        =   134
      Top             =   4755
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      Caption         =   "다운로드"
   End
   Begin prjDownloadBooster.TygemButton tygOpen 
      Height          =   330
      Left            =   7320
      TabIndex        =   132
      Top             =   3960
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      Enabled         =   0   'False
      Caption         =   "열기"
   End
   Begin prjDownloadBooster.TygemButton tygStopbatch 
      Height          =   375
      Left            =   7560
      TabIndex        =   131
      Top             =   6960
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Enabled         =   0   'False
      Caption         =   "중지"
   End
   Begin prjDownloadBooster.TygemButton tygStartBatch 
      Height          =   375
      Left            =   5880
      TabIndex        =   130
      Top             =   6960
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Enabled         =   0   'False
      Caption         =   "시작"
   End
   Begin prjDownloadBooster.TygemButton tygDelete 
      Height          =   375
      Left            =   4200
      TabIndex        =   129
      Top             =   6960
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Enabled         =   0   'False
      Caption         =   "제거"
   End
   Begin prjDownloadBooster.TygemButton tygOpenBatch 
      Height          =   375
      Left            =   240
      TabIndex        =   128
      Top             =   6960
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      Enabled         =   0   'False
      Caption         =   "열기"
   End
   Begin prjDownloadBooster.TygemButton tygAdd 
      Height          =   375
      Left            =   2520
      TabIndex        =   127
      Top             =   6960
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "추가..."
   End
   Begin prjDownloadBooster.ProgressBar pbTotalProgressMarquee 
      Height          =   255
      Left            =   360
      Top             =   1560
      Visible         =   0   'False
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   450
      Step            =   10
      MarqueeSpeed    =   35
      Scrolling       =   2
   End
   Begin prjDownloadBooster.ProgressBar pbTotalProgress 
      Height          =   255
      Left            =   360
      Top             =   1560
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   450
      Step            =   10
      MarqueeSpeed    =   35
   End
   Begin prjDownloadBooster.TygemButton tygBatch 
      Height          =   330
      Left            =   7320
      TabIndex        =   124
      Top             =   5565
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      Caption         =   "일괄 처리 >>"
   End
   Begin prjDownloadBooster.TygemButton tygAddToQueue 
      Height          =   330
      Left            =   7320
      TabIndex        =   123
      Top             =   5130
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      Caption         =   "목록에 추가"
   End
   Begin prjDownloadBooster.TygemButton tygOpenFolder 
      Height          =   330
      Left            =   7320
      TabIndex        =   122
      Top             =   4320
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      Caption         =   "폴더 열기"
   End
   Begin prjDownloadBooster.TygemButton tygOptions 
      Height          =   300
      Left            =   6960
      TabIndex        =   121
      Top             =   3090
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   529
      Caption         =   "추가 옵션..."
   End
   Begin prjDownloadBooster.TygemButton tygAbout 
      Height          =   300
      Left            =   6960
      TabIndex        =   120
      Top             =   3435
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   529
      Caption         =   "프로그램 정보..."
   End
   Begin prjDownloadBooster.TygemButton tygBrowse 
      Height          =   330
      Left            =   7440
      TabIndex        =   119
      Top             =   435
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      Caption         =   "찾아보기..."
   End
   Begin prjDownloadBooster.TygemButton tygReset 
      Height          =   330
      Left            =   7440
      TabIndex        =   118
      Top             =   90
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      Caption         =   "초기화"
   End
   Begin prjDownloadBooster.ImageList imgWrench 
      Left            =   9240
      Top             =   6600
      _ExtentX        =   1005
      _ExtentY        =   1005
      ImageWidth      =   16
      ImageHeight     =   16
      ColorDepth      =   4
      InitListImages  =   "frmMain.frx":030A
   End
   Begin prjDownloadBooster.ImageList imgErase 
      Left            =   9840
      Top             =   6600
      _ExtentX        =   1005
      _ExtentY        =   1005
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      InitListImages  =   "frmMain.frx":07F2
   End
   Begin prjDownloadBooster.StatusBar sbStatusBar 
      Align           =   2  '아래 맞춤
      Height          =   330
      Left            =   0
      Top             =   7410
      Width           =   11115
      _ExtentX        =   19606
      _ExtentY        =   582
      InitPanels      =   "frmMain.frx":0BDA
   End
   Begin prjDownloadBooster.ListView lvBatchFiles 
      Height          =   870
      Left            =   240
      TabIndex        =   19
      Top             =   6030
      Visible         =   0   'False
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   1535
      View            =   3
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      LabelEdit       =   2
      Checkboxes      =   -1  'True
      HideSelection   =   0   'False
      ClickableColumnHeaders=   0   'False
      AutoSelectFirstItem=   0   'False
   End
   Begin prjDownloadBooster.CommandButtonW cmdAbout 
      Height          =   300
      Left            =   7080
      TabIndex        =   113
      Top             =   3435
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   529
      Caption         =   "프로그램 정보(&U)..."
      Transparent     =   -1  'True
   End
   Begin prjDownloadBooster.CommandButtonW cmdOptions 
      Height          =   300
      Left            =   7080
      TabIndex        =   112
      Top             =   3090
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   529
      ImageList       =   "imgWrench"
      Caption         =   "추가 옵션(&I)..."
      Transparent     =   -1  'True
   End
   Begin prjDownloadBooster.CheckBoxW chkAutoRetry 
      Height          =   255
      Left            =   6840
      TabIndex        =   60
      Top             =   2805
      Width           =   2210
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "오류 시 자동 재시도(&G)"
   End
   Begin prjDownloadBooster.CommandButtonW cmdStop 
      Height          =   330
      Left            =   7320
      TabIndex        =   17
      Top             =   4760
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      Enabled         =   0   'False
      ImageList       =   "imgStopRed"
      Caption         =   "중지(&P) "
      Transparent     =   -1  'True
   End
   Begin prjDownloadBooster.CheckBoxW chkContinueDownload 
      Height          =   255
      Left            =   6840
      TabIndex        =   59
      Top             =   2580
      Width           =   1935
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "항상 이어받기(&J)"
   End
   Begin prjDownloadBooster.CommandButtonW cmdOpenBatch 
      Height          =   375
      Left            =   240
      TabIndex        =   20
      Top             =   6960
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Enabled         =   0   'False
      ImageList       =   "imgOpenFile"
      Caption         =   "열기(&W) "
      Transparent     =   -1  'True
   End
   Begin prjDownloadBooster.CommandButtonW cmdDelete 
      Height          =   375
      Left            =   4200
      TabIndex        =   22
      Top             =   6960
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Enabled         =   0   'False
      ImageList       =   "imgMinus"
      Caption         =   "제거(&V) "
      Transparent     =   -1  'True
   End
   Begin prjDownloadBooster.ImageList imgDropdownReverse 
      Left            =   9840
      Top             =   5880
      _ExtentX        =   1005
      _ExtentY        =   1005
      ImageWidth      =   13
      ImageHeight     =   5
      MaskColor       =   16711935
      InitListImages  =   "frmMain.frx":0EDE
   End
   Begin prjDownloadBooster.CommandButtonW cmdOpenDropdown 
      Height          =   375
      Left            =   1800
      TabIndex        =   58
      Top             =   6960
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      Enabled         =   0   'False
      ImageList       =   "imgDropdown"
      ImageListAlignment=   4
      Transparent     =   -1  'True
   End
   Begin prjDownloadBooster.ImageList imgDropdown 
      Left            =   9840
      Top             =   5160
      _ExtentX        =   1005
      _ExtentY        =   1005
      ImageWidth      =   13
      ImageHeight     =   5
      MaskColor       =   16711935
      InitListImages  =   "frmMain.frx":15CE
   End
   Begin prjDownloadBooster.CommandButtonW cmdDeleteDropdown 
      Height          =   375
      Left            =   5520
      TabIndex        =   57
      Top             =   6960
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      Enabled         =   0   'False
      ImageList       =   "imgDropdown"
      ImageListAlignment=   4
      Transparent     =   -1  'True
   End
   Begin prjDownloadBooster.ImageList imgPlusYellow 
      Left            =   9840
      Top             =   4440
      _ExtentX        =   1005
      _ExtentY        =   1005
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      InitListImages  =   "frmMain.frx":1CBE
   End
   Begin prjDownloadBooster.CommandButtonW cmdAddToQueue 
      Height          =   330
      Left            =   7320
      TabIndex        =   56
      Top             =   5130
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      ImageList       =   "imgPlusYellow"
      Caption         =   "목록에 추가(&Q)"
      Transparent     =   -1  'True
   End
   Begin prjDownloadBooster.CommandButtonW cmdStartBatch 
      Height          =   375
      Left            =   5880
      TabIndex        =   23
      Top             =   6960
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Enabled         =   0   'False
      ImageList       =   "imgPlay"
      Caption         =   "시작(&S) "
      Transparent     =   -1  'True
   End
   Begin prjDownloadBooster.ImageList imgStopRed 
      Left            =   9840
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      InitListImages  =   "frmMain.frx":2EC6
   End
   Begin prjDownloadBooster.ImageList imgPlay 
      Left            =   9840
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      InitListImages  =   "frmMain.frx":40CE
   End
   Begin prjDownloadBooster.ImageList imgDownload 
      Left            =   9840
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      InitListImages  =   "frmMain.frx":52D6
   End
   Begin prjDownloadBooster.ImageList imgMinus 
      Left            =   9840
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      InitListImages  =   "frmMain.frx":64DE
   End
   Begin prjDownloadBooster.ImageList imgOpenFile 
      Left            =   9840
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      InitListImages  =   "frmMain.frx":76E6
   End
   Begin prjDownloadBooster.ImageList imgOpenFolder 
      Left            =   9840
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      InitListImages  =   "frmMain.frx":88EE
   End
   Begin prjDownloadBooster.CheckBoxW chkPlaySound 
      Height          =   255
      Left            =   6840
      TabIndex        =   10
      Top             =   2010
      Width           =   2205
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "완료 시 신호음(&M)"
   End
   Begin prjDownloadBooster.FrameW fTabThreads 
      Height          =   165
      Left            =   1545
      TabIndex        =   49
      Top             =   2070
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   291
      Caption         =   " 스레드 "
      Alignment       =   2
      Transparent     =   -1  'True
   End
   Begin prjDownloadBooster.FrameW fTabDownload 
      Height          =   165
      Left            =   660
      TabIndex        =   48
      Top             =   2070
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   291
      Caption         =   " 요약  "
      Alignment       =   2
      Transparent     =   -1  'True
   End
   Begin prjDownloadBooster.OptionButtonW optTabThreads2 
      Height          =   195
      Left            =   1320
      TabIndex        =   47
      Top             =   2055
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   344
      Value           =   -1  'True
      Transparent     =   -1  'True
   End
   Begin prjDownloadBooster.OptionButtonW optTabDownload2 
      Height          =   195
      Left            =   435
      TabIndex        =   46
      Top             =   2055
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   344
      Transparent     =   -1  'True
   End
   Begin prjDownloadBooster.FrameW fDownloadInfo 
      Height          =   2775
      Left            =   1440
      TabIndex        =   30
      Top             =   2520
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   4895
      BorderStyle     =   0
      Caption         =   " "
      Transparent     =   -1  'True
      Begin VB.Label lblRemaining 
         BackStyle       =   0  '투명
         Caption         =   "-"
         Height          =   255
         Left            =   1320
         TabIndex        =   115
         Top             =   2520
         Width           =   4335
      End
      Begin VB.Label Label10 
         BackStyle       =   0  '투명
         Caption         =   "남은 시간:"
         Height          =   255
         Left            =   0
         TabIndex        =   114
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label8 
         BackStyle       =   0  '투명
         Caption         =   "파일 이름:"
         Height          =   255
         Left            =   0
         TabIndex        =   55
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label lblFilename 
         BackStyle       =   0  '투명
         Caption         =   "-"
         Height          =   180
         Left            =   1320
         TabIndex        =   54
         Top             =   0
         Width           =   4335
      End
      Begin VB.Label lblTotalSizeThread 
         BackStyle       =   0  '투명
         Caption         =   "-"
         Height          =   255
         Left            =   1320
         TabIndex        =   53
         Top             =   2160
         Width           =   4335
      End
      Begin VB.Label Label7 
         BackStyle       =   0  '투명
         Caption         =   "스레드당 크기:"
         Height          =   255
         Left            =   0
         TabIndex        =   52
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label lblThreadCount2 
         BackStyle       =   0  '투명
         Caption         =   "-"
         Height          =   255
         Left            =   1320
         TabIndex        =   51
         Top             =   1800
         Width           =   4335
      End
      Begin VB.Label Label6 
         BackStyle       =   0  '투명
         Caption         =   "스레드 수:"
         Height          =   255
         Left            =   0
         TabIndex        =   50
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label5 
         BackStyle       =   0  '투명
         Caption         =   "속도:"
         Height          =   255
         Left            =   0
         TabIndex        =   44
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label lblSpeed 
         BackStyle       =   0  '투명
         Caption         =   "-"
         Height          =   255
         Left            =   1320
         TabIndex        =   43
         Top             =   1440
         Width           =   4335
      End
      Begin VB.Label lblElapsed 
         BackStyle       =   0  '투명
         Caption         =   "-"
         Height          =   255
         Left            =   1320
         TabIndex        =   36
         Top             =   1080
         Width           =   4335
      End
      Begin VB.Label Label4 
         BackStyle       =   0  '투명
         Caption         =   "경과 시간:"
         Height          =   255
         Left            =   0
         TabIndex        =   35
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblDownloadedBytes 
         BackStyle       =   0  '투명
         Caption         =   "-"
         Height          =   255
         Left            =   1320
         TabIndex        =   34
         Top             =   720
         Width           =   4335
      End
      Begin VB.Label Label3 
         BackStyle       =   0  '투명
         Caption         =   "받은 크기:"
         Height          =   255
         Left            =   0
         TabIndex        =   33
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblTotalBytes 
         BackStyle       =   0  '투명
         Caption         =   "-"
         Height          =   255
         Left            =   1320
         TabIndex        =   32
         Top             =   360
         Width           =   4335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '투명
         Caption         =   "총 크기:"
         Height          =   255
         Left            =   0
         TabIndex        =   31
         Top             =   360
         Width           =   975
      End
   End
   Begin prjDownloadBooster.FrameW fThreadInfo 
      Height          =   3495
      Left            =   360
      TabIndex        =   15
      Top             =   2310
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   6165
      BorderStyle     =   0
      Caption         =   " 스레드 현황 "
      Transparent     =   -1  'True
      Begin VB.VScrollBar vsProgressScroll 
         Height          =   3495
         Left            =   5760
         Max             =   15
         TabIndex        =   27
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
      Begin prjDownloadBooster.FrameW fDummyScroll 
         Height          =   3495
         Left            =   5760
         TabIndex        =   42
         Top             =   0
         Visible         =   0   'False
         Width           =   255
         _ExtentX        =   0
         _ExtentY        =   0
         BorderStyle     =   0
      End
      Begin VB.PictureBox pbProgressOuterContainer 
         BorderStyle     =   0  '없음
         Height          =   3495
         Left            =   0
         ScaleHeight     =   3495
         ScaleWidth      =   5775
         TabIndex        =   28
         Top             =   0
         Width           =   5775
         Begin VB.PictureBox pbProgressContainer 
            BorderStyle     =   0  '없음
            Height          =   9015
            Left            =   0
            ScaleHeight     =   9015
            ScaleWidth      =   5775
            TabIndex        =   61
            Top             =   0
            Width           =   5775
            Begin prjDownloadBooster.ProgressBar pbProgressMarquee 
               Height          =   255
               Index           =   1
               Left            =   840
               Top             =   0
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   450
               Step            =   10
               MarqueeSpeed    =   35
               Scrolling       =   2
            End
            Begin prjDownloadBooster.ProgressBar pbProgressMarquee 
               Height          =   255
               Index           =   2
               Left            =   840
               Top             =   360
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   450
               Step            =   10
               MarqueeSpeed    =   35
               Scrolling       =   2
            End
            Begin prjDownloadBooster.ProgressBar pbProgressMarquee 
               Height          =   255
               Index           =   3
               Left            =   840
               Top             =   720
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   450
               Step            =   10
               MarqueeSpeed    =   35
               Scrolling       =   2
            End
            Begin prjDownloadBooster.ProgressBar pbProgressMarquee 
               Height          =   255
               Index           =   4
               Left            =   840
               Top             =   1080
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   450
               Step            =   10
               MarqueeSpeed    =   35
               Scrolling       =   2
            End
            Begin prjDownloadBooster.ProgressBar pbProgressMarquee 
               Height          =   255
               Index           =   5
               Left            =   840
               Top             =   1440
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   450
               Step            =   10
               MarqueeSpeed    =   35
               Scrolling       =   2
            End
            Begin prjDownloadBooster.ProgressBar pbProgressMarquee 
               Height          =   255
               Index           =   6
               Left            =   840
               Top             =   1800
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   450
               Step            =   10
               MarqueeSpeed    =   35
               Scrolling       =   2
            End
            Begin prjDownloadBooster.ProgressBar pbProgressMarquee 
               Height          =   255
               Index           =   7
               Left            =   840
               Top             =   2160
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   450
               Step            =   10
               MarqueeSpeed    =   35
               Scrolling       =   2
            End
            Begin prjDownloadBooster.ProgressBar pbProgressMarquee 
               Height          =   255
               Index           =   8
               Left            =   840
               Top             =   2520
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   450
               Step            =   10
               MarqueeSpeed    =   35
               Scrolling       =   2
            End
            Begin prjDownloadBooster.ProgressBar pbProgressMarquee 
               Height          =   255
               Index           =   9
               Left            =   840
               Top             =   2880
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   450
               Step            =   10
               MarqueeSpeed    =   35
               Scrolling       =   2
            End
            Begin prjDownloadBooster.ProgressBar pbProgressMarquee 
               Height          =   255
               Index           =   10
               Left            =   840
               Top             =   3240
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   450
               Step            =   10
               MarqueeSpeed    =   35
               Scrolling       =   2
            End
            Begin prjDownloadBooster.ProgressBar pbProgressMarquee 
               Height          =   255
               Index           =   11
               Left            =   840
               Top             =   3600
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   450
               Step            =   10
               MarqueeSpeed    =   35
               Scrolling       =   2
            End
            Begin prjDownloadBooster.ProgressBar pbProgressMarquee 
               Height          =   255
               Index           =   12
               Left            =   840
               Top             =   3960
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   450
               Step            =   10
               MarqueeSpeed    =   35
               Scrolling       =   2
            End
            Begin prjDownloadBooster.ProgressBar pbProgressMarquee 
               Height          =   255
               Index           =   13
               Left            =   840
               Top             =   4320
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   450
               Step            =   10
               MarqueeSpeed    =   35
               Scrolling       =   2
            End
            Begin prjDownloadBooster.ProgressBar pbProgressMarquee 
               Height          =   255
               Index           =   14
               Left            =   840
               Top             =   4680
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   450
               Step            =   10
               MarqueeSpeed    =   35
               Scrolling       =   2
            End
            Begin prjDownloadBooster.ProgressBar pbProgressMarquee 
               Height          =   255
               Index           =   15
               Left            =   840
               Top             =   5040
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   450
               Step            =   10
               MarqueeSpeed    =   35
               Scrolling       =   2
            End
            Begin prjDownloadBooster.ProgressBar pbProgressMarquee 
               Height          =   255
               Index           =   16
               Left            =   840
               Top             =   5400
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   450
               Step            =   10
               MarqueeSpeed    =   35
               Scrolling       =   2
            End
            Begin prjDownloadBooster.ProgressBar pbProgressMarquee 
               Height          =   255
               Index           =   17
               Left            =   840
               Top             =   5760
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   450
               Step            =   10
               MarqueeSpeed    =   35
               Scrolling       =   2
            End
            Begin prjDownloadBooster.ProgressBar pbProgressMarquee 
               Height          =   255
               Index           =   18
               Left            =   840
               Top             =   6120
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   450
               Step            =   10
               MarqueeSpeed    =   35
               Scrolling       =   2
            End
            Begin prjDownloadBooster.ProgressBar pbProgressMarquee 
               Height          =   255
               Index           =   19
               Left            =   840
               Top             =   6480
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   450
               Step            =   10
               MarqueeSpeed    =   35
               Scrolling       =   2
            End
            Begin prjDownloadBooster.ProgressBar pbProgressMarquee 
               Height          =   255
               Index           =   20
               Left            =   840
               Top             =   6840
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   450
               Step            =   10
               MarqueeSpeed    =   35
               Scrolling       =   2
            End
            Begin prjDownloadBooster.ProgressBar pbProgressMarquee 
               Height          =   255
               Index           =   21
               Left            =   840
               Top             =   7200
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   450
               Step            =   10
               MarqueeSpeed    =   35
               Scrolling       =   2
            End
            Begin prjDownloadBooster.ProgressBar pbProgressMarquee 
               Height          =   255
               Index           =   22
               Left            =   840
               Top             =   7560
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   450
               Step            =   10
               MarqueeSpeed    =   35
               Scrolling       =   2
            End
            Begin prjDownloadBooster.ProgressBar pbProgressMarquee 
               Height          =   255
               Index           =   23
               Left            =   840
               Top             =   7920
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   450
               Step            =   10
               MarqueeSpeed    =   35
               Scrolling       =   2
            End
            Begin prjDownloadBooster.ProgressBar pbProgressMarquee 
               Height          =   255
               Index           =   24
               Left            =   840
               Top             =   8280
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   450
               Step            =   10
               MarqueeSpeed    =   35
               Scrolling       =   2
            End
            Begin prjDownloadBooster.ProgressBar pbProgressMarquee 
               Height          =   255
               Index           =   25
               Left            =   840
               Top             =   8640
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   450
               Step            =   10
               MarqueeSpeed    =   35
               Scrolling       =   2
            End
            Begin prjDownloadBooster.ProgressBar pbProgress 
               Height          =   255
               Index           =   1
               Left            =   840
               Top             =   0
               Visible         =   0   'False
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   450
               Step            =   10
               MarqueeSpeed    =   35
            End
            Begin prjDownloadBooster.ProgressBar pbProgress 
               Height          =   255
               Index           =   2
               Left            =   840
               Top             =   360
               Visible         =   0   'False
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   450
               Step            =   10
            End
            Begin prjDownloadBooster.ProgressBar pbProgress 
               Height          =   255
               Index           =   3
               Left            =   840
               Top             =   720
               Visible         =   0   'False
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   450
               Step            =   10
            End
            Begin prjDownloadBooster.ProgressBar pbProgress 
               Height          =   255
               Index           =   4
               Left            =   840
               Top             =   1080
               Visible         =   0   'False
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   450
               Step            =   10
            End
            Begin prjDownloadBooster.ProgressBar pbProgress 
               Height          =   255
               Index           =   5
               Left            =   840
               Top             =   1440
               Visible         =   0   'False
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   450
               Step            =   10
            End
            Begin prjDownloadBooster.ProgressBar pbProgress 
               Height          =   255
               Index           =   6
               Left            =   840
               Top             =   1800
               Visible         =   0   'False
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   450
               Step            =   10
            End
            Begin prjDownloadBooster.ProgressBar pbProgress 
               Height          =   255
               Index           =   7
               Left            =   840
               Top             =   2160
               Visible         =   0   'False
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   450
               Step            =   10
            End
            Begin prjDownloadBooster.ProgressBar pbProgress 
               Height          =   255
               Index           =   8
               Left            =   840
               Top             =   2520
               Visible         =   0   'False
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   450
               Step            =   10
            End
            Begin prjDownloadBooster.ProgressBar pbProgress 
               Height          =   255
               Index           =   9
               Left            =   840
               Top             =   2880
               Visible         =   0   'False
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   450
               Step            =   10
            End
            Begin prjDownloadBooster.ProgressBar pbProgress 
               Height          =   255
               Index           =   10
               Left            =   840
               Top             =   3240
               Visible         =   0   'False
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   450
               Step            =   10
            End
            Begin prjDownloadBooster.ProgressBar pbProgress 
               Height          =   255
               Index           =   11
               Left            =   840
               Top             =   3600
               Visible         =   0   'False
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   450
               Step            =   10
            End
            Begin prjDownloadBooster.ProgressBar pbProgress 
               Height          =   255
               Index           =   12
               Left            =   840
               Top             =   3960
               Visible         =   0   'False
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   450
               Step            =   10
            End
            Begin prjDownloadBooster.ProgressBar pbProgress 
               Height          =   255
               Index           =   13
               Left            =   840
               Top             =   4320
               Visible         =   0   'False
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   450
               Step            =   10
            End
            Begin prjDownloadBooster.ProgressBar pbProgress 
               Height          =   255
               Index           =   14
               Left            =   840
               Top             =   4680
               Visible         =   0   'False
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   450
               Step            =   10
            End
            Begin prjDownloadBooster.ProgressBar pbProgress 
               Height          =   255
               Index           =   15
               Left            =   840
               Top             =   5040
               Visible         =   0   'False
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   450
               Step            =   10
            End
            Begin prjDownloadBooster.ProgressBar pbProgress 
               Height          =   255
               Index           =   16
               Left            =   840
               Top             =   5400
               Visible         =   0   'False
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   450
               Step            =   10
            End
            Begin prjDownloadBooster.ProgressBar pbProgress 
               Height          =   255
               Index           =   17
               Left            =   840
               Top             =   5760
               Visible         =   0   'False
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   450
               Step            =   10
            End
            Begin prjDownloadBooster.ProgressBar pbProgress 
               Height          =   255
               Index           =   18
               Left            =   840
               Top             =   6120
               Visible         =   0   'False
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   450
               Step            =   10
            End
            Begin prjDownloadBooster.ProgressBar pbProgress 
               Height          =   255
               Index           =   19
               Left            =   840
               Top             =   6480
               Visible         =   0   'False
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   450
               Step            =   10
            End
            Begin prjDownloadBooster.ProgressBar pbProgress 
               Height          =   255
               Index           =   20
               Left            =   840
               Top             =   6840
               Visible         =   0   'False
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   450
               Step            =   10
            End
            Begin prjDownloadBooster.ProgressBar pbProgress 
               Height          =   255
               Index           =   21
               Left            =   840
               Top             =   7200
               Visible         =   0   'False
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   450
               Step            =   10
            End
            Begin prjDownloadBooster.ProgressBar pbProgress 
               Height          =   255
               Index           =   22
               Left            =   840
               Top             =   7560
               Visible         =   0   'False
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   450
               Step            =   10
            End
            Begin prjDownloadBooster.ProgressBar pbProgress 
               Height          =   255
               Index           =   23
               Left            =   840
               Top             =   7920
               Visible         =   0   'False
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   450
               Step            =   10
            End
            Begin prjDownloadBooster.ProgressBar pbProgress 
               Height          =   255
               Index           =   24
               Left            =   840
               Top             =   8280
               Visible         =   0   'False
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   450
               Step            =   10
            End
            Begin prjDownloadBooster.ProgressBar pbProgress 
               Height          =   255
               Index           =   25
               Left            =   840
               Top             =   8640
               Visible         =   0   'False
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   450
               Step            =   10
            End
            Begin VB.Label lblDownloader 
               Caption         =   "스레드 0:"
               Height          =   255
               Index           =   25
               Left            =   0
               TabIndex        =   111
               Top             =   8685
               Width           =   855
            End
            Begin VB.Label lblPercentage 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "(100%)"
               Height          =   255
               Index           =   25
               Left            =   5040
               TabIndex        =   110
               Top             =   8700
               Width           =   615
            End
            Begin VB.Label lblDownloader 
               Caption         =   "스레드 0:"
               Height          =   255
               Index           =   24
               Left            =   0
               TabIndex        =   109
               Top             =   8325
               Width           =   855
            End
            Begin VB.Label lblPercentage 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "(100%)"
               Height          =   255
               Index           =   24
               Left            =   5040
               TabIndex        =   108
               Top             =   8325
               Width           =   615
            End
            Begin VB.Label lblDownloader 
               Caption         =   "스레드 0:"
               Height          =   255
               Index           =   23
               Left            =   0
               TabIndex        =   107
               Top             =   7965
               Width           =   855
            End
            Begin VB.Label lblPercentage 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "(100%)"
               Height          =   255
               Index           =   23
               Left            =   5040
               TabIndex        =   106
               Top             =   7965
               Width           =   615
            End
            Begin VB.Label lblDownloader 
               Caption         =   "스레드 0:"
               Height          =   255
               Index           =   22
               Left            =   0
               TabIndex        =   105
               Top             =   7605
               Width           =   855
            End
            Begin VB.Label lblPercentage 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "(100%)"
               Height          =   255
               Index           =   22
               Left            =   5040
               TabIndex        =   104
               Top             =   7605
               Width           =   615
            End
            Begin VB.Label lblDownloader 
               Caption         =   "스레드 0:"
               Height          =   255
               Index           =   21
               Left            =   0
               TabIndex        =   103
               Top             =   7245
               Width           =   855
            End
            Begin VB.Label lblPercentage 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "(100%)"
               Height          =   255
               Index           =   21
               Left            =   5040
               TabIndex        =   102
               Top             =   7245
               Width           =   615
            End
            Begin VB.Label lblDownloader 
               Caption         =   "스레드 0:"
               Height          =   255
               Index           =   20
               Left            =   0
               TabIndex        =   101
               Top             =   6885
               Width           =   855
            End
            Begin VB.Label lblPercentage 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "(100%)"
               Height          =   255
               Index           =   20
               Left            =   5040
               TabIndex        =   100
               Top             =   6885
               Width           =   615
            End
            Begin VB.Label lblDownloader 
               Caption         =   "스레드 0:"
               Height          =   255
               Index           =   19
               Left            =   0
               TabIndex        =   99
               Top             =   6525
               Width           =   855
            End
            Begin VB.Label lblPercentage 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "(100%)"
               Height          =   255
               Index           =   19
               Left            =   5040
               TabIndex        =   98
               Top             =   6525
               Width           =   615
            End
            Begin VB.Label lblDownloader 
               Caption         =   "스레드 0:"
               Height          =   255
               Index           =   18
               Left            =   0
               TabIndex        =   97
               Top             =   6165
               Width           =   855
            End
            Begin VB.Label lblPercentage 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "(100%)"
               Height          =   255
               Index           =   18
               Left            =   5040
               TabIndex        =   96
               Top             =   6165
               Width           =   615
            End
            Begin VB.Label lblDownloader 
               Caption         =   "스레드 0:"
               Height          =   255
               Index           =   17
               Left            =   0
               TabIndex        =   95
               Top             =   5805
               Width           =   855
            End
            Begin VB.Label lblPercentage 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "(100%)"
               Height          =   255
               Index           =   17
               Left            =   5040
               TabIndex        =   94
               Top             =   5805
               Width           =   615
            End
            Begin VB.Label lblDownloader 
               Caption         =   "스레드 0:"
               Height          =   255
               Index           =   16
               Left            =   0
               TabIndex        =   93
               Top             =   5445
               Width           =   855
            End
            Begin VB.Label lblPercentage 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "(100%)"
               Height          =   255
               Index           =   16
               Left            =   5040
               TabIndex        =   92
               Top             =   5445
               Width           =   615
            End
            Begin VB.Label lblDownloader 
               Caption         =   "스레드 0:"
               Height          =   255
               Index           =   15
               Left            =   0
               TabIndex        =   91
               Top             =   5085
               Width           =   855
            End
            Begin VB.Label lblPercentage 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "(100%)"
               Height          =   255
               Index           =   15
               Left            =   5040
               TabIndex        =   90
               Top             =   5085
               Width           =   615
            End
            Begin VB.Label lblDownloader 
               Caption         =   "스레드 0:"
               Height          =   255
               Index           =   14
               Left            =   0
               TabIndex        =   89
               Top             =   4725
               Width           =   855
            End
            Begin VB.Label lblPercentage 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "(100%)"
               Height          =   255
               Index           =   14
               Left            =   5040
               TabIndex        =   88
               Top             =   4725
               Width           =   615
            End
            Begin VB.Label lblDownloader 
               Caption         =   "스레드 0:"
               Height          =   255
               Index           =   13
               Left            =   0
               TabIndex        =   87
               Top             =   4365
               Width           =   855
            End
            Begin VB.Label lblPercentage 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "(100%)"
               Height          =   255
               Index           =   13
               Left            =   5040
               TabIndex        =   86
               Top             =   4365
               Width           =   615
            End
            Begin VB.Label lblDownloader 
               Caption         =   "스레드 0:"
               Height          =   255
               Index           =   12
               Left            =   0
               TabIndex        =   85
               Top             =   4005
               Width           =   855
            End
            Begin VB.Label lblPercentage 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "(100%)"
               Height          =   255
               Index           =   12
               Left            =   5040
               TabIndex        =   84
               Top             =   4005
               Width           =   615
            End
            Begin VB.Label lblDownloader 
               Caption         =   "스레드 0:"
               Height          =   255
               Index           =   11
               Left            =   0
               TabIndex        =   83
               Top             =   3645
               Width           =   855
            End
            Begin VB.Label lblPercentage 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "(100%)"
               Height          =   255
               Index           =   11
               Left            =   5040
               TabIndex        =   82
               Top             =   3645
               Width           =   615
            End
            Begin VB.Label lblDownloader 
               Caption         =   "스레드 0:"
               Height          =   255
               Index           =   10
               Left            =   0
               TabIndex        =   81
               Top             =   3285
               Width           =   855
            End
            Begin VB.Label lblPercentage 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "(100%)"
               Height          =   255
               Index           =   10
               Left            =   5040
               TabIndex        =   80
               Top             =   3285
               Width           =   615
            End
            Begin VB.Label lblDownloader 
               Caption         =   "스레드 0:"
               Height          =   255
               Index           =   9
               Left            =   0
               TabIndex        =   79
               Top             =   2925
               Width           =   855
            End
            Begin VB.Label lblPercentage 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "(100%)"
               Height          =   255
               Index           =   9
               Left            =   5040
               TabIndex        =   78
               Top             =   2925
               Width           =   615
            End
            Begin VB.Label lblDownloader 
               Caption         =   "스레드 0:"
               Height          =   255
               Index           =   8
               Left            =   0
               TabIndex        =   77
               Top             =   2565
               Width           =   855
            End
            Begin VB.Label lblPercentage 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "(100%)"
               Height          =   255
               Index           =   8
               Left            =   5040
               TabIndex        =   76
               Top             =   2565
               Width           =   615
            End
            Begin VB.Label lblDownloader 
               Caption         =   "스레드 0:"
               Height          =   255
               Index           =   7
               Left            =   0
               TabIndex        =   75
               Top             =   2205
               Width           =   855
            End
            Begin VB.Label lblPercentage 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "(100%)"
               Height          =   255
               Index           =   7
               Left            =   5040
               TabIndex        =   74
               Top             =   2205
               Width           =   615
            End
            Begin VB.Label lblDownloader 
               Caption         =   "스레드 0:"
               Height          =   255
               Index           =   6
               Left            =   0
               TabIndex        =   73
               Top             =   1845
               Width           =   855
            End
            Begin VB.Label lblPercentage 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "(100%)"
               Height          =   255
               Index           =   6
               Left            =   5040
               TabIndex        =   72
               Top             =   1845
               Width           =   615
            End
            Begin VB.Label lblDownloader 
               Caption         =   "스레드 0:"
               Height          =   255
               Index           =   5
               Left            =   0
               TabIndex        =   71
               Top             =   1485
               Width           =   855
            End
            Begin VB.Label lblPercentage 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "(100%)"
               Height          =   255
               Index           =   5
               Left            =   5040
               TabIndex        =   70
               Top             =   1485
               Width           =   615
            End
            Begin VB.Label lblDownloader 
               Caption         =   "스레드 0:"
               Height          =   255
               Index           =   4
               Left            =   0
               TabIndex        =   69
               Top             =   1125
               Width           =   855
            End
            Begin VB.Label lblPercentage 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "(100%)"
               Height          =   255
               Index           =   4
               Left            =   5040
               TabIndex        =   68
               Top             =   1125
               Width           =   615
            End
            Begin VB.Label lblDownloader 
               Caption         =   "스레드 0:"
               Height          =   255
               Index           =   3
               Left            =   0
               TabIndex        =   67
               Top             =   765
               Width           =   855
            End
            Begin VB.Label lblPercentage 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "(100%)"
               Height          =   255
               Index           =   3
               Left            =   5040
               TabIndex        =   66
               Top             =   765
               Width           =   615
            End
            Begin VB.Label lblDownloader 
               Caption         =   "스레드 0:"
               Height          =   255
               Index           =   2
               Left            =   0
               TabIndex        =   65
               Top             =   405
               Width           =   855
            End
            Begin VB.Label lblPercentage 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "(100%)"
               Height          =   255
               Index           =   2
               Left            =   5040
               TabIndex        =   64
               Top             =   405
               Width           =   615
            End
            Begin VB.Label lblDownloader 
               Caption         =   "스레드 0:"
               Height          =   255
               Index           =   1
               Left            =   0
               TabIndex        =   63
               Top             =   45
               Width           =   855
            End
            Begin VB.Label lblPercentage 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "(100%)"
               Height          =   255
               Index           =   1
               Left            =   5040
               TabIndex        =   62
               Top             =   45
               Width           =   615
            End
         End
      End
      Begin prjDownloadBooster.TextBoxW txtDummyScroll 
         Height          =   3450
         Left            =   5640
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   0
         _ExtentY        =   0
         Enabled         =   0   'False
         BorderStyle     =   0
         MultiLine       =   -1  'True
         ScrollBars      =   2
      End
      Begin prjDownloadBooster.ListBoxW lvDummyScroll 
         Height          =   3450
         Left            =   5400
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   0
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   6085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BorderStyle     =   0
      End
   End
   Begin VB.DirListBox CurDir 
      Height          =   510
      Left            =   9240
      TabIndex        =   41
      Top             =   5280
      Visible         =   0   'False
      Width           =   495
   End
   Begin prjDownloadBooster.CommandButtonW cmdIncreaseThreads 
      Height          =   315
      Left            =   6960
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   795
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   556
      ImageListAlignment=   4
      Caption         =   ">"
      Transparent     =   -1  'True
   End
   Begin prjDownloadBooster.CommandButtonW cmdDecreaseThreads 
      Height          =   315
      Left            =   1560
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   795
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   556
      ImageListAlignment=   4
      Caption         =   "<"
      Transparent     =   -1  'True
   End
   Begin prjDownloadBooster.ComboBoxW cbWhenExist 
      Height          =   300
      Left            =   7590
      TabIndex        =   12
      Top             =   2265
      Width           =   1425
      _ExtentX        =   0
      _ExtentY        =   0
      Style           =   2
   End
   Begin prjDownloadBooster.CheckBoxW chkOpenAfterComplete 
      Height          =   255
      Left            =   6840
      TabIndex        =   8
      Top             =   1560
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   450
      Caption         =   "완료 후 열기(&C)"
   End
   Begin prjDownloadBooster.CheckBoxW chkOpenFolder 
      Height          =   255
      Left            =   6840
      TabIndex        =   9
      Top             =   1785
      Width           =   2175
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "완료 후 폴더 열기(&L)"
   End
   Begin prjDownloadBooster.CommandButtonW cmdClear 
      Height          =   330
      Left            =   7440
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   90
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      ImageList       =   "imgErase"
      Caption         =   "초기화(&Y) "
      Transparent     =   -1  'True
   End
   Begin prjDownloadBooster.CommandButtonW cmdAdd 
      Height          =   375
      Left            =   2520
      TabIndex        =   21
      Top             =   6960
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      ImageList       =   "imgPlusYellow"
      Caption         =   " 추가(&R)..."
      Transparent     =   -1  'True
   End
   Begin prjDownloadBooster.CommandButtonW cmdStopBatch 
      Height          =   375
      Left            =   7560
      TabIndex        =   24
      Top             =   6960
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Enabled         =   0   'False
      ImageList       =   "imgStopRed"
      Caption         =   "중지(&Z) "
      Transparent     =   -1  'True
   End
   Begin prjDownloadBooster.CommandButtonW cmdBatch 
      Height          =   330
      Left            =   7320
      TabIndex        =   18
      Top             =   5565
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      ImageList       =   "imgDropdown"
      ImageListAlignment=   1
      Caption         =   "  일괄 처리(&H)"
      Transparent     =   -1  'True
   End
   Begin prjDownloadBooster.FrameW fTotal 
      Height          =   615
      Left            =   240
      TabIndex        =   29
      Top             =   1320
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   1085
      Caption         =   " 전체 다운로드 진행률 "
      Transparent     =   -1  'True
      Begin VB.Label lblOverlay 
         Alignment       =   2  '가운데 맞춤
         BackStyle       =   0  '투명
         Caption         =   "전체 다운로드 현황"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   117
         Top             =   30
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Shape pgOverlay 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  '투명하지 않음
         Height          =   225
         Index           =   1
         Left            =   120
         Shape           =   4  '둥근 사각형
         Top             =   0
         Visible         =   0   'False
         Width           =   2775
      End
   End
   Begin prjDownloadBooster.FrameW fOptions 
      Height          =   2490
      Left            =   6720
      TabIndex        =   26
      Top             =   1320
      Width           =   2415
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   " 옵션 "
      Transparent     =   -1  'True
      Begin VB.Label lblOverlay 
         Alignment       =   2  '가운데 맞춤
         BackStyle       =   0  '투명
         Caption         =   "설정"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   116
         Top             =   30
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Shape pgOverlay 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  '투명하지 않음
         Height          =   225
         Index           =   0
         Left            =   120
         Shape           =   4  '둥근 사각형
         Top             =   0
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "중복(&K):"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Tag             =   "nocolorchange"
         Top             =   990
         Width           =   735
      End
      Begin VB.Shape pgSettingsBackground 
         BackColor       =   &H8000000F&
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H80000010&
         Height          =   1545
         Left            =   60
         Top             =   210
         Visible         =   0   'False
         Width           =   2295
      End
   End
   Begin prjDownloadBooster.CommandButtonW cmdOpen 
      Height          =   330
      Left            =   7320
      TabIndex        =   13
      Top             =   3945
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      Enabled         =   0   'False
      ImageList       =   "imgOpenFile"
      Caption         =   "열기(&O) "
      Transparent     =   -1  'True
   End
   Begin prjDownloadBooster.CommandButtonW cmdOpenFolder 
      Height          =   330
      Left            =   7320
      TabIndex        =   14
      Top             =   4320
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      ImageList       =   "imgOpenFolder"
      Caption         =   "폴더 열기(&E) "
      Transparent     =   -1  'True
   End
   Begin VB.Timer timElapsed 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   9240
      Top             =   4560
   End
   Begin prjDownloadBooster.Slider trThreadCount 
      Height          =   495
      Left            =   1935
      TabIndex        =   6
      Top             =   750
      Width           =   5025
      _ExtentX        =   8864
      _ExtentY        =   873
      Min             =   1
      Max             =   25
      Value           =   1
      TickFrequency   =   2
      TipSide         =   1
      SelStart        =   1
      Transparent     =   -1  'True
   End
   Begin prjDownloadBooster.CommandButtonW cmdBrowse 
      Height          =   330
      Left            =   7440
      TabIndex        =   4
      Top             =   435
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      ImageList       =   "imgOpenFolder"
      Caption         =   " 찾아보기(&B)..."
      Transparent     =   -1  'True
   End
   Begin prjDownloadBooster.TextBoxW txtFileName 
      Height          =   300
      Left            =   1560
      TabIndex        =   3
      Top             =   450
      Width           =   5775
      _ExtentX        =   0
      _ExtentY        =   0
   End
   Begin prjDownloadBooster.TextBoxW txtURL 
      Height          =   300
      Left            =   1560
      TabIndex        =   1
      Top             =   105
      Width           =   5775
      _ExtentX        =   0
      _ExtentY        =   0
   End
   Begin prjDownloadBooster.CommandButtonW cmdGo 
      Default         =   -1  'True
      Height          =   330
      Left            =   7320
      TabIndex        =   16
      Top             =   4760
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      ImageList       =   "imgDownload"
      Caption         =   "다운로드(&D) "
      Transparent     =   -1  'True
   End
   Begin prjDownloadBooster.FrameW Frame4 
      Height          =   3885
      Left            =   240
      TabIndex        =   45
      Top             =   2040
      Width           =   6255
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "                               "
      Transparent     =   -1  'True
      Begin VB.Shape pgOverlay 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  '투명하지 않음
         Height          =   225
         Index           =   2
         Left            =   120
         Shape           =   4  '둥근 사각형
         Top             =   0
         Visible         =   0   'False
         Width           =   1935
      End
   End
   Begin VB.Label lblLBCaption 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "현   황"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   126
      Tag             =   "nocolorsizechange"
      Top             =   1320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image imgTopLeft 
      Height          =   435
      Left            =   120
      Picture         =   "frmMain.frx":8CD6
      Top             =   1200
      Width           =   1725
   End
   Begin VB.Image imgTop 
      Height          =   435
      Left            =   1845
      Picture         =   "frmMain.frx":B484
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   3585
   End
   Begin VB.Image imgTopRight 
      Height          =   435
      Left            =   5430
      Picture         =   "frmMain.frx":10656
      Top             =   1200
      Width           =   150
   End
   Begin VB.Image imgLeft 
      Height          =   2310
      Left            =   120
      Picture         =   "frmMain.frx":10A38
      Stretch         =   -1  'True
      Top             =   1635
      Width           =   1725
   End
   Begin VB.Image imgBottomLeft 
      Height          =   180
      Left            =   120
      Picture         =   "frmMain.frx":1DBD2
      Top             =   3945
      Width           =   1725
   End
   Begin VB.Image imgBottom 
      Height          =   180
      Left            =   1845
      Picture         =   "frmMain.frx":1EC64
      Stretch         =   -1  'True
      Top             =   3945
      Width           =   3585
   End
   Begin VB.Image imgBottomRight 
      Height          =   180
      Left            =   5430
      Picture         =   "frmMain.frx":20E66
      Top             =   3945
      Width           =   150
   End
   Begin VB.Image imgRight 
      Height          =   2310
      Left            =   5415
      Picture         =   "frmMain.frx":21028
      Stretch         =   -1  'True
      Top             =   1635
      Width           =   165
   End
   Begin VB.Image imgCenter 
      Height          =   2310
      Left            =   1845
      Stretch         =   -1  'True
      Top             =   1635
      Width           =   3570
   End
   Begin VB.Label lblState 
      BackStyle       =   0  '투명
      Caption         =   "중지됨"
      Height          =   255
      Left            =   10560
      TabIndex        =   125
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label lblThreadCount 
      BackStyle       =   0  '투명
      Caption         =   "(일반 다운로드)"
      Height          =   255
      Left            =   7440
      TabIndex        =   7
      Top             =   870
      Width           =   1455
   End
   Begin VB.Label lblThreadCountLabel 
      BackStyle       =   0  '투명
      Caption         =   "강도(&T):"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   870
      Width           =   1215
   End
   Begin VB.Label lblFilePath 
      BackStyle       =   0  '투명
      Caption         =   "저장 경로(&F):"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   490
      Width           =   1215
   End
   Begin VB.Label lblURL 
      BackStyle       =   0  '투명
      Caption         =   "파일 주소(&A):"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   135
      Width           =   1215
   End
   Begin prjDownloadBooster.ShellPipe SP 
      Left            =   9240
      Top             =   4920
      _ExtentX        =   635
      _ExtentY        =   635
   End
   Begin VB.Image imgBackground 
      Height          =   135
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   135
   End
   Begin VB.Menu mnuListContext 
      Caption         =   "mnuListContext"
      Visible         =   0   'False
      Begin VB.Menu mnuDeleteItem 
         Caption         =   "제거(&R)"
      End
      Begin VB.Menu mnuClearBatch3 
         Caption         =   "모두 제거(&C)"
      End
   End
   Begin VB.Menu mnuListContext2 
      Caption         =   "mnuListContext2"
      Visible         =   0   'False
      Begin VB.Menu mnuAddItem 
         Caption         =   "추가(&A)..."
      End
      Begin VB.Menu mnuClearBatch2 
         Caption         =   "모두 제거(&C)"
      End
   End
   Begin VB.Menu mnuDeleteDropdown 
      Caption         =   "mnuDeleteDropdown"
      Visible         =   0   'False
      Begin VB.Menu mnuClearBatch 
         Caption         =   "모두 제거(&C)"
      End
   End
   Begin VB.Menu mnuOpenDropdown 
      Caption         =   "mnuOpenDropdown"
      Visible         =   0   'False
      Begin VB.Menu mnuOpenFolder 
         Caption         =   "폴더 열기(&F)"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Elapsed As Long
Dim BatchStarted As Boolean
Dim CurrentBatchIdx As Integer
Dim DownloadPath As String
Dim IsDownloading As Boolean
Dim BatchErrorCount As Integer
Dim TahomaAvailable As Boolean
Dim PrevDownloadedBytes As Double
Dim SpeedCount As Integer
Dim HttpStatusCode As String
Dim ResumeUnsupported As Boolean
Public ImagePosition As Integer

Sub OnData(Data As String)
    Dim output$
    Dim idx%
    Dim progress%
    Dim DownloadedBytes As Double
    If Left$(Data, 7) = "STATUS " Then
        Select Case Replace(Right$(Data, Len(Data) - 7), " ", "")
            Case "CHECKREDIRECT"
                sbStatusBar.Panels(1).Text = t("서버를 찾는 중...", "Finding server...")
            Case "CHECKFILE"
                sbStatusBar.Panels(1).Text = t("가용성 확인 중...", "Checking availability...")
            Case "DOWNLOADING"
                sbStatusBar.Panels(1).Text = t("다운로드 중...", "Downloading...")
            Case "MERGING"
                sbStatusBar.Panels(1).Text = t("파일 조각 결합 중...", "Merging segments...")
                'pbTotalProgress.Scrolling = PrbScrollingMarquee
                pbTotalProgressMarquee.Visible = -1
                pbTotalProgressMarquee.MarqueeAnimation = -1
            Case "COMPLETE"
                sbStatusBar.Panels(1).Text = t("완료", "Complete")
                sbStatusBar.Panels(2).Text = ""
                sbStatusBar.Panels(3).Text = ""
                sbStatusBar.Panels(4).Text = ""
                'pbTotalProgress.Scrolling = PrbScrollingStandard
                pbTotalProgressMarquee.MarqueeAnimation = 0
                pbTotalProgressMarquee.Visible = 0
                pbTotalProgress.Value = 100
            Case "UNABLETOCONTINUE"
                Alert t("이어받기가 불가능합니다. 처음부터 다시 다운로드합니다.", "Unable to resume. Starting over..."), App.Title, , 48, 5000
            Case "RESUMEUNSUPPORTED"
                ResumeUnsupported = True
        End Select
    ElseIf Left$(Data, 11) = "STATUSCODE " Then
        output = Right$(Data, Len(Data) - 11)
        HttpStatusCode = Trim$(output)
    ElseIf Left$(Data, 5) = "DATA " Then
        output = Right$(Data, Len(Data) - 5)
        idx = CInt(Split(output, ",")(0))
        If CLng(Split(output, ",")(1)) > 100 Then
            progress = -1
        Else
            progress = CInt(Split(output, ",")(1))
        End If
        If progress < 0 Then
            If Not pbProgressMarquee(idx).Visible Then
                'pbProgress(idx).Scrolling = PrbScrollingMarquee
                pbProgressMarquee(idx).Visible = -1
                pbProgressMarquee(idx).MarqueeAnimation = -1
            End If
            lblPercentage(idx).Caption = ""
        Else
            If pbProgressMarquee(idx).Visible Then
                'pbProgress(idx).Scrolling = PrbScrollingStandard
                pbProgressMarquee(idx).MarqueeAnimation = 0
                pbProgressMarquee(idx).Visible = 0
            End If
            pbProgress(idx).Value = progress
            lblPercentage(idx).Caption = "(" & progress & "%)"
        End If
        
        If trThreadCount.Value > 1 And idx = 1 And (CDbl(Split(output, ",")(2)) > 0 Or lblTotalBytes.Caption = "0 바이트") Then lblTotalSizeThread.Caption = ParseSize(CDbl(Split(output, ",")(2)), True)
    ElseIf Left$(Data, 6) = "TOTAL " Then
        output = Right$(Data, Len(Data) - 6)
        Dim strTotal As String
        Dim total As Double
        strTotal = Split(output, ",")(0)
        If IsNumeric(strTotal) Then
            total = CDbl(strTotal)
        Else
            total = -1
        End If
        If (Not IsNumeric(Split(output, ",")(2))) Or CLng(Split(output, ",")(2)) > 100 Then
            progress = -1
        Else
            progress = CInt(Split(output, ",")(2))
        End If
        
        DownloadedBytes = CDbl(Split(output, ",")(1))
        
        If progress < 0 Then
            If total > 0 And DownloadedBytes > 0 And DownloadedBytes <= total Then
                progress = Floor(DownloadedBytes / total * 100)
                GoTo progressAvailable
            Else
                If Not pbTotalProgressMarquee.Visible Then
                    pbTotalProgressMarquee.Visible = -1
                    pbTotalProgressMarquee.MarqueeAnimation = -1
                End If
            End If
            If fTotal.Caption <> t(" 전체 다운로드 진행률 ", " Total Progress ") Then fTotal.Caption = t(" 전체 다운로드 진행률 ", " Total Progress ")
            If pbTotalProgress.Value <> 0 Then pbTotalProgress.Value = 0
            If DownloadedBytes = -1 Then
                sbStatusBar.Panels(2).Text = ""
            ElseIf total <= 0 Then
                sbStatusBar.Panels(2).Text = ParseSize(DownloadedBytes)
            Else
                sbStatusBar.Panels(2).Text = t(ParseSize(total) & " 중 " & ParseSize(DownloadedBytes), ParseSize(DownloadedBytes) & " of " & ParseSize(total))
            End If
            If total <= 0 Then
                If lblTotalBytes.Caption <> t("알 수 없음", "Unknown") Then lblTotalBytes.Caption = t("알 수 없음", "Unknown")
            Else
                lblTotalBytes.Caption = ParseSize(total, True)
            End If
            lblDownloadedBytes.Caption = ParseSize(DownloadedBytes, True)
        Else
progressAvailable:
            If pbTotalProgressMarquee.Visible Then
                pbTotalProgressMarquee.MarqueeAnimation = 0
                pbTotalProgressMarquee.Visible = 0
            End If
            If strTotal = "-1" Then
                sbStatusBar.Panels(2).Text = ParseSize(DownloadedBytes)
            Else
                sbStatusBar.Panels(2).Text = t(ParseSize(strTotal) & " 중 " & ParseSize(DownloadedBytes), ParseSize(DownloadedBytes) & " of " & ParseSize(strTotal))
            End If
            If strTotal = "NaN" Or strTotal = "-1" Then
                lblTotalBytes.Caption = t("알 수 없음", "Unknown")
            Else
                lblTotalBytes.Caption = ParseSize(total, True)
            End If
            lblDownloadedBytes.Caption = ParseSize(DownloadedBytes, True)
            pbTotalProgress.Value = progress
            fTotal.Caption = t(" 전체 다운로드 진행률 (" & progress & "%) ", " Total Progress (" & progress & "%) ")
        End If
        
        Dim Speed As Double
        SpeedCount = SpeedCount + 1
        If SpeedCount >= 10 Then
            Speed = (DownloadedBytes - PrevDownloadedBytes)
            lblSpeed.Caption = ParseSize(Speed, True, "/" & t("초", "sec"))
            sbStatusBar.Panels(3).Text = ParseSize(Speed, False, "/" & t("초", "sec"))
            PrevDownloadedBytes = DownloadedBytes
            SpeedCount = 0
            
            If progress >= 0 And strTotal <> "-1" And IsNumeric(strTotal) Then
                lblRemaining = FormatTime((CDbl(strTotal) - CDbl(DownloadedBytes)) / Speed)
            End If
        End If
    ElseIf Left$(Data, 17) = "MODIFIEDFILENAME " Then
        output = Right$(Data, Len(Data) - 17)
        DownloadPath = output
        lblFilename.Caption = fso.GetFilename(output)
        If BatchStarted Then
            lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(1).Text = output
            lvBatchFiles.ListItems(CurrentBatchIdx).Text = lblFilename.Caption
        End If
        If Len(lblFilename.Caption) > 22 Then lblFilename.Caption = Left$(lblFilename.Caption, 22) & "..."
    End If
End Sub

Sub NextBatchDownload()
    Dim i%
    If Not BatchStarted Then Exit Sub
    
    If lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(3).Text = t("완료", "Done") Then _
        lvBatchFiles.ListItems(CurrentBatchIdx).Checked = False
    
    If CurrentBatchIdx = lvBatchFiles.ListItems.Count Then
        BatchStarted = False
        CurrentBatchIdx = 1
        cmdStartBatch.Enabled = -1
        tygStartBatch.Enabled = -1
        cmdStopBatch.Enabled = 0
        tygStopbatch.Enabled = 0
        timElapsed.Enabled = 0
        sbStatusBar.Panels(3).Text = ""
        sbStatusBar.Panels(4).Text = ""
        chkOpenAfterComplete.Enabled = -1
        If chkOpenFolder.Value Then
            cmdOpenFolder_Click
        End If
        cmdGo.Enabled = -1
        tygGo.Enabled = cmdGo.Enabled
        
        If lvBatchFiles.ListItems.Count > 0 Then
            Dim Enable As Boolean
            For i = 1 To lvBatchFiles.ListItems.Count
                If lvBatchFiles.ListItems(i).Checked Then
                    Enable = True
                    Exit For
                End If
            Next i
            If Not Enable Then
                cmdStartBatch.Enabled = 0
                tygStartBatch.Enabled = 0
            Else
                cmdStartBatch.Enabled = -1
                tygStartBatch.Enabled = -1
            End If
        Else
            cmdStartBatch.Enabled = 0
            tygStartBatch.Enabled = 0
        End If
        
        If BatchErrorCount Then
            Alert t("하나 이상의 오류가 발생했습니다. 오류 코드 정보는 다음과 같습니다." & vbCrLf & vbCrLf & "1: 알 수 없는 오류가 발생했습니다. 유효하지 않은 주소를 입력했거나 프로그램 내부 오류입니다." & vbCrLf & "102: 주소나 파일 이름을 지정하지 않았습니다." & vbCrLf & "3: 저장 경로가 존재하지 않습니다." & vbCrLf & "104: 저장할 파일명이 사용 중입니다. 다른 이름을 선택하십시오." & vbCrLf & "106: 파일 서버가 다운로드 부스트를 지원하지 않습니다. 강도를 1로 변경해 보십시오." & vbCrLf & "107: 파일의 크기를 알 수 없어서 다운로드를 부스트할 수 없습니다. 강도를 1로 변경해 보십시오." & vbCrLf & "108: 서버가 요청을 거부했습니다. 서버 측 오류이거나 페이지가 존재하지 않거나 접근 권한이 없을 수 있습니다.", _
                     "One or more errors have occurred." & vbCrLf & vbCrLf & "1: Network error" & vbCrLf & "103: Save path doesn't exist." & vbCrLf & "104: File name already exists" & vbCrLf & "106: Download boosting not supported. Try changing the thread count to 1." & vbCrLf & "107: Unable to boost download because the file size is not provided. Try changing the thread count to 1." & vbCrLf & "108: Server has denied your request. The file may not exist or have insufficient permissions to access it."), App.Title, Me, 48
        ElseIf chkPlaySound.Value Then
            MessageBeep 64
        End If
        
        If lblState.Caption = t("완료됨", "Done") Then
            pbTotalProgress.Value = 100
            For i = 1 To trThreadCount.Value
                pbProgress(i).Value = 100
                lblPercentage(i).Caption = "(100%)"
            Next i
        End If
        
        Exit Sub
    End If
    
    CurrentBatchIdx = CurrentBatchIdx + 1
    StartDownload lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(2), lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(1)
End Sub

Sub OnExit(RetVal As Long)
    timElapsed.Enabled = 0
    If Not BatchStarted Then
        Select Case RetVal
            Case 0
                '정상 종료
                GoTo nextln
            Case 999
                GoTo nextln
            Case 1
                If chkAutoRetry.Value <> 1 Then
                    If pbTotalProgressMarquee.Visible And (lblDownloadedBytes.Caption = "-" Or lblDownloadedBytes.Caption = "대기 중...") Then
                        Alert t("해당 파일 주소에 연결할 수 없습니다. 주소가 유효하지 않거나 서버가 응답하지 않습니다.", "The server does not respond or the file URL is invalid."), App.Title, Me, 16
                    Else
                        Alert t("서버와의 접속이 끊겼습니다. 다운로드 도중에 네트워크 오류가 발생했을 수 있습니다.", "Network error while downloading."), App.Title, Me, 16
                    End If
                End If
            Case 102
                Alert "주소나 파일 이름을 지정하지 않았습니다.", App.Title, Me, 16
            Case 3, 103
                Alert t("저장 경로가 존재하지 않습니다.", "Save path doesn't exist."), App.Title, Me, 16
            Case 104
                Alert t("저장할 파일명이 사용 중입니다. 다른 이름을 선택하십시오.", "File name already exists."), App.Title, Me, 16
            Case 106
                Alert t("파일 서버가 다운로드 부스트를 지원하지 않습니다. 강도를 1로 변경해 보십시오.", "Download boosting not supported. Try changing the thread count to 1."), App.Title, Me, 16
            Case 107
                Alert t("파일의 크기를 알 수 없어서 다운로드를 부스트할 수 없습니다. 강도를 1로 변경해 보십시오.", "Unable to boost download because the file size is not provided. Try changing the thread count to 1."), App.Title, Me, 16
            Case 108
                Dim statusMsg As String
                statusMsg = ""
                Dim ErrDesc As String
                Dim Icon As VbMsgBoxStyle
                Icon = vbCritical
                If Len(HttpStatusCode) > 0 And LangID = 1042 Then
                    Select Case HttpStatusCode
                        Case "400"
                            ErrDesc = "요청이 잘못되었습니다."
                        Case "401"
                            ErrDesc = "접근하려면 인증 정보가 필요합니다."
                        Case "402"
                            ErrDesc = "접근하려면 결제가 필요합니다."
                        Case "403"
                            ErrDesc = "접근 권한이 없습니다."
                        Case "404"
                            ErrDesc = "서버에 파일이 존재하지 않습니다."
                        Case "405"
                            ErrDesc = "파일을 받으려면 데이타를 전송해야 합니다."
                        Case "406"
                            ErrDesc = "요청을 받아들일 수 없습니다."
                        Case "407"
                            ErrDesc = "프록시 인증이 필요합니다."
                        Case "408"
                            ErrDesc = "요청이 제시간 안에 마무리되지 않았습니다."
                        Case "409"
                            ErrDesc = "요청이 서버와 충돌했습니다."
                        Case "410"
                            If Month(Now) = 4 And Day(Now) = 1 Then
                                ErrDesc = "파일이 있었는데 없었습니다."
                                Icon = vbInformation
                            Else
                                ErrDesc = "파일이 더 이상 서버에 없습니다."
                            End If
                        Case "414"
                            ErrDesc = "주소가 너무 깁니다."
                        Case "418"
                            ErrDesc = "서버가 자신은 찻주전자라서 커피를 만들 수 없다고 합니다 ㅎ   "
                            Icon = vbInformation
                        Case "451"
                            ErrDesc = "법적인 이유로 파일을 다운로드 받을 수 없습니다."
                        Case "500"
                            ErrDesc = "서버 측에서 오류가 발생했습니다."
                        Case "502"
                            ErrDesc = "게이트웨이가 불량입니다."
                        Case "503"
                            ErrDesc = "서버가 일시적으로 응답할 수 없는 상태입니다."
                        Case "504"
                            ErrDesc = "게이트웨이 시간이 초과되었습니다."
                        Case "505"
                            ErrDesc = "HTTP 버전이 지원되지 않습니다."
                        Case Else
                            ErrDesc = "서버 측 오류이거나 페이지가 존재하지 않거나 접근 권한이 없을 수 있습니다."
                            statusMsg = " HTTP 응답 코드는 ( " & HttpStatusCode & " ) 입니다."
                    End Select
                End If
                Alert t("서버가 요청을 거부했습니다. " & ErrDesc & statusMsg, "Server denied your request. The file may not exist or have insufficient permissions to access it."), App.Title, Me, Icon
            Case Else
                Alert t("내부 오류가 발생했습니다. 프로세스 반환 값은 ( " & RetVal & " ) 입니다.", "Internal error. Process returned ( " & RetVal & " )."), App.Title, Me, 16
        End Select
    End If
    
nextln:
    
    If Not BatchStarted Then
        cmdGo.Enabled = -1
        tygGo.Enabled = cmdGo.Enabled
    End If
    cmdStop.Enabled = 0
    cmdStop.Left = Me.Width + 1200
    tygStop.Enabled = cmdStop.Enabled
    tygStop.Left = cmdStop.Left
    cmdGo.Visible = -1
    OnStop (RetVal = 0)
    Dim i%
    If BatchStarted Then
        pbTotalProgress.Value = 0
        For i = 1 To lblDownloader.UBound
            pbProgress(i).Value = 0
            lblPercentage(i).Caption = ""
        Next i
        
        If RetVal <> 0 Then
            lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(3).Text = t("오류", "Error") & " (" & RetVal & ")"
            lvBatchFiles.ListItems(CurrentBatchIdx).ForeColor = 255
            lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(1).ForeColor = 255
            lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(2).ForeColor = 255
            lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(3).ForeColor = 255
            BatchErrorCount = BatchErrorCount + 1
        Else
            lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(3).Text = t("완료", "Done")
            'lvBatchFiles.ListItems(CurrentBatchIdx).Checked = False
            lvBatchFiles.ListItems(CurrentBatchIdx).ForeColor = &H8000&
            lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(1).ForeColor = &H8000&
            lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(2).ForeColor = &H8000&
            lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(3).ForeColor = &H8000&
        End If
    
        NextBatchDownload
    ElseIf RetVal = 0 Then
        cmdOpen.Enabled = -1
        tygOpen.Enabled = -1
        If chkOpenAfterComplete.Value Then
            cmdOpen_Click
        End If
        If chkOpenFolder.Value Then
            cmdOpenFolder_Click
        End If
    ElseIf RetVal = 1 And chkAutoRetry.Value Then
        MessageBeep 48
        cmdGo_Click
    End If
End Sub

Sub OnStart()
    IsDownloading = True
    ResumeUnsupported = False
    cmdGo.Enabled = 0
    tygGo.Enabled = cmdGo.Enabled
    If Not BatchStarted Then
        cmdStop.Enabled = -1
        cmdStop.Left = cmdGo.Left
        tygStop.Enabled = cmdStop.Enabled
        tygStop.Left = cmdStop.Left
        cmdStop.Refresh
        cmdGo.Visible = 0
    Else
        cmdStop.Enabled = 0
        cmdStop.Left = Me.Width + 1200
        tygStop.Enabled = cmdStop.Enabled
        tygStop.Left = cmdStop.Left
        cmdGo.Visible = -1
    End If
    
    lblURL.Enabled = 0
    txtURL.Enabled = 0
    cmdClear.Enabled = 0
    
    lblFilePath.Enabled = 0
    lblThreadCountLabel.Enabled = 0
    
    txtFileName.Enabled = 0
    cmdBrowse.Enabled = 0
    
    trThreadCount.Enabled = 0
    cmdDecreaseThreads.Enabled = 0
    cmdIncreaseThreads.Enabled = 0
    
    lblThreadCount.Enabled = 0
    
    cmdStartBatch.Enabled = 0
    tygStartBatch.Enabled = 0
    
    cmdOpen.Enabled = 0
    tygOpen.Enabled = 0
    
    lblTotalBytes.Caption = t("대기 중...", "Pending...")
    lblDownloadedBytes.Caption = t("대기 중...", "Pending...")
    If trThreadCount.Value > 1 Then
        lblTotalSizeThread.Caption = t("대기 중...", "Pending...")
        lblThreadCount2.Caption = trThreadCount.Value
    Else
        lblTotalSizeThread.Caption = "-"
        lblThreadCount2.Caption = "-"
    End If
    lblElapsed.Caption = "0" & t("초", " seconds")
    lblSpeed.Caption = "-"
    lblRemaining.Caption = "-"
    
    fTotal.Caption = t(" 전체 다운로드 진행률 ", " Total Progress ")
    pbTotalProgress.Value = 0
    For i = 1 To trThreadCount.Value
        lblPercentage(i).Caption = ""
        pbProgress(i).Value = 0
    Next i
    
    For i = 1 To trThreadCount.Value
        'pbProgress(i).MarqueeSpeed = 35
        'pbProgress(i).Scrolling = PrbScrollingMarquee
        pbProgressMarquee(i).Visible = -1
        pbProgressMarquee(i).MarqueeAnimation = -1
    Next i
    
    pbTotalProgressMarquee.Visible = -1
    pbTotalProgressMarquee.MarqueeAnimation = -1
    
    lblState.Caption = t("진행 중", "Working")
    sbStatusBar.Panels(1).Text = t("시작 중...", "Starting...")
End Sub

Sub OnStop(Optional PlayBeep As Boolean = True)
    IsDownloading = False
    If Not BatchStarted Then
        cmdGo.Enabled = -1
        tygGo.Enabled = cmdGo.Enabled
    End If
    cmdStop.Enabled = 0
    cmdStop.Left = Me.Width + 1200
    tygStop.Enabled = cmdStop.Enabled
    tygStop.Left = cmdStop.Left
    cmdGo.Visible = -1
    
    lblURL.Enabled = -1
    lblFilePath.Enabled = -1
    lblThreadCountLabel.Enabled = -1
    
    txtURL.Enabled = -1
    txtFileName.Enabled = -1
    cmdBrowse.Enabled = -1
    cmdClear.Enabled = -1
    
    trThreadCount.Enabled = -1
    If trThreadCount.Value > trThreadCount.Min Then cmdDecreaseThreads.Enabled = -1
    If trThreadCount.Value < trThreadCount.Max Then cmdIncreaseThreads.Enabled = -1
    
    lblThreadCount.Enabled = -1
    
    SP.FinishChild 0, 0
    
    Dim i%
    For i = 1 To trThreadCount.Value
        'pbProgress(i).Scrolling = PrbScrollingStandard
        pbProgressMarquee(i).MarqueeAnimation = 0
        pbProgressMarquee(i).Visible = 0
    Next i
    
    If pbTotalProgressMarquee.Visible Then
        pbTotalProgressMarquee.MarqueeAnimation = 0
        pbTotalProgressMarquee.Visible = 0
    End If
    
    If pbTotalProgress.Value < 100 Then
        pbTotalProgress.Value = 0
    End If
    
    If pbTotalProgress.Value < 100 Then
        lblState.Caption = t("중지됨", "Stopped")
        sbStatusBar.Panels(1).Text = t("준비", "Ready")
    
        fTotal.Caption = t(" 전체 다운로드 진행률 ", " Total Progress ")
        For i = 1 To lblDownloader.UBound
            pbProgress(i).Value = 0
            lblPercentage(i).Caption = ""
        Next i
    Else
        lblState.Caption = t("완료됨", "Done")
        sbStatusBar.Panels(1).Text = t("완료", "Done")
        sbStatusBar.Panels(2).Text = ""
        sbStatusBar.Panels(3).Text = ""
        sbStatusBar.Panels(4).Text = ""
    End If
    
    If Not BatchStarted Then
        timElapsed.Enabled = 0
        sbStatusBar.Panels(3).Text = ""
        sbStatusBar.Panels(4).Text = ""
        
        If lvBatchFiles.ListItems.Count > 0 Then
            Dim Enable As Boolean
            For i = 1 To lvBatchFiles.ListItems.Count
                If lvBatchFiles.ListItems(i).Checked Then
                    Enable = True
                    Exit For
                End If
            Next i
            If Not Enable Then
                cmdStartBatch.Enabled = 0
                tygStartBatch.Enabled = 0
            Else
                cmdStartBatch.Enabled = -1
                tygStartBatch.Enabled = -1
            End If
        Else
            cmdStartBatch.Enabled = 0
            tygStartBatch.Enabled = 0
        End If
        
        If PlayBeep And chkPlaySound.Value Then
            MessageBeep 64
            lblState.Caption = t("완료됨", "Done")
            sbStatusBar.Panels(1).Text = t("완료", "Done")
            sbStatusBar.Panels(2).Text = ""
        End If
    End If
    
    If lblTotalBytes.Caption = t("대기 중...", "Pending...") Then lblTotalBytes.Caption = "-"
    If lblDownloadedBytes.Caption = t("대기 중...", "Pending...") Then
        lblDownloadedBytes.Caption = "-"
    End If
    If PlayBeep And lblDownloadedBytes.Caption <> "-" Then
        lblTotalBytes.Caption = lblDownloadedBytes.Caption
    End If
    If lblTotalSizeThread.Caption = t("대기 중...", "Pending...") Then lblTotalSizeThread.Caption = "-"
    lblRemaining.Caption = "-"
End Sub

Private Sub cbWhenExist_Click()
    SaveSetting "DownloadBooster", "Options", "WhenFileExists", cbWhenExist.ListIndex
End Sub

Private Sub chkAutoRetry_Click()
    SaveSetting "DownloadBooster", "Options", "AutoRetry", chkAutoRetry.Value
End Sub

Private Sub chkContinueDownload_Click()
    SaveSetting "DownloadBooster", "Options", "ContinueDownload", chkContinueDownload.Value
End Sub

Private Sub chkOpenAfterComplete_Click()
    SaveSetting "DownloadBooster", "Options", "OpenWhenComplete", chkOpenAfterComplete.Value
End Sub

Private Sub chkOpenFolder_Click()
    SaveSetting "DownloadBooster", "Options", "OpenFolderWhenComplete", chkOpenFolder.Value
End Sub

Private Sub chkPlaySound_Click()
    SaveSetting "DownloadBooster", "Options", "PlaySound", chkPlaySound.Value
End Sub

Private Sub cmdAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub cmdAdd_Click()
    On Error Resume Next
    If Replace(txtURL.Text, " ", "") <> "" Then
        frmBatchAdd.txtURLs.Text = Trim$(txtURL.Text) & vbCrLf
        frmBatchAdd.txtURLs.SelStart = 0
        frmBatchAdd.txtURLs.SelLength = Len(Trim$(txtURL.Text)) + 2
    End If
    frmBatchAdd.Show vbModal, Me
End Sub

Sub AddBatchURLs(URL As String)
    If Left$(URL, 7) <> "http://" And Left$(URL, 8) <> "https://" Then
        Alert URL & " - " & t("주소가 올바르지 않습니다. 'http://' 또는 'https://'로 시작해야 합니다.", "Invalid address. Must start with 'http://' or 'https://'."), App.Title, Me, 16
        Exit Sub
    End If

    Dim idx%
    Dim FileName$
    Dim ServerName$
    FileName = Trim$(txtFileName.Text)
    If FolderExists(FileName) Then
        If Not (Right$(FileName, 1) = "\") Then FileName = FileName & "\"
        ServerName = FilterFilename(URLDecode(Split(URL, "/")(UBound(Split(URL, "/")))))
        If Replace(ServerName, " ", "") = "" Then ServerName = "download_" & CStr(Rnd * 1E+15)
        FileName = FileName & ServerName
    Else
        ServerName = FilterFilename(URLDecode(Split(URL, "/")(UBound(Split(URL, "/")))))
        If Replace(ServerName, " ", "") = "" Then
            ServerName = "download_" & CStr(Rnd * 1E+15)
        Else
            ServerName = CStr(Rnd * 1E+15) & "_" & ServerName
        End If
        FileName = fso.GetParentFolderName(txtFileName.Text) & "\"
        FileName = Replace(FileName, "\\", "\") & ServerName
    End If
    idx = lvBatchFiles.ListItems.Add(, , ServerName).Index
    lvBatchFiles.ListItems(idx).ListSubItems.Add , , FileName
    lvBatchFiles.ListItems(idx).ListSubItems.Add , , URL
    lvBatchFiles.ListItems(idx).ListSubItems.Add , , t("대기", "Queued")
    lvBatchFiles.ListItems(idx).Checked = -1
    If IsDownloading Or cmdStop.Enabled Or BatchStarted Then
        cmdStartBatch.Enabled = 0
        tygStartBatch.Enabled = 0
    Else
        cmdStartBatch.Enabled = -1
        tygStartBatch.Enabled = -1
    End If
End Sub

Private Sub cmdAddToQueue_Click()
    If Replace(txtURL.Text, " ", "") = "" Then
        Alert t("파일 주소를 입력하십시오.", "Specify the file URL."), App.Title, Me, 64
        Exit Sub
    End If
    On Error GoTo justadd
    Dim i%
    If lvBatchFiles.ListItems.Count Then
        For i = 1 To lvBatchFiles.ListItems.Count
            If lvBatchFiles.ListItems(i).ListSubItems(2).Text = Trim$(txtURL.Text) Then
                Alert t("해당 주소는 이미 대기열에 추가되었습니다.", "That URL is already added"), App.Title, Me, 64
                Exit Sub
            End If
        Next i
    End If
justadd:
    AddBatchURLs txtURL.Text
End Sub

Private Sub cmdBatch_Click()
    On Error Resume Next
    If Me.Height <= 6930 + PaddedBorderWidth * 15 * 2 Then
        cmdBatch.ImageList = imgDropdownReverse
        tygBatch.Caption = t("<< 일괄 처리", "<< Batch")
        lvBatchFiles.Visible = -1
        cmdAddToQueue.Visible = -1
        SetWindowSizeLimit Me.hWnd, Me.Width, Me.Width, 8220 + PaddedBorderWidth * 15 * 2, Screen.Height + 1200
        
        Dim formHeight As Integer
        formHeight = GetSetting("DownloadBooster", "UserData", "FormHeight", 8985)
        If formHeight < 8220 Then
            Me.Height = 8985 + PaddedBorderWidth * 15 * 2
        Else
            Me.Height = formHeight + PaddedBorderWidth * 15 * 2
        End If
    Else
        SaveSetting "DownloadBooster", "UserData", "FormHeight", Me.Height - PaddedBorderWidth * 15 * 2
        SetWindowSizeLimit Me.hWnd, Me.Width, Me.Width, 6930 + PaddedBorderWidth * 15, 6930 + PaddedBorderWidth * 15 * 2
        Me.Height = 6930 + PaddedBorderWidth * 15 * 2
        cmdBatch.ImageList = imgDropdown
        lvBatchFiles.Visible = 0
        cmdAddToQueue.Visible = 0
        tygBatch.Caption = t("일괄 처리 >>", "Batch >>")
    End If
End Sub

Private Sub cmdBrowse_Click()
    frmBrowse.Show vbModal, Me
End Sub

Private Sub cmdClear_Click()
    txtURL.Text = ""
End Sub

Private Sub cmdDecreaseThreads_Click()
    If trThreadCount.Value > trThreadCount.Min Then trThreadCount.Value = trThreadCount.Value - 1
    If trThreadCount.Value = trThreadCount.Min Then
        cmdDecreaseThreads.Enabled = 0
    Else
        cmdDecreaseThreads.Enabled = -1
    End If
End Sub

Private Sub cmdDelete_Click()
    If BatchStarted And CurrentBatchIdx = lvBatchFiles.SelectedItem.Index Then Exit Sub

    If BatchStarted And CurrentBatchIdx > lvBatchFiles.SelectedItem.Index Then
        CurrentBatchIdx = CurrentBatchIdx - 1
    End If
    lvBatchFiles.ListItems.Remove lvBatchFiles.SelectedItem.Index
    If lvBatchFiles.ListItems.Count < 1 Or cmdStop.Enabled Or BatchStarted Then
        cmdStartBatch.Enabled = 0
        tygStartBatch.Enabled = 0
        Exit Sub
    End If
    
    Dim i%
    Dim Enable As Boolean
    For i = 1 To lvBatchFiles.ListItems.Count
        If lvBatchFiles.ListItems(i).Checked Then
            Enable = True
            Exit For
        End If
    Next i
    If Not Enable Then
        cmdStartBatch.Enabled = 0
        tygStartBatch.Enabled = 0
    ElseIf Not IsDownloading Then
        cmdStartBatch.Enabled = -1
        tygStartBatch.Enabled = -1
    End If
End Sub

Sub StartDownload(ByVal URL As String, ByVal FileName As String)
    If BatchStarted Then
        If Not lvBatchFiles.ListItems(CurrentBatchIdx).Checked Then
            lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(3).Text = t("통과", "Skip")
            lvBatchFiles.ListItems(CurrentBatchIdx).ForeColor = &H808080
            lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(1).ForeColor = &H808080
            lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(2).ForeColor = &H808080
            lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(3).ForeColor = &H808080
            NextBatchDownload
            Exit Sub
        End If
        
        If lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(3).Text = t("완료", "Done") Then
            NextBatchDownload
            Exit Sub
        End If
    
        lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(3).Text = t("진행 중...", "Working...")
        lvBatchFiles.ListItems(CurrentBatchIdx).ForeColor = &HFF0000
        lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(1).ForeColor = &HFF0000
        lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(2).ForeColor = &HFF0000
        lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(3).ForeColor = &HFF0000
        
        On Error GoTo L1
        If lvBatchFiles.SelectedItem.Index = CurrentBatchIdx Then
            cmdDelete.Enabled = 0
            tygDelete.Enabled = 0
            cmdDeleteDropdown.Enabled = 0
        ElseIf lvBatchFiles.SelectedItem.Text <> "" And lvBatchFiles.SelectedItem.Selected Then
            cmdDelete.Enabled = -1
            tygDelete.Enabled = -1
            cmdDeleteDropdown.Enabled = -1
        Else
            cmdDelete.Enabled = 0
            tygDelete.Enabled = 0
            cmdDeleteDropdown.Enabled = 0
        End If
        GoTo L2
L1:
        cmdDelete.Enabled = 0
        tygDelete.Enabled = 0
        cmdDeleteDropdown.Enabled = 0
L2:
        On Error GoTo 0
    End If
    
    URL = Trim$(URL)
    FileName = Trim$(FileName)
    OnStart
    If Replace(FileName, " ", "") = "" Then
        FileName = Replace(CurDir.Path & "\", "\\", "\")
    End If
    Dim ServerName$
    If FolderExists(FileName) Then
        If Not (Right$(FileName, 1) = "\") Then FileName = FileName & "\"
        ServerName = FilterFilename(URLDecode(Split(URL, "/")(UBound(Split(URL, "/")))))
        If Replace(ServerName, " ", "") = "" Then ServerName = "download_" & CStr(Rnd * 1E+15)
        FileName = FileName & ServerName
    End If
    If Right$(FileName, 1) = "." Then FileName = Left$(FileName, Len(FileName) - 1) & "_"
    DownloadPath = FileName
    PrevDownloadedBytes = 0
    SpeedCount = 0
    lblFilename.Caption = fso.GetFilename(DownloadPath)
    If Len(lblFilename.Caption) > 22 Then lblFilename.Caption = Left$(lblFilename.Caption, 22) & "..."
    
    Dim ContinueDownload As Integer
    ContinueDownload = chkContinueDownload.Value
    If (Not BatchStarted) And chkContinueDownload.Value <> 1 Then
        Dim PrevPartialDownload As Boolean
        PrevPartialDownload = (trThreadCount.Value <= 1 And FileExists(FileName & ".part.tmp")) Or _
                              (trThreadCount.Value > 1 And FileExists(FileName & ".part_" & trThreadCount.Value & ".tmp") And (Not FileExists(FileName & ".part_" & (trThreadCount.Value + 1) & ".tmp")))
        If PrevPartialDownload Then
            Dim ContinueMsgboxResult As VbMsgBoxResult
            ContinueMsgboxResult = ConfirmCancel(t("기존에 다운로드 받다가 중지한 파일입니다. 다운로드받은 지점부터 이어서 받으시겠습니까?" & vbCrLf & "　[아니요]를 누를 경우 처음부터 다시 다운로드됩니다.", "This file was previously downloaded partially. Would you like to resume?" & vbCrLf & "  We will download from the start if you choose No."), App.Title, Me)
            If ContinueMsgboxResult = vbYes Then
                ContinueDownload = 1
            ElseIf ContinueMsgboxResult = vbCancel Then
                OnExit 999
                Exit Sub
            End If
        End If
    End If
    
    Dim NodePath$, ScriptPath$
    NodePath = GetSetting("DownloadBooster", "Options", "NodePath", "")
    ScriptPath = GetSetting("DownloadBooster", "Options", "ScriptPath", "")
    If NodePath = "" Then NodePath = CachePath & "node_v0_11_11.exe"
    If ScriptPath = "" Then ScriptPath = CachePath & "booster_v" & App.Major & "_" & App.Minor & "_" & App.Revision & ".js"
    SPResult = SP.Run("""" & NodePath & """ """ & ScriptPath & """ """ & Replace(Replace(URL, " ", "%20"), """", "%22") & """ """ & FileName & """ " & trThreadCount.Value & " " & GetSetting("DownloadBooster", "Options", "NoCleanup", 0) & " " & cbWhenExist.ListIndex & " " & ContinueDownload)
    Select Case SPResult
        Case SP_SUCCESS
            SP.ClosePipe
        Case SP_CREATEPIPEFAILED
            Alert t("다운로드 시작에 실패했습니다. 다운로더 프로세스로부터 정보를 받아올 수 없습니다. 디렉토리 설정에서 올바른 프로그램을 지정했는지 확인하십시오.", "Failed to receieve data from the downloader process. Check if the directory settings are valid."), App.Title, Me, 16
            If Not BatchStarted Then cmdGo.Enabled = -1
            tygGo.Enabled = cmdGo.Enabled
            cmdStop.Enabled = 0
            cmdStop.Left = Me.Width + 1200
            tygStop.Enabled = cmdStop.Enabled
            tygStop.Left = cmdStop.Left
            cmdGo.Enabled = -1
            cmdGo.Visible = -1
            tygGo.Enabled = cmdGo.Enabled
            OnStop False
        Case SP_CREATEPROCFAILED
            Alert t("다운로드 시작에 실패했습니다. 다운로더 프로세스를 생성할 수 없습니다. 디렉토리 설정에서 올바른 프로그램을 지정했는지 확인하십시오.", "Failed to create the downloader process. Check if the directory settings are valid."), App.Title, Me, 16
            If Not BatchStarted Then cmdGo.Enabled = -1
            tygGo.Enabled = cmdGo.Enabled
            cmdStop.Enabled = 0
            cmdStop.Left = Me.Width + 1200
            tygStop.Enabled = cmdStop.Enabled
            tygStop.Left = cmdStop.Left
            cmdGo.Enabled = -1
            cmdGo.Visible = -1
            tygGo.Enabled = cmdGo.Enabled
            OnStop False
    End Select
End Sub

Private Sub cmdDelete_DropDown()
    cmdDeleteDropdown_Click
End Sub

Private Sub cmdDeleteDropdown_Click()
    Me.PopupMenu mnuDeleteDropdown, , cmdDelete.Left, cmdDelete.Top + cmdDelete.Height
End Sub

Private Sub cmdDeleteDropdown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdDeleteDropdown_Click
End Sub

Private Sub cmdGo_Click()
    Dim SPResult As SP_RESULTS
    Dim TextLine As String
    
    If Replace(txtURL.Text, " ", "") = "" Then
        Alert t("파일 주소를 입력하십시오.", "Specify the file URL."), App.Title, Me, 64
        Exit Sub
    End If
    
    If Left$(txtURL.Text, 7) <> "http://" And Left$(txtURL.Text, 8) <> "https://" Then
        Alert t("주소가 올바르지 않습니다. 'http://' 또는 'https://'로 시작해야 합니다.", "Invalid address. Must start with 'http://' or 'https://'."), App.Title, Me, 16
        Exit Sub
    End If
    
    Dim SplittedPath() As String
    SplittedPath = Split(Trim$(txtFileName.Text), "\")
    Dim i%
    For i = LBound(SplittedPath) To UBound(SplittedPath)
        If Trim$(SplittedPath(i)) <> "" And Replace(Trim$(SplittedPath(i)), ".", "") = "" Then
            Alert t("저장 경로가 유효하지 않습니다.", "Invalid save path."), App.Title, Me, 16
            Exit Sub
        End If
    Next i
    
    If (Not FolderExists(Trim$(txtFileName.Text))) And (Not FolderExists(fso.GetParentFolderName(Trim$(txtFileName.Text)))) Then
        Alert t("저장 경로가 존재하지 않습니다.", "Save path does not exist."), App.Title, Me, 16
        Exit Sub
    End If

    Elapsed = 0
    timElapsed.Enabled = -1
    StartDownload txtURL.Text, txtFileName.Text
End Sub

Private Sub cmdIncreaseThreads_Click()
    If trThreadCount.Value < trThreadCount.Max Then trThreadCount.Value = trThreadCount.Value + 1
    If trThreadCount.Value = trThreadCount.Max Then
        cmdIncreaseThreads.Enabled = 0
    Else
        cmdIncreaseThreads.Enabled = -1
    End If
End Sub

Private Sub cmdOpen_Click()
    Shell "cmd /c start """" """ & DownloadPath & """"
End Sub

Private Sub cmdOpenBatch_Click()
    On Error Resume Next
    Shell "cmd /c start """" """ & lvBatchFiles.SelectedItem.ListSubItems(1).Text & """"
End Sub

Private Sub cmdOpenBatch_DropDown()
    cmdOpenDropdown_Click
End Sub

Private Sub cmdOpenDropdown_Click()
    Me.PopupMenu mnuOpenDropdown, , cmdOpenBatch.Left, cmdOpenBatch.Top + cmdOpenBatch.Height
End Sub

Private Sub cmdOpenDropdown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdOpenDropdown_Click
End Sub

Private Sub cmdOpenFolder_Click()
    Dim pth$
    pth = DownloadPath
    If DownloadPath = "" Then pth = txtFileName.Text
    If FolderExists(pth) Then
        Shell "cmd /c start """" explorer.exe """ & pth & """"
    Else
        Shell "cmd /c start """" explorer.exe """ & fso.GetParentFolderName(pth) & """"
    End If
End Sub

Private Sub cmdOpenFolderBatch_Click()
    cmdOpenFolder_Click
End Sub

Private Sub cmdOptions_Click()
    frmOptions.Show vbModal, Me
End Sub

Private Sub cmdStartBatch_Click()
    If lvBatchFiles.ListItems.Count <= 0 Then
        cmdStartBatch.Enabled = 0
        tygStartBatch.Enabled = 0
        Exit Sub
    End If
    
    Dim SplittedPath() As String
    SplittedPath = Split(Trim$(txtFileName.Text), "\")
    Dim i%
    For i = LBound(SplittedPath) To UBound(SplittedPath)
        If Trim$(SplittedPath(i)) <> "" And Replace(Trim$(SplittedPath(i)), ".", "") = "" Then
            Alert t("저장 경로가 유효하지 않습니다.", "Invalid save path."), App.Title, Me, 16
            Exit Sub
        End If
    Next i
    
    If (Not FolderExists(Trim$(txtFileName.Text))) And (Not FolderExists(fso.GetParentFolderName(Trim$(txtFileName.Text)))) Then
        Alert t("저장 경로가 존재하지 않습니다.", "Save path does not exist."), App.Title, Me, 16
        Exit Sub
    End If
    
    BatchErrorCount = 0
    CurrentBatchIdx = 1
    BatchStarted = True
    cmdStartBatch.Enabled = 0
    tygStartBatch.Enabled = 0
    cmdStopBatch.Enabled = -1
    tygStopbatch.Enabled = -1
    Elapsed = 0
    timElapsed.Enabled = -1
    chkOpenAfterComplete.Enabled = 0
    cmdOpen.Enabled = 0
    tygOpen.Enabled = 0
    cmdGo.Enabled = 0
    tygGo.Enabled = cmdGo.Enabled
    StartDownload lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(2), lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(1)
End Sub

Private Sub cmdStop_Click()
    Dim IsMarquee As Boolean
    IsMarquee = pbTotalProgressMarquee.Visible
    Dim ConfirmResult As VbMsgBoxResult
    If IsMarquee Or ResumeUnsupported Then
        ConfirmResult = ConfirmEx(t("다운로드를 중지하시겠습니까? 현재 파일은 이어받기가 지원되지 않으므로 처음부터 다시 다운로드받아야 합니다.", "Cancel download? Resuming is not supported for this file."), t("다운로드 취소", "Cancel download"), Me, 48)
    Else
        ConfirmResult = Confirm(t("다운로드를 중지하시겠습니까? 이어받기 기능을 통해 중단한 곳부터 계속 다운로드받을 수 있습니다.", "Cancel download? You can resume later."), t("다운로드 취소", "Cancel download"), Me)
    End If
    If ConfirmResult = vbYes Then
        Dim CurrentProgress As Integer
        CurrentProgress = pbTotalProgress.Value
        
        OnStop False
        cmdOpen.Enabled = 0
        tygOpen.Enabled = 0
        
        If IsMarquee Or (CurrentProgress > 0 And CurrentProgress < 100) Then
            Dim KillTemp As Boolean
            KillTemp = False
            If IsMarquee Or ResumeUnsupported Then
                KillTemp = True
            Else
                KillTemp = Confirm(t("나중에 계속 이어서 다운로드받을 수 있도록 다운로드한 데이타를 저장하시겠습니까?", "Would you like to keep the partially downloaded data to resume later?"), App.Title, Me) <> vbYes
            End If
            If KillTemp Then
                On Error Resume Next
                If trThreadCount.Value <= 1 Then
                    Kill DownloadPath & ".part.tmp"
                Else
                    Dim i%
                    For i = 1 To trThreadCount.Value
                        Kill DownloadPath & ".part_" & i & ".tmp"
                    Next i
                End If
            End If
        End If
    End If
End Sub

Private Sub cmdStopBatch_Click()
    Dim IsMarquee As Boolean
    IsMarquee = pbTotalProgressMarquee.Visible
    Dim ConfirmResult As VbMsgBoxResult
    If IsMarquee Or ResumeUnsupported Then
        ConfirmResult = ConfirmEx(t("다운로드를 중지하시겠습니까? 현재 파일은 이어받기가 지원되지 않으므로 처음부터 다시 다운로드받아야 합니다.", "Cancel download? Resuming is not supported for this file."), t("다운로드 취소", "Cancel download"), Me, 48)
    Else
        ConfirmResult = Confirm(t("다운로드를 중지하시겠습니까? 이어받기 기능을 통해 중단한 곳부터 계속 다운로드받을 수 있습니다.", "Cancel download? You can resume later."), t("다운로드 취소", "Cancel download"), Me)
    End If
    If ConfirmResult = vbYes Then
        Dim CurrentProgress As Integer
        CurrentProgress = pbTotalProgress.Value
        
        lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(3).Text = t("중지", "Stopped")
        lvBatchFiles.ListItems(CurrentBatchIdx).ForeColor = 255
        lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(1).ForeColor = 255
        lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(2).ForeColor = 255
        lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(3).ForeColor = 255
        BatchStarted = False
        CurrentBatchIdx = 1
        cmdStartBatch.Enabled = -1
        tygStartBatch.Enabled = -1
        cmdStopBatch.Enabled = 0
        tygStopbatch.Enabled = 0
        OnStop False
        cmdGo.Enabled = 0
        tygGo.Enabled = cmdGo.Enabled
        timElapsed.Enabled = 0
        sbStatusBar.Panels(3).Text = ""
        sbStatusBar.Panels(4).Text = ""
        chkOpenAfterComplete.Enabled = -1
        cmdGo.Enabled = -1
        tygGo.Enabled = cmdGo.Enabled
        
        If IsMarquee Or (CurrentProgress > 0 And CurrentProgress < 100) Then
            Dim KillTemp As Boolean
            KillTemp = False
            If IsMarquee Or ResumeUnsupported Then
                KillTemp = True
            Else
                KillTemp = Confirm(t("나중에 계속 이어서 다운로드받을 수 있도록 다운로드한 데이타를 저장하시겠습니까?", "Would you like to keep the partially downloaded data to resume later?"), App.Title, Me) <> vbYes
            End If
            If KillTemp Then
                On Error Resume Next
                If trThreadCount.Value <= 1 Then
                    Kill DownloadPath & ".part.tmp"
                Else
                    Dim i%
                    For i = 1 To trThreadCount.Value
                        Kill DownloadPath & ".part_" & i & ".tmp"
                    Next i
                End If
            End If
        End If
        
        If BatchErrorCount Then Alert t("하나 이상의 오류가 발생했습니다. 오류 코드 정보는 다음과 같습니다." & vbCrLf & vbCrLf & "1: 알 수 없는 오류가 발생했습니다. 유효하지 않은 주소를 입력했거나 프로그램 내부 오류입니다." & vbCrLf & "102: 주소나 파일 이름을 지정하지 않았습니다." & vbCrLf & "3: 저장 경로가 존재하지 않습니다." & vbCrLf & "104: 저장할 파일명이 사용 중입니다. 다른 이름을 선택하십시오." & vbCrLf & "106: 파일 서버가 다운로드 부스트를 지원하지 않습니다. 강도를 1로 변경해 보십시오." & vbCrLf & "107: 파일의 크기를 알 수 없어서 다운로드를 부스트할 수 없습니다. 강도를 1로 변경해 보십시오." & vbCrLf & "108: 서버가 요청을 거부했습니다. 서버 측 오류이거나 페이지가 존재하지 않거나 접근 권한이 없을 수 있습니다.", _
                                         "One or more errors have occurred." & vbCrLf & vbCrLf & "1: Network error" & vbCrLf & "103: Save path doesn't exist." & vbCrLf & "104: File name already exists" & vbCrLf & "106: Download boosting not supported. Try changing the thread count to 1." & vbCrLf & "107: Unable to boost download because the file size is not provided. Try changing the thread count to 1." & vbCrLf & "108: Server has denied your request. The file may not exist or have insufficient permissions to access it."), App.Title, Me, 48
    End If
End Sub

Sub SetBackgroundPosition(Optional ByVal ForceRefresh As Boolean = False)
    On Error Resume Next
    If imgBackground.Visible Then
        Select Case ImagePosition
            Case 0 '늘이기
                If imgBackground.Stretch <> True Then imgBackground.Stretch = True
                imgBackground.Width = Me.Width
                imgBackground.Height = Me.Height
            Case 1 '높이에 맞추기
                If imgBackground.Stretch <> True Then imgBackground.Stretch = True
                imgBackground.Height = Me.Height
                imgBackground.Width = imgBackground.Picture.Width / imgBackground.Picture.Height * Me.Height
            Case 2 '너비에 맞추기
                If imgBackground.Stretch <> True Then imgBackground.Stretch = True
                imgBackground.Width = Me.Width
                imgBackground.Height = imgBackground.Picture.Height / imgBackground.Picture.Width * Me.Width
            Case 3 '원본 크기
                If imgBackground.Stretch = True Then imgBackground.Stretch = False
        End Select
        If ImagePosition < 2 Or ForceRefresh Then
            On Error Resume Next
            Dim ctrl As Control
            For Each ctrl In Me.Controls
                If TypeName(ctrl) = "FrameW" Or TypeName(ctrl) = "CheckBoxW" Or TypeName(ctrl) = "OptionButtonW" Or TypeName(ctrl) = "CommandButtonW" Or TypeName(ctrl) = "Slider" Then ctrl.Refresh
            Next ctrl
            trThreadCount.VisualStyles = False
            trThreadCount.VisualStyles = True
        End If
    End If
End Sub

Sub SetBackgroundImage()
    On Error Resume Next
    If GetSetting("DownloadBooster", "Options", "UseBackgroundImage", 0) = 1 And Trim$(GetSetting("DownloadBooster", "Options", "BackgroundImagePath", "")) <> "" Then
        If LCase(Right$(GetSetting("DownloadBooster", "Options", "BackgroundImagePath", ""), 4)) = ".png" Then
            Set imgBackground.Picture = LoadPngIntoPictureWithAlpha(GetSetting("DownloadBooster", "Options", "BackgroundImagePath", ""))
        Else
            imgBackground.Picture = LoadPicture(GetSetting("DownloadBooster", "Options", "BackgroundImagePath", ""))
        End If
        imgBackground.Visible = -1
        pbProgressOuterContainer.BorderStyle = 1
        SetBackgroundPosition True
    Else
        imgBackground.Visible = 0
        pbProgressOuterContainer.BorderStyle = 0
    End If
    
    On Error Resume Next
    Dim ctrl As Control
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "FrameW" Or TypeName(ctrl) = "CheckBoxW" Or TypeName(ctrl) = "OptionButtonW" Or TypeName(ctrl) = "CommandButtonW" Or TypeName(ctrl) = "Slider" Then
            ctrl.Refresh
        End If
    Next ctrl
    trThreadCount.VisualStyles = False
    trThreadCount.VisualStyles = True
End Sub

Sub LoadLiveBadukSkin()
    If CInt(GetSetting("DownloadBooster", "Options", "EnableLiveBadukMemoSkin", 0)) Then
        imgTopLeft.Picture = LoadPngIntoPictureWithAlpha(CachePath & "topleft.png")
        imgTopRight.Picture = LoadPngIntoPictureWithAlpha(CachePath & "topright.png")
        imgTop.Picture = LoadPngIntoPictureWithAlpha(CachePath & "top.png")
        imgLeft.Picture = LoadPngIntoPictureWithAlpha(CachePath & "left.png")
        imgRight.Picture = LoadPngIntoPictureWithAlpha(CachePath & "right.png")
        imgBottom.Picture = LoadPngIntoPictureWithAlpha(CachePath & "bottom.png")
        imgBottomLeft.Picture = LoadPngIntoPictureWithAlpha(CachePath & "bottomleft.png")
        imgBottomRight.Picture = LoadPngIntoPictureWithAlpha(CachePath & "bottomright.png")
        imgCenter.Picture = LoadPngIntoPictureWithAlpha(CachePath & "center.png")
        
        imgTop.Width = 6495 - imgTopLeft.Width - imgTopRight.Width - 30
        imgBottom.Width = imgTop.Width
        imgLeft.Height = 4620 - imgBottomLeft.Height - imgTopLeft.Height - 30 + 240
        imgRight.Height = imgLeft.Height
        
        imgTopRight.Left = 6495 - imgTopRight.Width - 30 + 120
        imgBottomRight.Left = imgTopRight.Left
        imgBottomLeft.Top = 4620 - imgBottomLeft.Height - 30 + imgTop.Top + 240
        imgBottomRight.Top = imgBottomLeft.Top
        imgBottomRight.Left = imgTopRight.Left
        
        imgRight.Left = imgBottomRight.Left - 15
        imgBottom.Top = imgBottomLeft.Top
        
        imgCenter.Width = imgRight.Left - (imgLeft.Left + imgLeft.Width)
        imgCenter.Height = imgBottom.Top - (imgTop.Top + imgTop.Height)
        
        imgTopLeft.Visible = -1
        imgTopRight.Visible = -1
        imgTop.Visible = -1
        imgLeft.Visible = -1
        imgRight.Visible = -1
        imgBottom.Visible = -1
        imgBottomLeft.Visible = -1
        imgBottomRight.Visible = -1
        imgCenter.Visible = -1
        
        fTotal.Visible = 0
        Frame4.Visible = 0
        lblLBCaption.Visible = -1
        
        pbTotalProgress.Top = 1800 - 90
        pbTotalProgressMarquee.Top = 1800 - 90
    Else
        imgTopLeft.Visible = 0
        imgTopRight.Visible = 0
        imgTop.Visible = 0
        imgLeft.Visible = 0
        imgRight.Visible = 0
        imgBottom.Visible = 0
        imgBottomLeft.Visible = 0
        imgBottomRight.Visible = 0
        imgCenter.Visible = 0
        
        fTotal.Visible = -1
        Frame4.Visible = -1
        lblLBCaption.Visible = 0

        pbTotalProgress.Top = 1560
        pbTotalProgressMarquee.Top = 1560
    End If
End Sub

Private Sub Form_Load()
    On Error Resume Next

    ResumeUnsupported = False
    sbStatusBar.Panels(1).Text = t("준비", "Ready")
    Me.Caption = t(Me.Caption, "Download Booster") & " v" & App.Major & "." & App.Minor & "." & App.Revision
    TahomaAvailable = FontExists("Tahoma")
    
    ImagePosition = GetSetting("DownloadBooster", "Options", "ImagePosition", 1)
    SetBackgroundImage
    imgBackground.Width = Me.Width
    imgBackground.Height = Me.Height
    LoadLiveBadukSkin
    
    Dim Lft%
    Dim Top%
    Top = GetSetting("DownloadBooster", "UserData", "FormTop", -1)
    Lft = GetSetting("DownloadBooster", "UserData", "FormLeft", -1)
    If Top >= 0 And Lft >= 0 Then
        Me.Top = Top
        Me.Left = Lft
    End If
    
    Dim i%
    For i = 1 To lblDownloader.UBound
        lblDownloader(i).Caption = t("스레드", "Thread") & " " & i & ":"
        lblDownloader(i).BackStyle = 0
        pbProgress(i).Left = pbProgress(i).Left + 60
        pbProgress(i).Width = pbProgress(i).Width - 60
        pbProgressMarquee(i).Left = pbProgressMarquee(i).Left + 60
        pbProgressMarquee(i).Width = pbProgressMarquee(i).Width - 60
        pbProgressMarquee(i).MarqueeAnimation = 0
        pbProgressMarquee(i).Visible = 0
        lblPercentage(i).Caption = ""
        lblPercentage(i).BackStyle = 0
    Next i
    fDownloadInfo.Top = fThreadInfo.Top + 60
    fDownloadInfo.Left = fThreadInfo.Left
    fDownloadInfo.Width = fThreadInfo.Width '5925
    fDownloadInfo.Height = fThreadInfo.Height - 60
    
    Me.Width = 9450 + PaddedBorderWidth * 15 * 2
    cmdStop.Left = Me.Width + 1200
    tygStop.Left = cmdStop.Left
    
    If GetSetting("DownloadBooster", "UserData", "LastTab", 1) = 1 Then
        fTabDownload_Click
    Else
        fTabThreads_Click
    End If
    
    lvDummyScroll.Clear
    For i = 1 To 25
        lvDummyScroll.AddItem CStr(i)
    Next i
    lvDummyScroll.ListIndex = 0
    txtDummyScroll.Height = lvDummyScroll.Height
    
    trThreadCount.Value = GetSetting("DownloadBooster", "UserData", "ThreadCount", GetSetting("DownloadBooster", "Options", "ThreadCount", 1))
    trThreadCount_Scroll
    
    lvBatchFiles.ColumnHeaders.Add , "filename", t("파일 이름", "File Name")
    lvBatchFiles.ColumnHeaders.Add , "fullpath", t("전체 경로", "Full Path")
    lvBatchFiles.ColumnHeaders.Add , "url", t("파일 주소", "File URL")
    lvBatchFiles.ColumnHeaders.Add , "status", t("상태", "Status")
    lvBatchFiles.ColumnHeaders(1).Width = 2895
    lvBatchFiles.ColumnHeaders(2).Width = 0
    lvBatchFiles.ColumnHeaders(3).Width = 4495
    lvBatchFiles.ColumnHeaders(4).Width = 1105
    lvBatchFiles.ColumnHeaders(4).Alignment = LvwColumnHeaderAlignmentCenter
    Me.Height = 6930
    
    BatchStarted = False
    
    txtFileName.Text = GetSetting("DownloadBooster", "UserData", "SavePath", CurDir.Path)
    
    Me.Height = 6930 + PaddedBorderWidth * 15 * 2
    If GetSetting("DownloadBooster", "UserData", "BatchExpanded", 1) <> 0 Then
        cmdBatch_Click
    End If
    
    chkOpenAfterComplete.Value = GetSetting("DownloadBooster", "Options", "OpenWhenComplete", 0)
    chkOpenFolder.Value = GetSetting("DownloadBooster", "Options", "OpenFolderWhenComplete", 0)
    If GetSetting("DownloadBooster", "Options", "RememberURL", 0) Then
        txtURL.Text = GetSetting("DownloadBooster", "UserData", "FileURL", "")
        txtURL.SelStart = 0
        txtURL.SelLength = Len(txtURL.Text)
    End If
    chkPlaySound.Value = GetSetting("DownloadBooster", "Options", "PlaySound", 1)
    chkContinueDownload.Value = GetSetting("DownloadBooster", "Options", "ContinueDownload", 0)
    chkAutoRetry.Value = GetSetting("DownloadBooster", "Options", "AutoRetry", 0)
    
    cbWhenExist.Clear
    cbWhenExist.AddItem t("중단", "Abort")
    cbWhenExist.AddItem t("덮어쓰기", "Overwrite")
    cbWhenExist.AddItem t("이름 변경", "Rename")
    cbWhenExist.ListIndex = GetSetting("DownloadBooster", "Options", "WhenFileExists", 0)
    
    If WinVer >= 6.1 Then
        cmdOpenBatch.SplitButton = True
        cmdOpenBatch.Width = 1815
        cmdOpenDropdown.Visible = 0
        
        cmdDelete.SplitButton = True
        cmdDelete.Width = 1575
        cmdDeleteDropdown.Visible = 0
    End If

    '언어설정
    lblURL.Caption = t(lblURL.Caption, "File &address:")
    lblFilePath.Caption = t(lblFilePath.Caption, "Save &file to:")
    lblThreadCountLabel.Caption = t(lblThreadCountLabel.Caption, "Threads:")
    cmdClear.Caption = t(cmdClear.Caption, "Clear(&Y)")
    cmdBrowse.Caption = t(cmdBrowse.Caption, "&Browse...")
    fTotal.Caption = t(fTotal.Caption, " Total Progress ")
    fTabDownload.Caption = t(fTabDownload.Caption, " Total ")
    fTabThreads.Caption = t(fTabThreads.Caption, " Threads ")
    cmdOptions.Caption = t(cmdOptions.Caption, "More opt&ions...")
    cmdOpen.Caption = t(cmdOpen.Caption, "&Open")
    cmdOpenFolder.Caption = t(cmdOpenFolder.Caption, "Op&en folder")
    cmdGo.Caption = t(cmdGo.Caption, "&Download")
    cmdStop.Caption = t(cmdStop.Caption, "Sto&p")
    cmdAddToQueue.Caption = t(cmdAddToQueue.Caption, "Add to &queue")
    cmdBatch.Caption = t(cmdBatch.Caption, "Batc&h download")
    lblState.Caption = t(lblState.Caption, "Stopped")
    cmdOpenBatch.Caption = t(cmdOpenBatch.Caption, "Open(&W)")
    cmdAdd.Caption = t(cmdAdd.Caption, "Add U&RL")
    cmdDelete.Caption = t(cmdDelete.Caption, "Remo&ve")
    cmdStartBatch.Caption = t(cmdStartBatch.Caption, "&Start")
    cmdStopBatch.Caption = t(cmdStopBatch.Caption, "Stop(&Z)")
    Label8.Caption = t(Label8.Caption, "File name:")
    Label2.Caption = t(Label2.Caption, "Total:")
    Label3.Caption = t(Label3.Caption, "Recieved:")
    Label4.Caption = t(Label4.Caption, "Elapsed:")
    Label5.Caption = t(Label5.Caption, "Speed:")
    Label6.Caption = t(Label6.Caption, "Threads:")
    Label7.Caption = t(Label7.Caption, "Size per thread:")
    fOptions.Caption = t(fOptions.Caption, " Settings ")
    
    chkOpenAfterComplete.Caption = t(chkOpenAfterComplete.Caption, "Open when &complete")
    chkOpenFolder.Caption = t(chkOpenFolder.Caption, "Open fo&lder when done")
    chkPlaySound.Caption = t(chkPlaySound.Caption, "Beep when co&mplete")
    chkContinueDownload.Caption = t(chkContinueDownload.Caption, "Always resume(&J)")
    chkAutoRetry.Caption = t(chkAutoRetry.Caption, "Auto retry on error(&G)")
    
    Label1.Caption = t(Label1.Caption, "Exists(&K):")
    mnuAddItem.Caption = t(mnuAddItem.Caption, "&Add URL...")
    mnuClearBatch.Caption = t(mnuClearBatch.Caption, "&Clear list")
    mnuClearBatch2.Caption = t(mnuClearBatch.Caption, "&Clear list")
    mnuClearBatch3.Caption = t(mnuClearBatch.Caption, "&Clear list")
    mnuDeleteItem.Caption = t(mnuDeleteItem.Caption, "&Remove")
    mnuOpenFolder.Caption = t(mnuOpenFolder.Caption, "Open &folder")
    cmdAbout.Caption = t(cmdAbout.Caption, "Abo&ut application...")
    Label10.Caption = t(Label10.Caption, "Remaining:")
    lblOverlay(0).Caption = fOptions.Caption
    lblOverlay(1).Caption = fTotal.Caption
    lblLBCaption.Caption = t(lblLBCaption.Caption, "Progress")
    
    tygReset.Caption = t("초기화", "Clear")
    tygBrowse.Caption = t("찾아보기...", "Browse...")
    tygOptions.Caption = t("추가 옵션...", "More options...")
    tygAbout.Caption = t("프로그램 정보...", "About application...")
    tygOpenFolder.Caption = t("폴더 열기", "Open folder")
    tygAddToQueue.Caption = t("목록에 추가", "Add to queue")
    tygOpenBatch.Caption = t("열기", "Open")
    tygAdd.Caption = t("추가...", "Add...")
    tygDelete.Caption = t("제거", "Remove")
    tygStartBatch.Caption = t("시작", "Start")
    tygStopbatch.Caption = t("중지", "Stop")
    tygOpen.Caption = t("열기", "Open")
    tygGo.Caption = t("다운로드", "Download")
    tygStop.Caption = t("중지", "Stop")
    '언어설정끝
    
    If GetSetting("DownloadBooster", "Options", "DisableDWMWindow", DefaultDisableDWMWindow) = 1 Then DisableDWMWindow Me.hWnd
    If GetSetting("DownloadBooster", "Options", "ForeColor", -1) <> -1 Then
        pgSettingsBackground.Visible = -1
        chkOpenAfterComplete.Tag = "nobackcolorchange"
        chkOpenFolder.Tag = "nobackcolorchange"
        chkPlaySound.Tag = "nobackcolorchange"
        chkContinueDownload.Tag = "nobackcolorchange"
        chkAutoRetry.Tag = "nobackcolorchange"
        
        chkOpenAfterComplete.Transparent = 0
        chkOpenFolder.Transparent = 0
        chkPlaySound.Transparent = 0
        chkContinueDownload.Transparent = 0
        chkAutoRetry.Transparent = 0
    Else
        chkOpenAfterComplete.Transparent = -1
        chkOpenFolder.Transparent = -1
        chkPlaySound.Transparent = -1
        chkContinueDownload.Transparent = -1
        chkAutoRetry.Transparent = -1
    End If
    
    SetFormBackgroundColor Me
    
    If GetSetting("DownloadBooster", "Options", "ForeColor", -1) <> -1 Or GetSetting("DownloadBooster", "Options", "UseBackgroundImage", 0) = 1 Then
        For i = pgOverlay.LBound To pgOverlay.UBound
            pgOverlay(i).Visible = -1
            lblOverlay(i).Visible = -1
        Next i
        optTabDownload2.Transparent = 0
        optTabDownload2.BackColor = pgOverlay(0).BackColor
        optTabDownload2.Refresh
        optTabThreads2.Transparent = 0
        optTabThreads2.BackColor = pgOverlay(0).BackColor
        optTabThreads2.Refresh
        fTabDownload.Transparent = 0
        fTabDownload.BackColor = pgOverlay(0).BackColor
        fTabDownload.Refresh
        fTabThreads.Transparent = 0
        fTabThreads.BackColor = pgOverlay(0).BackColor
        fTabThreads.Refresh
    End If
    
    SetFont Me
End Sub

Private Sub Form_Resize()
    If Me.Height <= 6930 + PaddedBorderWidth * 15 * 2 Then Exit Sub
    If Me.Height - lvBatchFiles.Top - 1320 < 870 + PaddedBorderWidth * 15 * 2 Then Exit Sub
    lvBatchFiles.Height = Me.Height - PaddedBorderWidth * 15 * 2 - lvBatchFiles.Top - 1320
    cmdOpenBatch.Top = lvBatchFiles.Top + lvBatchFiles.Height + 45
    cmdOpenDropdown.Top = lvBatchFiles.Top + lvBatchFiles.Height + 45
    cmdAdd.Top = lvBatchFiles.Top + lvBatchFiles.Height + 45
    cmdDelete.Top = lvBatchFiles.Top + lvBatchFiles.Height + 45
    cmdDeleteDropdown.Top = lvBatchFiles.Top + lvBatchFiles.Height + 45
    cmdStartBatch.Top = lvBatchFiles.Top + lvBatchFiles.Height + 45
    cmdStopBatch.Top = lvBatchFiles.Top + lvBatchFiles.Height + 45
    tygOpenBatch.Top = cmdOpenBatch.Top
    tygAdd.Top = cmdAdd.Top
    tygDelete.Top = cmdDelete.Top
    tygStartBatch.Top = cmdStartBatch.Top
    tygStopbatch.Top = cmdStopBatch.Top
    SetBackgroundPosition
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If cmdStop.Enabled = -1 Or BatchStarted Then
        Dim IsMarquee As Boolean
        IsMarquee = pbTotalProgressMarquee.Visible
        Dim ConfirmResult As VbMsgBoxResult
        If IsMarquee Or ResumeUnsupported Then
            ConfirmResult = ConfirmEx(t("다운로드를 중지하시겠습니까? 현재 파일은 이어받기가 지원되지 않으므로 처음부터 다시 다운로드받아야 합니다.", "Cancel download? Resuming is not supported for this file."), t("다운로드 취소", "Cancel download"), Me, 48)
        Else
            ConfirmResult = Confirm(t("다운로드를 중지하시겠습니까? 이어받기 기능을 통해 중단한 곳부터 계속 다운로드받을 수 있습니다.", "Cancel download? You can resume later."), t("다운로드 취소", "Cancel download"), Me)
        End If
        If ConfirmResult <> vbYes Then
            Cancel = 1
            Exit Sub
        Else
            Dim CurrentProgress As Integer
            CurrentProgress = pbTotalProgress.Value
            
            BatchStarted = False
            SP.FinishChild 0, 0
            
            If IsMarquee Or (CurrentProgress > 0 And CurrentProgress < 100) Then
                Dim KillTemp As Boolean
                KillTemp = False
                If IsMarquee Or ResumeUnsupported Then
                    KillTemp = True
                Else
                    KillTemp = Confirm(t("나중에 계속 이어서 다운로드받을 수 있도록 다운로드한 데이타를 저장하시겠습니까?", "Would you like to keep the partially downloaded data to resume later?"), App.Title, Me) <> vbYes
                End If
                If KillTemp Then
                    On Error Resume Next
                    If trThreadCount.Value <= 1 Then
                        Kill DownloadPath & ".part.tmp"
                    Else
                        Dim i%
                        For i = 1 To trThreadCount.Value
                            Kill DownloadPath & ".part_" & i & ".tmp"
                        Next i
                    End If
                End If
            End If
        End If
    Else
        BatchStarted = False
        SP.FinishChild 0, 0
    End If
    
    SaveSetting "DownloadBooster", "UserData", "SavePath", Trim$(txtFileName.Text)
    SaveSetting "DownloadBooster", "UserData", "BatchExpanded", CInt(Me.Height > 6931) * -1
    
'    SaveSetting "DownloadBooster", "Options", "OpenWhenComplete", chkOpenAfterComplete.Value
'    SaveSetting "DownloadBooster", "Options", "OpenFolderWhenComplete", chkOpenFolder.Value
'    SaveSetting "DownloadBooster", "Options", "PlaySound", chkPlaySound.Value
'    SaveSetting "DownloadBooster", "Options", "ContinueDownload", chkContinueDownload.Value
'    SaveSetting "DownloadBooster", "Options", "AutoRetry", chkAutoRetry.Value
    
    SaveSetting "DownloadBooster", "Options", "WhenFileExists", cbWhenExist.ListIndex
    If GetSetting("DownloadBooster", "Options", "RememberURL", 0) Then
        SaveSetting "DownloadBooster", "UserData", "FileURL", Trim$(txtURL.Text)
    End If
    SaveSetting "DownloadBooster", "UserData", "FormTop", Me.Top
    SaveSetting "DownloadBooster", "UserData", "FormLeft", Me.Left
    If Me.Height >= 8220 Then SaveSetting "DownloadBooster", "UserData", "FormHeight", Me.Height - PaddedBorderWidth * 15 * 2
    SaveSetting "DownloadBooster", "UserData", "LastTab", (CInt(optTabThreads2.Value) * -1) + 1
    On Error Resume Next
    Unload ConfirmMsgBox
    Unload OKMsgBox
    Unload YesNoMsgBox
    Unload YesNoCancelMsgBox
    Unload frmBatchAdd
    Unload frmBrowse
    Unload frmOptions
    Unload frmGame
    Unload frmGameVista
    Unload frmGameWin95
    Unload frmGameWinXP
    Unload frmCustomBackground
End Sub

Private Sub fTabDownload_Click()
    optTabDownload2.Value = True
    optTabDownload2_Click
End Sub

Private Sub fTabDownload_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fTabDownload_Click
End Sub

Private Sub fTabThreads_Click()
    optTabThreads2.Value = True
    optTabThreads2_Click
End Sub

Private Sub fTabThreads_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fTabThreads_Click
End Sub

Private Sub lvBatchFiles_ContextMenu(ByVal X As Single, ByVal Y As Single)
    On Error GoTo ErrLn
    If lvBatchFiles.SelectedItem.Selected Then
        If cmdDelete.Enabled Then Me.PopupMenu mnuListContext
    Else
        GoTo ErrLn
    End If
    
    Exit Sub
ErrLn:
    Me.PopupMenu mnuListContext2
End Sub

Private Sub lvBatchFiles_ItemCheck(ByVal Item As LvwListItem, ByVal Checked As Boolean)
    If BatchStarted And Item.Index = CurrentBatchIdx And (Not Checked) Then
        Item.Checked = True
        Exit Sub
    End If
    
    If Not (BatchStarted And Item.Index = CurrentBatchIdx) Then
        If Not Checked Then
            Item.ListSubItems(3).Text = t("통과", "Skip")
            Item.ForeColor = &H808080
            Item.ListSubItems(1).ForeColor = &H808080
            Item.ListSubItems(2).ForeColor = &H808080
            Item.ListSubItems(3).ForeColor = &H808080
        Else
            Item.ListSubItems(3).Text = t("대기", "Queued")
            Item.ForeColor = 0
            Item.ListSubItems(1).ForeColor = 0
            Item.ListSubItems(2).ForeColor = 0
            Item.ListSubItems(3).ForeColor = 0
        End If
    End If
    
    If IsDownloading Or BatchStarted Then
        cmdStartBatch.Enabled = 0
        tygStartBatch.Enabled = 0
        Exit Sub
    End If
    
    If Checked Then
        cmdStartBatch.Enabled = -1
        tygStartBatch.Enabled = -1
        Exit Sub
    End If

    Dim i%
    Dim Enable As Boolean
    For i = 1 To lvBatchFiles.ListItems.Count
        If lvBatchFiles.ListItems(i).Checked Then
            Enable = True
            Exit For
        End If
    Next i
    If Not Enable Then
        cmdStartBatch.Enabled = 0
        tygStartBatch.Enabled = 0
    End If
End Sub

Private Sub lvBatchFiles_ItemDblClick(ByVal Item As LvwListItem, ByVal Button As Integer)
    On Error Resume Next
    If cmdOpenBatch.Enabled And Item.Selected Then
        cmdOpenBatch_Click
    End If
End Sub

Private Sub lvBatchFiles_ItemSelect(ByVal Item As LvwListItem, ByVal Selected As Boolean)
    If Selected Then
        If BatchStarted And Item.Index = CurrentBatchIdx Then
            cmdDelete.Enabled = 0
            tygDelete.Enabled = 0
            cmdDeleteDropdown.Enabled = 0
        Else
            cmdDelete.Enabled = -1
            tygDelete.Enabled = -1
            cmdDeleteDropdown.Enabled = -1
        End If
        
        If Item.ListSubItems(3).Text = t("완료", "Done") Then
            cmdOpenBatch.Enabled = -1
            cmdOpenDropdown.Enabled = -1
            tygOpenBatch.Enabled = -1
        Else
            cmdOpenBatch.Enabled = 0
            cmdOpenDropdown.Enabled = 0
            tygOpenBatch.Enabled = 0
        End If
    Else
        cmdDelete.Enabled = 0
        tygDelete.Enabled = 0
        cmdDeleteDropdown.Enabled = 0
        cmdOpenBatch.Enabled = 0
        cmdOpenDropdown.Enabled = 0
        tygOpenBatch.Enabled = 0
    End If
End Sub

Private Sub lvBatchFiles_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ErrLn2
    If KeyCode = 46 Then
        If lvBatchFiles.SelectedItem.Selected Then cmdDelete_Click
    End If
    Exit Sub
    
ErrLn2:
End Sub

Private Sub lvDummyScroll_Click()
    If lvDummyScroll.ListCount Then _
        lvDummyScroll.ListIndex = lvDummyScroll.TopIndex
End Sub

Private Sub lvDummyScroll_Scroll()
    If Not TahomaAvailable Then Exit Sub
    If lvDummyScroll.ListCount Then _
        lvDummyScroll.ListIndex = lvDummyScroll.TopIndex
    pbProgressContainer.Top = lvDummyScroll.TopIndex * 255 * -1 - (105 * lvDummyScroll.TopIndex)
End Sub

Private Sub mnuAddItem_Click()
    cmdAdd_Click
End Sub

Private Sub mnuClearBatch_Click()
    If lvBatchFiles.ListItems.Count Then
        If Confirm(t("대기열을 비우시겠습니까?", "Clear the queue?"), App.Title, Me) <> vbYes Then Exit Sub
        Dim i%
        i = 1
        Do While i <= lvBatchFiles.ListItems.Count
            If Not (BatchStarted And CurrentBatchIdx = i) Then
                lvBatchFiles.ListItems.Remove i
                If BatchStarted And CurrentBatchIdx > i Then
                    CurrentBatchIdx = CurrentBatchIdx - 1
                End If
            ElseIf BatchStarted And CurrentBatchIdx = i Then
                i = i + 1
            End If
        Loop
    End If
End Sub

Private Sub mnuClearBatch2_Click()
    mnuClearBatch_Click
End Sub

Private Sub mnuClearBatch3_Click()
    mnuClearBatch_Click
End Sub

Private Sub mnuDeleteItem_Click()
    cmdDelete_Click
End Sub

Private Sub mnuDeleteItem2_Click()
    mnuDeleteItem_Click
End Sub

Private Sub mnuOpen_Click()
    cmdOpenBatch_Click
End Sub

Private Sub mnuOpenFolder_Click()
    cmdOpenFolder_Click
End Sub

Private Sub optTabDownload2_Click()
    fDownloadInfo.Visible = -1
    fThreadInfo.Visible = 0
End Sub

Private Sub optTabDownload2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fTabDownload_Click
End Sub

Private Sub optTabThreads2_Click()
    fThreadInfo.Visible = -1
    fDownloadInfo.Visible = 0
End Sub

Private Sub optTabThreads2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fTabThreads_Click
End Sub

Private Sub SP_ChildFinished()
    If SP.Length > 0 Then OnData SP.GetData()
    OnExit SP.FinishChild(0)
End Sub

Private Sub SP_DataArrival(ByVal CharsTotal As Long)
    Do While SP.HasLine
        OnData SP.GetLine()
    Loop
End Sub

Private Sub SP_EOF(ByVal EOFType As SPEOF_TYPES)
    'Pick up any leftover output prior to EOF.
    If SP.Length > 0 Then OnData SP.GetData()
End Sub

Private Sub SP_Error(ByVal Number As Long, ByVal Source As String, CancelDisplay As Boolean)
    MsgBox "Error " & CStr(Number) & " in " & Source, _
           vbOKOnly Or vbExclamation, _
           Caption
    CancelDisplay = True
    SP.FinishChild 0
    OnStop
End Sub

Private Sub timElapsed_Timer()
    Elapsed = Elapsed + 1
    sbStatusBar.Panels(4).Text = FormatTime(Elapsed) & t(" 경과", " elapsed")
    lblElapsed.Caption = Replace(sbStatusBar.Panels(4).Text, " " & t("경과", "elapsed"), "")
End Sub

Private Sub trThreadCount_Change()
    trThreadCount_Scroll
    SaveSetting "DownloadBooster", "UserData", "ThreadCount", trThreadCount.Value
End Sub

Private Sub trThreadCount_KeyDown(KeyCode As Integer, Shift As Integer)
    trThreadCount_Scroll
End Sub

Private Sub trThreadCount_Scroll()
    If trThreadCount.Value = 1 Then
        lblThreadCount.Caption = "(" & t("부스트 없음", "No boost") & ")"
    Else
        lblThreadCount.Caption = "(" & trThreadCount.Value & t("개 스레드)", " threads)")
    End If
    Dim i%
    For i = 1 To trThreadCount.Value
        lblDownloader(i).Visible = -1
        pbProgress(i).Visible = -1
        lblPercentage(i).Visible = -1
        'If Not pbProgress(i).MarqueeAnimation Then pbProgress(i).MarqueeAnimation = True
    Next i
    For i = trThreadCount.Value + 1 To lblDownloader.UBound
        lblDownloader(i).Visible = 0
        pbProgress(i).Visible = 0
        lblPercentage(i).Visible = 0
    Next i
    
    If trThreadCount.Value - 10 > 0 Then
        vsProgressScroll.Max = trThreadCount.Value - 10
        vsProgressScroll.Enabled = -1
        
        '------------
        If lvDummyScroll.ListCount > trThreadCount.Value Then
            Do While lvDummyScroll.ListCount > trThreadCount.Value
                lvDummyScroll.RemoveItem lvDummyScroll.ListCount - 1
            Loop
            If lvDummyScroll.TopIndex > trThreadCount.Value - 10 Then _
                lvDummyScroll.TopIndex = trThreadCount.Value - 10
        ElseIf lvDummyScroll.ListCount < trThreadCount.Value Then
            Do While lvDummyScroll.ListCount < trThreadCount.Value
                lvDummyScroll.AddItem lvDummyScroll.ListCount + 1
            Loop
        End If
        
        If TahomaAvailable Then
            'txtDummyScroll.Visible = 0
            'fDummyScroll.Visible = 0
            lvDummyScroll.Visible = -1
        Else
            vsProgressScroll.Visible = -1
        End If
        'fThreadInfo.Width = 6255
    Else
        If vsProgressScroll.Max <> 0 Then vsProgressScroll.Max = 0
        If vsProgressScroll.Enabled Then vsProgressScroll.Enabled = 0
        
        '------------
        Do While lvDummyScroll.ListCount > 10
            lvDummyScroll.RemoveItem lvDummyScroll.ListCount - 1
        Loop
        
        'txtDummyScroll.Visible = -1
        'fDummyScroll.Visible = -1
        lvDummyScroll.ListIndex = 0
        lvDummyScroll.Visible = 0
        vsProgressScroll.Visible = 0
        'fThreadInfo.Width = 5925
        pbProgressContainer.Top = 0
    End If
    
    If trThreadCount.Value <= 1 Then
'        fDownloadInfo.Visible = -1
'        fThreadInfo.Visible = 0
'        optTabDownload2.Value = True
    Else
'        fThreadInfo.Visible = -1
'        fDownloadInfo.Visible = 0
'        optTabThreads2.Value = True
    End If
    
    If trThreadCount.Value = trThreadCount.Min Then
        cmdDecreaseThreads.Enabled = 0
    Else
        cmdDecreaseThreads.Enabled = -1
    End If
    If trThreadCount.Value = trThreadCount.Max Then
        cmdIncreaseThreads.Enabled = 0
    Else
        cmdIncreaseThreads.Enabled = -1
    End If
End Sub

Private Sub tygAbout_Click()
    cmdAbout_Click
End Sub

Private Sub tygAdd_Click()
    cmdAdd_Click
End Sub

Private Sub tygAddToQueue_Click()
    cmdAddToQueue_Click
End Sub

Private Sub tygBatch_Click()
    cmdBatch_Click
End Sub

Private Sub tygBrowse_Click()
    cmdBrowse_Click
End Sub

Private Sub tygDelete_Click()
    cmdDelete_Click
End Sub

Private Sub tygGo_Click()
    cmdGo_Click
End Sub

Private Sub tygOpen_Click()
    cmdOpen_Click
End Sub

Private Sub tygOpenBatch_Click()
    cmdOpenBatch_Click
End Sub

Private Sub tygOpenFolder_Click()
    cmdOpenFolder_Click
End Sub

Private Sub tygOptions_Click()
    cmdOptions_Click
End Sub

Private Sub tygReset_Click()
    cmdClear_Click
End Sub

Private Sub tygStartBatch_Click()
    cmdStartBatch_Click
End Sub

Private Sub tygStop_Click()
    cmdStop_Click
End Sub

Private Sub tygStopbatch_Click()
    cmdStopBatch_Click
End Sub

Private Sub vsProgressScroll_Change()
    vsProgressScroll_Scroll
End Sub

Private Sub vsProgressScroll_Scroll()
    'pbProgressContainer.Top = pbProgressOuterContainer.Height * vsProgressScroll.Value * -1 - (105 * vsProgressScroll.Value)
    pbProgressContainer.Top = vsProgressScroll.Value * 255 * -1 - (105 * vsProgressScroll.Value)
End Sub
