VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  '단일 고정
   Caption         =   "다운로드 부스터"
   ClientHeight    =   8925
   ClientLeft      =   45
   ClientTop       =   570
   ClientWidth     =   9360
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
   ScaleHeight     =   8925
   ScaleWidth      =   9360
   StartUpPosition =   3  'Windows 기본값
   Begin VB.Frame fDummyUI4 
      BorderStyle     =   0  '없음
      Height          =   345
      Left            =   9120
      TabIndex        =   116
      Top             =   6000
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame fDummyUI3 
      BorderStyle     =   0  '없음
      Height          =   2535
      Left            =   9180
      TabIndex        =   115
      Top             =   6000
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame fDummyUI1 
      BorderStyle     =   0  '없음
      Height          =   110
      Left            =   0
      TabIndex        =   114
      Top             =   8470
      Width           =   9375
   End
   Begin VB.Frame fDummyUI2 
      BorderStyle     =   0  '없음
      Height          =   2500
      Left            =   0
      TabIndex        =   113
      Top             =   6000
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.CheckBox chkPlaySound 
      Caption         =   "완료 시 신호음 재생(&N)"
      Height          =   255
      Left            =   6840
      TabIndex        =   12
      Top             =   2520
      Width           =   2175
   End
   Begin prjDownloadBooster.FrameW fTabThreads 
      Height          =   165
      Left            =   1560
      TabIndex        =   112
      Top             =   2055
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   291
      Caption         =   " 스레드 "
      Alignment       =   2
   End
   Begin prjDownloadBooster.FrameW fTabDownload 
      Height          =   165
      Left            =   690
      TabIndex        =   111
      Top             =   2055
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   291
      Caption         =   " 요약  "
      Alignment       =   2
   End
   Begin VB.OptionButton optTabThreads2 
      Height          =   255
      Left            =   1320
      TabIndex        =   109
      Top             =   2010
      Value           =   -1  'True
      Width           =   255
   End
   Begin VB.Frame Frame6 
      Height          =   420
      Left            =   2760
      TabIndex        =   110
      Top             =   1920
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.OptionButton optTabDownload2 
      Height          =   255
      Left            =   435
      TabIndex        =   108
      Top             =   2010
      Width           =   255
   End
   Begin VB.Frame Frame5 
      Height          =   420
      Left            =   360
      TabIndex        =   107
      Top             =   1920
      Visible         =   0   'False
      Width           =   1125
   End
   Begin prjDownloadBooster.OptionButtonW optTabThreads 
      Height          =   255
      Left            =   1725
      TabIndex        =   106
      Top             =   2010
      Visible         =   0   'False
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   450
      Transparent     =   -1  'True
   End
   Begin prjDownloadBooster.OptionButtonW optTabDownload 
      Height          =   255
      Left            =   495
      TabIndex        =   105
      Top             =   2010
      Visible         =   0   'False
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   450
      Transparent     =   -1  'True
   End
   Begin VB.CommandButton cmdTabThreads 
      Caption         =   "    스레드"
      Height          =   330
      Left            =   2760
      TabIndex        =   104
      Top             =   1980
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdTabDownload 
      Caption         =   "    요약"
      Height          =   330
      Left            =   450
      TabIndex        =   103
      Top             =   1980
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  '없음
      Height          =   375
      Left            =   2520
      TabIndex        =   102
      Top             =   1980
      Visible         =   0   'False
      Width           =   135
   End
   Begin prjDownloadBooster.TabStrip tsTabs 
      Height          =   315
      Left            =   360
      TabIndex        =   100
      Top             =   1995
      Visible         =   0   'False
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   556
      MultiRow        =   0   'False
      Style           =   2
      TabWidthStyle   =   2
      TabFixedWidth   =   53
      TabMinWidth     =   45
      TabScrollWheel  =   0   'False
      InitTabs        =   "frmMain.frx":030A
   End
   Begin VB.Frame fDownloadInfo 
      BorderStyle     =   0  '없음
      Caption         =   " "
      Height          =   2055
      Left            =   1320
      TabIndex        =   84
      Top             =   2880
      Visible         =   0   'False
      Width           =   3495
      Begin VB.Label lblTotalSizeThread 
         Caption         =   "-"
         Height          =   255
         Left            =   1320
         TabIndex        =   120
         Top             =   1800
         Width           =   4335
      End
      Begin VB.Label Label7 
         Caption         =   "스레드당 크기:"
         Height          =   255
         Left            =   0
         TabIndex        =   119
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label lblThreadCount2 
         Caption         =   "-"
         Height          =   255
         Left            =   1320
         TabIndex        =   118
         Top             =   1440
         Width           =   4335
      End
      Begin VB.Label Label6 
         Caption         =   "스레드 수:"
         Height          =   255
         Left            =   0
         TabIndex        =   117
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "속도:"
         Height          =   255
         Left            =   0
         TabIndex        =   99
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblSpeed 
         Caption         =   "-"
         Height          =   255
         Left            =   1320
         TabIndex        =   98
         Top             =   1080
         Width           =   4335
      End
      Begin VB.Label lblElapsed 
         Caption         =   "-"
         Height          =   255
         Left            =   1320
         TabIndex        =   90
         Top             =   720
         Width           =   4335
      End
      Begin VB.Label Label4 
         Caption         =   "경과 시간:"
         Height          =   255
         Left            =   0
         TabIndex        =   89
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblDownloadedBytes 
         Caption         =   "-"
         Height          =   255
         Left            =   1320
         TabIndex        =   88
         Top             =   360
         Width           =   4335
      End
      Begin VB.Label Label3 
         Caption         =   "받은 크기:"
         Height          =   255
         Left            =   0
         TabIndex        =   87
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblTotalBytes 
         Caption         =   "-"
         Height          =   255
         Left            =   1320
         TabIndex        =   86
         Top             =   0
         Width           =   4335
      End
      Begin VB.Label Label2 
         Caption         =   "총 크기:"
         Height          =   255
         Left            =   0
         TabIndex        =   85
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.Frame fThreadInfo 
      BorderStyle     =   0  '없음
      Caption         =   " 스레드 현황 "
      Height          =   3495
      Left            =   360
      TabIndex        =   17
      Top             =   2310
      Width           =   6015
      Begin VB.VScrollBar vsProgressScroll 
         Height          =   3495
         Left            =   5760
         Max             =   15
         TabIndex        =   29
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Frame fDummyScroll 
         BorderStyle     =   0  '없음
         Height          =   3495
         Left            =   5760
         TabIndex        =   96
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox pbProgressOuterContainer 
         BorderStyle     =   0  '없음
         Height          =   3495
         Left            =   0
         ScaleHeight     =   3495
         ScaleWidth      =   5775
         TabIndex        =   30
         Top             =   0
         Width           =   5775
         Begin VB.PictureBox pbProgressContainer 
            BorderStyle     =   0  '없음
            Height          =   9015
            Left            =   0
            ScaleHeight     =   9015
            ScaleWidth      =   5775
            TabIndex        =   31
            Top             =   0
            Width           =   5775
            Begin prjDownloadBooster.ProgressBar pbProgress 
               Height          =   255
               Index           =   1
               Left            =   840
               Top             =   0
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
               TabIndex        =   81
               Top             =   8685
               Width           =   855
            End
            Begin VB.Label lblPercentage 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "(100%)"
               Height          =   255
               Index           =   25
               Left            =   5040
               TabIndex        =   80
               Top             =   8700
               Width           =   615
            End
            Begin VB.Label lblDownloader 
               Caption         =   "스레드 0:"
               Height          =   255
               Index           =   24
               Left            =   0
               TabIndex        =   79
               Top             =   8325
               Width           =   855
            End
            Begin VB.Label lblPercentage 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "(100%)"
               Height          =   255
               Index           =   24
               Left            =   5040
               TabIndex        =   78
               Top             =   8325
               Width           =   615
            End
            Begin VB.Label lblDownloader 
               Caption         =   "스레드 0:"
               Height          =   255
               Index           =   23
               Left            =   0
               TabIndex        =   77
               Top             =   7965
               Width           =   855
            End
            Begin VB.Label lblPercentage 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "(100%)"
               Height          =   255
               Index           =   23
               Left            =   5040
               TabIndex        =   76
               Top             =   7965
               Width           =   615
            End
            Begin VB.Label lblDownloader 
               Caption         =   "스레드 0:"
               Height          =   255
               Index           =   22
               Left            =   0
               TabIndex        =   75
               Top             =   7605
               Width           =   855
            End
            Begin VB.Label lblPercentage 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "(100%)"
               Height          =   255
               Index           =   22
               Left            =   5040
               TabIndex        =   74
               Top             =   7605
               Width           =   615
            End
            Begin VB.Label lblDownloader 
               Caption         =   "스레드 0:"
               Height          =   255
               Index           =   21
               Left            =   0
               TabIndex        =   73
               Top             =   7245
               Width           =   855
            End
            Begin VB.Label lblPercentage 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "(100%)"
               Height          =   255
               Index           =   21
               Left            =   5040
               TabIndex        =   72
               Top             =   7245
               Width           =   615
            End
            Begin VB.Label lblDownloader 
               Caption         =   "스레드 0:"
               Height          =   255
               Index           =   20
               Left            =   0
               TabIndex        =   71
               Top             =   6885
               Width           =   855
            End
            Begin VB.Label lblPercentage 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "(100%)"
               Height          =   255
               Index           =   20
               Left            =   5040
               TabIndex        =   70
               Top             =   6885
               Width           =   615
            End
            Begin VB.Label lblDownloader 
               Caption         =   "스레드 0:"
               Height          =   255
               Index           =   19
               Left            =   0
               TabIndex        =   69
               Top             =   6525
               Width           =   855
            End
            Begin VB.Label lblPercentage 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "(100%)"
               Height          =   255
               Index           =   19
               Left            =   5040
               TabIndex        =   68
               Top             =   6525
               Width           =   615
            End
            Begin VB.Label lblDownloader 
               Caption         =   "스레드 0:"
               Height          =   255
               Index           =   18
               Left            =   0
               TabIndex        =   67
               Top             =   6165
               Width           =   855
            End
            Begin VB.Label lblPercentage 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "(100%)"
               Height          =   255
               Index           =   18
               Left            =   5040
               TabIndex        =   66
               Top             =   6165
               Width           =   615
            End
            Begin VB.Label lblDownloader 
               Caption         =   "스레드 0:"
               Height          =   255
               Index           =   17
               Left            =   0
               TabIndex        =   65
               Top             =   5805
               Width           =   855
            End
            Begin VB.Label lblPercentage 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "(100%)"
               Height          =   255
               Index           =   17
               Left            =   5040
               TabIndex        =   64
               Top             =   5805
               Width           =   615
            End
            Begin VB.Label lblDownloader 
               Caption         =   "스레드 0:"
               Height          =   255
               Index           =   16
               Left            =   0
               TabIndex        =   63
               Top             =   5445
               Width           =   855
            End
            Begin VB.Label lblPercentage 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "(100%)"
               Height          =   255
               Index           =   16
               Left            =   5040
               TabIndex        =   62
               Top             =   5445
               Width           =   615
            End
            Begin VB.Label lblDownloader 
               Caption         =   "스레드 0:"
               Height          =   255
               Index           =   15
               Left            =   0
               TabIndex        =   61
               Top             =   5085
               Width           =   855
            End
            Begin VB.Label lblPercentage 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "(100%)"
               Height          =   255
               Index           =   15
               Left            =   5040
               TabIndex        =   60
               Top             =   5085
               Width           =   615
            End
            Begin VB.Label lblDownloader 
               Caption         =   "스레드 0:"
               Height          =   255
               Index           =   14
               Left            =   0
               TabIndex        =   59
               Top             =   4725
               Width           =   855
            End
            Begin VB.Label lblPercentage 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "(100%)"
               Height          =   255
               Index           =   14
               Left            =   5040
               TabIndex        =   58
               Top             =   4725
               Width           =   615
            End
            Begin VB.Label lblDownloader 
               Caption         =   "스레드 0:"
               Height          =   255
               Index           =   13
               Left            =   0
               TabIndex        =   57
               Top             =   4365
               Width           =   855
            End
            Begin VB.Label lblPercentage 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "(100%)"
               Height          =   255
               Index           =   13
               Left            =   5040
               TabIndex        =   56
               Top             =   4365
               Width           =   615
            End
            Begin VB.Label lblDownloader 
               Caption         =   "스레드 0:"
               Height          =   255
               Index           =   12
               Left            =   0
               TabIndex        =   55
               Top             =   4005
               Width           =   855
            End
            Begin VB.Label lblPercentage 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "(100%)"
               Height          =   255
               Index           =   12
               Left            =   5040
               TabIndex        =   54
               Top             =   4005
               Width           =   615
            End
            Begin VB.Label lblDownloader 
               Caption         =   "스레드 0:"
               Height          =   255
               Index           =   11
               Left            =   0
               TabIndex        =   53
               Top             =   3645
               Width           =   855
            End
            Begin VB.Label lblPercentage 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "(100%)"
               Height          =   255
               Index           =   11
               Left            =   5040
               TabIndex        =   52
               Top             =   3645
               Width           =   615
            End
            Begin VB.Label lblDownloader 
               Caption         =   "스레드 0:"
               Height          =   255
               Index           =   10
               Left            =   0
               TabIndex        =   51
               Top             =   3285
               Width           =   855
            End
            Begin VB.Label lblPercentage 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "(100%)"
               Height          =   255
               Index           =   10
               Left            =   5040
               TabIndex        =   50
               Top             =   3285
               Width           =   615
            End
            Begin VB.Label lblDownloader 
               Caption         =   "스레드 0:"
               Height          =   255
               Index           =   9
               Left            =   0
               TabIndex        =   49
               Top             =   2925
               Width           =   855
            End
            Begin VB.Label lblPercentage 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "(100%)"
               Height          =   255
               Index           =   9
               Left            =   5040
               TabIndex        =   48
               Top             =   2925
               Width           =   615
            End
            Begin VB.Label lblDownloader 
               Caption         =   "스레드 0:"
               Height          =   255
               Index           =   8
               Left            =   0
               TabIndex        =   47
               Top             =   2565
               Width           =   855
            End
            Begin VB.Label lblPercentage 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "(100%)"
               Height          =   255
               Index           =   8
               Left            =   5040
               TabIndex        =   46
               Top             =   2565
               Width           =   615
            End
            Begin VB.Label lblDownloader 
               Caption         =   "스레드 0:"
               Height          =   255
               Index           =   7
               Left            =   0
               TabIndex        =   45
               Top             =   2205
               Width           =   855
            End
            Begin VB.Label lblPercentage 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "(100%)"
               Height          =   255
               Index           =   7
               Left            =   5040
               TabIndex        =   44
               Top             =   2205
               Width           =   615
            End
            Begin VB.Label lblDownloader 
               Caption         =   "스레드 0:"
               Height          =   255
               Index           =   6
               Left            =   0
               TabIndex        =   43
               Top             =   1845
               Width           =   855
            End
            Begin VB.Label lblPercentage 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "(100%)"
               Height          =   255
               Index           =   6
               Left            =   5040
               TabIndex        =   42
               Top             =   1845
               Width           =   615
            End
            Begin VB.Label lblDownloader 
               Caption         =   "스레드 0:"
               Height          =   255
               Index           =   5
               Left            =   0
               TabIndex        =   41
               Top             =   1485
               Width           =   855
            End
            Begin VB.Label lblPercentage 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "(100%)"
               Height          =   255
               Index           =   5
               Left            =   5040
               TabIndex        =   40
               Top             =   1485
               Width           =   615
            End
            Begin VB.Label lblDownloader 
               Caption         =   "스레드 0:"
               Height          =   255
               Index           =   4
               Left            =   0
               TabIndex        =   39
               Top             =   1125
               Width           =   855
            End
            Begin VB.Label lblPercentage 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "(100%)"
               Height          =   255
               Index           =   4
               Left            =   5040
               TabIndex        =   38
               Top             =   1125
               Width           =   615
            End
            Begin VB.Label lblDownloader 
               Caption         =   "스레드 0:"
               Height          =   255
               Index           =   3
               Left            =   0
               TabIndex        =   37
               Top             =   765
               Width           =   855
            End
            Begin VB.Label lblPercentage 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "(100%)"
               Height          =   255
               Index           =   3
               Left            =   5040
               TabIndex        =   36
               Top             =   765
               Width           =   615
            End
            Begin VB.Label lblDownloader 
               Caption         =   "스레드 0:"
               Height          =   255
               Index           =   2
               Left            =   0
               TabIndex        =   35
               Top             =   405
               Width           =   855
            End
            Begin VB.Label lblPercentage 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "(100%)"
               Height          =   255
               Index           =   2
               Left            =   5040
               TabIndex        =   34
               Top             =   405
               Width           =   615
            End
            Begin VB.Label lblDownloader 
               Caption         =   "스레드 0:"
               Height          =   255
               Index           =   1
               Left            =   0
               TabIndex        =   33
               Top             =   45
               Width           =   855
            End
            Begin VB.Label lblPercentage 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "(100%)"
               Height          =   255
               Index           =   1
               Left            =   5040
               TabIndex        =   32
               Top             =   45
               Width           =   615
            End
         End
      End
      Begin VB.TextBox txtDummyScroll 
         BorderStyle     =   0  '없음
         Enabled         =   0   'False
         Height          =   3450
         Left            =   5640
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   94
         TabStop         =   0   'False
         Top             =   15
         Visible         =   0   'False
         Width           =   375
      End
      Begin prjDownloadBooster.ListBoxW lvDummyScroll 
         Height          =   3450
         Left            =   5400
         TabIndex        =   93
         TabStop         =   0   'False
         Top             =   15
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
   Begin prjDownloadBooster.ListView lvBatchFiles 
      Height          =   1635
      Left            =   240
      TabIndex        =   21
      Top             =   6390
      Visible         =   0   'False
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   2884
      View            =   3
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      LabelEdit       =   2
      Checkboxes      =   -1  'True
      HideSelection   =   0   'False
      ClickableColumnHeaders=   0   'False
      AutoSelectFirstItem=   0   'False
   End
   Begin VB.CheckBox chkRememberURL 
      Caption         =   "파일 주소 기억(&M)"
      Height          =   255
      Left            =   6840
      TabIndex        =   11
      Top             =   2280
      Width           =   2055
   End
   Begin VB.DirListBox CurDir 
      Height          =   510
      Left            =   6720
      TabIndex        =   95
      Top             =   5400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdIncreaseThreads 
      Caption         =   ">"
      Height          =   375
      Left            =   7200
      TabIndex        =   92
      TabStop         =   0   'False
      Top             =   765
      Width           =   375
   End
   Begin VB.CommandButton cmdDecreaseThreads 
      Caption         =   "<"
      Height          =   375
      Left            =   1560
      TabIndex        =   91
      TabStop         =   0   'False
      Top             =   765
      Width           =   375
   End
   Begin VB.ComboBox cbWhenExist 
      Height          =   300
      Left            =   7590
      Style           =   2  '드롭다운 목록
      TabIndex        =   14
      Top             =   2790
      Width           =   1425
   End
   Begin VB.CheckBox chkOpenAfterComplete 
      Caption         =   "완료 후 열기(&C)"
      Height          =   255
      Left            =   6840
      TabIndex        =   9
      Top             =   1800
      Width           =   1935
   End
   Begin VB.CheckBox chkNoCleanup 
      Caption         =   "조각 파일 삭제 안 함(&D)"
      Height          =   255
      Left            =   6840
      TabIndex        =   8
      Top             =   1560
      Width           =   2250
   End
   Begin VB.CheckBox chkOpenFolder 
      Caption         =   "완료 후 폴더 열기(&L)"
      Height          =   255
      Left            =   6840
      TabIndex        =   10
      Top             =   2040
      Width           =   2055
   End
   Begin VB.CommandButton cmdOpenBatch 
      Caption         =   "파일 열기(&Q)"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   22
      Top             =   8070
      Width           =   1575
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "지우기(&Y)"
      Height          =   300
      Left            =   7680
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   105
      Width           =   1455
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "추가(&R)..."
      Height          =   375
      Left            =   2520
      TabIndex        =   23
      Top             =   8070
      Width           =   1575
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "삭제(&V)"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4200
      TabIndex        =   24
      Top             =   8070
      Width           =   1575
   End
   Begin VB.CommandButton cmdStopBatch 
      Caption         =   "중지(&Z)"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7560
      TabIndex        =   26
      Top             =   8070
      Width           =   1575
   End
   Begin VB.CommandButton cmdStartBatch 
      Caption         =   "시작(&A)"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5880
      TabIndex        =   25
      Top             =   8070
      Width           =   1575
   End
   Begin VB.CommandButton cmdBatch 
      Caption         =   "일괄 처리(&W) >>"
      Height          =   375
      Left            =   7320
      TabIndex        =   20
      Top             =   5520
      Width           =   1815
   End
   Begin VB.Frame fTotal 
      Caption         =   " 전체 다운로드 진행률 "
      Height          =   615
      Left            =   240
      TabIndex        =   82
      Top             =   1320
      Width           =   6255
      Begin prjDownloadBooster.ProgressBar pbTotalProgress 
         Height          =   255
         Left            =   840
         Top             =   240
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   450
         Step            =   10
         MarqueeAnimation=   -1  'True
         MarqueeSpeed    =   35
      End
      Begin VB.Label lblState 
         Caption         =   "중지됨"
         Height          =   255
         Left            =   120
         TabIndex        =   83
         Top             =   285
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " 옵션 "
      Height          =   1905
      Left            =   6720
      TabIndex        =   28
      Top             =   1320
      Width           =   2415
      Begin VB.Label Label1 
         Caption         =   "중복(&K):"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1530
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "파일 열기(&O)"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7320
      TabIndex        =   15
      Top             =   3540
      Width           =   1815
   End
   Begin VB.CommandButton cmdOpenFolder 
      Caption         =   "폴더 열기(&E)"
      Height          =   375
      Left            =   7320
      TabIndex        =   16
      Top             =   3960
      Width           =   1815
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "중지(&P)"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7320
      TabIndex        =   19
      Top             =   4980
      Width           =   1815
   End
   Begin prjDownloadBooster.StatusBar sbStatusBar 
      Align           =   2  '아래 맞춤
      Height          =   330
      Left            =   0
      Top             =   8595
      Width           =   9360
      _ExtentX        =   16510
      _ExtentY        =   582
      InitPanels      =   "frmMain.frx":03D6
   End
   Begin VB.Timer timElapsed 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6720
      Top             =   4440
   End
   Begin prjDownloadBooster.Slider trThreadCount 
      Height          =   495
      Left            =   1935
      TabIndex        =   6
      Top             =   750
      Width           =   5265
      _ExtentX        =   9287
      _ExtentY        =   873
      Min             =   1
      Max             =   25
      Value           =   1
      TickFrequency   =   2
      TipSide         =   1
      SelStart        =   1
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "찾아보기(&B)..."
      Height          =   300
      Left            =   7680
      TabIndex        =   4
      Top             =   465
      Width           =   1455
   End
   Begin VB.TextBox txtFileName 
      Height          =   270
      Left            =   1560
      TabIndex        =   3
      Top             =   480
      Width           =   6015
   End
   Begin VB.TextBox txtURL 
      Height          =   270
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   6015
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "다운로드(&S)"
      Height          =   375
      Left            =   7320
      TabIndex        =   18
      Top             =   4560
      Width           =   1815
   End
   Begin VB.Frame fBatchDownload 
      Caption         =   "일괄 다운로드 "
      Height          =   2415
      Left            =   120
      TabIndex        =   97
      Top             =   6120
      Width           =   9135
   End
   Begin VB.Frame Frame4 
      Caption         =   "                               "
      Height          =   3885
      Left            =   240
      TabIndex        =   101
      Top             =   2040
      Width           =   6255
   End
   Begin VB.Label lblThreadCount 
      Caption         =   "(일반 다운로드)"
      Height          =   255
      Left            =   7680
      TabIndex        =   7
      Top             =   870
      Width           =   1455
   End
   Begin VB.Label lblThreadCountLabel 
      Caption         =   "강도(&T):"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   870
      Width           =   1215
   End
   Begin VB.Label lblFilePath 
      Caption         =   "저장 경로(&F):"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   510
      Width           =   1215
   End
   Begin VB.Label lblURL 
      Caption         =   "파일 주소(&U):"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   150
      Width           =   1215
   End
   Begin prjDownloadBooster.ShellPipe SP 
      Left            =   6720
      Top             =   4920
      _ExtentX        =   635
      _ExtentY        =   635
   End
   Begin VB.Menu mnuListContext 
      Caption         =   "mnuListContext"
      Visible         =   0   'False
      Begin VB.Menu mnuDeleteItem 
         Caption         =   "삭제(&D)"
      End
   End
   Begin VB.Menu mnuListContext2 
      Caption         =   "mnuListContext2"
      Visible         =   0   'False
      Begin VB.Menu mnuAddItem 
         Caption         =   "추가(&A)..."
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

Sub OnData(Data As String)
    Dim output$
    Dim idx%
    Dim progress%
    Dim DownloadedBytes As Double
    If Left$(Data, 7) = "STATUS " Then
        Select Case Replace(Right$(Data, Len(Data) - 7), " ", "")
            Case "CHECKREDIRECT"
                sbStatusBar.Panels(1).Text = "리다이렉트 확인 중..."
            Case "CHECKFILE"
                sbStatusBar.Panels(1).Text = "가용성 확인 중..."
            Case "DOWNLOADING"
                sbStatusBar.Panels(1).Text = "다운로드 중..."
            Case "MERGING"
                sbStatusBar.Panels(1).Text = "파일 조각 결합 중..."
                pbTotalProgress.Scrolling = PrbScrollingMarquee
            Case "COMPLETE"
                sbStatusBar.Panels(1).Text = "완료"
                sbStatusBar.Panels(2).Text = ""
                sbStatusBar.Panels(3).Text = ""
                sbStatusBar.Panels(4).Text = ""
                pbTotalProgress.Scrolling = PrbScrollingStandard
                pbTotalProgress.Value = 100
        End Select
    ElseIf Left$(Data, 5) = "DATA " Then
        output = Right$(Data, Len(Data) - 5)
        idx = CInt(Split(output, ",")(0))
        If CLng(Split(output, ",")(1)) > 100 Then
            progress = -1
        Else
            progress = CInt(Split(output, ",")(1))
        End If
        If progress < 0 Then
            If pbProgress(idx).Scrolling <> PrbScrollingMarquee Then
                pbProgress(idx).Scrolling = PrbScrollingMarquee
            End If
            lblPercentage(idx).Caption = ""
        Else
            If pbProgress(idx).Scrolling = PrbScrollingMarquee Then
                pbProgress(idx).Scrolling = PrbScrollingStandard
            End If
            pbProgress(idx).Value = progress
            lblPercentage(idx).Caption = "(" & progress & "%)"
        End If
        
        If trThreadCount.Value > 1 And idx = 1 And (CDbl(Split(output, ",")(2)) > 0 Or lblTotalBytes.Caption = "0 바이트") Then lblTotalSizeThread.Caption = ParseSize(CDbl(Split(output, ",")(2)), True)
    ElseIf Left$(Data, 6) = "TOTAL " Then
        output = Right$(Data, Len(Data) - 6)
        If CLng(Split(output, ",")(2)) > 100 Then
            progress = -1
        Else
            progress = CInt(Split(output, ",")(2))
        End If
        
        DownloadedBytes = CDbl(Split(output, ",")(1))
        
        If progress < 0 Then
            If pbTotalProgress.Scrolling <> PrbScrollingMarquee Then
                pbTotalProgress.Scrolling = PrbScrollingMarquee
            End If
            If fTotal.Caption <> " 전체 다운로드 진행률 " Then fTotal.Caption = " 전체 다운로드 진행률 "
            If pbTotalProgress.Value <> 0 Then pbTotalProgress.Value = 0
            If DownloadedBytes = -1 Then
                sbStatusBar.Panels(2).Text = ""
            Else
                sbStatusBar.Panels(2).Text = DownloadedBytes & " 바이트"
            End If
            If lblTotalBytes.Caption <> "알 수 없음" Then lblTotalBytes.Caption = "알 수 없음"
            lblDownloadedBytes.Caption = ParseSize(DownloadedBytes, True)
        Else
            If pbTotalProgress.Scrolling = PrbScrollingMarquee Then
                pbTotalProgress.Scrolling = PrbScrollingStandard
            End If
            If Split(output, ",")(0) = "-1" Then
                sbStatusBar.Panels(2).Text = DownloadedBytes & " 바이트"
            Else
                sbStatusBar.Panels(2).Text = Split(output, ",")(0) & " 중 " & DownloadedBytes
            End If
            If Split(output, ",")(0) = "NaN" Or Split(output, ",")(0) = "-1" Then
                lblTotalBytes.Caption = "알 수 없음"
            Else
                lblTotalBytes.Caption = ParseSize(CStr(Split(output, ",")(0)), True)
            End If
            lblDownloadedBytes.Caption = ParseSize(DownloadedBytes, True)
            pbTotalProgress.Value = progress
            fTotal.Caption = " 전체 다운로드 진행률 (" & progress & "%) "
        End If
        
        Dim Speed As Double
        SpeedCount = SpeedCount + 1
        If SpeedCount >= 10 Then
            Speed = (DownloadedBytes - PrevDownloadedBytes)
            lblSpeed.Caption = ParseSize(Speed, True, "/초")
            sbStatusBar.Panels(3).Text = ParseSize(Speed, False, "/초")
            PrevDownloadedBytes = DownloadedBytes
            SpeedCount = 0
        End If
    ElseIf Left$(Data, 17) = "MODIFIEDFILENAME " Then
        output = Right$(Data, Len(Data) - 17)
        DownloadPath = output
        If BatchStarted Then
            lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(1).Text = output
            lvBatchFiles.ListItems(CurrentBatchIdx).Text = fso.GetFilename(output)
        End If
    End If
End Sub

Sub NextBatchDownload()
    If Not BatchStarted Then Exit Sub
    
    If lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(3).Text = "완료" Then _
        lvBatchFiles.ListItems(CurrentBatchIdx).Checked = False
    
    If CurrentBatchIdx = lvBatchFiles.ListItems.Count Then
        BatchStarted = False
        CurrentBatchIdx = 1
        cmdStartBatch.Enabled = -1
        cmdStopBatch.Enabled = 0
        timElapsed.Enabled = 0
        sbStatusBar.Panels(3).Text = ""
        sbStatusBar.Panels(4).Text = ""
        chkOpenAfterComplete.Enabled = -1
        If chkOpenFolder.Value Then
            cmdOpenFolder_Click
        End If
        cmdGo.Enabled = -1
        
        If lvBatchFiles.ListItems.Count > 0 Then
            Dim Enable As Boolean
            Dim i%
            For i = 1 To lvBatchFiles.ListItems.Count
                If lvBatchFiles.ListItems(i).Checked Then
                    Enable = True
                    Exit For
                End If
            Next i
            If Not Enable Then
                cmdStartBatch.Enabled = 0
            Else
                cmdStartBatch.Enabled = -1
            End If
        Else
            cmdStartBatch.Enabled = 0
        End If
        
        If BatchErrorCount Then
            MsgBox "하나 이상의 오류가 발생했습니다. 오류 코드 정보는 다음과 같습니다." & vbCrLf & vbCrLf & "1: 알 수 없는 오류가 발생했습니다. 유효하지 않은 주소를 입력했거나 프로그램 내부 오류입니다." & vbCrLf & "2: 주소나 파일 이름을 지정하지 않았습니다." & vbCrLf & "3: 저장 경로가 존재하지 않습니다." & vbCrLf & "4: 저장할 파일명이 사용 중입니다. 다른 이름을 선택하십시오." & vbCrLf & "5: 내부 작업을 위한 파일명이 사용 중입니다. 다른 이름을 선택하십시오." & vbCrLf & "6: 파일 서버가 다운로드 부스트를 지원하지 않습니다. 강도를 1로 변경해 보십시오." & vbCrLf & "7: 파일의 크기를 알 수 없어서 다운로드를 부스트할 수 없습니다. 강도를 1로 변경해 보십시오.", 48
        ElseIf chkPlaySound.Value Then
            MessageBeep 64
        End If
        
        Exit Sub
    End If
    
    CurrentBatchIdx = CurrentBatchIdx + 1
    StartDownload lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(2), lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(1)
End Sub

Sub OnExit(RetVal As Long)
    If Not BatchStarted Then
        Select Case RetVal
            Case 1
                MsgBox "유효하지 않은 주소를 입력했거나 내부 오류가 발생했습니다.", 16
            Case 2
                MsgBox "주소나 파일 이름을 지정하지 않았습니다.", 16
            Case 3
                MsgBox "저장 경로가 존재하지 않습니다.", 16
            Case 4
                MsgBox "저장할 파일명이 사용 중입니다. 다른 이름을 선택하십시오.", 16
            Case 5
                MsgBox "내부 작업을 위한 파일명이 사용 중입니다. 다른 이름을 선택하십시오.", 16
            Case 6
                MsgBox "파일 서버가 다운로드 부스트를 지원하지 않습니다. 강도를 1로 변경해 보십시오.", 16
            Case 7
                MsgBox "파일의 크기를 알 수 없어서 다운로드를 부스트할 수 없습니다. 강도를 1로 변경해 보십시오.", 16
        End Select
    End If
    
    If Not BatchStarted Then cmdGo.Enabled = -1
    cmdStop.Enabled = 0
    OnStop (RetVal = 0)
    Dim i%
    If BatchStarted Then
        pbTotalProgress.Value = 0
        For i = 1 To lblDownloader.UBound
            pbProgress(i).Value = 0
            lblPercentage(i).Caption = ""
        Next i
        
        If RetVal <> 0 Then
            lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(3).Text = "오류 (" & RetVal & ")"
            lvBatchFiles.ListItems(CurrentBatchIdx).ForeColor = 255
            lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(1).ForeColor = 255
            lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(2).ForeColor = 255
            lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(3).ForeColor = 255
            BatchErrorCount = BatchErrorCount + 1
        Else
            lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(3).Text = "완료"
            lvBatchFiles.ListItems(CurrentBatchIdx).ForeColor = &H8000&
            lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(1).ForeColor = &H8000&
            lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(2).ForeColor = &H8000&
            lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(3).ForeColor = &H8000&
        End If
    
        NextBatchDownload
    ElseIf RetVal = 0 Then
        cmdOpen.Enabled = -1
        If chkOpenAfterComplete.Value Then
            cmdOpen_Click
        End If
        If chkOpenFolder.Value Then
            cmdOpenFolder_Click
        End If
    End If
End Sub

Sub OnStart()
    IsDownloading = True
    cmdGo.Enabled = 0
    If Not BatchStarted Then
        cmdStop.Enabled = -1
    Else
        cmdStop.Enabled = 0
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
    
    chkNoCleanup.Enabled = 0
    
    lblThreadCount.Enabled = 0
    
    cmdBatch.Enabled = 0
    
    cmdStartBatch.Enabled = 0
    
    cmdOpen.Enabled = 0
    
    lblTotalBytes.Caption = "대기 중..."
    lblDownloadedBytes.Caption = "대기 중..."
    If trThreadCount.Value > 1 Then
        lblTotalSizeThread.Caption = "대기 중..."
        lblThreadCount2.Caption = trThreadCount.Value
    Else
        lblTotalSizeThread.Caption = "-"
        lblThreadCount2.Caption = "-"
    End If
    lblElapsed.Caption = "0초"
    
    fTotal.Caption = " 전체 다운로드 진행률 "
    pbTotalProgress.Value = 0
    For i = 1 To trThreadCount.Value
        lblPercentage(i).Caption = ""
        pbProgress(i).Value = 0
    Next i
    
    For i = 1 To trThreadCount.Value
        pbProgress(i).MarqueeSpeed = 35
        pbProgress(i).Scrolling = PrbScrollingMarquee
    Next i
    
    pbTotalProgress.Scrolling = PrbScrollingMarquee
    
    lblState.Caption = "진행 중"
    sbStatusBar.Panels(1).Text = "시작 중..."
End Sub

Sub OnStop(Optional PlayBeep As Boolean = True)
    IsDownloading = False
    If Not BatchStarted Then cmdGo.Enabled = -1
    cmdStop.Enabled = 0
    
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
    
    chkNoCleanup.Enabled = -1
    
    lblThreadCount.Enabled = -1
    
    SP.FinishChild 0, 0
    
    Dim i%
    For i = 1 To trThreadCount.Value
        pbProgress(i).Scrolling = PrbScrollingStandard
    Next i
    
    If pbTotalProgress.Scrolling = PrbScrollingMarquee Then
        pbTotalProgress.Scrolling = PrbScrollingStandard
    End If
    
    If pbTotalProgress.Value < 100 Then
        pbTotalProgress.Value = 0
    End If
    
    If pbTotalProgress.Value < 100 Then
        lblState.Caption = "중지됨"
        sbStatusBar.Panels(1).Text = "준비"
    
        fTotal.Caption = " 전체 다운로드 진행률 "
        For i = 1 To lblDownloader.UBound
            pbProgress(i).Value = 0
            lblPercentage(i).Caption = ""
        Next i
    Else
        lblState.Caption = "완료됨"
        sbStatusBar.Panels(1).Text = "완료"
        sbStatusBar.Panels(2).Text = ""
        sbStatusBar.Panels(3).Text = ""
        sbStatusBar.Panels(4).Text = ""
    End If
    
    cmdBatch.Enabled = -1
    
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
            Else
                cmdStartBatch.Enabled = -1
            End If
        Else
            cmdStartBatch.Enabled = 0
        End If
        
        If PlayBeep And chkPlaySound.Value Then MessageBeep 64
    End If
    
    If lblTotalBytes.Caption = "대기 중..." Then lblTotalBytes.Caption = "-"
    If lblDownloadedBytes.Caption = "대기 중..." Then
        lblDownloadedBytes.Caption = "-"
    Else
        lblTotalBytes.Caption = lblDownloadedBytes.Caption
    End If
    If lblTotalSizeThread.Caption = "대기 중..." Then lblTotalSizeThread.Caption = "-"
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
        MsgBox URL & " - 주소가 올바르지 않습니다. 'http://' 또는 'https://'로 시작해야 합니다.", 16
        Exit Sub
    End If

    Dim idx%
    
    Dim FileName$
    Dim ServerName$
    FileName = Trim$(txtFileName.Text)
    If FolderExists(FileName) Then
        If Not (Right$(FileName, 1) = "\") Then FileName = FileName & "\"
        ServerName = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Split(URL, "/")(UBound(Split(URL, "/"))), "\", "_"), "?", "_"), "*", "_"), "|", "_"), """", "_"), ":", "_"), "<", "_"), ">", "_")
        If Replace(ServerName, " ", "") = "" Then ServerName = "download_" & CStr(Rnd * 1E+15)
        FileName = FileName & ServerName
    Else
        ServerName = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Split(URL, "/")(UBound(Split(URL, "/"))), "\", "_"), "?", "_"), "*", "_"), "|", "_"), """", "_"), ":", "_"), "<", "_"), ">", "_")
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
    lvBatchFiles.ListItems(idx).ListSubItems.Add , , "대기"
    lvBatchFiles.ListItems(idx).Checked = -1
    If IsDownloading Or cmdStop.Enabled Or BatchStarted Then
        cmdStartBatch.Enabled = 0
    Else
        cmdStartBatch.Enabled = -1
    End If
End Sub

Private Sub cmdBatch_Click()
    On Error Resume Next
    If Me.Height = 6930 Then
        cmdBatch.Caption = "<< 일괄 처리(&W)"
        lvBatchFiles.Visible = -1
        fBatchDownload.Visible = -1
        
        fDummyUI1.Visible = -1
        fDummyUI2.Visible = -1
        fDummyUI3.Visible = -1
        fDummyUI4.Visible = -1
        
        Me.Height = 9405
    Else
        Me.Height = 6930
        cmdBatch.Caption = "일괄 처리(&W) >>"
        lvBatchFiles.Visible = 0
        fBatchDownload.Visible = 0
        
        fDummyUI1.Visible = 0
        fDummyUI2.Visible = 0
        fDummyUI3.Visible = 0
        fDummyUI4.Visible = 0
    End If
End Sub

Private Sub cmdBrowse_Click()
    frmBrowse.Show vbModal, Me
End Sub

Private Sub cmdClear_Click()
    txtURL.Text = ""
End Sub

Private Sub cmdDecreaseThreads_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
    ElseIf Not IsDownloading Then
        cmdStartBatch.Enabled = -1
    End If
End Sub

Sub StartDownload(URL As String, FileName As String)
    If BatchStarted Then
        If Not lvBatchFiles.ListItems(CurrentBatchIdx).Checked Then
            lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(3).Text = "통과"
            lvBatchFiles.ListItems(CurrentBatchIdx).ForeColor = &H808080
            lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(1).ForeColor = &H808080
            lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(2).ForeColor = &H808080
            lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(3).ForeColor = &H808080
            NextBatchDownload
            Exit Sub
        End If
        
        If lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(3).Text = "완료" Then
            NextBatchDownload
            Exit Sub
        End If
    
        lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(3).Text = "진행 중..."
        lvBatchFiles.ListItems(CurrentBatchIdx).ForeColor = &HFF0000
        lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(1).ForeColor = &HFF0000
        lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(2).ForeColor = &HFF0000
        lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(3).ForeColor = &HFF0000
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
        ServerName = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Split(URL, "/")(UBound(Split(URL, "/"))), "\", "_"), "?", "_"), "*", "_"), "|", "_"), """", "_"), ":", "_"), "<", "_"), ">", "_")
        If Replace(ServerName, " ", "") = "" Then ServerName = "download_" & CStr(Rnd * 1E+15)
        FileName = FileName & ServerName
    End If
    DownloadPath = FileName
    PrevDownloadedBytes = 0
    SpeedCount = 0
    SPResult = SP.Run("""" & CachePath & "node.exe"" """ & CachePath & "booster_v" & App.Major & "_" & App.Minor & "_" & App.Revision & ".js"" " & Replace(Replace(URL, " ", "%20"), """", "%22") & " """ & FileName & """ " & trThreadCount.Value & " " & (chkNoCleanup.Value * -1) & " " & cbWhenExist.ListIndex)
    Select Case SPResult
        Case SP_SUCCESS
            SP.ClosePipe
        Case SP_CREATEPIPEFAILED
            MsgBox "Run failed, could not create pipe", _
                   vbOKOnly Or vbExclamation, _
                   Caption
            If Not BatchStarted Then cmdGo.Enabled = -1
            cmdStop.Enabled = 0
            OnStop
        Case SP_CREATEPROCFAILED
            MsgBox "Run failed, could not create process", _
                   vbOKOnly Or vbExclamation, _
                   Caption
            If Not BatchStarted Then cmdGo.Enabled = -1
            cmdStop.Enabled = 0
            OnStop
    End Select
End Sub

Private Sub cmdGo_Click()
    Dim SPResult As SP_RESULTS
    Dim TextLine As String
    
    If Replace(txtURL.Text, " ", "") = "" Then
        MsgBox "파일 주소를 입력하십시오.", 64
        Exit Sub
    End If
    
    If Left$(txtURL.Text, 7) <> "http://" And Left$(txtURL.Text, 8) <> "https://" Then
        MsgBox "주소가 올바르지 않습니다. 'http://' 또는 'https://'로 시작해야 합니다.", 16
        Exit Sub
    End If
    
    Dim SplittedPath() As String
    SplittedPath = Split(Trim$(txtFileName.Text), "\")
    Dim i%
    For i = LBound(SplittedPath) To UBound(SplittedPath)
        If Trim$(SplittedPath(i)) <> "" And Replace(Trim$(SplittedPath(i)), ".", "") = "" Then
            MsgBox "저장 경로가 유효하지 않습니다.", 16
            Exit Sub
        End If
    Next i
    
    If (Not FolderExists(Trim$(txtFileName.Text))) And (Not FolderExists(fso.GetParentFolderName(Trim$(txtFileName.Text)))) Then
        MsgBox "저장 경로가 존재하지 않습니다.", 16
        Exit Sub
    End If

    Elapsed = 0
    timElapsed.Enabled = -1
    StartDownload txtURL.Text, txtFileName.Text
End Sub

Private Sub cmdIncreaseThreads_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
    Shell "cmd /c start """" """ & lvBatchFiles.SelectedItem.ListSubItems(1).Text & """"
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

Private Sub cmdStartBatch_Click()
    If lvBatchFiles.ListItems.Count <= 0 Then
        cmdStartBatch.Enabled = 0
        Exit Sub
    End If
    
    Dim SplittedPath() As String
    SplittedPath = Split(Trim$(txtFileName.Text), "\")
    Dim i%
    For i = LBound(SplittedPath) To UBound(SplittedPath)
        If Trim$(SplittedPath(i)) <> "" And Replace(Trim$(SplittedPath(i)), ".", "") = "" Then
            MsgBox "저장 경로가 유효하지 않습니다.", 16
            Exit Sub
        End If
    Next i
    
    If (Not FolderExists(Trim$(txtFileName.Text))) And (Not FolderExists(fso.GetParentFolderName(Trim$(txtFileName.Text)))) Then
        MsgBox "저장 경로가 존재하지 않습니다.", 16
        Exit Sub
    End If
    
    BatchErrorCount = 0
    CurrentBatchIdx = 1
    BatchStarted = True
    cmdStartBatch.Enabled = 0
    cmdStopBatch.Enabled = -1
    Elapsed = 0
    timElapsed.Enabled = -1
    chkOpenAfterComplete.Enabled = 0
    cmdOpen.Enabled = 0
    cmdGo.Enabled = 0
    StartDownload lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(2), lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(1)
End Sub

Private Sub cmdStop_Click()
    If ConfirmEx("다운로드를 중지하시겠습니까? 이어받기는 불가능합니다.", App.Title, Me, 48) = vbYes Then
        OnStop False
        cmdOpen.Enabled = 0
    End If
End Sub

Private Sub cmdStopBatch_Click()
    If ConfirmEx("다운로드를 중지하시겠습니까? 이어받기는 불가능합니다.", App.Title, Me, 48) = vbYes Then
        lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(3).Text = "중지"
        lvBatchFiles.ListItems(CurrentBatchIdx).ForeColor = 255
        lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(1).ForeColor = 255
        lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(2).ForeColor = 255
        lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(3).ForeColor = 255
        BatchStarted = False
        CurrentBatchIdx = 1
        cmdStartBatch.Enabled = -1
        cmdStopBatch.Enabled = 0
        OnStop False
        cmdGo.Enabled = 0
        timElapsed.Enabled = 0
        sbStatusBar.Panels(3).Text = ""
        sbStatusBar.Panels(4).Text = ""
        chkOpenAfterComplete.Enabled = -1
        cmdGo.Enabled = -1
        If BatchErrorCount Then MsgBox "하나 이상의 오류가 발생했습니다. 오류 코드 정보는 다음과 같습니다." & vbCrLf & vbCrLf & "1: 알 수 없는 오류가 발생했습니다. 유효하지 않은 주소를 입력했거나 프로그램 내부 오류입니다." & vbCrLf & "2: 주소나 파일 이름을 지정하지 않았습니다." & vbCrLf & "3: 저장 경로가 존재하지 않습니다." & vbCrLf & "4: 저장할 파일명이 사용 중입니다. 다른 이름을 선택하십시오." & vbCrLf & "5: 내부 작업을 위한 파일명이 사용 중입니다. 다른 이름을 선택하십시오." & vbCrLf & "6: 파일 서버가 다운로드 부스트를 지원하지 않습니다. 강도를 1로 변경해 보십시오." & vbCrLf & "7: 파일의 크기를 알 수 없어서 다운로드를 부스트할 수 없습니다. 강도를 1로 변경해 보십시오.", 48
    End If
End Sub

Private Sub cmdTabDownload_Click()
    optTabDownload.Value = True
    optTabDownload2.Value = True
    optTabDownload_Click
End Sub

Private Sub cmdTabThreads_Click()
    optTabThreads.Value = True
    optTabThreads2.Value = True
    optTabThreads_Click
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor & "." & App.Revision
    TahomaAvailable = FontExists("Tahoma")
    
    Dim Left%
    Dim Top%
    Top = GetSetting("DownloadBooster", "UserData", "FormTop", -1)
    Left = GetSetting("DownloadBooster", "UserData", "FormLeft", -1)
    If Top >= 0 And Left >= 0 Then
        Me.Top = Top
        Me.Left = Left
    End If
    
    Dim i%
    For i = 1 To lblDownloader.UBound
        lblDownloader(i).Caption = "스레드 " & i & ":"
        pbProgress(i).Left = pbProgress(i).Left + 60
        pbProgress(i).Width = pbProgress(i).Width - 60
        lblPercentage(i).Caption = ""
    Next i
    fDownloadInfo.Top = fThreadInfo.Top + 60
    fDownloadInfo.Left = fThreadInfo.Left
    fDownloadInfo.Width = fThreadInfo.Width '5925
    fDownloadInfo.Height = fThreadInfo.Height - 60
    
    If GetSetting("DownloadBooster", "UserData", "LastTab", 1) = 1 Then
        fTabDownload_Click
    Else
        fTabThreads_Click
    End If
    
    lvDummyScroll.AddItem "1"
    lvDummyScroll.AddItem "2"
    lvDummyScroll.AddItem "3"
    lvDummyScroll.AddItem "4"
    lvDummyScroll.AddItem "5"
    lvDummyScroll.AddItem "6"
    lvDummyScroll.AddItem "7"
    lvDummyScroll.AddItem "8"
    lvDummyScroll.AddItem "9"
    lvDummyScroll.AddItem "10"
    lvDummyScroll.AddItem "11"
    lvDummyScroll.AddItem "12"
    lvDummyScroll.AddItem "13"
    lvDummyScroll.AddItem "14"
    lvDummyScroll.AddItem "15"
    lvDummyScroll.AddItem "16"
    lvDummyScroll.AddItem "17"
    lvDummyScroll.AddItem "18"
    lvDummyScroll.AddItem "19"
    lvDummyScroll.AddItem "20"
    lvDummyScroll.AddItem "21"
    lvDummyScroll.AddItem "22"
    lvDummyScroll.AddItem "23"
    lvDummyScroll.AddItem "24"
    lvDummyScroll.AddItem "25"
    lvDummyScroll.ListIndex = 0
    txtDummyScroll.Height = lvDummyScroll.Height
    
    trThreadCount.Value = GetSetting("DownloadBooster", "UserData", "ThreadCount", GetSetting("DownloadBooster", "Options", "ThreadCount", 1))
    trThreadCount_Scroll
    
    lvBatchFiles.ColumnHeaders.Add , "filename", "파일 이름"
    lvBatchFiles.ColumnHeaders.Add , "fullpath", "전체 경로"
    lvBatchFiles.ColumnHeaders.Add , "url", "파일 주소"
    lvBatchFiles.ColumnHeaders.Add , "status", "상태"
    lvBatchFiles.ColumnHeaders(1).Width = 2895
    lvBatchFiles.ColumnHeaders(2).Width = 0
    lvBatchFiles.ColumnHeaders(3).Width = 4595
    lvBatchFiles.ColumnHeaders(4).Width = 1005
    lvBatchFiles.ColumnHeaders(4).Alignment = LvwColumnHeaderAlignmentCenter
    Me.Height = 6930
    
    BatchStarted = False
    
    txtFileName.Text = GetSetting("DownloadBooster", "UserData", "SavePath", CurDir.Path)
    
    If GetSetting("DownloadBooster", "UserData", "BatchExpanded", 1) <> 0 Then
        cmdBatch_Click
    End If
    
    chkNoCleanup.Value = GetSetting("DownloadBooster", "Options", "NoCleanup", 0)
    chkOpenAfterComplete.Value = GetSetting("DownloadBooster", "Options", "OpenWhenComplete", 0)
    chkOpenFolder.Value = GetSetting("DownloadBooster", "Options", "OpenFolderWhenComplete", 0)
    chkRememberURL.Value = GetSetting("DownloadBooster", "Options", "RememberURL", 0)
    If chkRememberURL.Value Then
        txtURL.Text = GetSetting("DownloadBooster", "UserData", "FileURL", "")
        txtURL.SelStart = 0
        txtURL.SelLength = Len(txtURL.Text)
    End If
    chkPlaySound.Value = GetSetting("DownloadBooster", "Options", "PlaySound", 1)
    
    cbWhenExist.Clear
    cbWhenExist.AddItem "작업 중단"
    cbWhenExist.AddItem "덮어쓰기"
    cbWhenExist.AddItem "이름 변경"
    cbWhenExist.ListIndex = GetSetting("DownloadBooster", "Options", "WhenFileExists", 0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If cmdStop.Enabled = -1 Or BatchStarted Then
        If ConfirmEx("다운로드를 중지하시겠습니까? 이어받기는 불가능합니다.", App.Title, Me, 48) <> vbYes Then
            Cancel = 1
            Exit Sub
        Else
            BatchStarted = False
            SP.FinishChild 0, 0
        End If
    Else
        BatchStarted = False
        SP.FinishChild 0, 0
    End If
    
    SaveSetting "DownloadBooster", "UserData", "SavePath", Trim$(txtFileName.Text)
    SaveSetting "DownloadBooster", "UserData", "BatchExpanded", CInt(Me.Height > 6931) * -1
    SaveSetting "DownloadBooster", "Options", "NoCleanup", chkNoCleanup.Value
    SaveSetting "DownloadBooster", "Options", "OpenWhenComplete", chkOpenAfterComplete.Value
    SaveSetting "DownloadBooster", "Options", "OpenFolderWhenComplete", chkOpenFolder.Value
    SaveSetting "DownloadBooster", "Options", "WhenFileExists", cbWhenExist.ListIndex
    SaveSetting "DownloadBooster", "Options", "RememberURL", chkRememberURL.Value
    If chkRememberURL.Value Then
        SaveSetting "DownloadBooster", "UserData", "FileURL", Trim$(txtURL.Text)
    End If
    SaveSetting "DownloadBooster", "Options", "PlaySound", chkPlaySound.Value
    SaveSetting "DownloadBooster", "UserData", "FormTop", Me.Top
    SaveSetting "DownloadBooster", "UserData", "FormLeft", Me.Left
    SaveSetting "DownloadBooster", "UserData", "LastTab", (CInt(optTabThreads2.Value) * -1) + 1
    Unload Me
End Sub

Private Sub fTabDownload_Click()
    cmdTabDownload_Click
End Sub

Private Sub fTabThreads_Click()
    cmdTabThreads_Click
End Sub

Private Sub lvBatchFiles_ContextMenu(ByVal X As Single, ByVal Y As Single)
    On Error GoTo ErrLn
    If lvBatchFiles.SelectedItem.Selected And cmdDelete.Enabled Then _
        Me.PopupMenu mnuListContext
    
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
            Item.ListSubItems(3).Text = "통과"
            Item.ForeColor = &H808080
            Item.ListSubItems(1).ForeColor = &H808080
            Item.ListSubItems(2).ForeColor = &H808080
            Item.ListSubItems(3).ForeColor = &H808080
        Else
            Item.ListSubItems(3).Text = "대기"
            Item.ForeColor = 0
            Item.ListSubItems(1).ForeColor = 0
            Item.ListSubItems(2).ForeColor = 0
            Item.ListSubItems(3).ForeColor = 0
        End If
    End If
    
    If IsDownloading Or BatchStarted Then
        cmdStartBatch.Enabled = 0
        Exit Sub
    End If
    
    If Checked Then
        cmdStartBatch.Enabled = -1
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
    End If
End Sub

Private Sub lvBatchFiles_ItemSelect(ByVal Item As LvwListItem, ByVal Selected As Boolean)
    If Selected Then
        If BatchStarted And Item.Index = CurrentBatchIdx Then
            cmdDelete.Enabled = 0
        Else
            cmdDelete.Enabled = -1
        End If
        
        If Item.ListSubItems(3).Text = "완료" Then
            cmdOpenBatch.Enabled = -1
        Else
            cmdOpenBatch.Enabled = 0
        End If
    Else
        cmdDelete.Enabled = 0
        cmdOpenBatch.Enabled = 0
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

Private Sub mnuDeleteItem_Click()
    cmdDelete_Click
End Sub

Private Sub optTabDownload_Click()
    tsTabs.Tabs(1).Selected = True
End Sub

Private Sub optTabDownload2_Click()
    optTabDownload_Click
End Sub

Private Sub optTabThreads_Click()
    tsTabs.Tabs(2).Selected = True
End Sub

Private Sub optTabThreads2_Click()
    optTabThreads_Click
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
    Dim Hour As Integer
    Dim Minutes As Integer
    Dim Seconds As Integer
    If Elapsed >= 3600 Then
        sbStatusBar.Panels(4).Text = CStr(Floor(Elapsed / 3600)) & "시간 "
    Else
        sbStatusBar.Panels(4).Text = ""
    End If
    
    If Elapsed >= 60 Then
        sbStatusBar.Panels(4).Text = sbStatusBar.Panels(4).Text & Floor((Elapsed Mod 3600) / 60) & "분 "
    End If
    sbStatusBar.Panels(4).Text = sbStatusBar.Panels(4).Text & (Elapsed Mod 60) & "초 경과"
    
    lblElapsed.Caption = Replace(sbStatusBar.Panels(4).Text, " 경과", "")
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
        lblThreadCount.Caption = "(부스트 없음)"
    Else
        lblThreadCount.Caption = "(" & trThreadCount.Value & "개 스레드)"
    End If
    Dim i%
    For i = 1 To trThreadCount.Value
        lblDownloader(i).Visible = -1
        pbProgress(i).Visible = -1
        lblPercentage(i).Visible = -1
        If Not pbProgress(i).MarqueeAnimation Then pbProgress(i).MarqueeAnimation = True
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
'        tsTabs.Tabs(1).Selected = True
'        optTabDownload.Value = True
'        optTabDownload2.Value = True
        chkNoCleanup.Enabled = 0
    Else
'        fThreadInfo.Visible = -1
'        fDownloadInfo.Visible = 0
'        tsTabs.Tabs(2).Selected = True
'        optTabThreads.Value = True
'        optTabThreads2.Value = True
        chkNoCleanup.Enabled = -1
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

Private Sub tsTabs_TabClick(ByVal TabItem As TbsTab)
    If TabItem.Index = 1 Then
        fDownloadInfo.Visible = -1
        fThreadInfo.Visible = 0
    Else
        fThreadInfo.Visible = -1
        fDownloadInfo.Visible = 0
    End If
End Sub

Private Sub vsProgressScroll_Change()
    vsProgressScroll_Scroll
End Sub

Private Sub vsProgressScroll_Scroll()
    'pbProgressContainer.Top = pbProgressOuterContainer.Height * vsProgressScroll.Value * -1 - (105 * vsProgressScroll.Value)
    pbProgressContainer.Top = vsProgressScroll.Value * 255 * -1 - (105 * vsProgressScroll.Value)
End Sub
