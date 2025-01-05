VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  '단일 고정
   Caption         =   "다운로드 부스터"
   ClientHeight    =   8460
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9345
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
   ScaleHeight     =   8460
   ScaleWidth      =   9345
   StartUpPosition =   3  'Windows 기본값
   Begin VB.Frame fDownloadInfo 
      Caption         =   "다운로드 정보"
      Height          =   3855
      Left            =   240
      TabIndex        =   82
      Top             =   2040
      Visible         =   0   'False
      Width           =   6255
      Begin VB.Label lblElapsed 
         Height          =   255
         Left            =   1440
         TabIndex        =   88
         Top             =   1080
         Width           =   4095
      End
      Begin VB.Label Label4 
         Caption         =   "경과 시간:"
         Height          =   255
         Left            =   240
         TabIndex        =   87
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblDownloadedBytes 
         Height          =   255
         Left            =   1440
         TabIndex        =   86
         Top             =   720
         Width           =   4215
      End
      Begin VB.Label Label3 
         Caption         =   "받은 바이트:"
         Height          =   255
         Left            =   240
         TabIndex        =   85
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblTotalBytes 
         Height          =   255
         Left            =   1440
         TabIndex        =   84
         Top             =   360
         Width           =   3855
      End
      Begin VB.Label Label2 
         Caption         =   "총 바이트:"
         Height          =   255
         Left            =   240
         TabIndex        =   83
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.ComboBox cbWhenExist 
      Height          =   300
      Left            =   7590
      Style           =   2  '드롭다운 목록
      TabIndex        =   80
      Top             =   2640
      Width           =   1425
   End
   Begin VB.CheckBox chkOpenAfterComplete 
      Caption         =   "완료 후 열기(&C)"
      Height          =   255
      Left            =   6840
      TabIndex        =   79
      Top             =   1920
      Width           =   1935
   End
   Begin VB.CheckBox chkNoCleanup 
      Caption         =   "조각 파일 삭제 안 함(&D)"
      Height          =   255
      Left            =   6840
      TabIndex        =   78
      Top             =   1560
      Width           =   2250
   End
   Begin VB.CheckBox chkOpenFolder 
      Caption         =   "완료 후 폴더 열기(&L)"
      Height          =   255
      Left            =   6840
      TabIndex        =   77
      Top             =   2280
      Width           =   2055
   End
   Begin VB.CommandButton cmdOpenBatch 
      Caption         =   "열기(&Q)"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   76
      Top             =   7680
      Width           =   1575
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "초기화(&Y)"
      Height          =   300
      Left            =   7680
      TabIndex        =   75
      Top             =   105
      Width           =   1455
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "추가(&R)..."
      Height          =   375
      Left            =   2520
      TabIndex        =   74
      Top             =   7680
      Width           =   1575
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "삭제(&V)"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4200
      TabIndex        =   73
      Top             =   7680
      Width           =   1575
   End
   Begin VB.CommandButton cmdStopBatch 
      Caption         =   "중지(&Z)"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7560
      TabIndex        =   72
      Top             =   7680
      Width           =   1575
   End
   Begin VB.CommandButton cmdStartBatch 
      Caption         =   "다운로드(&A)"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5880
      TabIndex        =   71
      Top             =   7680
      Width           =   1575
   End
   Begin DownloadBooster.ListView lvBatchFiles 
      Height          =   1635
      Left            =   240
      TabIndex        =   70
      Top             =   6000
      Visible         =   0   'False
      Width           =   8890
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
   Begin VB.CommandButton cmdBatch 
      Caption         =   "일괄처리(&W) >>"
      Height          =   375
      Left            =   7440
      TabIndex        =   69
      Top             =   5520
      Width           =   1695
   End
   Begin VB.Frame fTotal 
      Caption         =   "전체 다운로드 현황"
      Height          =   615
      Left            =   240
      TabIndex        =   67
      Top             =   1320
      Width           =   6255
      Begin DownloadBooster.ProgressBar pbTotalProgress 
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
         TabIndex        =   68
         Top             =   285
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "옵션"
      Height          =   1695
      Left            =   6720
      TabIndex        =   13
      Top             =   1320
      Width           =   2415
      Begin VB.Label Label1 
         Caption         =   "존재하면"
         Height          =   255
         Left            =   120
         TabIndex        =   81
         Top             =   1380
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "열기(&O)"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7440
      TabIndex        =   12
      Top             =   3360
      Width           =   1695
   End
   Begin VB.CommandButton cmdOpenFolder 
      Caption         =   "폴더 열기(&E)"
      Height          =   375
      Left            =   7440
      TabIndex        =   11
      Top             =   3840
      Width           =   1695
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "중지(&P)"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7440
      TabIndex        =   10
      Top             =   4920
      Width           =   1695
   End
   Begin DownloadBooster.StatusBar sbStatusBar 
      Align           =   2  '아래 맞춤
      Height          =   330
      Left            =   0
      Top             =   8130
      Width           =   9345
      _ExtentX        =   16484
      _ExtentY        =   582
      InitPanels      =   "frmMain.frx":0442
   End
   Begin VB.Timer timElapsed 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6840
      Top             =   3840
   End
   Begin DownloadBooster.Slider trThreadCount 
      Height          =   495
      Left            =   1560
      TabIndex        =   9
      Top             =   750
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   873
      Min             =   1
      Max             =   25
      Value           =   1
      TickFrequency   =   2
      TipSide         =   1
      SelStart        =   1
   End
   Begin VB.Frame fThreadInfo 
      Caption         =   "스레드 현황"
      Height          =   3855
      Left            =   240
      TabIndex        =   6
      Top             =   2040
      Width           =   6255
      Begin VB.PictureBox pbProgressOuterContainer 
         BorderStyle     =   0  '없음
         Height          =   3495
         Left            =   120
         ScaleHeight     =   3495
         ScaleWidth      =   5775
         TabIndex        =   15
         Top             =   240
         Width           =   5775
         Begin VB.PictureBox pbProgressContainer 
            BorderStyle     =   0  '없음
            Height          =   9015
            Left            =   0
            ScaleHeight     =   9015
            ScaleWidth      =   5775
            TabIndex        =   16
            Top             =   0
            Width           =   5775
            Begin DownloadBooster.ProgressBar pbProgress 
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
            Begin DownloadBooster.ProgressBar pbProgress 
               Height          =   255
               Index           =   2
               Left            =   840
               Top             =   360
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   450
               Step            =   10
            End
            Begin DownloadBooster.ProgressBar pbProgress 
               Height          =   255
               Index           =   3
               Left            =   840
               Top             =   720
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   450
               Step            =   10
            End
            Begin DownloadBooster.ProgressBar pbProgress 
               Height          =   255
               Index           =   4
               Left            =   840
               Top             =   1080
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   450
               Step            =   10
            End
            Begin DownloadBooster.ProgressBar pbProgress 
               Height          =   255
               Index           =   5
               Left            =   840
               Top             =   1440
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   450
               Step            =   10
            End
            Begin DownloadBooster.ProgressBar pbProgress 
               Height          =   255
               Index           =   6
               Left            =   840
               Top             =   1800
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   450
               Step            =   10
            End
            Begin DownloadBooster.ProgressBar pbProgress 
               Height          =   255
               Index           =   7
               Left            =   840
               Top             =   2160
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   450
               Step            =   10
            End
            Begin DownloadBooster.ProgressBar pbProgress 
               Height          =   255
               Index           =   8
               Left            =   840
               Top             =   2520
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   450
               Step            =   10
            End
            Begin DownloadBooster.ProgressBar pbProgress 
               Height          =   255
               Index           =   9
               Left            =   840
               Top             =   2880
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   450
               Step            =   10
            End
            Begin DownloadBooster.ProgressBar pbProgress 
               Height          =   255
               Index           =   10
               Left            =   840
               Top             =   3240
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   450
               Step            =   10
            End
            Begin DownloadBooster.ProgressBar pbProgress 
               Height          =   255
               Index           =   11
               Left            =   840
               Top             =   3600
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   450
               Step            =   10
            End
            Begin DownloadBooster.ProgressBar pbProgress 
               Height          =   255
               Index           =   12
               Left            =   840
               Top             =   3960
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   450
               Step            =   10
            End
            Begin DownloadBooster.ProgressBar pbProgress 
               Height          =   255
               Index           =   13
               Left            =   840
               Top             =   4320
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   450
               Step            =   10
            End
            Begin DownloadBooster.ProgressBar pbProgress 
               Height          =   255
               Index           =   14
               Left            =   840
               Top             =   4680
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   450
               Step            =   10
            End
            Begin DownloadBooster.ProgressBar pbProgress 
               Height          =   255
               Index           =   15
               Left            =   840
               Top             =   5040
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   450
               Step            =   10
            End
            Begin DownloadBooster.ProgressBar pbProgress 
               Height          =   255
               Index           =   16
               Left            =   840
               Top             =   5400
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   450
               Step            =   10
            End
            Begin DownloadBooster.ProgressBar pbProgress 
               Height          =   255
               Index           =   17
               Left            =   840
               Top             =   5760
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   450
               Step            =   10
            End
            Begin DownloadBooster.ProgressBar pbProgress 
               Height          =   255
               Index           =   18
               Left            =   840
               Top             =   6120
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   450
               Step            =   10
            End
            Begin DownloadBooster.ProgressBar pbProgress 
               Height          =   255
               Index           =   19
               Left            =   840
               Top             =   6480
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   450
               Step            =   10
            End
            Begin DownloadBooster.ProgressBar pbProgress 
               Height          =   255
               Index           =   20
               Left            =   840
               Top             =   6840
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   450
               Step            =   10
            End
            Begin DownloadBooster.ProgressBar pbProgress 
               Height          =   255
               Index           =   21
               Left            =   840
               Top             =   7200
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   450
               Step            =   10
            End
            Begin DownloadBooster.ProgressBar pbProgress 
               Height          =   255
               Index           =   22
               Left            =   840
               Top             =   7560
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   450
               Step            =   10
            End
            Begin DownloadBooster.ProgressBar pbProgress 
               Height          =   255
               Index           =   23
               Left            =   840
               Top             =   7920
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   450
               Step            =   10
            End
            Begin DownloadBooster.ProgressBar pbProgress 
               Height          =   255
               Index           =   24
               Left            =   840
               Top             =   8280
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   450
               Step            =   10
            End
            Begin DownloadBooster.ProgressBar pbProgress 
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
               TabIndex        =   66
               Top             =   8685
               Width           =   855
            End
            Begin VB.Label lblPercentage 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "(100%)"
               Height          =   255
               Index           =   25
               Left            =   5040
               TabIndex        =   65
               Top             =   8700
               Width           =   615
            End
            Begin VB.Label lblDownloader 
               Caption         =   "스레드 0:"
               Height          =   255
               Index           =   24
               Left            =   0
               TabIndex        =   64
               Top             =   8325
               Width           =   855
            End
            Begin VB.Label lblPercentage 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "(100%)"
               Height          =   255
               Index           =   24
               Left            =   5040
               TabIndex        =   63
               Top             =   8325
               Width           =   615
            End
            Begin VB.Label lblDownloader 
               Caption         =   "스레드 0:"
               Height          =   255
               Index           =   23
               Left            =   0
               TabIndex        =   62
               Top             =   7965
               Width           =   855
            End
            Begin VB.Label lblPercentage 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "(100%)"
               Height          =   255
               Index           =   23
               Left            =   5040
               TabIndex        =   61
               Top             =   7965
               Width           =   615
            End
            Begin VB.Label lblDownloader 
               Caption         =   "스레드 0:"
               Height          =   255
               Index           =   22
               Left            =   0
               TabIndex        =   60
               Top             =   7605
               Width           =   855
            End
            Begin VB.Label lblPercentage 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "(100%)"
               Height          =   255
               Index           =   22
               Left            =   5040
               TabIndex        =   59
               Top             =   7605
               Width           =   615
            End
            Begin VB.Label lblDownloader 
               Caption         =   "스레드 0:"
               Height          =   255
               Index           =   21
               Left            =   0
               TabIndex        =   58
               Top             =   7245
               Width           =   855
            End
            Begin VB.Label lblPercentage 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "(100%)"
               Height          =   255
               Index           =   21
               Left            =   5040
               TabIndex        =   57
               Top             =   7245
               Width           =   615
            End
            Begin VB.Label lblDownloader 
               Caption         =   "스레드 0:"
               Height          =   255
               Index           =   20
               Left            =   0
               TabIndex        =   56
               Top             =   6885
               Width           =   855
            End
            Begin VB.Label lblPercentage 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "(100%)"
               Height          =   255
               Index           =   20
               Left            =   5040
               TabIndex        =   55
               Top             =   6885
               Width           =   615
            End
            Begin VB.Label lblDownloader 
               Caption         =   "스레드 0:"
               Height          =   255
               Index           =   19
               Left            =   0
               TabIndex        =   54
               Top             =   6525
               Width           =   855
            End
            Begin VB.Label lblPercentage 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "(100%)"
               Height          =   255
               Index           =   19
               Left            =   5040
               TabIndex        =   53
               Top             =   6525
               Width           =   615
            End
            Begin VB.Label lblDownloader 
               Caption         =   "스레드 0:"
               Height          =   255
               Index           =   18
               Left            =   0
               TabIndex        =   52
               Top             =   6165
               Width           =   855
            End
            Begin VB.Label lblPercentage 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "(100%)"
               Height          =   255
               Index           =   18
               Left            =   5040
               TabIndex        =   51
               Top             =   6165
               Width           =   615
            End
            Begin VB.Label lblDownloader 
               Caption         =   "스레드 0:"
               Height          =   255
               Index           =   17
               Left            =   0
               TabIndex        =   50
               Top             =   5805
               Width           =   855
            End
            Begin VB.Label lblPercentage 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "(100%)"
               Height          =   255
               Index           =   17
               Left            =   5040
               TabIndex        =   49
               Top             =   5805
               Width           =   615
            End
            Begin VB.Label lblDownloader 
               Caption         =   "스레드 0:"
               Height          =   255
               Index           =   16
               Left            =   0
               TabIndex        =   48
               Top             =   5445
               Width           =   855
            End
            Begin VB.Label lblPercentage 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "(100%)"
               Height          =   255
               Index           =   16
               Left            =   5040
               TabIndex        =   47
               Top             =   5445
               Width           =   615
            End
            Begin VB.Label lblDownloader 
               Caption         =   "스레드 0:"
               Height          =   255
               Index           =   15
               Left            =   0
               TabIndex        =   46
               Top             =   5085
               Width           =   855
            End
            Begin VB.Label lblPercentage 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "(100%)"
               Height          =   255
               Index           =   15
               Left            =   5040
               TabIndex        =   45
               Top             =   5085
               Width           =   615
            End
            Begin VB.Label lblDownloader 
               Caption         =   "스레드 0:"
               Height          =   255
               Index           =   14
               Left            =   0
               TabIndex        =   44
               Top             =   4725
               Width           =   855
            End
            Begin VB.Label lblPercentage 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "(100%)"
               Height          =   255
               Index           =   14
               Left            =   5040
               TabIndex        =   43
               Top             =   4725
               Width           =   615
            End
            Begin VB.Label lblDownloader 
               Caption         =   "스레드 0:"
               Height          =   255
               Index           =   13
               Left            =   0
               TabIndex        =   42
               Top             =   4365
               Width           =   855
            End
            Begin VB.Label lblPercentage 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "(100%)"
               Height          =   255
               Index           =   13
               Left            =   5040
               TabIndex        =   41
               Top             =   4365
               Width           =   615
            End
            Begin VB.Label lblDownloader 
               Caption         =   "스레드 0:"
               Height          =   255
               Index           =   12
               Left            =   0
               TabIndex        =   40
               Top             =   4005
               Width           =   855
            End
            Begin VB.Label lblPercentage 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "(100%)"
               Height          =   255
               Index           =   12
               Left            =   5040
               TabIndex        =   39
               Top             =   4005
               Width           =   615
            End
            Begin VB.Label lblDownloader 
               Caption         =   "스레드 0:"
               Height          =   255
               Index           =   11
               Left            =   0
               TabIndex        =   38
               Top             =   3645
               Width           =   855
            End
            Begin VB.Label lblPercentage 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "(100%)"
               Height          =   255
               Index           =   11
               Left            =   5040
               TabIndex        =   37
               Top             =   3645
               Width           =   615
            End
            Begin VB.Label lblDownloader 
               Caption         =   "스레드 0:"
               Height          =   255
               Index           =   10
               Left            =   0
               TabIndex        =   36
               Top             =   3285
               Width           =   855
            End
            Begin VB.Label lblPercentage 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "(100%)"
               Height          =   255
               Index           =   10
               Left            =   5040
               TabIndex        =   35
               Top             =   3285
               Width           =   615
            End
            Begin VB.Label lblDownloader 
               Caption         =   "스레드 0:"
               Height          =   255
               Index           =   9
               Left            =   0
               TabIndex        =   34
               Top             =   2925
               Width           =   855
            End
            Begin VB.Label lblPercentage 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "(100%)"
               Height          =   255
               Index           =   9
               Left            =   5040
               TabIndex        =   33
               Top             =   2925
               Width           =   615
            End
            Begin VB.Label lblDownloader 
               Caption         =   "스레드 0:"
               Height          =   255
               Index           =   8
               Left            =   0
               TabIndex        =   32
               Top             =   2565
               Width           =   855
            End
            Begin VB.Label lblPercentage 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "(100%)"
               Height          =   255
               Index           =   8
               Left            =   5040
               TabIndex        =   31
               Top             =   2565
               Width           =   615
            End
            Begin VB.Label lblDownloader 
               Caption         =   "스레드 0:"
               Height          =   255
               Index           =   7
               Left            =   0
               TabIndex        =   30
               Top             =   2205
               Width           =   855
            End
            Begin VB.Label lblPercentage 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "(100%)"
               Height          =   255
               Index           =   7
               Left            =   5040
               TabIndex        =   29
               Top             =   2205
               Width           =   615
            End
            Begin VB.Label lblDownloader 
               Caption         =   "스레드 0:"
               Height          =   255
               Index           =   6
               Left            =   0
               TabIndex        =   28
               Top             =   1845
               Width           =   855
            End
            Begin VB.Label lblPercentage 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "(100%)"
               Height          =   255
               Index           =   6
               Left            =   5040
               TabIndex        =   27
               Top             =   1845
               Width           =   615
            End
            Begin VB.Label lblDownloader 
               Caption         =   "스레드 0:"
               Height          =   255
               Index           =   5
               Left            =   0
               TabIndex        =   26
               Top             =   1485
               Width           =   855
            End
            Begin VB.Label lblPercentage 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "(100%)"
               Height          =   255
               Index           =   5
               Left            =   5040
               TabIndex        =   25
               Top             =   1485
               Width           =   615
            End
            Begin VB.Label lblDownloader 
               Caption         =   "스레드 0:"
               Height          =   255
               Index           =   4
               Left            =   0
               TabIndex        =   24
               Top             =   1125
               Width           =   855
            End
            Begin VB.Label lblPercentage 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "(100%)"
               Height          =   255
               Index           =   4
               Left            =   5040
               TabIndex        =   23
               Top             =   1125
               Width           =   615
            End
            Begin VB.Label lblDownloader 
               Caption         =   "스레드 0:"
               Height          =   255
               Index           =   3
               Left            =   0
               TabIndex        =   22
               Top             =   765
               Width           =   855
            End
            Begin VB.Label lblPercentage 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "(100%)"
               Height          =   255
               Index           =   3
               Left            =   5040
               TabIndex        =   21
               Top             =   765
               Width           =   615
            End
            Begin VB.Label lblDownloader 
               Caption         =   "스레드 0:"
               Height          =   255
               Index           =   2
               Left            =   0
               TabIndex        =   20
               Top             =   405
               Width           =   855
            End
            Begin VB.Label lblPercentage 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "(100%)"
               Height          =   255
               Index           =   2
               Left            =   5040
               TabIndex        =   19
               Top             =   405
               Width           =   615
            End
            Begin VB.Label lblDownloader 
               Caption         =   "스레드 0:"
               Height          =   255
               Index           =   1
               Left            =   0
               TabIndex        =   18
               Top             =   45
               Width           =   855
            End
            Begin VB.Label lblPercentage 
               Alignment       =   1  '오른쪽 맞춤
               Caption         =   "(100%)"
               Height          =   255
               Index           =   1
               Left            =   5040
               TabIndex        =   17
               Top             =   45
               Width           =   615
            End
         End
      End
      Begin VB.VScrollBar vsProgressScroll 
         Height          =   3495
         Left            =   5880
         Max             =   15
         TabIndex        =   14
         Top             =   240
         Width           =   255
      End
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
      TabIndex        =   2
      Top             =   120
      Width           =   6015
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "다운로드(&S)"
      Height          =   375
      Left            =   7440
      TabIndex        =   0
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Label lblThreadCount 
      Caption         =   "(일반 다운로드)"
      Height          =   255
      Left            =   7680
      TabIndex        =   8
      Top             =   870
      Width           =   1455
   End
   Begin VB.Label lblThreadCountLabel 
      Caption         =   "강도(&T):"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   870
      Width           =   1215
   End
   Begin VB.Label lblFilePath 
      Caption         =   "저장 경로(&F):"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   510
      Width           =   1215
   End
   Begin VB.Label lblURL 
      Caption         =   "파일 주소(&U):"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   150
      Width           =   1215
   End
   Begin DownloadBooster.ShellPipe SP 
      Left            =   6840
      Top             =   4440
      _ExtentX        =   635
      _ExtentY        =   635
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

Sub OnData(Data As String)
    Dim output$
    Dim idx%
    Dim progress%
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
    ElseIf Left$(Data, 6) = "TOTAL " Then
        output = Right$(Data, Len(Data) - 6)
        If CLng(Split(output, ",")(2)) > 100 Then
            progress = -1
        Else
            progress = CInt(Split(output, ",")(2))
        End If
        
        If progress < 0 Then
            If pbTotalProgress.Scrolling <> PrbScrollingMarquee Then
                pbTotalProgress.Scrolling = PrbScrollingMarquee
            End If
            If fTotal.Caption <> "전체 다운로드 현황" Then fTotal.Caption = "전체 다운로드 현황"
            If pbTotalProgress.Value <> 0 Then pbTotalProgress.Value = 0
            If Split(output, ",")(1) = "-1" Then
                sbStatusBar.Panels(2).Text = ""
            Else
                sbStatusBar.Panels(2).Text = Split(output, ",")(1) & " 바이트"
            End If
            If lblTotalBytes.Caption <> "알 수 없음" Then lblTotalBytes.Caption = "알 수 없음"
            lblDownloadedBytes.Caption = Split(output, ",")(1)
        Else
            If pbTotalProgress.Scrolling = PrbScrollingMarquee Then
                pbTotalProgress.Scrolling = PrbScrollingStandard
            End If
            If Split(output, ",")(0) = "-1" Then
                sbStatusBar.Panels(2).Text = Split(output, ",")(1) & " 바이트"
            Else
                sbStatusBar.Panels(2).Text = Split(output, ",")(0) & " 중 " & Split(output, ",")(1)
            End If
            If Split(output, ",")(0) = "NaN" Or Split(output, ",")(0) = "-1" Then
                lblTotalBytes.Caption = "알 수 없음"
            Else
                lblTotalBytes.Caption = Split(output, ",")(0)
            End If
            lblDownloadedBytes.Caption = Split(output, ",")(1)
            pbTotalProgress.Value = progress
            fTotal.Caption = "전체 다운로드 현황 (" & progress & "%)"
        End If
    ElseIf Left$(Data, 17) = "MODIFIEDFILENAME " Then
        output = Right$(Data, Len(Data) - 17)
        DownloadPath = output
    End If
End Sub

Sub NextBatchDownload()
    If Not BatchStarted Then Exit Sub
    
    If CurrentBatchIdx = lvBatchFiles.ListItems.Count Then
        BatchStarted = False
        CurrentBatchIdx = 1
        cmdStartBatch.Enabled = -1
        cmdStopBatch.Enabled = 0
        timElapsed.Enabled = 0
        sbStatusBar.Panels(3).Text = ""
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
        
        Exit Sub
    End If
    
    CurrentBatchIdx = CurrentBatchIdx + 1
    StartDownload lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(2), lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(1)
End Sub

Sub OnExit(RetVal As Long)
    If Not BatchStarted Then
        Select Case RetVal
            Case 1
                MsgBox "해당 주소에 접속하는 중 오류가 발생했습니다.", 16
            Case 2
                MsgBox "주소나 파일 이름을 지정하지 않았습니다.", 16
            Case 3
                MsgBox "다운로드 강도가 잘못되었습니다.", 16
            Case 4
                MsgBox "저장할 파일명이 사용 중입니다. 다른 이름을 선택하십시오.'", 16
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
    OnStop
    Dim i%
    If BatchStarted Then
        pbTotalProgress.Value = 0
        For i = 1 To lblDownloader.UBound
            pbProgress(i).Value = 0
            lblPercentage(i).Caption = ""
        Next i
        
        If RetVal <> 0 Then
            lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(3).Text = "오류 (" & RetVal & ")"
        Else
            lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(3).Text = "완료"
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
    
    chkNoCleanup.Enabled = 0
    
    lblThreadCount.Enabled = 0
    
    cmdBatch.Enabled = 0
    
    cmdStartBatch.Enabled = 0
    
    cmdOpen.Enabled = 0
    
    lblTotalBytes.Caption = "대기 중..."
    lblDownloadedBytes.Caption = "대기 중..."
    lblElapsed.Caption = "0초"
    
    fTotal.Caption = "전체 다운로드 현황"
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

Sub OnStop()
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
    
        fTotal.Caption = "전체 다운로드 현황"
        For i = 1 To lblDownloader.UBound
            pbProgress(i).Value = 0
            lblPercentage(i).Caption = ""
        Next i
    Else
        lblState.Caption = "완료됨"
        sbStatusBar.Panels(1).Text = "완료"
        sbStatusBar.Panels(2).Text = ""
        sbStatusBar.Panels(3).Text = ""
    End If
    
    cmdBatch.Enabled = -1
    
    If Not BatchStarted Then
        timElapsed.Enabled = 0
        sbStatusBar.Panels(3).Text = ""
        
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
    End If
    
    If lblTotalBytes.Caption = "대기 중..." Then lblTotalBytes.Caption = ""
    If lblDownloadedBytes.Caption = "대기 중..." Then lblDownloadedBytes.Caption = ""
End Sub

Private Sub cmdAdd_Click()
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
    FileName = txtFileName.Text
    If FolderExists(FileName) Then
        If Not (Right$(FileName, 1) = "\") Then FileName = FileName & "\"
        ServerName = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Split(URL, "/")(UBound(Split(URL, "/"))), "\", "_"), "?", "_"), "*", "_"), "|", "_"), """", "_"), ":", "_"), "<", "_"), ">", "_")
        If Replace(ServerName, " ", "") = "" Then ServerName = "download_" & CStr(Rnd * 1E+15)
        FileName = FileName & ServerName
    Else
        Dim fso
        Set fso = CreateObject("Scripting.FileSystemObject")
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
    If Me.Height = 6840 Then
        Me.Height = 8940
        cmdBatch.Caption = "<< 일괄처리(&W)"
        lvBatchFiles.Visible = -1
    Else
        Me.Height = 6840
        cmdBatch.Caption = "일괄처리(&W) >>"
        lvBatchFiles.Visible = 0
    End If
End Sub

Private Sub cmdBrowse_Click()
    frmBrowse.Show vbModal, Me
End Sub

Private Sub cmdClear_Click()
    txtURL.Text = ""
End Sub

Private Sub cmdDelete_Click()
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
            'lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(3).Text = "통과"
            lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(3).Text = "대기"
            NextBatchDownload
            Exit Sub
        End If
        
        If lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(3).Text = "완료" Then
            NextBatchDownload
            Exit Sub
        End If
    
        lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(3).Text = "진행 중..."
    End If
    
    OnStart
    Dim ServerName$
    If FolderExists(FileName) Then
        If Not (Right$(FileName, 1) = "\") Then FileName = FileName & "\"
        ServerName = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Split(URL, "/")(UBound(Split(URL, "/"))), "\", "_"), "?", "_"), "*", "_"), "|", "_"), """", "_"), ":", "_"), "<", "_"), ">", "_")
        If Replace(ServerName, " ", "") = "" Then ServerName = "download_" & CStr(Rnd * 1E+15)
        FileName = FileName & ServerName
    End If
    DownloadPath = FileName
    SPResult = SP.Run("""" & App.Path & "\node.exe"" """ & App.Path & "\booster.js"" " & Replace(Replace(URL, " ", "%20"), """", "%22") & " """ & FileName & """ " & trThreadCount.Value & " " & (chkNoCleanup.Value * -1) & " " & cbWhenExist.ListIndex)
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
    
    If Left$(txtURL, 7) <> "http://" And Left$(txtURL, 8) <> "https://" Then
        MsgBox "주소가 올바르지 않습니다. 'http://' 또는 'https://'로 시작해야 합니다.", 16
        Exit Sub
    End If

    On Error Resume Next
    On Error GoTo 0
    Elapsed = 0
    timElapsed.Enabled = -1
    StartDownload txtURL.Text, txtFileName.Text
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
        Dim fso
        Set fso = CreateObject("Scripting.FileSystemObject")
        Shell "cmd /c start """" explorer.exe """ & fso.GetParentFolderName(pth) & """"
    End If
End Sub

Private Sub cmdStartBatch_Click()
    If lvBatchFiles.ListItems.Count <= 0 Then
        cmdStartBatch.Enabled = 0
        Exit Sub
    End If

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
    If MsgBox("다운로드를 중지하시겠습니까? 이어받기는 불가능합니다.", 48 + vbYesNo) = vbYes Then
        OnStop
        cmdOpen.Enabled = 0
    End If
End Sub

Private Sub cmdStopBatch_Click()
    If MsgBox("다운로드를 중지하시겠습니까? 이어받기는 불가능합니다.", 48 + vbYesNo) = vbYes Then
        lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(3).Text = "중지"
        BatchStarted = False
        CurrentBatchIdx = 1
        cmdStartBatch.Enabled = -1
        cmdStopBatch.Enabled = 0
        OnStop
        cmdGo.Enabled = 0
        timElapsed.Enabled = 0
        sbStatusBar.Panels(3).Text = ""
        chkOpenAfterComplete.Enabled = -1
        cmdGo.Enabled = -1
    End If
End Sub

Private Sub Command2_Click()
    Shell "cmd /c start """" """ & DownloadPath & """"
End Sub

Private Sub Form_Load()
    Dim i%
    For i = 1 To lblDownloader.UBound
        lblDownloader(i).Caption = "스레드" & i & ":"
        lblPercentage(i).Caption = ""
    Next i
    trThreadCount.Value = GetSetting("DownloadBooster", "Options", "ThreadCount", 1)
    trThreadCount_Scroll
    
    lvBatchFiles.ColumnHeaders.Add , "filename", "파일 이름"
    lvBatchFiles.ColumnHeaders.Add , "fullpath", "전체 경로"
    lvBatchFiles.ColumnHeaders.Add , "url", "파일 주소"
    lvBatchFiles.ColumnHeaders.Add , "status", "상태"
    lvBatchFiles.ColumnHeaders(1).Width = 2895
    lvBatchFiles.ColumnHeaders(2).Width = 0
    lvBatchFiles.ColumnHeaders(3).Width = 3975
    lvBatchFiles.ColumnHeaders(4).Width = 1455
    lvBatchFiles.ColumnHeaders(4).Alignment = LvwColumnHeaderAlignmentCenter
    Me.Height = 6840
    
    BatchStarted = False
    
    txtFileName.Text = GetSetting("DownloadBooster", "UserData", "SavePath", App.Path)
    
    If GetSetting("DownloadBooster", "UserData", "BatchExpanded", 1) <> 0 Then
        cmdBatch_Click
    End If
    
    chkNoCleanup.Value = GetSetting("DownloadBooster", "Options", "NoCleanup", 0)
    chkOpenAfterComplete.Value = GetSetting("DownloadBooster", "Options", "OpenWhenComplete", 0)
    chkOpenFolder.Value = GetSetting("DownloadBooster", "Options", "OpenFolderWhenComplete", 0)
    
    cbWhenExist.Clear
    cbWhenExist.AddItem "작업 중단"
    cbWhenExist.AddItem "덮어쓰기"
    cbWhenExist.AddItem "이름 변경"
    cbWhenExist.ListIndex = GetSetting("DownloadBooster", "Options", "WhenFileExists", 0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If cmdStop.Enabled = -1 Or BatchStarted Then
        If MsgBox("다운로드를 중지하시겠습니까? 이어받기는 불가능합니다.", 48 + vbYesNo) <> vbYes Then
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
    
    SaveSetting "DownloadBooster", "UserData", "SavePath", txtFileName.Text
    SaveSetting "DownloadBooster", "UserData", "BatchExpanded", CInt(Me.Height > 6840) * -1
    SaveSetting "DownloadBooster", "Options", "NoCleanup", chkNoCleanup.Value
    SaveSetting "DownloadBooster", "Options", "OpenWhenComplete", chkOpenAfterComplete.Value
    SaveSetting "DownloadBooster", "Options", "OpenFolderWhenComplete", chkOpenFolder.Value
    SaveSetting "DownloadBooster", "Options", "WhenFileExists", cbWhenExist.ListIndex
    Unload Me
End Sub

Private Sub lvBatchFiles_ItemCheck(ByVal Item As LvwListItem, ByVal Checked As Boolean)
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
        cmdDelete.Enabled = -1
        
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
        sbStatusBar.Panels(3).Text = CStr(Floor(Elapsed / 3600)) & "시간 "
    Else
        sbStatusBar.Panels(3).Text = ""
    End If
    
    If Elapsed >= 60 Then
        sbStatusBar.Panels(3).Text = sbStatusBar.Panels(3).Text & Floor((Elapsed Mod 3600) / 60) & "분 "
    End If
    sbStatusBar.Panels(3).Text = sbStatusBar.Panels(3).Text & (Elapsed Mod 60) & "초 경과"
    
    lblElapsed.Caption = Replace(sbStatusBar.Panels(3).Text, " 경과", "")
End Sub

Private Sub trThreadCount_Change()
    trThreadCount_Scroll
    SaveSetting "DownloadBooster", "Options", "ThreadCount", trThreadCount.Value
End Sub

Private Sub trThreadCount_Scroll()
    If trThreadCount.Value = 1 Then
        lblThreadCount.Caption = "(일반 다운로드)"
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
    Else
        If vsProgressScroll.Max <> 0 Then vsProgressScroll.Max = 0
        If vsProgressScroll.Enabled Then vsProgressScroll.Enabled = 0
    End If
    
    If trThreadCount.Value <= 1 Then
        fDownloadInfo.Visible = -1
        fThreadInfo.Visible = 0
    Else
        fDownloadInfo.Visible = 0
        fThreadInfo.Visible = -1
    End If
End Sub

Private Sub vsProgressScroll_Change()
    vsProgressScroll_Scroll
End Sub

Private Sub vsProgressScroll_Scroll()
    'pbProgressContainer.Top = pbProgressOuterContainer.Height * vsProgressScroll.Value * -1 - (105 * vsProgressScroll.Value)
    pbProgressContainer.Top = vsProgressScroll.Value * 255 * -1 - (105 * vsProgressScroll.Value)
End Sub
