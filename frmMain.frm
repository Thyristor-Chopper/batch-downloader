VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "다운로드 부스터"
   ClientHeight    =   7740
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11715
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
   ScaleWidth      =   11715
   StartUpPosition =   3  'Windows 기본값
   Begin VB.VScrollBar vsProgressScroll 
      Height          =   3495
      LargeChange     =   10
      Left            =   6120
      Max             =   5
      TabIndex        =   16
      Top             =   2280
      Width           =   255
   End
   Begin VB.PictureBox pbDummy 
      AutoRedraw      =   -1  'True
      Height          =   135
      Left            =   10680
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   84
      Top             =   1920
      Width           =   135
   End
   Begin prjDownloadBooster.FrameW fTygemFrameTransparent 
      Height          =   4845
      Left            =   10560
      TabIndex        =   81
      Tag             =   "nobackcolorchange"
      Top             =   2280
      Visible         =   0   'False
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   8546
      BackColor       =   12632256
      BorderStyle     =   0
      Caption         =   "FrameW1"
      Begin VB.Label lblLBCaption2 
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
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   82
         Tag             =   "nosizechange"
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label lblLBCaptionShadow2 
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
         Left            =   135
         TabIndex        =   83
         Tag             =   "nocolorsizechange"
         Top             =   135
         Width           =   1215
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00404040&
         X1              =   6345
         X2              =   6345
         Y1              =   435
         Y2              =   4650
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00404040&
         X1              =   6480
         X2              =   6480
         Y1              =   435
         Y2              =   4680
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00404040&
         X1              =   135
         X2              =   135
         Y1              =   435
         Y2              =   4650
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         X1              =   135
         X2              =   1725
         Y1              =   4650
         Y2              =   4650
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00404040&
         X1              =   0
         X2              =   0
         Y1              =   360
         Y2              =   4695
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00404040&
         X1              =   135
         X2              =   1725
         Y1              =   435
         Y2              =   435
      End
      Begin VB.Image imgBottomRight2 
         Height          =   180
         Left            =   6345
         Picture         =   "frmMain.frx":1782
         Top             =   4665
         Width           =   150
      End
      Begin VB.Image imgTopRight2 
         Height          =   435
         Left            =   6345
         Picture         =   "frmMain.frx":17D4
         Top             =   0
         Width           =   150
      End
      Begin VB.Image imgBottom2 
         Height          =   195
         Left            =   1725
         Picture         =   "frmMain.frx":182B
         Stretch         =   -1  'True
         Top             =   4650
         Width           =   4635
      End
      Begin VB.Image imgBottomLeft2 
         Height          =   180
         Left            =   0
         Picture         =   "frmMain.frx":18C8
         Top             =   4665
         Width           =   1725
      End
      Begin VB.Image imgTop2 
         Height          =   450
         Left            =   1725
         Picture         =   "frmMain.frx":194C
         Stretch         =   -1  'True
         Top             =   0
         Width           =   4620
      End
      Begin VB.Image imgTopLeft2 
         Height          =   435
         Left            =   0
         Picture         =   "frmMain.frx":1A16
         Top             =   0
         Width           =   1725
      End
      Begin VB.Image imgFrameTexture 
         Height          =   735
         Left            =   360
         Stretch         =   -1  'True
         Top             =   600
         Visible         =   0   'False
         Width           =   975
      End
   End
   Begin prjDownloadBooster.CommandButtonW cmdOpen 
      Height          =   330
      Left            =   7200
      TabIndex        =   27
      Top             =   4065
      Width           =   1695
      _extentx        =   2990
      _extenty        =   582
      enabled         =   0   'False
      imagelist       =   "imgOpenFile"
      caption         =   "열기(&O) "
      transparent     =   -1  'True
   End
   Begin prjDownloadBooster.CommandButtonW cmdOpenBatch 
      Height          =   375
      Left            =   240
      TabIndex        =   37
      Top             =   6960
      Width           =   1575
      _extentx        =   2778
      _extenty        =   661
      enabled         =   0   'False
      imagelist       =   "imgOpenFile"
      caption         =   "열기(&W) "
      transparent     =   -1  'True
   End
   Begin prjDownloadBooster.CommandButtonW cmdDelete 
      Height          =   375
      Left            =   4200
      TabIndex        =   40
      Top             =   6960
      Width           =   1335
      _extentx        =   2355
      _extenty        =   661
      enabled         =   0   'False
      imagelist       =   "imgMinus"
      caption         =   "제거(&V) "
      transparent     =   -1  'True
   End
   Begin VB.TextBox txtURL 
      Height          =   300
      Left            =   1560
      TabIndex        =   1
      Top             =   105
      Width           =   5745
   End
   Begin prjDownloadBooster.CommandButtonW cmdDownloadOptions 
      Height          =   330
      Left            =   7200
      TabIndex        =   26
      Top             =   3690
      Width           =   1935
      _extentx        =   3413
      _extenty        =   582
      caption         =   "다운로드 설정(&S)..."
      transparent     =   -1  'True
   End
   Begin prjDownloadBooster.ListView lvLogTest 
      Height          =   1335
      Left            =   6600
      TabIndex        =   35
      Top             =   4080
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   2355
      View            =   3
   End
   Begin prjDownloadBooster.CommandButtonW cmdYtdlTest 
      Height          =   375
      Left            =   6600
      TabIndex        =   34
      Top             =   5520
      Visible         =   0   'False
      Width           =   615
      _extentx        =   1085
      _extenty        =   661
      caption         =   "ㅇ"
   End
   Begin prjDownloadBooster.CommandButtonW cmdEdit 
      Height          =   375
      Left            =   5880
      TabIndex        =   42
      Top             =   6960
      Width           =   1575
      _extentx        =   2778
      _extenty        =   661
      enabled         =   0   'False
      imagelist       =   "imgEdit"
      caption         =   "편집(&N)..."
      transparent     =   -1  'True
   End
   Begin prjDownloadBooster.ImageList imgEdit 
      Left            =   9240
      Top             =   5880
      _ExtentX        =   1005
      _ExtentY        =   1005
      ImageWidth      =   16
      ImageHeight     =   16
      ColorDepth      =   4
      InitListImages  =   "frmMain.frx":1AED
   End
   Begin prjDownloadBooster.CommandButtonW cmdStopBatch 
      Height          =   375
      Left            =   7560
      TabIndex        =   44
      Top             =   6960
      Width           =   1575
      _extentx        =   2778
      _extenty        =   661
      enabled         =   0   'False
      imagelist       =   "imgStopRed"
      caption         =   "중지(&Z) "
      transparent     =   -1  'True
   End
   Begin prjDownloadBooster.ProgressBar pbTotalProgressMarquee 
      Height          =   255
      Left            =   1200
      Top             =   1560
      Visible         =   0   'False
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   450
      Step            =   10
      MarqueeSpeed    =   35
      Scrolling       =   2
   End
   Begin prjDownloadBooster.ProgressBar pbTotalProgress 
      Height          =   255
      Left            =   1200
      Top             =   1560
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   450
      Step            =   10
      MarqueeSpeed    =   35
   End
   Begin prjDownloadBooster.ImageList imgWrench 
      Left            =   9240
      Top             =   6600
      _ExtentX        =   1005
      _ExtentY        =   1005
      ImageWidth      =   16
      ImageHeight     =   16
      ColorDepth      =   4
      InitListImages  =   "frmMain.frx":1FD5
   End
   Begin prjDownloadBooster.ImageList imgErase 
      Left            =   9840
      Top             =   6600
      _ExtentX        =   1005
      _ExtentY        =   1005
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      InitListImages  =   "frmMain.frx":24BD
   End
   Begin prjDownloadBooster.StatusBar sbStatusBar 
      Align           =   2  '아래 맞춤
      Height          =   330
      Left            =   0
      Top             =   7410
      Width           =   11715
      _ExtentX        =   20664
      _ExtentY        =   582
      InitPanels      =   "frmMain.frx":28A5
   End
   Begin prjDownloadBooster.ListView lvBatchFiles 
      Height          =   870
      Left            =   240
      TabIndex        =   36
      Top             =   6075
      Visible         =   0   'False
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   1535
      VisualTheme     =   1
      View            =   3
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      LabelEdit       =   2
      Checkboxes      =   -1  'True
      HideSelection   =   0   'False
      ShowLabelTips   =   -1  'True
      HighlightColumnHeaders=   -1  'True
      AutoSelectFirstItem=   0   'False
   End
   Begin prjDownloadBooster.CommandButtonW cmdAbout 
      Height          =   300
      Left            =   7080
      TabIndex        =   25
      Top             =   3195
      Width           =   1935
      _extentx        =   3413
      _extenty        =   529
      caption         =   "프로그램 정보(&U)..."
      transparent     =   -1  'True
   End
   Begin prjDownloadBooster.CommandButtonW cmdOptions 
      Height          =   300
      Left            =   7080
      TabIndex        =   24
      Top             =   2865
      Width           =   1935
      _extentx        =   3413
      _extenty        =   529
      imagelist       =   "imgWrench"
      caption         =   "추가 옵션(&I)..."
      transparent     =   -1  'True
   End
   Begin prjDownloadBooster.CommandButtonW cmdStop 
      Height          =   330
      Left            =   7200
      TabIndex        =   31
      Top             =   4815
      Width           =   1935
      _extentx        =   3413
      _extenty        =   582
      enabled         =   0   'False
      imagelist       =   "imgStopRed"
      caption         =   "중지(&P) "
      transparent     =   -1  'True
   End
   Begin prjDownloadBooster.ImageList imgDropdownReverse 
      Left            =   9840
      Top             =   5880
      _ExtentX        =   1005
      _ExtentY        =   1005
      ImageWidth      =   13
      ImageHeight     =   5
      MaskColor       =   16711935
      InitListImages  =   "frmMain.frx":2BA9
   End
   Begin prjDownloadBooster.CommandButtonW cmdOpenDropdown 
      Height          =   375
      Left            =   1800
      TabIndex        =   38
      Top             =   6960
      Width           =   255
      _extentx        =   450
      _extenty        =   661
      enabled         =   0   'False
      imagelist       =   "imgDropdown"
      imagelistalignment=   4
      transparent     =   -1  'True
   End
   Begin prjDownloadBooster.ImageList imgDropdown 
      Left            =   9840
      Top             =   5160
      _ExtentX        =   1005
      _ExtentY        =   1005
      ImageWidth      =   13
      ImageHeight     =   5
      MaskColor       =   16711935
      InitListImages  =   "frmMain.frx":2EA9
   End
   Begin prjDownloadBooster.CommandButtonW cmdDeleteDropdown 
      Height          =   375
      Left            =   5520
      TabIndex        =   41
      Top             =   6960
      Width           =   255
      _extentx        =   450
      _extenty        =   661
      enabled         =   0   'False
      imagelist       =   "imgDropdown"
      imagelistalignment=   4
      transparent     =   -1  'True
   End
   Begin prjDownloadBooster.ImageList imgPlusYellow 
      Left            =   9840
      Top             =   4440
      _ExtentX        =   1005
      _ExtentY        =   1005
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      InitListImages  =   "frmMain.frx":31A9
   End
   Begin prjDownloadBooster.CommandButtonW cmdAddToQueue 
      Height          =   330
      Left            =   7200
      TabIndex        =   32
      Top             =   5190
      Visible         =   0   'False
      Width           =   1935
      _extentx        =   3413
      _extenty        =   582
      imagelist       =   "imgPlusYellow"
      caption         =   "목록에 추가(&Q)"
      transparent     =   -1  'True
   End
   Begin prjDownloadBooster.CommandButtonW cmdStartBatch 
      Height          =   375
      Left            =   7560
      TabIndex        =   43
      Top             =   6960
      Width           =   1575
      _extentx        =   2778
      _extenty        =   661
      enabled         =   0   'False
      imagelist       =   "imgPlay"
      caption         =   "시작(&S) "
      transparent     =   -1  'True
   End
   Begin prjDownloadBooster.ImageList imgStopRed 
      Left            =   9840
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      InitListImages  =   "frmMain.frx":3591
   End
   Begin prjDownloadBooster.ImageList imgPlay 
      Left            =   9840
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      InitListImages  =   "frmMain.frx":3979
   End
   Begin prjDownloadBooster.ImageList imgDownload 
      Left            =   9840
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      InitListImages  =   "frmMain.frx":40E9
   End
   Begin prjDownloadBooster.ImageList imgMinus 
      Left            =   9840
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      InitListImages  =   "frmMain.frx":44D1
   End
   Begin prjDownloadBooster.ImageList imgOpenFile 
      Left            =   9840
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      InitListImages  =   "frmMain.frx":4C41
   End
   Begin prjDownloadBooster.ImageList imgOpenFolder 
      Left            =   9840
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      InitListImages  =   "frmMain.frx":53B1
   End
   Begin prjDownloadBooster.OptionButtonW optTabThreads2 
      Height          =   195
      Left            =   1320
      TabIndex        =   14
      Top             =   2055
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   344
      Value           =   -1  'True
      Transparent     =   -1  'True
   End
   Begin prjDownloadBooster.OptionButtonW optTabDownload2 
      Height          =   195
      Left            =   435
      TabIndex        =   12
      Top             =   2055
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   344
      Transparent     =   -1  'True
   End
   Begin VB.PictureBox fDownloadInfo 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  '없음
      Height          =   3255
      Left            =   360
      ScaleHeight     =   3255
      ScaleWidth      =   5175
      TabIndex        =   47
      Tag             =   "forcebgchange"
      Top             =   2640
      Visible         =   0   'False
      Width           =   5175
      Begin prjDownloadBooster.ucPicBuddy pbdDownloadInfo 
         Height          =   1215
         Left            =   3000
         TabIndex        =   73
         Top             =   1440
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   2143
      End
      Begin VB.Label lblMergeStatus 
         BackStyle       =   0  '투명
         Caption         =   "-"
         Height          =   255
         Left            =   1320
         TabIndex        =   66
         Top             =   2940
         Width           =   4335
      End
      Begin VB.Label Label9 
         BackStyle       =   0  '투명
         Caption         =   "조각 결합 현황:"
         Height          =   255
         Left            =   0
         TabIndex        =   67
         Top             =   2940
         Width           =   1335
      End
      Begin VB.Label lblRemaining 
         BackStyle       =   0  '투명
         Caption         =   "-"
         Height          =   255
         Left            =   1320
         TabIndex        =   68
         Top             =   2580
         Width           =   4335
      End
      Begin VB.Label Label10 
         BackStyle       =   0  '투명
         Caption         =   "남은 시간:"
         Height          =   255
         Left            =   0
         TabIndex        =   69
         Top             =   2580
         Width           =   1215
      End
      Begin VB.Label Label8 
         BackStyle       =   0  '투명
         Caption         =   "파일 이름:"
         Height          =   255
         Left            =   0
         TabIndex        =   62
         Top             =   60
         Width           =   1095
      End
      Begin VB.Label lblFilename 
         BackStyle       =   0  '투명
         Caption         =   "-"
         Height          =   255
         Left            =   1320
         TabIndex        =   61
         Top             =   60
         Width           =   4335
      End
      Begin VB.Label lblTotalSizeThread 
         BackStyle       =   0  '투명
         Caption         =   "-"
         Height          =   255
         Left            =   1320
         TabIndex        =   60
         Top             =   2220
         Width           =   4335
      End
      Begin VB.Label Label7 
         BackStyle       =   0  '투명
         Caption         =   "스레드당 크기:"
         Height          =   255
         Left            =   0
         TabIndex        =   59
         Top             =   2220
         Width           =   1215
      End
      Begin VB.Label lblThreadCount2 
         BackStyle       =   0  '투명
         Caption         =   "-"
         Height          =   255
         Left            =   1320
         TabIndex        =   58
         Top             =   1860
         Width           =   4335
      End
      Begin VB.Label Label6 
         BackStyle       =   0  '투명
         Caption         =   "스레드 수:"
         Height          =   255
         Left            =   0
         TabIndex        =   57
         Top             =   1860
         Width           =   975
      End
      Begin VB.Label Label5 
         BackStyle       =   0  '투명
         Caption         =   "속도:"
         Height          =   255
         Left            =   0
         TabIndex        =   55
         Top             =   1500
         Width           =   975
      End
      Begin VB.Label lblSpeed 
         BackStyle       =   0  '투명
         Caption         =   "-"
         Height          =   255
         Left            =   1320
         TabIndex        =   54
         Top             =   1500
         Width           =   4335
      End
      Begin VB.Label lblElapsed 
         BackStyle       =   0  '투명
         Caption         =   "-"
         Height          =   255
         Left            =   1320
         TabIndex        =   53
         Top             =   1140
         Width           =   4335
      End
      Begin VB.Label Label4 
         BackStyle       =   0  '투명
         Caption         =   "경과 시간:"
         Height          =   255
         Left            =   0
         TabIndex        =   52
         Top             =   1140
         Width           =   975
      End
      Begin VB.Label lblDownloadedBytes 
         BackStyle       =   0  '투명
         Caption         =   "-"
         Height          =   255
         Left            =   1320
         TabIndex        =   51
         Top             =   780
         Width           =   4335
      End
      Begin VB.Label Label3 
         BackStyle       =   0  '투명
         Caption         =   "받은 크기:"
         Height          =   255
         Left            =   0
         TabIndex        =   50
         Top             =   780
         Width           =   1095
      End
      Begin VB.Label lblTotalBytes 
         BackStyle       =   0  '투명
         Caption         =   "-"
         Height          =   255
         Left            =   1320
         TabIndex        =   49
         Top             =   420
         Width           =   4335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '투명
         Caption         =   "총 크기:"
         Height          =   255
         Left            =   0
         TabIndex        =   48
         Top             =   420
         Width           =   975
      End
   End
   Begin VB.PictureBox fThreadInfo 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  '없음
      Height          =   3495
      Left            =   360
      ScaleHeight     =   3495
      ScaleWidth      =   5775
      TabIndex        =   45
      Tag             =   "forcebgchange"
      Top             =   2310
      Width           =   5775
      Begin VB.PictureBox pbProgressOuterContainer 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  '없음
         Height          =   3495
         Left            =   0
         ScaleHeight     =   3495
         ScaleWidth      =   5775
         TabIndex        =   46
         Tag             =   "forcebgchange"
         Top             =   0
         Width           =   5775
         Begin VB.PictureBox pbProgressContainer 
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  '없음
            Height          =   9015
            Left            =   0
            ScaleHeight     =   9015
            ScaleWidth      =   5775
            TabIndex        =   63
            Tag             =   "forcebgchange"
            Top             =   0
            Width           =   5775
            Begin prjDownloadBooster.ProgressBar pbProgressMarquee 
               Height          =   255
               Index           =   1
               Left            =   900
               Top             =   0
               Visible         =   0   'False
               Width           =   4035
               _ExtentX        =   7117
               _ExtentY        =   450
               Step            =   10
               MarqueeSpeed    =   35
               Scrolling       =   2
            End
            Begin prjDownloadBooster.ProgressBar pbProgress 
               Height          =   255
               Index           =   1
               Left            =   900
               Top             =   0
               Visible         =   0   'False
               Width           =   4035
               _ExtentX        =   7117
               _ExtentY        =   450
               Step            =   10
               MarqueeSpeed    =   35
            End
            Begin prjDownloadBooster.ucPicBuddy pbdProgressContainer 
               Height          =   255
               Left            =   5400
               TabIndex        =   74
               Top             =   0
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   450
            End
            Begin VB.Label lblDownloader 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "스레드 1:"
               Height          =   180
               Index           =   1
               Left            =   0
               TabIndex        =   65
               Top             =   45
               Width           =   750
            End
            Begin VB.Label lblPercentage 
               Alignment       =   1  '오른쪽 맞춤
               BackStyle       =   0  '투명
               Height          =   255
               Index           =   1
               Left            =   5040
               TabIndex        =   64
               Top             =   45
               Width           =   615
            End
         End
         Begin prjDownloadBooster.ucPicBuddy pbdProgressOuterContainer 
            Height          =   375
            Left            =   5640
            TabIndex        =   76
            Top             =   360
            Width           =   135
            _ExtentX        =   238
            _ExtentY        =   661
         End
      End
      Begin prjDownloadBooster.ucPicBuddy pbdThreadInfo 
         Height          =   255
         Left            =   5760
         TabIndex        =   75
         Top             =   360
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
      End
   End
   Begin prjDownloadBooster.CommandButtonW cmdIncreaseThreads 
      Height          =   330
      Left            =   6930
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   795
      Width           =   375
      _extentx        =   661
      _extenty        =   582
      imagelistalignment=   4
      caption         =   ">"
      transparent     =   -1  'True
   End
   Begin prjDownloadBooster.CommandButtonW cmdDecreaseThreads 
      Height          =   330
      Left            =   1560
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   795
      Width           =   375
      _extentx        =   661
      _extenty        =   582
      imagelistalignment=   4
      caption         =   "<"
      transparent     =   -1  'True
   End
   Begin VB.ComboBox cbWhenExist 
      Height          =   300
      Left            =   7830
      Style           =   2  '드롭다운 목록
      TabIndex        =   21
      Top             =   2025
      Width           =   1185
   End
   Begin prjDownloadBooster.CommandButtonW cmdClear 
      Height          =   330
      Left            =   7350
      TabIndex        =   2
      Top             =   90
      Width           =   1785
      _extentx        =   3149
      _extenty        =   582
      imagelist       =   "imgErase"
      caption         =   "초기화(&Y) "
      transparent     =   -1  'True
   End
   Begin prjDownloadBooster.CommandButtonW cmdAdd 
      Height          =   375
      Left            =   2520
      TabIndex        =   39
      Top             =   6960
      Width           =   1575
      _extentx        =   2778
      _extenty        =   661
      imagelist       =   "imgPlusYellow"
      caption         =   " 추가(&R)..."
      transparent     =   -1  'True
   End
   Begin prjDownloadBooster.CommandButtonW cmdBatch 
      Height          =   330
      Left            =   7200
      TabIndex        =   33
      Top             =   5565
      Width           =   1935
      _extentx        =   3413
      _extenty        =   582
      imagelist       =   "imgDropdown"
      imagelistalignment=   1
      caption         =   "  일괄 처리(&H)"
      transparent     =   -1  'True
   End
   Begin prjDownloadBooster.FrameW fTotal 
      Height          =   615
      Left            =   240
      TabIndex        =   11
      Top             =   1320
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   1085
      Caption         =   " 전체 다운로드 진행률 "
      Transparent     =   -1  'True
      Begin VB.Label lblState 
         BackStyle       =   0  '투명
         Caption         =   "중지됨"
         Height          =   255
         Left            =   120
         TabIndex        =   71
         Top             =   285
         Width           =   735
      End
   End
   Begin prjDownloadBooster.FrameW fOptions 
      Height          =   2220
      Left            =   6720
      TabIndex        =   17
      Top             =   1320
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   3916
      Caption         =   " 옵션 "
      Transparent     =   -1  'True
      Begin prjDownloadBooster.CheckBoxW chkOpenAfterComplete 
         Height          =   255
         Left            =   90
         TabIndex        =   18
         Top             =   240
         Width           =   2175
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "완료 후 열기(&C)"
         Transparent     =   -1  'True
      End
      Begin prjDownloadBooster.CheckBoxW chkOpenFolder 
         Height          =   255
         Left            =   90
         TabIndex        =   19
         Top             =   465
         Width           =   2280
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "완료 후 폴더 열기(&L)"
         Transparent     =   -1  'True
      End
      Begin prjDownloadBooster.CheckBoxW chkContinueDownload 
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   990
         Width           =   1935
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "항상 이어받기(&J)"
         Transparent     =   -1  'True
      End
      Begin prjDownloadBooster.CheckBoxW chkAutoRetry 
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1215
         Width           =   2205
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "오류 시 자동 재시도(&G)"
         Transparent     =   -1  'True
      End
      Begin VB.Line lbOptionsHeader 
         BorderColor     =   &H80000010&
         Visible         =   0   'False
         X1              =   600
         X2              =   2415
         Y1              =   90
         Y2              =   90
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "."
         Height          =   180
         Left            =   0
         TabIndex        =   72
         Top             =   0
         Visible         =   0   'False
         Width           =   60
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "중복(&K):"
         Height          =   180
         Left            =   330
         TabIndex        =   20
         Top             =   765
         Width           =   690
      End
   End
   Begin prjDownloadBooster.CommandButtonW cmdOpenFolder 
      Height          =   330
      Left            =   7200
      TabIndex        =   29
      Top             =   4440
      Width           =   1935
      _extentx        =   3413
      _extenty        =   582
      imagelist       =   "imgOpenFolder"
      caption         =   "폴더 열기(&E) "
      transparent     =   -1  'True
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
      TabIndex        =   8
      Top             =   750
      Width           =   4995
      _ExtentX        =   8811
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
      Left            =   7350
      TabIndex        =   5
      Top             =   435
      Width           =   1785
      _extentx        =   3149
      _extenty        =   582
      imagelist       =   "imgOpenFolder"
      caption         =   " 찾아보기(&B)..."
      transparent     =   -1  'True
   End
   Begin VB.TextBox txtFileName 
      Height          =   300
      Left            =   1560
      TabIndex        =   4
      Top             =   450
      Width           =   5745
   End
   Begin prjDownloadBooster.CommandButtonW cmdGo 
      Height          =   330
      Left            =   7200
      TabIndex        =   30
      Top             =   4815
      Width           =   1935
      _extentx        =   3413
      _extenty        =   582
      imagelist       =   "imgDownload"
      caption         =   "다운로드(&D) "
      transparent     =   -1  'True
   End
   Begin prjDownloadBooster.FrameW Frame4 
      Height          =   3885
      Left            =   240
      TabIndex        =   56
      Top             =   2040
      Width           =   6255
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "　　　　　　　　　　　"
      Transparent     =   -1  'True
      Begin VB.Label fTabDownload 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "요약"
         Height          =   180
         Left            =   450
         TabIndex        =   13
         Top             =   30
         Width           =   360
      End
      Begin VB.Label fTabThreads 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "스레드"
         Height          =   180
         Left            =   1335
         TabIndex        =   15
         Top             =   30
         Width           =   540
      End
   End
   Begin prjDownloadBooster.CommandButtonW cmdOpenFileDropdown 
      Height          =   330
      Left            =   8880
      TabIndex        =   28
      Top             =   4065
      Width           =   255
      _extentx        =   450
      _extenty        =   582
      enabled         =   0   'False
      imagelist       =   "imgDropdown"
      imagelistalignment=   4
      transparent     =   -1  'True
   End
   Begin VB.Image imgLBContentBackground 
      Height          =   375
      Left            =   10920
      Stretch         =   -1  'True
      Top             =   1440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgBorderBottomRight 
      Height          =   135
      Left            =   11040
      Picture         =   "frmMain.frx":5799
      Top             =   600
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image imgBorderBottomLeft 
      Height          =   135
      Left            =   10680
      Picture         =   "frmMain.frx":57E7
      Top             =   600
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image imgBorderTopRight 
      Height          =   135
      Left            =   11040
      Picture         =   "frmMain.frx":5833
      Top             =   240
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image imgBorderTopLeft 
      Height          =   135
      Left            =   10680
      Picture         =   "frmMain.frx":587F
      Top             =   240
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape pgBorderRight 
      BackColor       =   &H00C8D0D4&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H005F5F5C&
      Height          =   135
      Left            =   10680
      Top             =   1680
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Shape pgBorderLeft 
      BackColor       =   &H00C8D0D4&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H005F5F5C&
      Height          =   135
      Left            =   10680
      Top             =   1440
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Shape pgBorderBottom 
      BackColor       =   &H00C8D0D4&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H005F5F5C&
      Height          =   45
      Left            =   10680
      Top             =   1200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Shape pgBorderTop 
      BackColor       =   &H00C8D0D4&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H005F5F5C&
      Height          =   45
      Left            =   10680
      Top             =   960
      Visible         =   0   'False
      Width           =   615
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
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   240
      TabIndex        =   70
      Tag             =   "nosizechange"
      Top             =   1320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblLBCaptionShadow 
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
      Left            =   255
      TabIndex        =   77
      Tag             =   "nocolorsizechange"
      Top             =   1335
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Line lnTygemFrameBottom 
      Visible         =   0   'False
      X1              =   210
      X2              =   6495
      Y1              =   6030
      Y2              =   6030
   End
   Begin VB.Line lnTygemFrameRight 
      Visible         =   0   'False
      X1              =   6585
      X2              =   6585
      Y1              =   1575
      Y2              =   5940
   End
   Begin prjDownloadBooster.ShellPipe spYtdl 
      Left            =   9240
      Top             =   3360
      _ExtentX        =   635
      _ExtentY        =   635
   End
   Begin VB.Image imgTopLeft 
      Height          =   435
      Left            =   120
      Top             =   1200
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.Image imgTop 
      Height          =   435
      Left            =   1845
      Stretch         =   -1  'True
      Top             =   1200
      Visible         =   0   'False
      Width           =   3585
   End
   Begin VB.Image imgTopRight 
      Height          =   435
      Left            =   5430
      Top             =   1200
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Image imgLeft 
      Height          =   2310
      Left            =   120
      Stretch         =   -1  'True
      Top             =   1635
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.Image imgBottomLeft 
      Height          =   180
      Left            =   120
      Top             =   3945
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.Image imgBottom 
      Height          =   180
      Left            =   1845
      Stretch         =   -1  'True
      Top             =   3945
      Visible         =   0   'False
      Width           =   3585
   End
   Begin VB.Image imgBottomRight 
      Height          =   180
      Left            =   5430
      Top             =   3945
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Image imgRight 
      Height          =   2310
      Left            =   5415
      Stretch         =   -1  'True
      Top             =   1635
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Image imgCenter 
      Height          =   2310
      Left            =   1845
      Stretch         =   -1  'True
      Top             =   1635
      Visible         =   0   'False
      Width           =   3570
   End
   Begin VB.Label lblThreadCount 
      BackStyle       =   0  '투명
      Caption         =   "(스레드 1개)"
      Height          =   255
      Left            =   7350
      TabIndex        =   10
      Top             =   870
      Width           =   1695
   End
   Begin VB.Label lblThreadCountLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "강도(&T):"
      Height          =   180
      Left            =   240
      TabIndex        =   6
      Top             =   870
      Width           =   690
   End
   Begin VB.Label lblFilePath 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "저장 경로(&F):"
      Height          =   180
      Left            =   240
      TabIndex        =   3
      Top             =   495
      Width           =   1095
   End
   Begin VB.Label lblURL 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "파일 주소(&A):"
      Height          =   180
      Left            =   240
      TabIndex        =   0
      Top             =   150
      Width           =   1110
   End
   Begin prjDownloadBooster.ShellPipe SP 
      Left            =   9240
      Top             =   3960
      _ExtentX        =   635
      _ExtentY        =   635
   End
   Begin VB.Label lblURLShadow 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "파일 주소(&A):"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   255
      TabIndex        =   78
      Tag             =   "nocolorchange"
      Top             =   165
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.Label lblFilePathShadow 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "저장 경로(&F):"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   255
      TabIndex        =   79
      Tag             =   "nocolorchange"
      Top             =   510
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblThreadCountLabelShadow 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "강도(&T):"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   255
      TabIndex        =   80
      Tag             =   "nocolorchange"
      Top             =   885
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image imgBackgroundTile 
      Height          =   135
      Index           =   0
      Left            =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image imgBackground 
      Height          =   135
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   135
   End
   Begin VB.Shape pgPattern 
      BorderStyle     =   0  '투명
      Height          =   255
      Left            =   0
      Top             =   0
      Width           =   255
   End
   Begin VB.Menu mnuListContext 
      Caption         =   "mnuListContext"
      Visible         =   0   'False
      Begin VB.Menu mnuOpenBatch 
         Caption         =   "열기(&O)"
      End
      Begin VB.Menu mnuOpenFolder2 
         Caption         =   "폴더 열기(&F)"
      End
      Begin VB.Menu mnuErrorInfo 
         Caption         =   "오류 정보(&I)..."
      End
      Begin VB.Menu mnuSepOpen 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "편집(&E)..."
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddItem2 
         Caption         =   "새 주소 추가(&A)..."
      End
      Begin VB.Menu mnuDeleteItem 
         Caption         =   "제거(&R)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuClearBatch3 
         Caption         =   "모두 제거(&C)"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMoveUp 
         Caption         =   "위로 이동(&U)"
      End
      Begin VB.Menu mnuMoveDown 
         Caption         =   "아래로 이동(&D)"
      End
   End
   Begin VB.Menu mnuListContext2 
      Caption         =   "mnuListContext2"
      Visible         =   0   'False
      Begin VB.Menu mnuAddItem 
         Caption         =   "새 주소 추가(&A)..."
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
      Begin VB.Menu mnuPropertiesBatch 
         Caption         =   "속성 보기(&R)"
      End
   End
   Begin VB.Menu mnuOpenFileDropdown 
      Caption         =   "mnuOpenFileDropdown"
      Visible         =   0   'False
      Begin VB.Menu mnuProperties 
         Caption         =   "속성 보기(&R)"
      End
   End
   Begin VB.Menu mnuDownloadOptions 
      Caption         =   "mnuDownloadOptions"
      Visible         =   0   'False
      Begin VB.Menu mnuYtdlOptions 
         Caption         =   "&youtube-dl..."
      End
      Begin VB.Menu mnuHeaders 
         Caption         =   "헤더(&H)..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements ISubclass

Dim Elapsed As Long
Dim BatchStarted As Boolean
Dim CurrentBatchIdx As Integer
Dim DownloadPath As String
Dim IsDownloading As Boolean
Dim BatchErrorCount As Integer
Dim BatchErrorAllCount As Integer
Public ScrollOneScreen As Boolean
Dim PrevDownloadedBytes As Double
Dim SpeedCount As Integer
Dim HttpStatusCode As String
Dim ResumeUnsupported As Boolean
Public ImagePosition As Integer
Dim TotalSize As Double
Dim FormCaption$
Dim LBFrameEnabled As Boolean
Dim ErrorCodeDescription As Collection
Dim ThreadBuddyAttached As Boolean

Const MAIN_FORM_WIDTH As Long = 9450

'youtube-dl 관련 변수
Dim ytdlTotalFormatCount As Integer
Dim ytdlFileName As String
Public ytdlEnabled As Boolean
Public ytdlFormat As String
Public ytdlExtractAudio As Boolean
Public ytdlAudioFormat As AudioFormat
Public ytdlAudioBitrateType As AudioBitrateType
Public ytdlAudioCBR As Integer
Public ytdlAudioVBR As Byte

Dim MAX_THREAD_COUNT As Integer

Dim MaxLoadedTileBackgroundImage As Long

Dim FormWidth As Long
Dim FormMinHeight As Long
Dim FormMaxHeight As Long

#If HIDEYTDL Then
#Else
Sub StartYtdlDownload()
    If Not FileExists(GetSetting("DownloadBooster", "Options", "YtdlPath", "")) Then
        If Confirm(t("youtube-dl 실행 파일 경로가 지정되지 않았습니다. 지금 지정하시겠습니까?", "youtube-dl executable path is not specified. Would you like to specify it now?"), App.Title) = vbYes Then
            Load frmOptions
            frmOptions.tsTabStrip.Tabs(6).Selected = -1
            frmOptions.Show vbModal, Me
        End If
        Exit Sub
    End If

    If lvLogTest.ColumnHeaders.Count < 2 Then
        lvLogTest.ColumnHeaders.Add , , "주체", 1200
        lvLogTest.ColumnHeaders.Add , , "out", 7200
    End If
    
    ytdlTotalFormatCount = 1
    spYtdl.Run """" & GetSetting("DownloadBooster", "Options", "YtdlPath", "") & """ 8igShgEtHK8"
End Sub
#End If

Private Sub cmdDownloadOptions_Click()
    Tags.DownloadOptionsTargetForm = 0
    Set frmDownloadOptions.Headers = SessionHeaders
    Set frmDownloadOptions.HeaderKeys = SessionHeaderKeys
    frmDownloadOptions.Show vbModal, Me
End Sub

Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse)
    '
End Property

Private Property Get ISubclass_MsgResponse() As EMsgResponse
    ISubclass_MsgResponse = emrConsume
End Property

Private Function ISubclass_WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error Resume Next
    
    Dim hSysMenu As Long
    Dim MII As MENUITEMINFO
    
    Select Case uMsg
        Case WM_GETMINMAXINFO
            Dim lpMMI As MINMAXINFO
            CopyMemory lpMMI, ByVal lParam, Len(lpMMI)
            lpMMI.ptMinTrackSize.X = FormWidth * (DPI / 96)
            lpMMI.ptMinTrackSize.Y = FormMinHeight * (DPI / 96)
            lpMMI.ptMaxTrackSize.X = FormWidth * (DPI / 96)
            lpMMI.ptMaxTrackSize.Y = FormMaxHeight * (DPI / 96)
            CopyMemory ByVal lParam, lpMMI, Len(lpMMI)
            
            ISubclass_WindowProc = 1&
            Exit Function
        Case WM_INITMENU
            hSysMenu = GetSystemMenu(Me.hWnd, 0)
            With MII
                .cbSize = Len(MII)
                .fMask = MIIM_STATE
                .fState = MFS_ENABLED Or IIf(MainFormOnTop, MFS_CHECKED, 0)
            End With
            SetMenuItemInfo hSysMenu, 1000, 0, MII
            
            ISubclass_WindowProc = 1&
            Exit Function
        Case WM_SYSCOMMAND
            If wParam = 1000 Then '항상 위에 표시
                MainFormOnTop = Not MainFormOnTop
                SetWindowPos hWnd, IIf(MainFormOnTop, hWnd_TOPMOST, hWnd_NOTOPMOST), 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
                SaveSetting "DownloadBooster", "Options", "AlwaysOnTop", Abs(CInt(MainFormOnTop))
                
                ISubclass_WindowProc = 1&
                Exit Function
            ElseIf wParam = 1001 And (Not (frmMain.Height <= 6930 + PaddedBorderWidth * 15 * 2)) Then '일괄처리 접기
                cmdBatch_Click
                
                ISubclass_WindowProc = 1&
                Exit Function
            ElseIf wParam = 1002 And (frmMain.Height <= 6930 + PaddedBorderWidth * 15 * 2) Then '일괄처리 펼치기
                cmdBatch_Click
                
                ISubclass_WindowProc = 1&
                Exit Function
            ElseIf wParam = 1003 And (Not (frmMain.Height <= 6930 + PaddedBorderWidth * 15 * 2)) Then '창 크기 초기화
                Me.Height = 8985 + PaddedBorderWidth * 15 * 2
            
                ISubclass_WindowProc = 1&
                Exit Function
            End If
'        Case WM_DWMCOMPOSITIONCHANGED
        Case WM_SETTINGCHANGE
            Select Case GetStrFromPtr(lParam)
                Case "WindowMetrics"
                    UpdateBorderWidth
                    
                    FormWidth = (9450 + PaddedBorderWidth * 15 * 2) / 15
                    FormMinHeight = (8220 + PaddedBorderWidth * 15 * 2) / 15
                    
                    Me.Width = FormWidth * 15
                    Form_Resize
                    
                    On Error Resume Next
                    Dim ctrl As Control
                    For Each ctrl In frmMain.Controls
                        If TypeName(ctrl) = "FrameW" Or TypeName(ctrl) = "CheckBoxW" Or TypeName(ctrl) = "OptionButtonW" Or TypeName(ctrl) = "CommandButtonW" Or TypeName(ctrl) = "Slider" Then ctrl.Refresh
                    Next ctrl
                    Dim PrevTrackerVisualStyles As Boolean
                    PrevTrackerVisualStyles = trThreadCount.VisualStyles
                    trThreadCount.VisualStyles = False
                    trThreadCount.VisualStyles = True
                    trThreadCount.VisualStyles = PrevTrackerVisualStyles
                    
                    SetTextColors
            End Select
        Case WM_THEMECHANGED
            SetTextColors
        Case WM_CTLCOLORSCROLLBAR
            ISubclass_WindowProc = 0&
            Exit Function
    End Select
    
    ISubclass_WindowProc = CallOldWindowProc(hWnd, uMsg, wParam, lParam)
End Function

Private Sub mnuErrorInfo_Click()
    If lvBatchFiles.SelectedItem Is Nothing Then Exit Sub
    If lvBatchFiles.SelectedItem.ForeColor <> vbRed Then Exit Sub
    Dim StatusString$
    StatusString = lvBatchFiles.SelectedItem.ListSubItems(3).Text
    StatusString = Mid(StatusString, InStr(StatusString, "(") + 1)
    StatusString = Left$(StatusString, Len(StatusString) - 1)
    If Not IsNumeric(StatusString) Then
        MsgBox t("오류 정보를 표시할 수 없습니다.", "Unable to show the error information."), 16
        Exit Sub
    End If
    MsgBox t("오류 코드", "Error code") & ": " & StatusString & vbCrLf & t("설명", "Description") & ": " & IIf(Exists(ErrorCodeDescription, StatusString), ErrorCodeDescription(StatusString), t("설명이 없습니다.", "Description is unavailable")), 64, t("오류 정보", "Error information")
End Sub

Private Sub mnuHeaders_Click()
    Tags.DownloadOptionsTargetForm = 0
    Set frmDownloadOptions.Headers = SessionHeaders
    Set frmDownloadOptions.HeaderKeys = SessionHeaderKeys
#If HIDEYTDL Then
    frmDownloadOptions.Show vbModal, Me
    Exit Sub
#End If
    frmDownloadOptions.tsTabStrip.Tabs(2).Selected = True
    frmDownloadOptions.Show vbModal, Me
End Sub

Private Sub mnuYtdlOptions_Click()
    cmdDownloadOptions_Click
End Sub

#If HIDEYTDL Then
#Else
Private Sub spYtdl_DataArrival(ByVal CharsTotal As Long)
    Dim LinesLF() As String, LinesCR() As String, Data() As String
    LinesLF = Split(spYtdl.GetData(), vbLf)
    Dim Line As String
    Dim i%, k%
    For i = LBound(LinesLF) To UBound(LinesLF)
        LinesCR = Split(LinesLF(i), vbCr)
        For k = LBound(LinesCR) To UBound(LinesCR)
            Line = Trim$(LinesCR(k))
            If Line = "" Then GoTo nextLine
            Do While Replace(Line, "  ", " ") <> Line
                Line = Replace(Line, "  ", " ")
            Loop
            Data = Split(Line, " ")
            
            On Error Resume Next
            Select Case Data(0)
                Case "[info]"
                    '포맷 개수
                    If Includes(Line, "Downloading ") And Includes(Line, " format(s): ") Then
                        ytdlTotalFormatCount = UBound(Split(Data(5), "+")) + 1
                    End If
                    lvLogTest.ListItems.Add(, , "정보").ListSubItems.Add , , Line
                Case "[download]"
                    If UBound(Data) > 1 Then
                        If Data(1) = "Destination:" Then
                            ytdlFileName = Replace(Line, "[download] Destination: ", "")
                        End If
                    End If
                    lvLogTest.ListItems.Add(, , "다운로드").ListSubItems.Add , , Line
                Case "[Merger]"
                    If UBound(Data) > 3 Then
                        If Data(1) = "Merging" And Data(3) = "into" Then
                            ytdlFileName = Replace(Line, "[Merger] Merging formats into ", "")
                            If Left$(ytdlFileName, 1) = """" And Right$(ytdlFileName, 1) = """" Then
                                ytdlFileName = Mid$(ytdlFileName, 2, Len(ytdlFileName) - 2)
                            End If
                        End If
                    End If
                    lvLogTest.ListItems.Add(, , "합체").ListSubItems.Add , , Line
                Case Else
                    lvLogTest.ListItems.Add(, , Data(0)).ListSubItems.Add , , Line
            End Select
        
nextLine:
        Next k
    Next i
End Sub
#End If

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
                pbTotalProgressMarquee.Visible = -1
                pbTotalProgressMarquee.MarqueeAnimation = -1
                cmdStop.Enabled = 0
                If GetSetting("DownloadBooster", "Options", "ExcludeMergeFromElapsed", "0") = "1" Then timElapsed.Enabled = 0
            Case "COMPLETE"
                sbStatusBar.Panels(1).Text = t("완료", "Complete")
                sbStatusBar.Panels(2).Text = ""
                sbStatusBar.Panels(3).Text = ""
                sbStatusBar.Panels(4).Text = ""
                pbTotalProgressMarquee.MarqueeAnimation = 0
                pbTotalProgressMarquee.Visible = 0
                pbTotalProgress.Value = 100
            Case "UNABLETOCONTINUE"
                Alert t("이어받기가 불가능합니다. 처음부터 다시 다운로드합니다.", "Unable to resume. Starting over..."), App.Title, 48, False, 5000
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
                pbProgressMarquee(idx).Visible = -1
                pbProgressMarquee(idx).MarqueeAnimation = -1
            End If
            lblPercentage(idx).Caption = ""
        Else
            If pbProgressMarquee(idx).Visible Then
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
            If DownloadedBytes <> -1 Then timElapsed.Enabled = -1
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
                timElapsed.Enabled = -1
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
                TotalSize = total
            End If
            lblDownloadedBytes.Caption = ParseSize(DownloadedBytes, True)
            pbTotalProgress.Value = progress
            fTotal.Caption = t(" 전체 다운로드 진행률 (" & progress & "%) ", " Total Progress (" & progress & "%) ")
            If Not BatchStarted Then SetTitle progress & "% " & t("다운로드 중", "Downloading")
        End If
        
        Dim Speed As Double
        SpeedCount = SpeedCount + 1
        If SpeedCount >= 10 Then
            Speed = (DownloadedBytes - PrevDownloadedBytes)
            lblSpeed.Caption = ParseSize(Speed, True, "/" & t("초", "sec"))
            sbStatusBar.Panels(3).Text = ParseSize(Speed, False, "/" & t("초", "sec"))
            PrevDownloadedBytes = DownloadedBytes
            SpeedCount = 0
            
            If progress >= 0 And strTotal <> "-1" And IsNumeric(strTotal) And Speed > 0 Then
                lblRemaining = FormatTime((CDbl(strTotal) - CDbl(DownloadedBytes)) / Speed)
            End If
        End If
    ElseIf Left$(Data, 17) = "MODIFIEDFILENAME " Then
        output = StrConv(atob(Right$(Data, Len(Data) - 17)), vbUnicode)
        DownloadPath = output
        lblFilename.Caption = GetFilename(output)
        If BatchStarted Then
            lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(1).Text = output
            lvBatchFiles.ListItems(CurrentBatchIdx).Text = lblFilename.Caption
        End If
        If Len(lblFilename.Caption) > 22 Then lblFilename.Caption = Left$(lblFilename.Caption, 22) & "..."
    ElseIf Left$(Data, 10) = "MERGESIZE " Then
        On Error GoTo exitif
        Dim MergedSize As Double
        MergedSize = CDbl(Right$(Data, Len(Data) - 10))
        If TotalSize <= 0 Then GoTo exitif
        lblMergeStatus.Caption = t(ParseSize(TotalSize) & " 중 " & ParseSize(MergedSize), ParseSize(MergedSize) & " of " & ParseSize(TotalSize)) & " (" & Fix((MergedSize / TotalSize) * 100) & "%)"
exitif:
    ElseIf Left$(Data, 11) = "DELETEITEM " Then
        On Error Resume Next
        MsgBox Right$(Data, Len(Data) - 11)
        Kill Right$(Data, Len(Data) - 11)
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
        cmdStopBatch.Enabled = 0
        cmdStopBatch.Left = Me.Width + 1200
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
            Alert t("하나 이상의 오류가 발생했습니다. 해당 항목을 두 번 누르면 오류 정보를 볼 수 있습니다.", _
                    "One or more errors have occurred. Double click the error item to see details."), App.Title, 48
        ElseIf GetSetting("DownloadBooster", "Options", "PlaySound", 1) <> 0 And BatchErrorAllCount <= 0 Then
            PlayWave Trim$(GetSetting("DownloadBooster", "Options", "CompleteSoundPath", "")), FallbackSound:=Information
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
    
    If GetSetting("DownloadBooster", "Options", "LazyElapsed", "0") = "1" Then
        timElapsed.Enabled = 0
    Else
        timElapsed.Enabled = -1
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
                        Alert t("해당 파일 주소에 연결할 수 없습니다. 주소가 유효하지 않거나 서버가 응답하지 않습니다.", "The server does not respond or the file URL is invalid."), App.Title, 16
                    Else
                        Alert t("서버와의 접속이 끊겼습니다. 다운로드 도중에 네트워크 오류가 발생했을 수 있습니다.", "Network error while downloading."), App.Title, 16
                    End If
                End If
            Case 102
                Alert "주소나 파일 이름을 지정하지 않았습니다.", App.Title, 16
            Case 3, 103
                Alert t("저장 경로가 존재하지 않습니다.", "Save path doesn't exist."), App.Title, 16
            Case 104
                Alert t("저장할 파일명이 사용 중입니다. 다른 이름을 선택하십시오.", "File name already exists."), App.Title, 16
            Case 106
                Alert t("파일 서버가 다운로드 부스트를 지원하지 않습니다. 강도를 1로 변경해 보십시오.", "Download boosting not supported. Try changing the thread count to 1."), App.Title, 16
            Case 107
                Alert t("파일의 크기를 알 수 없어서 다운로드를 부스트할 수 없습니다. 강도를 1로 변경해 보십시오.", "Unable to boost download because the file size is not provided. Try changing the thread count to 1."), App.Title, 16
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
                Alert t("서버가 요청을 거부했습니다. " & ErrDesc & statusMsg, "Server denied your request. The file may not exist or have insufficient permissions to access it."), App.Title, Icon
            Case Else
                Alert t("내부 오류가 발생했습니다. 프로세스 반환 값은 ( " & RetVal & " ) 입니다.", "Internal error. Process returned ( " & RetVal & " )."), App.Title, 16
        End Select
    End If
    
nextln:
    
    If Not BatchStarted Then
        cmdGo.Enabled = -1
    End If
    cmdStop.Enabled = 0
    cmdStop.Left = Me.Width + 1200
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
            If RetVal <> 999& Then BatchErrorCount = BatchErrorCount + 1
            BatchErrorAllCount = BatchErrorAllCount + 1
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
        cmdOpenFileDropdown.Enabled = -1
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
    TotalSize = 0
    If Not BatchStarted Then
        cmdStop.Enabled = -1
        cmdStop.Left = cmdGo.Left
        cmdStop.Refresh
        cmdGo.Visible = 0
    Else
        cmdStop.Enabled = 0
        cmdStop.Left = Me.Width + 1200
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
    
    cmdDownloadOptions.Enabled = 0
    
    lblThreadCount.Enabled = 0
    
    cmdStartBatch.Enabled = 0
    
    cmdOpen.Enabled = 0
    cmdOpenFileDropdown.Enabled = 0
    
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
    lblMergeStatus.Caption = "-"
    
    fTotal.Caption = t(" 전체 다운로드 진행률 ", " Total Progress ")
    pbTotalProgress.Value = 0
    Dim i%
    For i = 1 To trThreadCount.Value
        lblPercentage(i).Caption = ""
        pbProgress(i).Value = 0
    Next i
    
    For i = 1 To trThreadCount.Value
        pbProgressMarquee(i).Visible = -1
        pbProgressMarquee(i).MarqueeAnimation = -1
    Next i
    
    pbTotalProgressMarquee.Visible = -1
    pbTotalProgressMarquee.MarqueeAnimation = -1
    
    lblState.Caption = t("진행 중", "Working")
    sbStatusBar.Panels(1).Text = t("시작 중...", "Starting...")
    
    If BatchStarted Then
'        Dim BatchCount%
'        BatchCount = 0
'        For i = 1 To lvBatchFiles.ListItems.Count
'            BatchCount = BatchCount + Abs(CInt(lvBatchFiles.ListItems(i).Checked))
'        Next i
        SetTitle t(lvBatchFiles.ListItems.Count & "개 중 " & CurrentBatchIdx & "번째 항목 다운로드 중", "Downloading " & CurrentBatchIdx & " of " & lvBatchFiles.ListItems.Count)
    Else
        SetTitle t("다운로드 중", "Downloading")
    End If
End Sub

Sub OnStop(Optional PlayBeep As Boolean = True)
    IsDownloading = False
    If Not BatchStarted Then
        cmdGo.Enabled = -1
    End If
    cmdStop.Enabled = 0
    cmdStop.Left = Me.Width + 1200
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
    
    cmdDownloadOptions.Enabled = -1
    
    lblThreadCount.Enabled = -1
    
    SP.FinishChild 0, 0
    
    Dim i%
    For i = 1 To trThreadCount.Value
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
            Else
                cmdStartBatch.Enabled = -1
            End If
        Else
            cmdStartBatch.Enabled = 0
        End If
        
        If PlayBeep And GetSetting("DownloadBooster", "Options", "PlaySound", 1) <> 0 Then
            PlayWave Trim$(GetSetting("DownloadBooster", "Options", "CompleteSoundPath", "")), FallbackSound:=Information
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
    
    SetTitle
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
    txtFileName.Text = Trim$(txtFileName.Text)
    If FolderExists(txtFileName.Text) Then
        frmBatchAdd.txtSavePath.Text = txtFileName.Text
    Else
        frmBatchAdd.txtSavePath.Text = GetParentFolderName(txtFileName.Text)
    End If
    frmBatchAdd.Show vbModal, Me
End Sub

Function AddBatchURLs(URL As String, Optional ByVal SavePath As String = "", Optional ByVal Headers As String = "") As Boolean
    If Left$(URL, 7) <> "http://" And Left$(URL, 8) <> "https://" Then
        Alert URL & " - " & t("주소가 올바르지 않습니다. 'http://' 또는 'https://'로 시작해야 합니다.", "Invalid address. Must start with 'http://' or 'https://'."), App.Title, 16
        AddBatchURLs = False
        Exit Function
    End If
    
    If Headers = "-" Then Headers = SessionHeaderCache
    
    If Trim$(SavePath) = "" Then SavePath = txtFileName.Text
    SavePath = Trim$(SavePath)
    Do While Replace(SavePath, "\\", "\") <> SavePath
        SavePath = Replace(SavePath, "\\", "\")
    Loop

    Dim idx%
    Dim FileName$
    Dim ServerName$
    FileName = SavePath
    If FolderExists(FileName) Then
        If Not (Right$(FileName, 1) = "\") Then FileName = FileName & "\"
        ServerName = FilterFilename(ExcludeParameters(URLDecode(Split(URL, "/")(UBound(Split(URL, "/"))))))
        If Replace(ServerName, " ", "") = "" Then ServerName = "download_" & CStr(Rnd * 1E+15)
        FileName = FileName & ServerName
    Else
        ServerName = FilterFilename(ExcludeParameters(URLDecode(Split(URL, "/")(UBound(Split(URL, "/"))))))
        If Replace(ServerName, " ", "") = "" Then
            ServerName = "download_" & CStr(Rnd * 1E+15)
        Else
            ServerName = CStr(Rnd * 1E+15) & "_" & ServerName
        End If
        FileName = GetParentFolderName(txtFileName.Text) & "\"
        FileName = Replace(FileName, "\\", "\") & ServerName
    End If
    idx = lvBatchFiles.ListItems.Add(, , ServerName).Index
    lvBatchFiles.ListItems(idx).ListSubItems.Add , , FileName
    lvBatchFiles.ListItems(idx).ListSubItems.Add , , URL
    lvBatchFiles.ListItems(idx).ListSubItems.Add , , t("대기", "Queued")
    lvBatchFiles.ListItems(idx).ListSubItems.Add , , "Y"
    lvBatchFiles.ListItems(idx).ListSubItems.Add , , Headers
#If HIDEYTDL Then
    GoTo afterheaderadd
#End If
    lvBatchFiles.ListItems(idx).ListSubItems.Add , , "N"
    lvBatchFiles.ListItems(idx).ListSubItems.Add , , ""
    lvBatchFiles.ListItems(idx).ListSubItems.Add , , "N"
    lvBatchFiles.ListItems(idx).ListSubItems.Add , , ""
    lvBatchFiles.ListItems(idx).ListSubItems.Add , , ""
    lvBatchFiles.ListItems(idx).ListSubItems.Add , , ""
    lvBatchFiles.ListItems(idx).ListSubItems.Add , , ""
afterheaderadd:
    lvBatchFiles.ListItems(idx).Checked = -1
    If IsDownloading Or cmdStop.Enabled Or BatchStarted Then
        cmdStartBatch.Enabled = 0
    Else
        cmdStartBatch.Enabled = -1
    End If
    AddBatchURLs = True
End Function

Private Sub cmdAddToQueue_Click()
    If Replace(txtURL.Text, " ", "") = "" Then
        Alert t("파일 주소를 입력하십시오.", "Specify the file URL."), App.Title, 64
        Exit Sub
    End If
    On Error GoTo justadd
    If GetSetting("DownloadBooster", "Options", "AllowDuplicatesInQueue", 0) <> 0 Then GoTo justadd
    Dim i%
    If lvBatchFiles.ListItems.Count Then
        For i = 1 To lvBatchFiles.ListItems.Count
            If lvBatchFiles.ListItems(i).ListSubItems(2).Text = Trim$(txtURL.Text) Then
                Alert t("해당 주소는 이미 대기열에 추가되었습니다.", "That URL is already added"), App.Title, 64
                Exit Sub
            End If
        Next i
    End If
justadd:
    AddBatchURLs txtURL.Text, , "-"
End Sub

Sub cmdBatch_Click()
    On Error Resume Next
    
    Dim hSysMenu As Long
    Dim MII As MENUITEMINFO
    hSysMenu = GetSystemMenu(Me.hWnd, 0)
    
    If Me.Height <= 6930 + PaddedBorderWidth * 15 * 2 Then
        cmdBatch.ImageList = imgDropdownReverse
        lvBatchFiles.Visible = -1
        cmdAddToQueue.Visible = -1
        FormWidth = (MAIN_FORM_WIDTH + PaddedBorderWidth * 15 * 2) / 15
        FormMinHeight = (8220 + PaddedBorderWidth * 15 * 2 + 45) / 15
        FormMaxHeight = (Screen.Height + 1200) / 15
        'sbStatusBar.AllowSizeGrip = True
        
        Dim formHeight As Integer
        formHeight = GetSetting("DownloadBooster", "UserData", "FormHeight", 8985)
        If formHeight < 8220 Then
            Me.Height = 8985 + PaddedBorderWidth * 15 * 2
        Else
            Me.Height = formHeight + PaddedBorderWidth * 15 * 2
        End If
        
        CheckMenuRadioItem hSysMenu, 1001, 1002, 1002, MF_BYCOMMAND
        
        With MII
            .cbSize = Len(MII)
            .fMask = MIIM_STATE
            .fState = MFS_ENABLED
        End With
        SetMenuItemInfo hSysMenu, 1003, 0, MII
    Else
        SaveSetting "DownloadBooster", "UserData", "FormHeight", Me.Height - PaddedBorderWidth * 15 * 2
        FormWidth = (MAIN_FORM_WIDTH + PaddedBorderWidth * 15 * 2) / 15
        FormMinHeight = (6930 + PaddedBorderWidth * 15 * 2) / 15
        FormMaxHeight = (6930 + PaddedBorderWidth * 15 * 2) / 15
        'sbStatusBar.AllowSizeGrip = False
        Me.Height = 6930 + PaddedBorderWidth * 15 * 2
        cmdBatch.ImageList = imgDropdown
        lvBatchFiles.Visible = 0
        cmdAddToQueue.Visible = 0
        
        CheckMenuRadioItem hSysMenu, 1001, 1002, 1001, MF_BYCOMMAND
        
        With MII
            .cbSize = Len(MII)
            .fMask = MIIM_STATE
            .fState = MFS_GRAYED
        End With
        SetMenuItemInfo hSysMenu, 1003, 0, MII
    End If
    SetBackgroundPosition
End Sub

Private Sub cmdBrowse_Click()
    Tags.BrowsePresetPath = ""
    Tags.BrowseTargetForm = 0
    If GetSetting("DownloadBooster", "Options", "ForceWin31Dialog", "0") = "1" Then
        frmBrowse.Show vbModal, Me
        Exit Sub
    End If
    frmExplorer.Show vbModal, Me
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
            cmdDeleteDropdown.Enabled = 0
            cmdEdit.Enabled = 0
        ElseIf lvBatchFiles.SelectedItem.Text <> "" And lvBatchFiles.SelectedItem.Selected Then
            cmdDelete.Enabled = -1
            cmdDeleteDropdown.Enabled = -1
            cmdEdit.Enabled = -1
        Else
            cmdDelete.Enabled = 0
            cmdDeleteDropdown.Enabled = 0
            cmdEdit.Enabled = 0
        End If
        GoTo L2
L1:
        cmdDelete.Enabled = 0
        cmdDeleteDropdown.Enabled = 0
        cmdEdit.Enabled = 0
L2:
        On Error GoTo 0
    End If
    
    URL = Trim$(URL)
    FileName = Trim$(FileName)
    
    OnStart
    
    Dim SplittedPath() As String
    SplittedPath = Split(Trim$(FileName), "\")
    Dim i%
    For i = LBound(SplittedPath) To UBound(SplittedPath)
        If Trim$(SplittedPath(i)) <> "" And Replace(Trim$(SplittedPath(i)), ".", "") = "" Then
            Alert t("저장 경로가 유효하지 않습니다.", "Invalid save path."), App.Title, 16
            OnExit 999
            Exit Sub
        End If
    Next i
    
    If (Not FolderExists(Trim$(FileName))) And ((Not FolderExists(GetParentFolderName(Trim$(FileName)))) Or Right$(FileName, 1) = "\") Then
        Alert t("저장 경로가 존재하지 않습니다.", "Save path does not exist."), App.Title, 16
        OnExit 999
        Exit Sub
    End If
    
    If Replace(FileName, " ", "") = "" Then
        FileName = Replace(CurDir() & "\", "\\", "\")
    End If
    Dim ServerName$
    Dim AutoName As Boolean
    AutoName = False
    If FolderExists(FileName) Then
        If Not (Right$(FileName, 1) = "\") Then FileName = FileName & "\"
        ServerName = FilterFilename(ExcludeParameters((Split(URL, "/")(UBound(Split(URL, "/"))))))
        If Replace(ServerName, " ", "") = "" Then ServerName = "download_" & CStr(Rnd * 1E+15)
        FileName = FileName & ServerName
        AutoName = True
    End If
    If BatchStarted And (AutoName = False) Then
        AutoName = (lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(4).Text = "Y")
    End If
    If Right$(FileName, 1) = "." Then FileName = Left$(FileName, Len(FileName) - 1) & "_"
    DownloadPath = FileName
    PrevDownloadedBytes = 0
    SpeedCount = 0
    lblFilename.Caption = GetFilename(DownloadPath)
    If Len(lblFilename.Caption) > 22 Then lblFilename.Caption = Left$(lblFilename.Caption, 22) & "..."
    
    Dim ContinueDownload As Integer
    ContinueDownload = chkContinueDownload.Value
    If (Not BatchStarted) And chkContinueDownload.Value <> 1 Then
        Dim PrevPartialDownload As Boolean
        PrevPartialDownload = (trThreadCount.Value <= 1 And FileExists(FileName & ".part.tmp")) Or _
                              (trThreadCount.Value > 1 And FileExists(FileName & ".part_" & trThreadCount.Value & ".tmp") And (Not FileExists(FileName & ".part_" & (trThreadCount.Value + 1) & ".tmp")))
        If PrevPartialDownload Then
            Dim ContinueMsgboxResult As VbMsgBoxResult
            ContinueMsgboxResult = ConfirmCancel(t("기존에 다운로드 받다가 중지한 파일입니다. 다운로드받은 지점부터 이어서 받으시겠습니까?" & vbCrLf & "　[아니요]를 누를 경우 처음부터 다시 다운로드됩니다.", "This file was previously downloaded partially. Would you like to resume?" & vbCrLf & "  We will download from the start if you choose No."), App.Title)
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
    Dim CurrentHeaderCache$
    If BatchStarted Then
        CurrentHeaderCache = lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(5).Text
    Else
        CurrentHeaderCache = Functions.SessionHeaderCache
    End If
    Dim SPResult As SP_RESULTS
    SPResult = SP.Run("""" & _
        NodePath & """ """ & _
        ScriptPath & """ """ & _
        Replace(Replace(URL, " ", "%20"), """", "%22") & """ """ & _
        FileName & """ " & _
        trThreadCount.Value & " " & _
        GetSetting("DownloadBooster", "Options", "NoCleanup", 0) & " " & _
        cbWhenExist.ListIndex & " " & _
        ContinueDownload & " " & _
        GetSetting("DownloadBooster", "Options", "NoRedirectCheck", 0) & " " & _
        GetSetting("DownloadBooster", "Options", "ForceGet", 1) & " " & _
        GetSetting("DownloadBooster", "Options", "Ignore300", 0) & " " & _
        Abs(CInt(AutoName)) & " " & _
        GetSetting("DownloadBooster", "Options", "ThreadRequestInterval", 100) & " " & _
        Col(Functions.HeaderCache, "-") & " " & _
        Col(CurrentHeaderCache, "-"))
    Select Case SPResult
        Case SP_SUCCESS
            SP.ClosePipe
        Case SP_CREATEPIPEFAILED
            Alert t("다운로드 시작에 실패했습니다. 다운로더 프로세스로부터 정보를 받아올 수 없습니다. 디렉토리 설정에서 올바른 프로그램을 지정했는지 확인하십시오.", "Failed to receieve data from the downloader process. Check if the directory settings are valid."), App.Title, 16
            If Not BatchStarted Then
                cmdGo.Enabled = -1
            End If
            cmdStop.Enabled = 0
            cmdStop.Left = Me.Width + 1200
            cmdGo.Enabled = -1
            cmdGo.Visible = -1
            OnStop False
        Case SP_CREATEPROCFAILED
            Alert t("다운로드 시작에 실패했습니다. 다운로더 프로세스를 생성할 수 없습니다. 디렉토리 설정에서 올바른 프로그램을 지정했는지 확인하십시오.", "Failed to create the downloader process. Check if the directory settings are valid."), App.Title, 16
            If Not BatchStarted Then
                cmdGo.Enabled = -1
            End If
            cmdStop.Enabled = 0
            cmdStop.Left = Me.Width + 1200
            cmdGo.Enabled = -1
            cmdGo.Visible = -1
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

Private Sub cmdEdit_Click()
    mnuEdit_Click
End Sub

Private Sub cmdGo_Click()
    Dim SPResult As SP_RESULTS
    Dim TextLine As String
    
    If Replace(txtURL.Text, " ", "") = "" Then
        Alert t("파일 주소를 입력하십시오.", "Specify the file URL."), App.Title, 64
        Exit Sub
    End If
    
    If Left$(txtURL.Text, 7) <> "http://" And Left$(txtURL.Text, 8) <> "https://" Then
        Alert t("주소가 올바르지 않습니다. 'http://' 또는 'https://'로 시작해야 합니다.", "Invalid address. Must start with 'http://' or 'https://'."), App.Title, 16
        Exit Sub
    End If
    
    txtFileName.Text = Trim$(txtFileName.Text)
    Do While Replace(txtFileName.Text, "\\", "\") <> txtFileName.Text
        txtFileName.Text = Replace(txtFileName.Text, "\\", "\")
    Loop
    
    Dim SplittedPath() As String
    SplittedPath = Split(txtFileName.Text, "\")
    Dim i%
    For i = LBound(SplittedPath) To UBound(SplittedPath)
        If Trim$(SplittedPath(i)) <> "" And Replace(Trim$(SplittedPath(i)), ".", "") = "" Then
            Alert t("저장 경로가 유효하지 않습니다.", "Invalid save path."), App.Title, 16
            Exit Sub
        End If
    Next i
    
    If (Not FolderExists(Trim$(txtFileName.Text))) And ((Not FolderExists(GetParentFolderName(Trim$(txtFileName.Text)))) Or Right$(txtFileName.Text, 1) = "\") Then
        Alert t("저장 경로가 존재하지 않습니다.", "Save path does not exist."), App.Title, 16
        Exit Sub
    End If
    
    txtURL.Text = Trim$(txtURL.Text)

    Elapsed = 0
    If GetSetting("DownloadBooster", "Options", "LazyElapsed", "0") <> "1" Then timElapsed.Enabled = -1
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

Private Sub cmdOpen_DropDown()
    cmdOpenFileDropdown_Click
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

Private Sub cmdOpenFileDropdown_Click()
    Me.PopupMenu mnuOpenFileDropdown, , cmdOpen.Left, cmdOpen.Top + cmdOpen.Height
End Sub

Private Sub cmdOpenFileDropdown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdOpenFileDropdown_Click
End Sub

Private Sub cmdOpenFolder_Click()
    Dim pth$
    pth = DownloadPath
    If DownloadPath = "" Then pth = txtFileName.Text
    If FolderExists(pth) Then
        Shell "cmd /c start """" explorer.exe """ & pth & """"
    Else
        Shell "cmd /c start """" explorer.exe """ & GetParentFolderName(pth) & """"
    End If
End Sub

Private Sub cmdOptions_Click()
    Unload frmOptions
    frmOptions.Show vbModal, Me
End Sub

Private Sub cmdStartBatch_Click()
    If lvBatchFiles.ListItems.Count <= 0 Then
        cmdStartBatch.Enabled = 0
        Exit Sub
    End If
    
    BatchErrorCount = 0
    BatchErrorAllCount = 0
    CurrentBatchIdx = 1
    BatchStarted = True
    cmdStartBatch.Enabled = 0
    cmdStopBatch.Enabled = -1
    cmdStopBatch.Left = cmdStartBatch.Left
    cmdStopBatch.Refresh
    Elapsed = 0
    timElapsed.Enabled = -1
    chkOpenAfterComplete.Enabled = 0
    cmdOpen.Enabled = 0
    cmdOpenFileDropdown.Enabled = 0
    cmdGo.Enabled = 0
    StartDownload lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(2), lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(1)
End Sub

Private Sub cmdStop_Click()
    Dim IsMarquee As Boolean
    IsMarquee = pbTotalProgressMarquee.Visible
    Dim ConfirmResult As VbMsgBoxResult
    If IsMarquee Or ResumeUnsupported Then
        ConfirmResult = ConfirmEx(t("다운로드를 중지하시겠습니까? 현재 파일은 이어받기가 지원되지 않으므로 처음부터 다시 다운로드받아야 합니다.", "Cancel download? Resuming is not supported for this file."), t("다운로드 취소", "Cancel download"), 48)
    Else
        ConfirmResult = Confirm(t("다운로드를 중지하시겠습니까? 이어받기 기능을 통해 중단한 곳부터 계속 다운로드받을 수 있습니다.", "Cancel download? You can resume later."), t("다운로드 취소", "Cancel download"))
    End If
    If ConfirmResult = vbYes Then
        Dim CurrentProgress As Integer
        CurrentProgress = pbTotalProgress.Value
        
        OnStop False
        cmdOpen.Enabled = 0
        cmdOpenFileDropdown.Enabled = 0
        
        If IsMarquee Or (CurrentProgress > 0 And CurrentProgress < 100) Then
            Dim KillTemp As Boolean
            KillTemp = False
            If IsMarquee Or ResumeUnsupported Then
                KillTemp = True
            Else
                KillTemp = Confirm(t("나중에 계속 이어서 다운로드받을 수 있도록 다운로드한 데이타를 저장하시겠습니까?", "Would you like to keep the partially downloaded data to resume later?"), App.Title) <> vbYes
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
        ConfirmResult = ConfirmEx(t("다운로드를 중지하시겠습니까? 현재 파일은 이어받기가 지원되지 않으므로 처음부터 다시 다운로드받아야 합니다.", "Cancel download? Resuming is not supported for this file."), t("다운로드 취소", "Cancel download"), 48)
    Else
        ConfirmResult = Confirm(t("다운로드를 중지하시겠습니까? 이어받기 기능을 통해 중단한 곳부터 계속 다운로드받을 수 있습니다.", "Cancel download? You can resume later."), t("다운로드 취소", "Cancel download"))
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
        cmdStopBatch.Enabled = 0
        cmdStopBatch.Left = Me.Width + 1200
        OnStop False
        cmdGo.Enabled = 0
        timElapsed.Enabled = 0
        sbStatusBar.Panels(3).Text = ""
        sbStatusBar.Panels(4).Text = ""
        chkOpenAfterComplete.Enabled = -1
        cmdGo.Enabled = -1
        
        If IsMarquee Or (CurrentProgress > 0 And CurrentProgress < 100) Then
            Dim KillTemp As Boolean
            KillTemp = False
            If IsMarquee Or ResumeUnsupported Then
                KillTemp = True
            Else
                KillTemp = Confirm(t("나중에 계속 이어서 다운로드받을 수 있도록 다운로드한 데이타를 저장하시겠습니까?", "Would you like to keep the partially downloaded data to resume later?"), App.Title) <> vbYes
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
        
        If BatchErrorCount Then _
            Alert t("하나 이상의 오류가 발생했습니다. 해당 항목을 두 번 누르면 오류 정보를 볼 수 있습니다.", _
                    "One or more errors have occurred. Double click the error item to see details."), App.Title, 48
    End If
End Sub

Sub SetBackgroundPosition(Optional ByVal ForceRefresh As Boolean = False)
    On Error Resume Next
    Dim i%, j%, k%
    If imgBackground.Visible Then
        Dim ImageCentered As Boolean
        Dim ImgPos As Integer
        ImgPos = ImagePosition
        ImageCentered = (GetSetting("DownloadBooster", "Options", "BackgroundImageCentered", 0) <> 0)
        Dim Width&, Height&
        Width = GetPictureWidth(imgBackground.Picture)
        Height = GetPictureHeight(imgBackground.Picture)
        Select Case ImgPos
            Case 0 '늘이기
                If imgBackground.Stretch <> True Then imgBackground.Stretch = True
                imgBackground.Move 0, 0, Me.Width, Me.Height
            Case 1 '높이에 맞추기
                If imgBackground.Stretch <> True Then imgBackground.Stretch = True
                imgBackground.Move IIf(ImageCentered, (Me.Width - imgBackground.Width) \ 2, 0), 0, Width / Height * Me.Height, Me.Height
            Case 2 '너비에 맞추기
                If imgBackground.Stretch <> True Then imgBackground.Stretch = True
                imgBackground.Move 0, IIf(ImageCentered, ((Me.Height - sbStatusBar.Height - CaptionHeight * 15 - 15) - imgBackground.Height) \ 2, 0), Me.Width, Height / Width * Me.Width
            Case 3 '원본 크기
                If imgBackground.Stretch = True Then imgBackground.Stretch = False
                imgBackground.Move IIf(ImageCentered, (Me.Width - imgBackground.Width) \ 2, 0), IIf(ImageCentered, ((Me.Height - sbStatusBar.Height - CaptionHeight * 15 - 15) - imgBackground.Height) \ 2, 0), Width, Height
            Case 4 '바둑판식
                If imgBackground.Stretch = True Then imgBackground.Stretch = False
                imgBackground.Move -Width, -Height, Width, Height
                k = 1
                If Width > 0 And Height > 0 Then
                    For i = 1 To Ceil(Me.Height / Height)
                        For j = 1 To Ceil(Me.Width / Width)
                            If k > imgBackgroundTile.UBound Then _
                                Load imgBackgroundTile(k)
                            If Not (imgBackgroundTile(k).Picture Is imgBackground.Picture) Then
                                Set imgBackgroundTile(k).Picture = imgBackground.Picture
                                imgBackgroundTile(k).Stretch = False
                                imgBackgroundTile(k).Width = Width
                                imgBackgroundTile(k).Height = Height
                                imgBackgroundTile(k).Top = (i - 1) * Height
                                imgBackgroundTile(k).Left = (j - 1) * Width
                                imgBackgroundTile(k).Visible = True
                            End If
                            k = k + 1
                        Next j
                    Next i
                    k = k - 1
                    If k > MaxLoadedTileBackgroundImage Then
                        MaxLoadedTileBackgroundImage = k
                    ElseIf k < MaxLoadedTileBackgroundImage Then
                        For i = (k + 1) To MaxLoadedTileBackgroundImage
                            Set imgBackgroundTile(i).Picture = Nothing
                            Unload imgBackgroundTile(i)
                            Set imgBackgroundTile(i) = Nothing
                        Next i
                        MaxLoadedTileBackgroundImage = k
                    End If
                ElseIf MaxLoadedTileBackgroundImage > 0 Then
                    For i = 1 To MaxLoadedTileBackgroundImage
                        Set imgBackgroundTile(i).Picture = Nothing
                        Unload imgBackgroundTile(i)
                        Set imgBackgroundTile(i) = Nothing
                    Next i
                    MaxLoadedTileBackgroundImage = 0
                End If
        End Select
        If ImgPos <> 4 And MaxLoadedTileBackgroundImage > 0 Then
            For i = 1 To MaxLoadedTileBackgroundImage
                Set imgBackgroundTile(i).Picture = Nothing
                Unload imgBackgroundTile(i)
                Set imgBackgroundTile(i) = Nothing
            Next i
        End If
        If ImagePosition < 2 Or (ImageCentered And ImagePosition = 1) Or ForceRefresh Or ImageCentered Then
dorefresh:
            On Error Resume Next
            fOptions.Refresh
            chkOpenAfterComplete.Refresh
            chkOpenFolder.Refresh
            chkContinueDownload.Refresh
            chkAutoRetry.Refresh
            If Not imgLeft.Visible Then
                Frame4.Refresh
                fTotal.Refresh
            End If
            Dim VSEnabled As Boolean
            VSEnabled = trThreadCount.VisualStyles
            trThreadCount.VisualStyles = Not VSEnabled
            trThreadCount.VisualStyles = VSEnabled
            optTabDownload2.Refresh
            optTabThreads2.Refresh
        End If
    Else
        GoTo dorefresh
    End If
End Sub

Sub SetBackgroundImage()
    On Error Resume Next
    Dim i%
    If GetSetting("DownloadBooster", "Options", "UseBackgroundImage", 0) = 1 And Trim$(GetSetting("DownloadBooster", "Options", "BackgroundImagePath", "")) <> "" Then
        If LCase(Right$(GetSetting("DownloadBooster", "Options", "BackgroundImagePath", ""), 4)) = ".png" Then
            Set imgBackground.Picture = LoadPngIntoPictureWithAlpha(GetSetting("DownloadBooster", "Options", "BackgroundImagePath", ""))
        Else
            imgBackground.Picture = LoadPicture(GetSetting("DownloadBooster", "Options", "BackgroundImagePath", ""))
        End If
        imgBackground.Visible = -1
        SetBackgroundPosition True
    Else
        imgBackground.Visible = 0
        If MaxLoadedTileBackgroundImage > 0 Then
            For i = 1 To MaxLoadedTileBackgroundImage
                Set imgBackgroundTile(i).Picture = Nothing
                Unload imgBackgroundTile(i)
                Set imgBackgroundTile(i) = Nothing
            Next i
        End If
        SetBackgroundPosition True
    End If
    
    SetPictureBuddy
End Sub

Sub SetPattern()
    pgPattern.FillStyle = CInt(GetSetting("DownloadBooster", "Options", "FormFillStyle", 0)) + 1
    pgPattern.FillColor = CLng(GetSetting("DownloadBooster", "Options", "FormFillColor", 0))
End Sub

Sub LoadLiveBadukSkin()
    On Error Resume Next
    Dim i%
    Dim LBEnabled As Boolean
    LBEnabled = (CInt(GetSetting("DownloadBooster", "Options", "EnableLiveBadukMemoSkin", 0)) <> 0 And DPI = 96)
    If LBEnabled Then
        Dim SolidFrameVisible As Boolean
        SolidFrameVisible = False
        
        LoadPNG
        imgLeft.Visible = -1
        imgCenter.Visible = -1
        imgTopRight.Left = 6495 - imgTopRight.Width - 30 + 120
        imgBottomRight.Left = imgTopRight.Left
        imgBottomLeft.Top = 4620 - imgBottomLeft.Height - 30 + imgTop.Top + 240
        imgBottomRight.Top = imgBottomLeft.Top
        imgBottomRight.Left = imgTopRight.Left

        imgRight.Left = imgBottomRight.Left - 15
        imgBottom.Top = imgBottomLeft.Top
        imgLeft.Height = 4620 - imgBottomLeft.Height - imgTopLeft.Height - 30 + 240
        imgCenter.Width = imgRight.Left - (imgLeft.Left + imgLeft.Width)
        imgCenter.Height = imgBottom.Top - (imgTop.Top + imgTop.Height)
        
        Dim FrameType$
        FrameType = LCase(GetSetting("DownloadBooster", "Options", "LiveBadukMemoSkinFrameType", "transparent"))
        Select Case FrameType
            Case "solidcolor", "texture"
                imgTopLeft.Visible = 0
                imgTopRight.Visible = 0
                imgTop.Visible = 0
                imgRight.Visible = 0
                imgBottom.Visible = 0
                imgBottomLeft.Visible = 0
                imgBottomRight.Visible = 0
                fTygemFrameTransparent.Visible = -1
                SolidFrameVisible = True
                If imgLeft.Picture.Width = 0 Then imgLeft.Picture = LoadPngIntoPictureWithAlpha(CachePath & "left.png")
                If imgCenter.Picture.Width = 0 Then imgCenter.Picture = LoadPngIntoPictureWithAlpha(CachePath & "center.png")
                imgLeft.Height = imgLeft.Height + 30
                imgLeft.Top = 1620
                imgCenter.Top = 1620
                imgCenter.Height = imgCenter.Height + 30
                imgCenter.Width = imgCenter.Width + 45
                fTygemFrameTransparent.BackColor = CLng(GetSetting("DownloadBooster", "Options", "LiveBadukMemoSkinFrameColor", 16777215))
                lnTygemFrameRight.Visible = False
                lnTygemFrameBottom.Visible = False
            Case Else
                If imgTopLeft.Picture.Width = 0 Then imgTopLeft.Picture = LoadPngIntoPictureWithAlpha(CachePath & "topleft.png")
                If imgTopRight.Picture.Width = 0 Then imgTopRight.Picture = LoadPngIntoPictureWithAlpha(CachePath & "topright.png")
                If imgTop.Picture.Width = 0 Then imgTop.Picture = LoadPngIntoPictureWithAlpha(CachePath & "top.png")
                If imgLeft.Picture.Width = 0 Then imgLeft.Picture = LoadPngIntoPictureWithAlpha(CachePath & "left.png")
                If imgRight.Picture.Width = 0 Then imgRight.Picture = LoadPngIntoPictureWithAlpha(CachePath & "right.png")
                If imgBottom.Picture.Width = 0 Then imgBottom.Picture = LoadPngIntoPictureWithAlpha(CachePath & "bottom.png")
                If imgBottomLeft.Picture.Width = 0 Then imgBottomLeft.Picture = LoadPngIntoPictureWithAlpha(CachePath & "bottomleft.png")
                If imgBottomRight.Picture.Width = 0 Then imgBottomRight.Picture = LoadPngIntoPictureWithAlpha(CachePath & "bottomright.png")
                If imgCenter.Picture.Width = 0 Then imgCenter.Picture = LoadPngIntoPictureWithAlpha(CachePath & "center.png")
                
                imgLeft.Top = 1635
                imgCenter.Top = 1635
        
                imgTop.Width = 6495 - imgTopLeft.Width - imgTopRight.Width - 30
                imgBottom.Width = imgTop.Width
                imgRight.Height = imgLeft.Height
        
                imgTopLeft.Visible = -1
                imgTopRight.Visible = -1
                imgTop.Visible = -1
                imgRight.Visible = -1
                imgBottom.Visible = -1
                imgBottomLeft.Visible = -1
                imgBottomRight.Visible = -1
                fTygemFrameTransparent.Visible = 0
                
                lnTygemFrameRight.Visible = True
                lnTygemFrameBottom.Visible = True
        End Select
        
        If FrameType = "texture" Then
            Dim TexturePath$
            TexturePath = GetSetting("DownloadBooster", "Options", "LiveBadukMemoSkinFrameTexture", "")
            If FileExists(TexturePath) Then
                If LCase(Right$(TexturePath, 4)) = ".png" Then
                    Set imgFrameTexture.Picture = LoadPngIntoPictureWithAlpha(TexturePath)
                Else
                    imgFrameTexture.Picture = LoadPicture(TexturePath)
                End If
            End If
            imgFrameTexture.Visible = True
        Else
            imgFrameTexture.Visible = False
        End If

        fTotal.Visible = 0
        Frame4.Visible = 0
        lblLBCaption.Visible = -1
        fDownloadInfo.Refresh
        pbProgressOuterContainer.Refresh
        pbProgressContainer.Refresh

        pbTotalProgress.Top = 1800 - 90
        pbTotalProgressMarquee.Top = 1800 - 90
        
        optTabDownload2.Width = 840
        optTabThreads2.Width = 1485
        optTabDownload2.Caption = fTabDownload.Caption
        optTabThreads2.Caption = fTabThreads.Caption
        
        pbTotalProgressMarquee.Left = 360
        pbTotalProgressMarquee.Width = 6015
        pbTotalProgress.Left = 360
        pbTotalProgress.Width = 6015
        lblState.Visible = False
        
        LBFrameEnabled = True
        fOptions.BorderStyle = 0
        cmdOptions.Left = 7200
        cmdAbout.Left = 7200
        Label11.Visible = True
        lbOptionsHeader.Visible = True
        
        lblURL.Font.Bold = True
        lblFilePath.Font.Bold = True
        lblThreadCountLabel.Font.Bold = True
        
        If GetSetting("DownloadBooster", "Options", "LiveBadukMemoSkinEnableShadow", 1) <> 0 Then
            Dim ShadowColor As Long
            ShadowColor = CLng(GetSetting("DownloadBooster", "Options", "LiveBadukMemoSkinShadowColor", 16777215))
            If ShadowColor < 0& Then ShadowColor = 16777215
            lblLBCaptionShadow.ForeColor = ShadowColor
            lblLBCaptionShadow2.ForeColor = ShadowColor
            lblURLShadow.ForeColor = ShadowColor
            lblFilePathShadow.ForeColor = ShadowColor
            lblThreadCountLabelShadow.ForeColor = ShadowColor
            lblLBCaptionShadow.Visible = True
            lblLBCaptionShadow2.Visible = True
            lblURLShadow.Visible = True
            lblFilePathShadow.Visible = True
            lblThreadCountLabelShadow.Visible = True
        Else
            lblLBCaptionShadow.Visible = False
            lblLBCaptionShadow2.Visible = False
            lblURLShadow.Visible = False
            lblFilePathShadow.Visible = False
            lblThreadCountLabelShadow.Visible = False
        End If
        
        cbWhenExist.Width = 1305
        
        If GetSetting("DownloadBooster", "Options", "LiveBadukMemoSkinEnableBorder", 1) <> 0 Then
            imgBorderTopLeft.Visible = True
            imgBorderTopRight.Visible = True
            imgBorderBottomLeft.Visible = True
            imgBorderBottomRight.Visible = True
            pgBorderLeft.Visible = True
            pgBorderTop.Visible = True
            pgBorderRight.Visible = True
            pgBorderBottom.Visible = True
        Else
            imgBorderTopLeft.Visible = False
            imgBorderTopRight.Visible = False
            imgBorderBottomLeft.Visible = False
            imgBorderBottomRight.Visible = False
            pgBorderLeft.Visible = False
            pgBorderTop.Visible = False
            pgBorderRight.Visible = False
            pgBorderBottom.Visible = False
        End If
        
        Dim FrameBackgroundType$
        FrameBackgroundType = LCase(GetSetting("DownloadBooster", "Options", "LiveBadukMemoSkinFrameBackgroundType", "transparent"))
        Select Case FrameBackgroundType
            Case "texture", "solidcolor"
                imgLBContentBackground.Visible = True
                If FrameBackgroundType = "texture" Then
                    Dim FrameBackgroundPath$
                    FrameBackgroundPath = GetSetting("DownloadBooster", "Options", "LiveBadukMemoSkinFrameBackground", "")
                    If FileExists(FrameBackgroundPath) Then
                        If LCase(Right$(FrameBackgroundPath, 4)) = ".png" Then
                            Set imgLBContentBackground.Picture = LoadPngIntoPictureWithAlpha(FrameBackgroundPath)
                        Else
                            imgLBContentBackground.Picture = LoadPicture(FrameBackgroundPath)
                        End If
                    Else
                        GoTo framecolorbackground
                    End If
                Else
framecolorbackground:
                    Set imgLBContentBackground.Picture = GenerateSolidColor(CLng(GetSetting("DownloadBooster", "Options", "LiveBadukMemoSkinFrameBackgroundColor", 16777215)))
                End If
                imgLBContentBackground.Top = imgTopLeft.Top + imgTopLeft.Height + IIf(fTygemFrameTransparent.Visible, 0, 15)
                imgLBContentBackground.Left = imgLeft.Left + 150
                imgLBContentBackground.Width = 6195 - IIf(SolidFrameVisible, 0, 45)
                imgLBContentBackground.Height = 4215 - IIf(SolidFrameVisible, 0, 30)
            Case Else
                imgLBContentBackground.Visible = False
        End Select
    Else
        fTygemFrameTransparent.Visible = 0
        imgTopLeft.Visible = 0
        imgTopRight.Visible = 0
        imgTop.Visible = 0
        imgLeft.Visible = 0
        imgRight.Visible = 0
        imgBottom.Visible = 0
        imgBottomLeft.Visible = 0
        imgBottomRight.Visible = 0
        imgCenter.Visible = 0

        lblLBCaption.Visible = 0
        fTotal.Visible = -1
        fTotal.Refresh
        Frame4.Visible = -1
        Frame4.Refresh
        fDownloadInfo.Refresh
        pbProgressContainer.Refresh

        pbTotalProgress.Top = 1560
        pbTotalProgressMarquee.Top = 1560
        
        optTabDownload2.Width = 195
        optTabThreads2.Width = 195
        optTabDownload2.Caption = ""
        optTabThreads2.Caption = ""
        
        lnTygemFrameRight.Visible = False
        lnTygemFrameBottom.Visible = False
        
        pbTotalProgressMarquee.Left = 1200
        pbTotalProgressMarquee.Width = 5175
        pbTotalProgress.Left = 1200
        pbTotalProgress.Width = 5175
        lblState.Visible = True
        
        LBFrameEnabled = False
        fOptions.BorderStyle = 1
        cmdOptions.Left = 7080
        cmdAbout.Left = 7080
        Label11.Visible = False
        lbOptionsHeader.Visible = False
        
        lblLBCaptionShadow.Visible = 0
        lblURLShadow.Visible = 0
        lblFilePathShadow.Visible = 0
        lblThreadCountLabelShadow.Visible = 0
        lblURL.Font.Bold = False
        lblFilePath.Font.Bold = False
        lblThreadCountLabel.Font.Bold = False
        
        cbWhenExist.Width = 1185
        
        imgBorderTopLeft.Visible = False
        imgBorderTopRight.Visible = False
        imgBorderBottomLeft.Visible = False
        imgBorderBottomRight.Visible = False
        pgBorderLeft.Visible = False
        pgBorderTop.Visible = False
        pgBorderRight.Visible = False
        pgBorderBottom.Visible = False
        
        imgLBContentBackground.Visible = False
    End If
    
    SetFormBackgroundColor Me
    If LBEnabled Then
        Dim ContentTextColor As Long
        ContentTextColor = CLng(GetSetting("DownloadBooster", "Options", "LiveBadukMemoSkinContentTextColor", 0))
        For i = 1 To MAX_THREAD_COUNT
            lblDownloader(i).ForeColor = ContentTextColor
            lblPercentage(i).ForeColor = ContentTextColor
        Next i
        optTabDownload2.ForeColor = ContentTextColor
        optTabThreads2.ForeColor = ContentTextColor
        Label8.ForeColor = ContentTextColor
        Label2.ForeColor = ContentTextColor
        Label3.ForeColor = ContentTextColor
        Label4.ForeColor = ContentTextColor
        Label5.ForeColor = ContentTextColor
        Label6.ForeColor = ContentTextColor
        Label7.ForeColor = ContentTextColor
        Label10.ForeColor = ContentTextColor
        Label9.ForeColor = ContentTextColor
        lblFilename.ForeColor = ContentTextColor
        lblTotalBytes.ForeColor = ContentTextColor
        lblDownloadedBytes.ForeColor = ContentTextColor
        lblElapsed.ForeColor = ContentTextColor
        lblSpeed.ForeColor = ContentTextColor
        lblThreadCount2.ForeColor = ContentTextColor
        lblTotalSizeThread.ForeColor = ContentTextColor
        lblRemaining.ForeColor = ContentTextColor
        lblMergeStatus.ForeColor = ContentTextColor
        
        If GetSetting("DownloadBooster", "Options", "LiveBadukMemoSkinEnableTextColor", 0) <> 0 Then
            Dim LBTextColor As Long
            LBTextColor = CLng(GetSetting("DownloadBooster", "Options", "LiveBadukMemoSkinTextColor", 0))
            lblURL.ForeColor = LBTextColor
            lblFilePath.ForeColor = LBTextColor
            lblThreadCountLabel.ForeColor = LBTextColor
            lblLBCaption.ForeColor = LBTextColor
            lblLBCaption2.ForeColor = LBTextColor
        End If
    End If

    Dim ctrl As Control
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "FrameW" Or TypeName(ctrl) = "CheckBoxW" Or TypeName(ctrl) = "OptionButtonW" Or TypeName(ctrl) = "CommandButtonW" Or TypeName(ctrl) = "Slider" Then
            ctrl.Refresh
        End If
    Next ctrl
    
    SetFont Me
    
    SetPictureBuddy
    ThreadBuddyAttached = False
End Sub

#If HIDEYTDL Then
#Else
Private Sub cmdYtdlTest_Click()
    StartYtdlDownload
End Sub
#End If

Function GenerateSolidColor(ByVal Color As Long) As IPictureDisp
    pbDummy.Cls
    pbDummy.DrawWidth = 5
    pbDummy.Width = 15
    pbDummy.Height = 15
    pbDummy.Line (0, 0)-(15, 15), Color
    pbDummy.Refresh
    Set GenerateSolidColor = pbDummy.Image
    pbDummy.Cls
End Function

Sub SetTitle(Optional ByVal Title As String = "")
    If Title = "" Then
        Me.Caption = FormCaption
    Else
        Me.Caption = Title & " - " & FormCaption
    End If
End Sub

Private Sub Form_Load()
    On Error Resume Next
    
    SetupVisualStylesFixes Me
    
    Set ErrorCodeDescription = New Collection
    ErrorCodeDescription.Add t("서버와의 접속이 끊겼습니다. 다운로드 중 네트워크 오류가 발생했거나 주소가 유효하지 않거나 서버가 응답하지 않습니다.", "Network error"), "1"
    ErrorCodeDescription.Add t("주소나 파일 이름을 지정하지 않았습니다.", "Address or file name unspecified"), "102"
    ErrorCodeDescription.Add t("저장 경로가 존재하지 않습니다.", "Save path doesn't exist"), "103"
    ErrorCodeDescription.Add t("저장할 파일명이 사용 중입니다. 다른 이름을 선택하십시오.", "File name already exists"), "104"
    ErrorCodeDescription.Add t("파일 서버가 다운로드 부스트를 지원하지 않습니다. 강도를 1로 변경해 보십시오.", "Download boosting not supported. Try changing the thread count to 1."), "106"
    ErrorCodeDescription.Add t("파일의 크기를 알 수 없어서 다운로드를 부스트할 수 없습니다. 강도를 1로 변경해 보십시오.", "Unable to boost download because the file size is not provided. Try changing the thread count to 1."), "107"
    ErrorCodeDescription.Add t("서버가 요청을 거부했습니다. 서버 측 오류이거나 페이지가 존재하지 않거나 접근 권한이 없을 수 있습니다.", "Server has denied your request. The file may not exist or have insufficient permissions to access it."), "108"
    
    MAX_THREAD_COUNT = CInt(GetSetting("DownloadBooster", "Options", "MaxThreadCount", 25))

    ResumeUnsupported = False
    LBFrameEnabled = False
    ThreadBuddyAttached = False
    sbStatusBar.Panels(1).Text = t("준비", "Ready")
    FormCaption = App.Title & IIf(InIDE, "*", "") & " " & App.Major & "." & App.Minor & IIf(App.Revision > 0, "." & App.Revision, "")
    SetTitle
    ScrollOneScreen = GetSetting("DownloadBooster", "Options", "ScrollOneScreen", 0) <> 0
    vsProgressScroll.LargeChange = IIf(ScrollOneScreen, 1, 10)
    
    MaxLoadedTileBackgroundImage = 0
    ImagePosition = GetSetting("DownloadBooster", "Options", "ImagePosition", 1)
    
    Dim Lft%
    Dim Top%
    Top = GetSetting("DownloadBooster", "UserData", "FormTop", -1)
    Lft = GetSetting("DownloadBooster", "UserData", "FormLeft", -1)
    If Top >= 0 And Lft >= 0 Then
        Me.Top = Top
        Me.Left = Lft
    End If
    
    Dim i%
    For i = 1 To MAX_THREAD_COUNT
        If i > 1 Then
            Load lblDownloader(i)
            Load lblPercentage(i)
            Load pbProgress(i)
            Load pbProgressMarquee(i)
        End If
        
        lblDownloader(i).Top = 360# * CDbl(i - 1) + 45#
        lblPercentage(i).Top = 360# * CDbl(i - 1) + 45#
        pbProgress(i).Top = 360# * CDbl(i - 1)
        pbProgress(i).ZOrder 1
        pbProgressMarquee(i).Top = 360# * CDbl(i - 1)
        pbProgressMarquee(i).ZOrder 0
        If MAX_THREAD_COUNT >= 100 Then
            pbProgress(i).Width = pbProgress(i).Width - 60
            pbProgress(i).Left = pbProgress(i).Left + 60
            pbProgressMarquee(i).Width = pbProgressMarquee(i).Width - 60
            pbProgressMarquee(i).Left = pbProgressMarquee(i).Left + 60
        End If
        lblDownloader(i).Caption = t("스레드", "Thread") & " " & i & ":"
    Next i
    If MAX_THREAD_COUNT >= 100 Then
        pbProgress(1).Width = pbProgress(1).Width - 60
        pbProgress(1).Left = pbProgress(1).Left + 60
        pbProgressMarquee(1).Width = pbProgressMarquee(1).Width - 60
        pbProgressMarquee(1).Left = pbProgressMarquee(1).Left + 60
        If MAX_THREAD_COUNT >= 250 Then
            trThreadCount.TickFrequency = 16
        Else
            trThreadCount.TickFrequency = 8
        End If
    ElseIf MAX_THREAD_COUNT >= 50 Then
        trThreadCount.TickFrequency = 4
    End If
    trThreadCount.Max = MAX_THREAD_COUNT
    If MAX_THREAD_COUNT <= 14 Then
        trThreadCount.TickFrequency = 1
    End If
    pbProgressContainer.Height = 360# * CDbl(MAX_THREAD_COUNT)
    fDownloadInfo.Top = fThreadInfo.Top '+ 60
    fDownloadInfo.Left = fThreadInfo.Left
    fDownloadInfo.Width = fThreadInfo.Width '5925
    fDownloadInfo.Height = fThreadInfo.Height '- 60
    
    LoadLiveBadukSkin
    
    Me.Width = MAIN_FORM_WIDTH + PaddedBorderWidth * 15 * 2 * (DPI / 96)
    cmdStop.Left = Me.Width + 1200
    
    cmdStopBatch.Left = Me.Width + 1200
    
    If GetSetting("DownloadBooster", "UserData", "LastTab", 1) = 1 Then
        fTabDownload_Click
    Else
        fTabThreads_Click
    End If
    
    trThreadCount.Value = GetSetting("DownloadBooster", "UserData", "ThreadCount", GetSetting("DownloadBooster", "Options", "ThreadCount", 1))
    trThreadCount_Scroll
    
    lvBatchFiles.ColumnHeaders.Add , "filename", t("파일 이름", "File Name"), 2895
    lvBatchFiles.ColumnHeaders.Add , "fullpath", t("전체 경로", "Full Path"), 0
    lvBatchFiles.ColumnHeaders.Add , "url", t("파일 주소", "File URL"), 4495
    lvBatchFiles.ColumnHeaders.Add , "status", t("상태", "Status"), 1105, LvwColumnHeaderAlignmentCenter
    lvBatchFiles.ColumnHeaders.Add , "autoname", t("파일 이름 자동 감지", "Autodetect File Name"), 0
    lvBatchFiles.ColumnHeaders.Add , "headers", t("인코딩된 헤더", "Encoded Headers"), 0
#If HIDEYTDL Then
    GoTo afterheaderadd
#End If
    lvBatchFiles.ColumnHeaders.Add , "useytdl", "youtube-dl " & t("사용", "used"), 0
    lvBatchFiles.ColumnHeaders.Add , "ytdlformat", "youtube-dl: " & t("포맷", "format"), 0
    lvBatchFiles.ColumnHeaders.Add , "ytdletractaudio", "youtube-dl: " & t("오디오 추출", "extract audio"), 0
    lvBatchFiles.ColumnHeaders.Add , "ytdlaudioformat", "youtube-dl: " & t("오디오 포맷", "audio format"), 0
    lvBatchFiles.ColumnHeaders.Add , "ytdlaudioqualitytype", "youtube-dl: " & t("오디오 음질 형식", "audio quality type"), 0
    lvBatchFiles.ColumnHeaders.Add , "ytdlcbr", "youtube-dl: CBR", 0
    lvBatchFiles.ColumnHeaders.Add , "ytdlvbr", "youtube-dl: VBR", 0
afterheaderadd:

    Me.Height = 6930
    
    BatchStarted = False
    
    txtFileName.Text = GetSetting("DownloadBooster", "UserData", "SavePath", CurDir())
    
    Me.Height = 6930 + PaddedBorderWidth * 15 * 2
    
    Dim hSysMenu As Long
    Dim MenuCount As Long
    hSysMenu = GetSystemMenu(Me.hWnd, 0)
    DeleteMenu hSysMenu, 0, MF_BYCOMMAND
    MenuCount = GetMenuItemCount(hSysMenu)
    Dim MII As MENUITEMINFO
    
    MII.cbSize = Len(MII)
    
    If GetSetting("DownloadBooster", "Options", "AlwaysOnTop", 0) = 1 Then
        MainFormOnTop = True
        SetWindowPos hWnd, hWnd_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    Else
        MainFormOnTop = False
    End If
    
    '항상 위에 표시
    With MII
        .fMask = MIIM_STATE Or MIIM_ID Or MIIM_TYPE
        .fType = MFT_STRING
        .fState = MFS_ENABLED
        .wID = 1000
        .dwTypeData = t("언제나 위(&A)", "&Always On Top")
        .cch = Len(.dwTypeData)
    End With
    InsertMenuItem hSysMenu, 0, 1, MII

    '높이 리셋
    With MII
        .fMask = MIIM_STATE Or MIIM_ID Or MIIM_TYPE
        .fType = MFT_STRING
        .fState = MFS_ENABLED
        .wID = 1003
        .dwTypeData = t("창 크기 초기화(&E)", "R&eset window size")
        .cch = Len(.dwTypeData)
    End With
    InsertMenuItem hSysMenu, 1, 1, MII

    '구분선
    With MII
        .cbSize = Len(MII)
        .fMask = MIIM_ID Or MIIM_TYPE
        .fType = MFT_SEPARATOR
        .wID = 2000
    End With
    InsertMenuItem hSysMenu, 2, 1, MII

    '일괄처리목록감추기
    With MII
        .fMask = MIIM_STATE Or MIIM_ID Or MIIM_TYPE
        .fType = MFT_STRING
        .fState = MFS_ENABLED
        .wID = 1001
        .dwTypeData = t("간단히 보기(&I)", "S&imple Mode")
        .cch = Len(.dwTypeData)
    End With
    InsertMenuItem hSysMenu, 3, 1, MII

    '일괄처리목록표시
    With MII
        .fMask = MIIM_STATE Or MIIM_ID Or MIIM_TYPE
        .fType = MFT_STRING
        .fState = MFS_ENABLED
        .wID = 1002
        .dwTypeData = t("일괄 처리 보기(&B)", "&Batch Mode")
        .cch = Len(.dwTypeData)
    End With
    InsertMenuItem hSysMenu, 4, 1, MII

    '구분선
    With MII
        .cbSize = Len(MII)
        .fMask = MIIM_ID Or MIIM_TYPE
        .fType = MFT_SEPARATOR
        .wID = 2001
    End With
    InsertMenuItem hSysMenu, 5, 1, MII
    
    If GetSetting("DownloadBooster", "UserData", "BatchExpanded", 1) <> 0 Then
        cmdBatch_Click
    Else
        CheckMenuRadioItem hSysMenu, 1001, 1002, 1001, MF_BYCOMMAND
        FormWidth = (MAIN_FORM_WIDTH + PaddedBorderWidth * 15 * 2) / 15
        FormMinHeight = (6930 + PaddedBorderWidth * 15 * 2) / 15
        FormMaxHeight = (6930 + PaddedBorderWidth * 15 * 2) / 15
        With MII
            .cbSize = Len(MII)
            .fMask = MIIM_STATE
            .fState = MFS_GRAYED
        End With
        SetMenuItemInfo hSysMenu, 1003, MF_BYCOMMAND, MII
    End If
    
    chkOpenAfterComplete.Value = GetSetting("DownloadBooster", "Options", "OpenWhenComplete", 0)
    chkOpenFolder.Value = GetSetting("DownloadBooster", "Options", "OpenFolderWhenComplete", 0)
    If GetSetting("DownloadBooster", "Options", "RememberURL", 0) <> 0 Then
        txtURL.Text = GetSetting("DownloadBooster", "UserData", "FileURL", "")
        txtURL.SelStart = 0
        txtURL.SelLength = Len(txtURL.Text)
    End If
    chkContinueDownload.Value = GetSetting("DownloadBooster", "Options", "ContinueDownload", 0)
    chkAutoRetry.Value = GetSetting("DownloadBooster", "Options", "AutoRetry", 0)
    
    cbWhenExist.Clear
    cbWhenExist.AddItem t("건너뛰기", "Skip")
    cbWhenExist.AddItem t("덮어쓰기", "Overwrite")
    cbWhenExist.AddItem t("이름 변경", "Rename")
    cbWhenExist.ListIndex = GetSetting("DownloadBooster", "Options", "WhenFileExists", 0)
    
    SetupSplitButtons

    '언어설정
    lblURL.Caption = t(lblURL.Caption, "File &address:")
    lblFilePath.Caption = t(lblFilePath.Caption, "Save &file to:")
    lblThreadCountLabel.Caption = t(lblThreadCountLabel.Caption, "&Threads:")
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
    cmdAdd.Caption = t(cmdAdd.Caption, "Add U&RL...")
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
    lblLBCaption.Caption = t(lblLBCaption.Caption, "Progress")
    
    mnuEdit.Caption = t(mnuEdit.Caption, "&Edit...")
    mnuMoveUp.Caption = t(mnuMoveUp.Caption, "Move &up")
    mnuMoveDown.Caption = t(mnuMoveDown.Caption, "Move &down")
    mnuAddItem2.Caption = t(mnuAddItem2.Caption, "&Add URL...")
    mnuOpenBatch.Caption = t(mnuOpenBatch.Caption, "&Open")
    mnuOpenFolder2.Caption = t(mnuOpenFolder2.Caption, "Open &folder")
    
    cmdEdit.Caption = t(cmdEdit.Caption, "Edit(&N)...")
    
    Label9.Caption = t(Label9.Caption, "Merge status:")
    
    mnuProperties.Caption = t(mnuProperties.Caption, "View p&roperties")
    mnuPropertiesBatch.Caption = t(mnuPropertiesBatch.Caption, "View p&roperties")
    
    cmdDownloadOptions.Caption = t(cmdDownloadOptions.Caption, "Download &settings...")
    
    tr mnuHeaders, "&Headers..."
    
    Label11.Caption = fOptions.Caption
    
    tr mnuErrorInfo, "Error &information..."
    
    lblURLShadow.Caption = lblURL.Caption
    lblFilePathShadow.Caption = lblFilePath.Caption
    lblThreadCountLabelShadow.Caption = lblThreadCountLabel.Caption
    lblLBCaptionShadow.Caption = lblLBCaption.Caption
    optTabDownload2.Caption = fTabDownload.Caption
    optTabThreads2.Caption = fTabThreads.Caption
    lblLBCaptionShadow2.Caption = lblLBCaption.Caption
    lblLBCaption2.Caption = lblLBCaption.Caption
    '언어설정끝
    lbOptionsHeader.X1 = Label11.Width + 60
    
    If GetSetting("DownloadBooster", "Options", "DisableDWMWindow", DefaultDisableDWMWindow) = 1 Then DisableDWMWindow Me.hWnd
    
    'SetFormBackgroundColor Me
    SetPattern
    SetBackgroundImage
    SetBackgroundPosition

#If HIDEYTDL Then
    mnuYtdlOptions.Visible = False
#End If
    
    SetTextColors
    
    SetFont Me
    
    'rgnset
    Dim RC As RECT
    GetWindowRect fTygemFrameTransparent.hWnd, RC
    Dim Rgn&, Rgn1&, Rgn2&, Rgn3&, Rgn4&, Rgn5&, Rgn6&, Rgn7&, Rgn8&, Rgn9&
    Rgn = CreateRectRgn(0, 0, RC.Right - RC.Left, RC.Bottom - RC.Top)
    Rgn1 = CreateRectRgn(0, 0, 6, 1)
    Rgn2 = CreateRectRgn(0, 1, 4, 2)
    Rgn3 = CreateRectRgn(0, 2, 3, 3)
    Rgn4 = CreateRectRgn(0, 3, 2, 4)
    Rgn5 = CreateRectRgn(0, 4, 1, 7)
    CombineRgn Rgn, Rgn, Rgn1, RGN_DIFF
    CombineRgn Rgn, Rgn, Rgn2, RGN_DIFF
    CombineRgn Rgn, Rgn, Rgn3, RGN_DIFF
    CombineRgn Rgn, Rgn, Rgn4, RGN_DIFF
    CombineRgn Rgn, Rgn, Rgn5, RGN_DIFF
    DeleteObject Rgn1
    DeleteObject Rgn2
    DeleteObject Rgn3
    DeleteObject Rgn4
    DeleteObject Rgn5
    Rgn1 = CreateRectRgn(0, RC.Bottom - RC.Top - 6, 1, RC.Bottom - RC.Top)
    Rgn2 = CreateRectRgn(1, RC.Bottom - RC.Top - 3, 2, RC.Bottom - RC.Top)
    Rgn3 = CreateRectRgn(2, RC.Bottom - RC.Top - 2, 3, RC.Bottom - RC.Top)
    Rgn4 = CreateRectRgn(3, RC.Bottom - RC.Top - 1, 5, RC.Bottom - RC.Top)
    CombineRgn Rgn, Rgn, Rgn1, RGN_DIFF
    CombineRgn Rgn, Rgn, Rgn2, RGN_DIFF
    CombineRgn Rgn, Rgn, Rgn3, RGN_DIFF
    CombineRgn Rgn, Rgn, Rgn4, RGN_DIFF
    DeleteObject Rgn1
    DeleteObject Rgn2
    DeleteObject Rgn3
    DeleteObject Rgn4
    Rgn1 = CreateRectRgn(RC.Right - RC.Left - 5, RC.Bottom - RC.Top - 1, RC.Right - RC.Left, RC.Bottom - RC.Top)
    Rgn2 = CreateRectRgn(RC.Right - RC.Left - 3, RC.Bottom - RC.Top - 2, RC.Right - RC.Left, RC.Bottom - RC.Top - 1)
    Rgn3 = CreateRectRgn(RC.Right - RC.Left - 2, RC.Bottom - RC.Top - 3, RC.Right - RC.Left, RC.Bottom - RC.Top - 2)
    Rgn4 = CreateRectRgn(RC.Right - RC.Left - 1, RC.Bottom - RC.Top - 6, RC.Right - RC.Left, RC.Bottom - RC.Top - 3)
    CombineRgn Rgn, Rgn, Rgn1, RGN_DIFF
    CombineRgn Rgn, Rgn, Rgn2, RGN_DIFF
    CombineRgn Rgn, Rgn, Rgn3, RGN_DIFF
    CombineRgn Rgn, Rgn, Rgn4, RGN_DIFF
    DeleteObject Rgn1
    DeleteObject Rgn2
    DeleteObject Rgn3
    DeleteObject Rgn4
    Rgn1 = CreateRectRgn(112, 0, RC.Right - RC.Left, 18)
    Rgn2 = CreateRectRgn(RC.Right - RC.Left - 6, 18, RC.Right - RC.Left, 19)
    Rgn3 = CreateRectRgn(RC.Right - RC.Left - 4, 19, RC.Right - RC.Left, 20)
    Rgn4 = CreateRectRgn(RC.Right - RC.Left - 3, 20, RC.Right - RC.Left, 21)
    Rgn5 = CreateRectRgn(RC.Right - RC.Left - 2, 21, RC.Right - RC.Left, 22)
    Rgn6 = CreateRectRgn(RC.Right - RC.Left - 1, 22, RC.Right - RC.Left, 24)
    CombineRgn Rgn, Rgn, Rgn1, RGN_DIFF
    CombineRgn Rgn, Rgn, Rgn2, RGN_DIFF
    CombineRgn Rgn, Rgn, Rgn3, RGN_DIFF
    CombineRgn Rgn, Rgn, Rgn4, RGN_DIFF
    CombineRgn Rgn, Rgn, Rgn5, RGN_DIFF
    CombineRgn Rgn, Rgn, Rgn6, RGN_DIFF
    DeleteObject Rgn1
    DeleteObject Rgn2
    DeleteObject Rgn3
    DeleteObject Rgn4
    DeleteObject Rgn5
    DeleteObject Rgn6
    Rgn1 = CreateRectRgn(94, 0, 112, 1)
    Rgn2 = CreateRectRgn(95, 1, 112, 2)
    Rgn3 = CreateRectRgn(96, 2, 112, 3)
    Rgn4 = CreateRectRgn(97, 3, 112, 4)
    Rgn5 = CreateRectRgn(98, 4, 112, 5)
    Rgn6 = CreateRectRgn(99, 5, 112, 6)
    Rgn7 = CreateRectRgn(100, 6, 112, 7)
    Rgn8 = CreateRectRgn(101, 7, 112, 8)
    Rgn9 = CreateRectRgn(102, 8, 112, 9)
    CombineRgn Rgn, Rgn, Rgn1, RGN_DIFF
    CombineRgn Rgn, Rgn, Rgn2, RGN_DIFF
    CombineRgn Rgn, Rgn, Rgn3, RGN_DIFF
    CombineRgn Rgn, Rgn, Rgn4, RGN_DIFF
    CombineRgn Rgn, Rgn, Rgn5, RGN_DIFF
    CombineRgn Rgn, Rgn, Rgn6, RGN_DIFF
    CombineRgn Rgn, Rgn, Rgn7, RGN_DIFF
    CombineRgn Rgn, Rgn, Rgn8, RGN_DIFF
    CombineRgn Rgn, Rgn, Rgn9, RGN_DIFF
    DeleteObject Rgn1
    DeleteObject Rgn2
    DeleteObject Rgn3
    DeleteObject Rgn4
    DeleteObject Rgn5
    DeleteObject Rgn6
    DeleteObject Rgn7
    DeleteObject Rgn8
    DeleteObject Rgn9
    Rgn1 = CreateRectRgn(103, 9, 112, 10)
    Rgn2 = CreateRectRgn(104, 10, 112, 11)
    Rgn3 = CreateRectRgn(105, 11, 112, 12)
    Rgn4 = CreateRectRgn(106, 12, 112, 13)
    Rgn5 = CreateRectRgn(107, 13, 112, 14)
    Rgn6 = CreateRectRgn(108, 14, 112, 15)
    Rgn7 = CreateRectRgn(109, 15, 112, 16)
    Rgn8 = CreateRectRgn(110, 16, 112, 17)
    Rgn9 = CreateRectRgn(111, 17, 112, 18)
    CombineRgn Rgn, Rgn, Rgn1, RGN_DIFF
    CombineRgn Rgn, Rgn, Rgn2, RGN_DIFF
    CombineRgn Rgn, Rgn, Rgn3, RGN_DIFF
    CombineRgn Rgn, Rgn, Rgn4, RGN_DIFF
    CombineRgn Rgn, Rgn, Rgn5, RGN_DIFF
    CombineRgn Rgn, Rgn, Rgn6, RGN_DIFF
    CombineRgn Rgn, Rgn, Rgn7, RGN_DIFF
    CombineRgn Rgn, Rgn, Rgn8, RGN_DIFF
    CombineRgn Rgn, Rgn, Rgn9, RGN_DIFF
    DeleteObject Rgn1
    DeleteObject Rgn2
    DeleteObject Rgn3
    DeleteObject Rgn4
    DeleteObject Rgn5
    DeleteObject Rgn6
    DeleteObject Rgn7
    DeleteObject Rgn8
    DeleteObject Rgn9
    Rgn1 = CreateRectRgn(10, 30, RC.Right - RC.Left - 10, RC.Bottom - RC.Top - 13)
    CombineRgn Rgn, Rgn, Rgn1, RGN_DIFF
    DeleteObject Rgn1
    SetWindowRgn fTygemFrameTransparent.hWnd, Rgn, True
    DeleteObject Rgn
    fTygemFrameTransparent.Top = imgTopLeft.Top
    fTygemFrameTransparent.Left = imgTopLeft.Left
    imgFrameTexture.Top = 0
    imgFrameTexture.Left = 0
    imgFrameTexture.Width = fTygemFrameTransparent.Width
    imgFrameTexture.Height = fTygemFrameTransparent.Height
    
    '이미지 리스트 로드
    imgDropdownReverse.ListImages.Add 1, Picture:=imgDropdownReverse.ListImages(1).ExtractIcon()
    imgDropdownReverse.ListImages.Add 1, Picture:=imgDropdownReverse.ListImages(1).ExtractIcon()
    imgDropdownReverse.ListImages.Add 5, Picture:=imgDropdownReverse.ListImages(1).ExtractIcon()
    
    imgDropdown.ListImages.Add 1, Picture:=imgDropdown.ListImages(1).ExtractIcon()
    imgDropdown.ListImages.Add 1, Picture:=imgDropdown.ListImages(1).ExtractIcon()
    imgDropdown.ListImages.Add 5, Picture:=imgDropdown.ListImages(1).ExtractIcon()
    
    imgPlay.ListImages.Add 1, Picture:=imgPlay.ListImages(1).ExtractIcon()
    imgPlay.ListImages.Add 1, Picture:=imgPlay.ListImages(1).ExtractIcon()
    imgPlay.ListImages.Add 5, Picture:=imgPlay.ListImages(1).ExtractIcon()
    
    imgMinus.ListImages.Add 1, Picture:=imgMinus.ListImages(1).ExtractIcon()
    imgMinus.ListImages.Add 1, Picture:=imgMinus.ListImages(1).ExtractIcon()
    imgMinus.ListImages.Add 5, Picture:=imgMinus.ListImages(1).ExtractIcon()
    
    imgOpenFile.ListImages.Add 1, Picture:=imgOpenFile.ListImages(1).ExtractIcon()
    imgOpenFile.ListImages.Add 1, Picture:=imgOpenFile.ListImages(1).ExtractIcon()
    imgOpenFile.ListImages.Add 5, Picture:=imgOpenFile.ListImages(1).ExtractIcon()
    
    AttachMessage Me, Me.hWnd, WM_GETMINMAXINFO
    AttachMessage Me, Me.hWnd, WM_INITMENU
    AttachMessage Me, Me.hWnd, WM_SYSCOMMAND
    'AttachMessage Me, Me.hWnd, WM_DWMCOMPOSITIONCHANGED
    AttachMessage Me, Me.hWnd, WM_SETTINGCHANGE
    AttachMessage Me, Me.hWnd, WM_THEMECHANGED
    AttachMessage Me, Me.hWnd, WM_CTLCOLORSCROLLBAR
    
    vsProgressScroll.Visible = (trThreadCount.Value > 10 And optTabThreads2.Value)
End Sub

Sub SetPictureBuddy()
    If imgBackground.Visible Or imgBackgroundTile(0).Visible Or imgTopLeft.Visible Or fTygemFrameTransparent.Visible Then
        pbdDownloadInfo.AttachBuddy fDownloadInfo
        pbdThreadInfo.AttachBuddy fThreadInfo
        pbdProgressOuterContainer.AttachBuddy pbProgressOuterContainer
        pbdProgressContainer.AttachBuddy pbProgressContainer
    Else
        pbdDownloadInfo.DetachBuddy
        pbdThreadInfo.DetachBuddy
        pbdProgressOuterContainer.DetachBuddy
        pbdProgressContainer.DetachBuddy
        ThreadBuddyAttached = False
    End If
End Sub

Sub SetTextColors()
    Dim DisableVisualStyle As Boolean
    DisableVisualStyle = CBool(CInt(GetSetting("DownloadBooster", "Options", "DisableVisualStyle", 0)))

    Dim StatusTextColor&, FrameCaptionColor&
    If DisableVisualStyle Then
        StatusTextColor = &H80000012
        FrameCaptionColor = 0&
    Else
        StatusTextColor = GetThemeColor(Me.hWnd, "STATUS", DefaultColor:=&H80000012)
        FrameCaptionColor = GetThemeColor(Me.hWnd, "BUTTON", 4)
    End If
    
    Dim i%
    For i = 1 To sbStatusBar.Panels.Count
        sbStatusBar.Panels(i).ForeColor = StatusTextColor
    Next i
    
    If optTabDownload2.VisualStyles Then
        fTabDownload.ForeColor = FrameCaptionColor
        fTabThreads.ForeColor = FrameCaptionColor
    End If
End Sub

Sub SetupSplitButtons()
    cmdOpenBatch.GetTygemButton().SplitLeft = True
    cmdOpenDropdown.GetTygemButton().SplitRight = True
    cmdOpenBatch.SetRgn
    cmdOpenDropdown.SetRgn
    
    cmdDelete.GetTygemButton().SplitLeft = True
    cmdDeleteDropdown.GetTygemButton().SplitRight = True
    cmdDelete.SetRgn
    cmdDeleteDropdown.SetRgn
    
    cmdOpen.GetTygemButton().SplitLeft = True
    cmdOpenFileDropdown.GetTygemButton().SplitRight = True
    cmdOpen.SetRgn
    cmdOpenFileDropdown.SetRgn

    If WinVer >= 6.1 Then
        If GetSetting("DownloadBooster", "Options", "EnableLiveBadukMemoSkin", 0) = 0 Then
            cmdOpenBatch.SplitButton = True
            cmdOpenBatch.Width = 1935
            cmdOpenDropdown.Visible = False
            
            cmdDelete.SplitButton = True
            cmdDelete.Width = 1575
            cmdDeleteDropdown.Visible = False
            
            cmdOpen.SplitButton = True
            cmdOpen.Width = 1935
            cmdOpenFileDropdown.Visible = False
        Else
            cmdOpenBatch.SplitButton = False
            cmdOpenBatch.Width = 1575
            cmdOpenDropdown.Visible = True
            
            cmdDelete.SplitButton = False
            cmdDelete.Width = 1335
            cmdDeleteDropdown.Visible = True
            
            cmdOpen.SplitButton = False
            cmdOpen.Width = 1695
            cmdOpenFileDropdown.Visible = True
        End If
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    imgBorderTopLeft.Left = 0
    imgBorderTopLeft.Top = 0
    imgBorderTopRight.Top = 0
    imgBorderTopRight.Left = Me.Width - imgBorderTopRight.Width - SizingBorderWidth * 15 * 2
    imgBorderBottomLeft.Left = 0
    imgBorderBottomLeft.Top = Me.Height - sbStatusBar.Height - imgBorderBottomLeft.Height - SizingBorderWidth * 15 * 2 - CaptionHeight * 15 - 15
    imgBorderBottomRight.Left = Me.Width - imgBorderBottomRight.Width - SizingBorderWidth * 15 * 2
    imgBorderBottomRight.Top = Me.Height - sbStatusBar.Height - imgBorderBottomRight.Height - SizingBorderWidth * 15 * 2 - CaptionHeight * 15 - 15
    pgBorderTop.Left = 0
    pgBorderTop.Top = -15
    pgBorderTop.Width = Me.Width
    pgBorderBottom.Left = 0
    pgBorderBottom.Top = Me.Height - sbStatusBar.Height - 30 - SizingBorderWidth * 15 * 2 - CaptionHeight * 15 - 15
    pgBorderBottom.Width = Me.Width
    pgBorderLeft.Left = -15
    pgBorderLeft.Top = 0
    pgBorderLeft.Height = Me.Height
    pgBorderRight.Top = 0
    pgBorderRight.Left = Me.Width - 30 - SizingBorderWidth * 15 * 2
    pgBorderRight.Height = Me.Height
    
    pgPattern.Width = Me.Width
    pgPattern.Height = Me.Height
    
    If Me.Height <= 6930 + PaddedBorderWidth * 15 * 2 Then Exit Sub
    If Me.Height - lvBatchFiles.Top - 1320 < 870 + PaddedBorderWidth * 15 * 2 Then Exit Sub
    If Me.WindowState = 1 Then Exit Sub
    
    lvBatchFiles.Height = Me.Height - PaddedBorderWidth * 15 * 2 - lvBatchFiles.Top - 1320
    cmdOpenBatch.Top = lvBatchFiles.Top + lvBatchFiles.Height + 45
    cmdOpenDropdown.Top = lvBatchFiles.Top + lvBatchFiles.Height + 45
    cmdAdd.Top = lvBatchFiles.Top + lvBatchFiles.Height + 45
    cmdDelete.Top = lvBatchFiles.Top + lvBatchFiles.Height + 45
    cmdDeleteDropdown.Top = lvBatchFiles.Top + lvBatchFiles.Height + 45
    cmdStartBatch.Top = lvBatchFiles.Top + lvBatchFiles.Height + 45
    cmdStopBatch.Top = lvBatchFiles.Top + lvBatchFiles.Height + 45
    cmdEdit.Top = lvBatchFiles.Top + lvBatchFiles.Height + 45
    
    SetBackgroundPosition
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i%
    If cmdStop.Enabled = -1 Or BatchStarted Then
        Dim IsMarquee As Boolean
        IsMarquee = pbTotalProgressMarquee.Visible
        Dim ConfirmResult As VbMsgBoxResult
        If IsMarquee Or ResumeUnsupported Then
            ConfirmResult = ConfirmEx(t("다운로드를 중지하시겠습니까? 현재 파일은 이어받기가 지원되지 않으므로 처음부터 다시 다운로드받아야 합니다.", "Cancel download? Resuming is not supported for this file."), t("다운로드 취소", "Cancel download"), 48)
        Else
            ConfirmResult = Confirm(t("다운로드를 중지하시겠습니까? 이어받기 기능을 통해 중단한 곳부터 계속 다운로드받을 수 있습니다.", "Cancel download? You can resume later."), t("다운로드 취소", "Cancel download"))
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
                    KillTemp = Confirm(t("나중에 계속 이어서 다운로드받을 수 있도록 다운로드한 데이타를 저장하시겠습니까?", "Would you like to keep the partially downloaded data to resume later?"), App.Title) <> vbYes
                End If
                If KillTemp Then
                    On Error Resume Next
                    If trThreadCount.Value <= 1 Then
                        Kill DownloadPath & ".part.tmp"
                    Else
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
    SaveSetting "DownloadBooster", "UserData", "BatchExpanded", CInt(Me.Height > 6930 + PaddedBorderWidth * 15 * 2) * -1
    SaveSetting "DownloadBooster", "Options", "WhenFileExists", cbWhenExist.ListIndex
    If GetSetting("DownloadBooster", "Options", "RememberURL", 0) <> 0 Then
        SaveSetting "DownloadBooster", "UserData", "FileURL", Trim$(txtURL.Text)
    End If
    SaveSetting "DownloadBooster", "UserData", "FormTop", Me.Top
    SaveSetting "DownloadBooster", "UserData", "FormLeft", Me.Left
    If Me.Height >= 8220 Then SaveSetting "DownloadBooster", "UserData", "FormHeight", Me.Height - PaddedBorderWidth * 15 * 2
    SaveSetting "DownloadBooster", "UserData", "LastTab", (CInt(optTabThreads2.Value) * -1) + 1
    
    On Error Resume Next
    Me.Hide
    For i = 1 To MAX_THREAD_COUNT
        Unload lblDownloader(i)
        Unload lblPercentage(i)
        Unload pbProgress(i)
        Unload pbProgressMarquee(i)
    Next i
    Unload frmBatchAdd
    Unload frmBrowse
    Unload frmOptions
    Unload frmExplorer
    Unload frmDummyForm
    Unload frmEditBatch
    
    DetachMessage Me, Me.hWnd, WM_GETMINMAXINFO
    DetachMessage Me, Me.hWnd, WM_INITMENU
    DetachMessage Me, Me.hWnd, WM_SYSCOMMAND
    'DetachMessage Me, Me.hWnd, WM_DWMCOMPOSITIONCHANGED
    DetachMessage Me, Me.hWnd, WM_SETTINGCHANGE
    DetachMessage Me, Me.hWnd, WM_THEMECHANGED
    DetachMessage Me, Me.hWnd, WM_CTLCOLORSCROLLBAR
    
    GetSystemMenu Me.hWnd, 1&
    Unload frmMessageBox
    Unload frmAbout
    Unload frmDownloadOptions
    Unload frmMain
    If Not InIDE Then ExitProcess 0&
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
        If cmdDelete.Enabled Then
            mnuOpenBatch.Visible = cmdOpenBatch.Enabled
            mnuOpenFolder2.Visible = cmdOpenBatch.Enabled
            mnuSepOpen.Visible = (lvBatchFiles.SelectedItem.ForeColor = vbRed Or cmdOpenBatch.Enabled)
            mnuMoveUp.Enabled = (lvBatchFiles.SelectedItem.Index <> 1) And (Not BatchStarted)
            mnuMoveDown.Enabled = (lvBatchFiles.SelectedItem.Index <> lvBatchFiles.ListItems.Count) And (Not BatchStarted)
            mnuErrorInfo.Visible = (lvBatchFiles.SelectedItem.ForeColor = vbRed)
            If cmdOpenBatch.Enabled Then
                Me.PopupMenu mnuListContext, , , , mnuOpenBatch
            ElseIf mnuErrorInfo.Visible Then
                Me.PopupMenu mnuListContext, , , , mnuErrorInfo
            Else
                Me.PopupMenu mnuListContext, , , , mnuEdit
            End If
        End If
    Else
        GoTo ErrLn
    End If
    
    Exit Sub
ErrLn:
    mnuClearBatch2.Enabled = (lvBatchFiles.ListItems.Count > 0)
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

Private Sub lvBatchFiles_ItemDblClick(ByVal Item As LvwListItem, ByVal Button As Integer)
    On Error Resume Next
    If Not Item.Selected Then Exit Sub
    If cmdOpenBatch.Enabled And Item.ListSubItems(3).Text = t("완료", "Done") Then
        cmdOpenBatch_Click
    ElseIf Item.ForeColor = vbRed Then
        mnuErrorInfo_Click
    ElseIf (Not BatchStarted) Or (BatchStarted And CurrentBatchIdx <> Item.Index) Then
        mnuEdit_Click
    End If
End Sub

Private Sub lvBatchFiles_ItemSelect(ByVal Item As LvwListItem, ByVal Selected As Boolean)
    If Selected Then
        If BatchStarted And Item.Index = CurrentBatchIdx Then
            cmdDelete.Enabled = 0
            cmdDeleteDropdown.Enabled = 0
            cmdEdit.Enabled = 0
        Else
            cmdDelete.Enabled = -1
            cmdDeleteDropdown.Enabled = -1
            cmdEdit.Enabled = -1
        End If
        
        If Item.ListSubItems(3).Text = t("완료", "Done") Then
            cmdOpenBatch.Enabled = -1
            cmdOpenDropdown.Enabled = -1
        Else
            cmdOpenBatch.Enabled = 0
            cmdOpenDropdown.Enabled = 0
        End If
    Else
        cmdDelete.Enabled = 0
        cmdDeleteDropdown.Enabled = 0
        cmdOpenBatch.Enabled = 0
        cmdOpenDropdown.Enabled = 0
        
        cmdEdit.Enabled = 0
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

Private Sub mnuAddItem_Click()
    cmdAdd_Click
End Sub

Private Sub mnuAddItem2_Click()
    mnuAddItem_Click
End Sub

Private Sub mnuClearBatch_Click()
    If lvBatchFiles.ListItems.Count Then
        If Confirm(t("대기열의 모든 항목을 삭제하시겠습니까?", "Are you sure you want to clear the queue?"), App.Title) <> vbYes Then Exit Sub
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
        
        If Not BatchStarted Then cmdStartBatch.Enabled = False
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

Private Sub mnuEdit_Click()
    frmEditBatch.EncodedHeaders = lvBatchFiles.SelectedItem.ListSubItems(5).Text
    On Error GoTo exitsub22
    frmEditBatch.txtURL.Text = lvBatchFiles.SelectedItem.ListSubItems(2).Text
    frmEditBatch.txtFilePath.Text = lvBatchFiles.SelectedItem.ListSubItems(1).Text
    Tags.FileNameOnly = lvBatchFiles.SelectedItem.Text
    On Error Resume Next
    frmEditBatch.txtURL.SelStart = 0
    frmEditBatch.txtURL.SelLength = Len(frmEditBatch.txtURL.Text)
    'frmEditBatch.Label2.Enabled = (lvBatchFiles.SelectedItem.ListSubItems(3).Text <> t("완료", "Done"))
    'frmEditBatch.txtFilePath.Enabled = frmEditBatch.Label2.Enabled
    frmEditBatch.OriginalURL = lvBatchFiles.SelectedItem.ListSubItems(2).Text
    frmEditBatch.OriginalPath = lvBatchFiles.SelectedItem.ListSubItems(1).Text
    frmEditBatch.Show vbModal, Me
exitsub22:
    Exit Sub
End Sub

Private Sub mnuMoveDown_Click()
    On Error GoTo exitsub
    Dim CurIdx, DownIdx, NewIdx As Integer
    CurIdx = lvBatchFiles.SelectedItem.Index
    If CurIdx >= lvBatchFiles.ListItems.Count Then Exit Sub
    DownIdx = CurIdx + 1
    NewIdx = lvBatchFiles.ListItems.Add(CurIdx, , lvBatchFiles.ListItems(DownIdx).Text).Index
    DownIdx = DownIdx + 1
    lvBatchFiles.ListItems(NewIdx).ListSubItems.Add , , lvBatchFiles.ListItems(DownIdx).ListSubItems(1).Text
    lvBatchFiles.ListItems(NewIdx).ListSubItems.Add , , lvBatchFiles.ListItems(DownIdx).ListSubItems(2).Text
    lvBatchFiles.ListItems(NewIdx).ListSubItems.Add , , lvBatchFiles.ListItems(DownIdx).ListSubItems(3).Text
    lvBatchFiles.ListItems(NewIdx).ListSubItems.Add , , lvBatchFiles.ListItems(DownIdx).ListSubItems(4).Text
    lvBatchFiles.ListItems(NewIdx).ListSubItems.Add , , lvBatchFiles.ListItems(DownIdx).ListSubItems(5).Text
#If HIDEYTDL Then
    GoTo afterheaderadd
#End If
    lvBatchFiles.ListItems(NewIdx).ListSubItems.Add , , lvBatchFiles.ListItems(DownIdx).ListSubItems(6).Text
    lvBatchFiles.ListItems(NewIdx).ListSubItems.Add , , lvBatchFiles.ListItems(DownIdx).ListSubItems(7).Text
    lvBatchFiles.ListItems(NewIdx).ListSubItems.Add , , lvBatchFiles.ListItems(DownIdx).ListSubItems(8).Text
    lvBatchFiles.ListItems(NewIdx).ListSubItems.Add , , lvBatchFiles.ListItems(DownIdx).ListSubItems(9).Text
    lvBatchFiles.ListItems(NewIdx).ListSubItems.Add , , lvBatchFiles.ListItems(DownIdx).ListSubItems(10).Text
    lvBatchFiles.ListItems(NewIdx).ListSubItems.Add , , lvBatchFiles.ListItems(DownIdx).ListSubItems(11).Text
    lvBatchFiles.ListItems(NewIdx).ListSubItems.Add , , lvBatchFiles.ListItems(DownIdx).ListSubItems(12).Text
afterheaderadd:
    lvBatchFiles.ListItems(NewIdx).Checked = lvBatchFiles.ListItems(DownIdx).Checked
    lvBatchFiles.ListItems(NewIdx).ForeColor = lvBatchFiles.ListItems(DownIdx).ForeColor
    lvBatchFiles.ListItems(NewIdx).ListSubItems(1).ForeColor = lvBatchFiles.ListItems(DownIdx).ListSubItems(1).ForeColor
    lvBatchFiles.ListItems(NewIdx).ListSubItems(2).ForeColor = lvBatchFiles.ListItems(DownIdx).ListSubItems(2).ForeColor
    lvBatchFiles.ListItems(NewIdx).ListSubItems(3).ForeColor = lvBatchFiles.ListItems(DownIdx).ListSubItems(3).ForeColor
    lvBatchFiles.ListItems(NewIdx).ListSubItems(3).Text = lvBatchFiles.ListItems(DownIdx).ListSubItems(3).Text
    
    lvBatchFiles.ListItems.Remove DownIdx
    
exitsub:
    Exit Sub
End Sub

Private Sub mnuMoveUp_Click()
    On Error GoTo exitsub
    Dim CurIdx, UpIdx, NewIdx As Integer
    CurIdx = lvBatchFiles.SelectedItem.Index
    If CurIdx <= 1 Then Exit Sub
    UpIdx = CurIdx - 1
    NewIdx = lvBatchFiles.ListItems.Add(CurIdx + 1, , lvBatchFiles.ListItems(UpIdx).Text).Index
    lvBatchFiles.ListItems(NewIdx).ListSubItems.Add , , lvBatchFiles.ListItems(UpIdx).ListSubItems(1).Text
    lvBatchFiles.ListItems(NewIdx).ListSubItems.Add , , lvBatchFiles.ListItems(UpIdx).ListSubItems(2).Text
    lvBatchFiles.ListItems(NewIdx).ListSubItems.Add , , lvBatchFiles.ListItems(UpIdx).ListSubItems(3).Text
    lvBatchFiles.ListItems(NewIdx).ListSubItems.Add , , lvBatchFiles.ListItems(UpIdx).ListSubItems(4).Text
    lvBatchFiles.ListItems(NewIdx).ListSubItems.Add , , lvBatchFiles.ListItems(UpIdx).ListSubItems(5).Text
#If HIDEYTDL Then
    GoTo afterheaderadd
#End If
    lvBatchFiles.ListItems(NewIdx).ListSubItems.Add , , lvBatchFiles.ListItems(UpIdx).ListSubItems(6).Text
    lvBatchFiles.ListItems(NewIdx).ListSubItems.Add , , lvBatchFiles.ListItems(UpIdx).ListSubItems(7).Text
    lvBatchFiles.ListItems(NewIdx).ListSubItems.Add , , lvBatchFiles.ListItems(UpIdx).ListSubItems(8).Text
    lvBatchFiles.ListItems(NewIdx).ListSubItems.Add , , lvBatchFiles.ListItems(UpIdx).ListSubItems(9).Text
    lvBatchFiles.ListItems(NewIdx).ListSubItems.Add , , lvBatchFiles.ListItems(UpIdx).ListSubItems(10).Text
    lvBatchFiles.ListItems(NewIdx).ListSubItems.Add , , lvBatchFiles.ListItems(UpIdx).ListSubItems(11).Text
    lvBatchFiles.ListItems(NewIdx).ListSubItems.Add , , lvBatchFiles.ListItems(UpIdx).ListSubItems(12).Text
afterheaderadd:
    lvBatchFiles.ListItems(NewIdx).Checked = lvBatchFiles.ListItems(UpIdx).Checked
    lvBatchFiles.ListItems(NewIdx).ForeColor = lvBatchFiles.ListItems(UpIdx).ForeColor
    lvBatchFiles.ListItems(NewIdx).ListSubItems(1).ForeColor = lvBatchFiles.ListItems(UpIdx).ListSubItems(1).ForeColor
    lvBatchFiles.ListItems(NewIdx).ListSubItems(2).ForeColor = lvBatchFiles.ListItems(UpIdx).ListSubItems(2).ForeColor
    lvBatchFiles.ListItems(NewIdx).ListSubItems(3).ForeColor = lvBatchFiles.ListItems(UpIdx).ListSubItems(3).ForeColor
    lvBatchFiles.ListItems(NewIdx).ListSubItems(3).Text = lvBatchFiles.ListItems(UpIdx).ListSubItems(3).Text
    
    lvBatchFiles.ListItems.Remove UpIdx
    
exitsub:
    Exit Sub
End Sub

Private Sub mnuOpenBatch_Click()
    cmdOpenBatch_Click
End Sub

Private Sub mnuOpenFolder_Click()
    Dim pth$
    pth = lvBatchFiles.SelectedItem.ListSubItems(1).Text
    If pth = "" Then pth = txtFileName.Text
    If FolderExists(pth) Then
        Shell "cmd /c start """" explorer.exe """ & pth & """"
    Else
        Shell "cmd /c start """" explorer.exe """ & GetParentFolderName(pth) & """"
    End If
End Sub

Private Sub mnuOpenFolder2_Click()
    mnuOpenFolder_Click
End Sub

Private Sub mnuProperties_Click()
    On Error Resume Next
    DisplayFileProperties DownloadPath
End Sub

Private Sub mnuPropertiesBatch_Click()
    On Error Resume Next
    DisplayFileProperties lvBatchFiles.SelectedItem.ListSubItems(1).Text
End Sub

Private Sub optTabDownload2_Click()
    fDownloadInfo.Visible = -1
    fThreadInfo.Visible = 0
    vsProgressScroll.Visible = False
End Sub

Private Sub optTabDownload2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fTabDownload_Click
End Sub

Private Sub optTabThreads2_Click()
    fThreadInfo.Visible = -1
    fDownloadInfo.Visible = 0
    If Not ThreadBuddyAttached Then
        SetPictureBuddy
        ThreadBuddyAttached = True
    End If
    vsProgressScroll.Visible = trThreadCount.Value > 10
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

Sub trThreadCount_Scroll()
    If trThreadCount.Value = 1 Then
        lblThreadCount.Caption = "(" & t("일반 다운로드", "No threading") & ")"
    Else
        lblThreadCount.Caption = "(" & trThreadCount.Value & t("개 스레드", " threads") & ")"
    End If
    Dim i%
    For i = 1 To trThreadCount.Value
        lblDownloader(i).Visible = -1
        pbProgress(i).Visible = -1
        lblPercentage(i).Visible = -1
    Next i
    For i = trThreadCount.Value + 1 To lblDownloader.UBound
        lblDownloader(i).Visible = 0
        pbProgress(i).Visible = 0
        lblPercentage(i).Visible = 0
    Next i
    
    If trThreadCount.Value - 10 > 0 Then
        If ScrollOneScreen Then
            vsProgressScroll.Max = Ceil(trThreadCount.Value / 10) - 1
        Else
            vsProgressScroll.Max = trThreadCount.Value - 10
        End If
        vsProgressScroll.Enabled = -1
        vsProgressScroll.Visible = fThreadInfo.Visible
    Else
        If vsProgressScroll.Max <> 0 Then vsProgressScroll.Max = 0
        If vsProgressScroll.Enabled Then vsProgressScroll.Enabled = 0
        
        vsProgressScroll.Visible = False
        pbProgressContainer.Top = 0
    End If
    
'    If trThreadCount.Value <= 1 Then
'        fDownloadInfo.Visible = -1
'        fThreadInfo.Visible = 0
'        optTabDownload2.Value = True
'    Else
'        fThreadInfo.Visible = -1
'        fDownloadInfo.Visible = 0
'        optTabThreads2.Value = True
'    End If
    
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
    
    pbProgressContainer.Refresh
End Sub

Private Sub vsProgressScroll_Change()
    vsProgressScroll_Scroll
End Sub

Private Sub vsProgressScroll_Scroll()
    If ScrollOneScreen Then
        pbProgressContainer.Top = CDbl(pbProgressOuterContainer.Height) * CDbl(vsProgressScroll.Value) * -1# - (105# * CDbl(vsProgressScroll.Value))
    Else
        pbProgressContainer.Top = CDbl(vsProgressScroll.Value) * 255# * -1# - (105# * CDbl(vsProgressScroll.Value))
    End If
    If LBFrameEnabled Or imgBackground.Visible Then
        pbdProgressContainer.Refresh
    End If
End Sub
