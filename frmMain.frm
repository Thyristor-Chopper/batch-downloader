VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "�ٿ�ε� �ν���"
   ClientHeight    =   7740
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   17295
   BeginProperty Font 
      Name            =   "����"
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
   ScaleWidth      =   17295
   StartUpPosition =   3  'Windows �⺻��
   Begin prjDownloadBooster.FrameW fProgress 
      Height          =   255
      Left            =   5640
      TabIndex        =   81
      Top             =   1800
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   450
      BorderStyle     =   0
      Transparent     =   -1  'True
      Begin VB.Label lblProgress 
         Alignment       =   1  '������ ����
         BackStyle       =   0  '����
         Height          =   255
         Left            =   0
         TabIndex        =   82
         Top             =   0
         Width           =   735
      End
   End
   Begin prjDownloadBooster.FrameW fState 
      Height          =   255
      Left            =   360
      TabIndex        =   79
      Top             =   1800
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   450
      BorderStyle     =   0
      Transparent     =   -1  'True
      Begin VB.Label lblLBState 
         BackStyle       =   0  '����
         Caption         =   "������"
         Height          =   255
         Left            =   0
         TabIndex        =   80
         Top             =   0
         Width           =   735
      End
   End
   Begin prjDownloadBooster.FrameW fTabs 
      Height          =   255
      Left            =   480
      TabIndex        =   13
      Top             =   2040
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   450
      BorderStyle     =   0
      Transparent     =   -1  'True
      Begin VB.OptionButton optTabDownload2 
         Height          =   195
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   195
      End
      Begin VB.OptionButton optTabThreads2 
         Height          =   195
         Left            =   1200
         TabIndex        =   17
         Top             =   0
         Value           =   -1  'True
         Width           =   195
      End
      Begin VB.Label fTabDownload 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "��� ����"
         Height          =   180
         Left            =   240
         TabIndex        =   14
         Top             =   15
         Width           =   780
      End
      Begin VB.Label fTabThreads 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "������"
         Height          =   180
         Left            =   1440
         TabIndex        =   16
         Top             =   15
         Width           =   540
      End
   End
   Begin VB.VScrollBar vsProgressScroll 
      Height          =   3495
      LargeChange     =   10
      Left            =   6120
      Max             =   5
      TabIndex        =   18
      Top             =   2280
      Width           =   255
   End
   Begin prjDownloadBooster.FrameW fTygemFrameTransparent 
      Height          =   4845
      Left            =   10560
      TabIndex        =   72
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
         Alignment       =   2  '��� ����
         BackStyle       =   0  '����
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   150
         TabIndex        =   73
         Tag             =   "nosizechange"
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label lblLBCaptionShadow2 
         Alignment       =   2  '��� ����
         BackStyle       =   0  '����
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   165
         TabIndex        =   74
         Tag             =   "nocolorsizechange"
         Top             =   135
         Width           =   1095
      End
      Begin VB.Image imgFrame 
         Height          =   4845
         Index           =   1
         Left            =   0
         Picture         =   "frmMain.frx":08CA
         Stretch         =   -1  'True
         Top             =   0
         Visible         =   0   'False
         Width           =   6495
      End
      Begin VB.Image imgFrame 
         Height          =   4845
         Index           =   2
         Left            =   0
         Picture         =   "frmMain.frx":1075
         Stretch         =   -1  'True
         Top             =   0
         Visible         =   0   'False
         Width           =   6495
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00404040&
         X1              =   135
         X2              =   1725
         Y1              =   435
         Y2              =   435
      End
      Begin VB.Image imgFrameTexture 
         Height          =   4845
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Visible         =   0   'False
         Width           =   6495
      End
   End
   Begin prjDownloadBooster.CommandButtonW cmdOpen 
      Height          =   330
      Left            =   7200
      TabIndex        =   29
      Top             =   4065
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      Enabled         =   0   'False
      ImageList       =   "imgOpenFile"
      Caption         =   "����(&O) "
      Transparent     =   -1  'True
   End
   Begin prjDownloadBooster.CommandButtonW cmdOpenBatch 
      Height          =   375
      Left            =   240
      TabIndex        =   39
      Top             =   6960
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Enabled         =   0   'False
      ImageList       =   "imgOpenFile"
      Caption         =   "����(&W) "
      Transparent     =   -1  'True
   End
   Begin prjDownloadBooster.CommandButtonW cmdDelete 
      Height          =   375
      Left            =   4200
      TabIndex        =   42
      Top             =   6960
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Enabled         =   0   'False
      ImageList       =   "imgMinus"
      Caption         =   "����(&V) "
      Transparent     =   -1  'True
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
      TabIndex        =   28
      Top             =   3690
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      Caption         =   "�ٿ�ε� ����(&S)..."
      Transparent     =   -1  'True
   End
   Begin prjDownloadBooster.ListView lvLogTest 
      Height          =   1335
      Left            =   6600
      TabIndex        =   37
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
      TabIndex        =   36
      Top             =   5520
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      Caption         =   "��"
   End
   Begin prjDownloadBooster.CommandButtonW cmdEdit 
      Height          =   375
      Left            =   5880
      TabIndex        =   44
      Top             =   6960
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Enabled         =   0   'False
      ImageList       =   "imgEdit"
      Caption         =   "����(&N)..."
      Transparent     =   -1  'True
   End
   Begin prjDownloadBooster.ImageList imgEdit 
      Left            =   9240
      Top             =   5880
      _ExtentX        =   1005
      _ExtentY        =   1005
      ImageWidth      =   16
      ImageHeight     =   16
      ColorDepth      =   4
      InitListImages  =   "frmMain.frx":182F
   End
   Begin prjDownloadBooster.CommandButtonW cmdStopBatch 
      Height          =   375
      Left            =   7560
      TabIndex        =   46
      Top             =   6960
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Enabled         =   0   'False
      ImageList       =   "imgStopRed"
      Caption         =   "����(&Z) "
      Transparent     =   -1  'True
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
      ShowInTaskBar   =   -1  'True
   End
   Begin prjDownloadBooster.ImageList imgWrench 
      Left            =   9240
      Top             =   6600
      _ExtentX        =   1005
      _ExtentY        =   1005
      ImageWidth      =   16
      ImageHeight     =   16
      ColorDepth      =   4
      InitListImages  =   "frmMain.frx":1D17
   End
   Begin prjDownloadBooster.ImageList imgErase 
      Left            =   9840
      Top             =   6600
      _ExtentX        =   1005
      _ExtentY        =   1005
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      InitListImages  =   "frmMain.frx":21FF
   End
   Begin prjDownloadBooster.StatusBar sbStatusBar 
      Align           =   2  '�Ʒ� ����
      Height          =   330
      Left            =   0
      Top             =   7410
      Width           =   17295
      _ExtentX        =   30506
      _ExtentY        =   582
      InitPanels      =   "frmMain.frx":25E7
   End
   Begin prjDownloadBooster.ListView lvBatchFiles 
      Height          =   870
      Left            =   240
      TabIndex        =   38
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
      TabIndex        =   27
      Top             =   3195
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   529
      Caption         =   "���α׷� ����(&U)"
      Transparent     =   -1  'True
   End
   Begin prjDownloadBooster.CommandButtonW cmdOptions 
      Height          =   300
      Left            =   7080
      TabIndex        =   26
      Top             =   2865
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   529
      ImageList       =   "imgWrench"
      Caption         =   "�߰� �ɼ�(&I)..."
      Transparent     =   -1  'True
   End
   Begin prjDownloadBooster.CommandButtonW cmdStop 
      Height          =   330
      Left            =   7200
      TabIndex        =   33
      Top             =   4815
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      Enabled         =   0   'False
      ImageList       =   "imgStopRed"
      Caption         =   "����(&P) "
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
      InitListImages  =   "frmMain.frx":28EB
   End
   Begin prjDownloadBooster.CommandButtonW cmdOpenDropdown 
      Height          =   375
      Left            =   1800
      TabIndex        =   40
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
      InitListImages  =   "frmMain.frx":2A9B
   End
   Begin prjDownloadBooster.CommandButtonW cmdDeleteDropdown 
      Height          =   375
      Left            =   5520
      TabIndex        =   43
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
      InitListImages  =   "frmMain.frx":2D9B
   End
   Begin prjDownloadBooster.CommandButtonW cmdAddToQueue 
      Height          =   330
      Left            =   7200
      TabIndex        =   34
      Top             =   5190
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      ImageList       =   "imgPlusYellow"
      Caption         =   "��Ͽ� �߰�(&Q)"
      Transparent     =   -1  'True
   End
   Begin prjDownloadBooster.CommandButtonW cmdStartBatch 
      Height          =   375
      Left            =   7560
      TabIndex        =   45
      Top             =   6960
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Enabled         =   0   'False
      ImageList       =   "imgPlay"
      Caption         =   "����(&S) "
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
      InitListImages  =   "frmMain.frx":3183
   End
   Begin prjDownloadBooster.ImageList imgPlay 
      Left            =   9840
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      InitListImages  =   "frmMain.frx":356B
   End
   Begin prjDownloadBooster.ImageList imgDownload 
      Left            =   9840
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      InitListImages  =   "frmMain.frx":3CDB
   End
   Begin prjDownloadBooster.ImageList imgMinus 
      Left            =   9840
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      InitListImages  =   "frmMain.frx":40C3
   End
   Begin prjDownloadBooster.ImageList imgOpenFile 
      Left            =   9840
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      InitListImages  =   "frmMain.frx":4833
   End
   Begin prjDownloadBooster.ImageList imgOpenFolder 
      Left            =   9840
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      InitListImages  =   "frmMain.frx":4FA3
   End
   Begin prjDownloadBooster.FrameW fDownloadInfo 
      Height          =   3255
      Left            =   360
      TabIndex        =   49
      Tag             =   "forcebgchange"
      Top             =   2640
      Visible         =   0   'False
      Width           =   5175
      _ExtentX        =   0
      _ExtentY        =   0
      BorderStyle     =   0
      Transparent     =   -1  'True
      Begin VB.Label lblRemaining 
         BackStyle       =   0  '����
         Caption         =   "-"
         Height          =   255
         Left            =   1320
         TabIndex        =   67
         Top             =   2580
         Width           =   4335
      End
      Begin VB.Label Label10 
         BackStyle       =   0  '����
         Caption         =   "���� �ð�:"
         Height          =   255
         Left            =   0
         TabIndex        =   68
         Top             =   2580
         Width           =   1215
      End
      Begin VB.Label Label8 
         BackStyle       =   0  '����
         Caption         =   "���� �̸�:"
         Height          =   255
         Left            =   0
         TabIndex        =   63
         Top             =   60
         Width           =   1095
      End
      Begin VB.Label lblFilename 
         BackStyle       =   0  '����
         Caption         =   "-"
         Height          =   255
         Left            =   1320
         TabIndex        =   62
         Top             =   60
         Width           =   4335
      End
      Begin VB.Label lblTotalSizeThread 
         BackStyle       =   0  '����
         Caption         =   "-"
         Height          =   255
         Left            =   1320
         TabIndex        =   61
         Top             =   2220
         Width           =   4335
      End
      Begin VB.Label Label7 
         BackStyle       =   0  '����
         Caption         =   "������� ũ��:"
         Height          =   255
         Left            =   0
         TabIndex        =   60
         Top             =   2220
         Width           =   1215
      End
      Begin VB.Label lblThreadCount2 
         BackStyle       =   0  '����
         Caption         =   "-"
         Height          =   255
         Left            =   1320
         TabIndex        =   59
         Top             =   1860
         Width           =   4335
      End
      Begin VB.Label Label6 
         BackStyle       =   0  '����
         Caption         =   "������ ��:"
         Height          =   255
         Left            =   0
         TabIndex        =   58
         Top             =   1860
         Width           =   975
      End
      Begin VB.Label Label5 
         BackStyle       =   0  '����
         Caption         =   "�ӵ�:"
         Height          =   255
         Left            =   0
         TabIndex        =   57
         Top             =   1500
         Width           =   975
      End
      Begin VB.Label lblSpeed 
         BackStyle       =   0  '����
         Caption         =   "-"
         Height          =   255
         Left            =   1320
         TabIndex        =   56
         Top             =   1500
         Width           =   4335
      End
      Begin VB.Label lblElapsed 
         BackStyle       =   0  '����
         Caption         =   "-"
         Height          =   255
         Left            =   1320
         TabIndex        =   55
         Top             =   1140
         Width           =   4335
      End
      Begin VB.Label Label4 
         BackStyle       =   0  '����
         Caption         =   "��� �ð�:"
         Height          =   255
         Left            =   0
         TabIndex        =   54
         Top             =   1140
         Width           =   975
      End
      Begin VB.Label lblDownloadedBytes 
         BackStyle       =   0  '����
         Caption         =   "-"
         Height          =   255
         Left            =   1320
         TabIndex        =   53
         Top             =   780
         Width           =   4335
      End
      Begin VB.Label Label3 
         BackStyle       =   0  '����
         Caption         =   "���� ũ��:"
         Height          =   255
         Left            =   0
         TabIndex        =   52
         Top             =   780
         Width           =   1095
      End
      Begin VB.Label lblTotalBytes 
         BackStyle       =   0  '����
         Caption         =   "-"
         Height          =   255
         Left            =   1320
         TabIndex        =   51
         Top             =   420
         Width           =   4335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '����
         Caption         =   "�� ũ��:"
         Height          =   255
         Left            =   0
         TabIndex        =   50
         Top             =   420
         Width           =   975
      End
   End
   Begin prjDownloadBooster.FrameW fThreadInfo 
      Height          =   3495
      Left            =   360
      TabIndex        =   47
      Tag             =   "forcebgchange"
      Top             =   2310
      Width           =   5775
      _ExtentX        =   0
      _ExtentY        =   0
      BorderStyle     =   0
      Transparent     =   -1  'True
      Begin prjDownloadBooster.FrameW pbProgressOuterContainer 
         Height          =   3495
         Left            =   0
         TabIndex        =   48
         Tag             =   "forcebgchange"
         Top             =   0
         Width           =   5775
         _ExtentX        =   0
         _ExtentY        =   0
         BorderStyle     =   0
         Transparent     =   -1  'True
         Begin prjDownloadBooster.FrameW pbProgressContainer 
            Height          =   9015
            Left            =   0
            TabIndex        =   64
            Tag             =   "forcebgchange"
            Top             =   0
            Width           =   5775
            _ExtentX        =   0
            _ExtentY        =   0
            BorderStyle     =   0
            Transparent     =   -1  'True
            Begin prjDownloadBooster.ProgressBar pbProgressMarquee 
               Height          =   255
               Index           =   1
               Left            =   960
               Top             =   0
               Visible         =   0   'False
               Width           =   3975
               _ExtentX        =   7011
               _ExtentY        =   450
               Step            =   10
               MarqueeSpeed    =   35
               Scrolling       =   2
            End
            Begin prjDownloadBooster.ProgressBar pbProgress 
               Height          =   255
               Index           =   1
               Left            =   960
               Top             =   0
               Visible         =   0   'False
               Width           =   3975
               _ExtentX        =   7011
               _ExtentY        =   450
               Step            =   10
               MarqueeSpeed    =   35
            End
            Begin VB.Label lblDownloader 
               AutoSize        =   -1  'True
               BackStyle       =   0  '����
               Caption         =   "������ 1:"
               Height          =   180
               Index           =   1
               Left            =   0
               TabIndex        =   66
               Top             =   45
               Width           =   750
            End
            Begin VB.Label lblPercentage 
               Alignment       =   1  '������ ����
               BackStyle       =   0  '����
               Height          =   255
               Index           =   1
               Left            =   5040
               TabIndex        =   65
               Top             =   45
               Width           =   615
            End
         End
      End
   End
   Begin prjDownloadBooster.CommandButtonW cmdIncreaseThreads 
      Height          =   330
      Left            =   6930
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   795
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   582
      ImageListAlignment=   4
      Caption         =   ">"
      Transparent     =   -1  'True
   End
   Begin prjDownloadBooster.CommandButtonW cmdDecreaseThreads 
      Height          =   330
      Left            =   1560
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   795
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   582
      ImageListAlignment=   4
      Caption         =   "<"
      Transparent     =   -1  'True
   End
   Begin VB.ComboBox cbWhenExist 
      Height          =   300
      Left            =   7830
      Style           =   2  '��Ӵٿ� ���
      TabIndex        =   23
      Top             =   2025
      Width           =   1185
   End
   Begin prjDownloadBooster.CommandButtonW cmdClear 
      Height          =   330
      Left            =   7350
      TabIndex        =   2
      Top             =   90
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   582
      ImageList       =   "imgErase"
      Caption         =   "�ʱ�ȭ(&Y) "
      Transparent     =   -1  'True
   End
   Begin prjDownloadBooster.CommandButtonW cmdAdd 
      Height          =   375
      Left            =   2520
      TabIndex        =   41
      Top             =   6960
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      ImageList       =   "imgPlusYellow"
      Caption         =   " �߰�(&R)..."
      Transparent     =   -1  'True
   End
   Begin prjDownloadBooster.CommandButtonW cmdBatch 
      Height          =   330
      Left            =   7200
      TabIndex        =   35
      Top             =   5565
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      ImageList       =   "imgDropdown"
      ImageListAlignment=   1
      Caption         =   "  �ϰ� ó��(&H)"
      Transparent     =   -1  'True
   End
   Begin prjDownloadBooster.FrameW fTotal 
      Height          =   615
      Left            =   240
      TabIndex        =   11
      Top             =   1320
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   1085
      Caption         =   " ��ü �ٿ�ε� ��Ȳ "
      Transparent     =   -1  'True
      Begin VB.Label lblState 
         BackStyle       =   0  '����
         Caption         =   "������"
         Height          =   255
         Left            =   120
         TabIndex        =   70
         Top             =   285
         Width           =   735
      End
   End
   Begin prjDownloadBooster.FrameW fOptions 
      Height          =   2220
      Left            =   6720
      TabIndex        =   19
      Top             =   1320
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   3916
      Caption         =   " �ɼ� "
      Transparent     =   -1  'True
      Begin prjDownloadBooster.CheckBoxW chkOpenAfterComplete 
         Height          =   255
         Left            =   90
         TabIndex        =   20
         Top             =   240
         Width           =   2175
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "�Ϸ� �� ����(&C)"
         Transparent     =   -1  'True
      End
      Begin prjDownloadBooster.CheckBoxW chkOpenFolder 
         Height          =   255
         Left            =   90
         TabIndex        =   21
         Top             =   465
         Width           =   2280
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "�Ϸ� �� ���� ����(&L)"
         Transparent     =   -1  'True
      End
      Begin prjDownloadBooster.CheckBoxW chkContinueDownload 
         Height          =   255
         Left            =   90
         TabIndex        =   24
         Top             =   990
         Width           =   1935
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "�׻� �̾�ޱ�(&J)"
         Transparent     =   -1  'True
      End
      Begin prjDownloadBooster.CheckBoxW chkAutoRetry 
         Height          =   255
         Left            =   90
         TabIndex        =   25
         Top             =   1215
         Width           =   2205
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "���� �� �ڵ� ��õ�(&G)"
         Transparent     =   -1  'True
      End
      Begin VB.Line lbOptionsHeader3D 
         BorderColor     =   &H80000014&
         Visible         =   0   'False
         X1              =   615
         X2              =   2415
         Y1              =   105
         Y2              =   105
      End
      Begin VB.Line lbOptionsHeader 
         BorderColor     =   &H80000010&
         Visible         =   0   'False
         X1              =   600
         X2              =   2400
         Y1              =   90
         Y2              =   90
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "."
         Height          =   180
         Left            =   0
         TabIndex        =   71
         Top             =   0
         Visible         =   0   'False
         Width           =   60
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "�ߺ�(&K):"
         Height          =   180
         Left            =   330
         TabIndex        =   22
         Top             =   765
         Width           =   690
      End
   End
   Begin prjDownloadBooster.CommandButtonW cmdOpenFolder 
      Height          =   330
      Left            =   7200
      TabIndex        =   31
      Top             =   4440
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      ImageList       =   "imgOpenFolder"
      Caption         =   "���� ����(&E) "
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
      _ExtentX        =   3149
      _ExtentY        =   582
      ImageList       =   "imgOpenFolder"
      Caption         =   " ã�ƺ���(&B)..."
      Transparent     =   -1  'True
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
      TabIndex        =   32
      Top             =   4815
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      ImageList       =   "imgDownload"
      Caption         =   "�ٿ�ε�(&D) "
      Transparent     =   -1  'True
   End
   Begin prjDownloadBooster.FrameW Frame4 
      Height          =   3885
      Left            =   240
      TabIndex        =   12
      Top             =   2040
      Width           =   6255
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "                  "
      Transparent     =   -1  'True
   End
   Begin prjDownloadBooster.CommandButtonW cmdOpenFileDropdown 
      Height          =   330
      Left            =   8880
      TabIndex        =   30
      Top             =   4065
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   582
      Enabled         =   0   'False
      ImageList       =   "imgDropdown"
      ImageListAlignment=   4
      Transparent     =   -1  'True
   End
   Begin VB.Image imgLBContentBackground 
      Height          =   4215
      Left            =   270
      Stretch         =   -1  'True
      Top             =   1635
      Visible         =   0   'False
      Width           =   6195
   End
   Begin VB.Image imgBorderBottomRight 
      Height          =   135
      Left            =   11040
      Picture         =   "frmMain.frx":538B
      Top             =   600
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image imgBorderBottomLeft 
      Height          =   135
      Left            =   10680
      Picture         =   "frmMain.frx":53D9
      Top             =   600
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image imgBorderTopRight 
      Height          =   135
      Left            =   11040
      Picture         =   "frmMain.frx":5425
      Top             =   240
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image imgBorderTopLeft 
      Height          =   135
      Left            =   10680
      Picture         =   "frmMain.frx":5471
      Top             =   240
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape pgBorderRight 
      BackColor       =   &H00C8D0D4&
      BackStyle       =   1  '�������� ����
      BorderColor     =   &H005F5F5C&
      Height          =   135
      Left            =   10680
      Top             =   1680
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Shape pgBorderLeft 
      BackColor       =   &H00C8D0D4&
      BackStyle       =   1  '�������� ����
      BorderColor     =   &H005F5F5C&
      Height          =   135
      Left            =   10680
      Top             =   1440
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Shape pgBorderBottom 
      BackColor       =   &H00C8D0D4&
      BackStyle       =   1  '�������� ����
      BorderColor     =   &H005F5F5C&
      Height          =   45
      Left            =   10680
      Top             =   1200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Shape pgBorderTop 
      BackColor       =   &H00C8D0D4&
      BackStyle       =   1  '�������� ����
      BorderColor     =   &H005F5F5C&
      Height          =   45
      Left            =   10680
      Top             =   960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblLBCaption 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "��    Ȳ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   240
      TabIndex        =   69
      Tag             =   "nosizechange"
      Top             =   1320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblLBCaptionShadow 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "��    Ȳ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   255
      TabIndex        =   75
      Tag             =   "nocolorsizechange"
      Top             =   1335
      Visible         =   0   'False
      Width           =   1215
   End
   Begin prjDownloadBooster.ShellPipe spYtdl 
      Left            =   9240
      Top             =   3360
      _ExtentX        =   635
      _ExtentY        =   635
   End
   Begin VB.Label lblThreadCount 
      BackStyle       =   0  '����
      Caption         =   "(������ 1��)"
      Height          =   255
      Left            =   7350
      TabIndex        =   10
      Top             =   870
      Width           =   1695
   End
   Begin VB.Label lblThreadCountLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "����(&T):"
      Height          =   180
      Left            =   240
      TabIndex        =   6
      Top             =   870
      Width           =   690
   End
   Begin VB.Label lblFilePath 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "���� ���(&F):"
      Height          =   180
      Left            =   240
      TabIndex        =   3
      Top             =   495
      Width           =   1095
   End
   Begin VB.Label lblURL 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "���� �ּ�(&A):"
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
      BackStyle       =   0  '����
      Caption         =   "���� �ּ�(&A):"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   255
      TabIndex        =   76
      Tag             =   "nocolorchange"
      Top             =   165
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.Label lblFilePathShadow 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "���� ���(&F):"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   255
      TabIndex        =   77
      Tag             =   "nocolorchange"
      Top             =   510
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblThreadCountLabelShadow 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "����(&T):"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   255
      TabIndex        =   78
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
      BorderStyle     =   0  '����
      Height          =   255
      Left            =   0
      Top             =   0
      Width           =   255
   End
   Begin VB.Menu mnuListContext 
      Caption         =   "mnuListContext"
      Visible         =   0   'False
      Begin VB.Menu mnuOpenBatch 
         Caption         =   "����(&O)"
      End
      Begin VB.Menu mnuOpenFolder2 
         Caption         =   "���� ����(&F)"
      End
      Begin VB.Menu mnuErrorInfo 
         Caption         =   "���� ����(&I)..."
      End
      Begin VB.Menu mnuSepOpen 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "����(&E)..."
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddItem2 
         Caption         =   "�� �ּ� �߰�(&A)..."
      End
      Begin VB.Menu mnuDeleteItem 
         Caption         =   "����(&R)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuClearBatch3 
         Caption         =   "��� ����(&C)"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMoveUp 
         Caption         =   "���� �̵�(&U)"
      End
      Begin VB.Menu mnuMoveDown 
         Caption         =   "�Ʒ��� �̵�(&D)"
      End
   End
   Begin VB.Menu mnuListContext2 
      Caption         =   "mnuListContext2"
      Visible         =   0   'False
      Begin VB.Menu mnuAddItem 
         Caption         =   "�� �ּ� �߰�(&A)..."
      End
      Begin VB.Menu mnuClearBatch2 
         Caption         =   "��� ����(&C)"
      End
   End
   Begin VB.Menu mnuDeleteDropdown 
      Caption         =   "mnuDeleteDropdown"
      Visible         =   0   'False
      Begin VB.Menu mnuClearBatch 
         Caption         =   "��� ����(&C)"
      End
   End
   Begin VB.Menu mnuOpenDropdown 
      Caption         =   "mnuOpenDropdown"
      Visible         =   0   'False
      Begin VB.Menu mnuOpenFolder 
         Caption         =   "���� ����(&F)"
      End
      Begin VB.Menu mnuPropertiesBatch 
         Caption         =   "�Ӽ� ����(&R)"
      End
   End
   Begin VB.Menu mnuOpenFileDropdown 
      Caption         =   "mnuOpenFileDropdown"
      Visible         =   0   'False
      Begin VB.Menu mnuProperties 
         Caption         =   "�Ӽ� ����(&R)"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements IBSSubclass

Dim Elapsed As Long
Dim BatchStarted As Boolean
Dim CurrentBatchIdx As Long
Dim DownloadPath As String
Dim IsDownloading As Boolean
Dim BatchErrorCount As Long
Dim BatchErrorAllCount As Long
Public ScrollOneScreen As Boolean
Dim PrevDownloadedBytes As Double
Dim SpeedCount As Byte
Dim HttpStatusCode As String
Dim ResumeUnsupported As Boolean
Public ImagePosition As Byte
Dim TotalSize As Double
Dim FormCaption$
Dim LBFrameEnabled As Boolean
Dim ErrorCodeDescription As Collection

Const MAIN_FORM_WIDTH As Long = 9450

#If HIDEYTDL Then
#Else
'youtube-dl ���� ����
Dim ytdlTotalFormatCount As Integer
Dim ytdlFileName As String
Public ytdlEnabled As Boolean
Public ytdlFormat As String
Public ytdlExtractAudio As Boolean
Public ytdlAudioFormat As AudioFormat
Public ytdlAudioBitrateType As AudioBitrateType
Public ytdlAudioCBR As Integer
Public ytdlAudioVBR As Byte
#End If

Dim MAX_THREAD_COUNT As Integer

Dim MaxLoadedTileBackgroundImage As Long

Dim FormWidth As Long
Dim FormMinHeight As Long
Dim FormMaxHeight As Long

Private Enum DownloadStopMode
    NormalStop = 1
    BatchStop = 2
    ExitApplication = 3
End Enum

#If HIDEYTDL Then
#Else
Sub StartYtdlDownload()
    If Not FileExists(GetSetting("DownloadBooster", "Options", "YtdlPath", "")) Then
        If MsgBox(t("youtube-dl ���� ���� ��ΰ� �������� �ʾҽ��ϴ�. ���� �����Ͻðڽ��ϱ�?", "youtube-dl executable path is not specified. Would you like to specify it now?"), vbQuestion + vbYesNo) = vbYes Then
            frmOptions.tsTabStrip.Tabs(5).Selected = -1
            frmOptions.Show vbModal, Me
        End If
        Exit Sub
    End If

    If lvLogTest.ColumnHeaders.Count < 2 Then
        lvLogTest.ColumnHeaders.Add , , "��ü", 1200
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

Private Function IBSSubclass_MsgResponse(ByVal hWnd As Long, ByVal uMsg As Long) As EMsgResponse
    IBSSubclass_MsgResponse = emrConsume
End Function

Private Sub IBSSubclass_UnsubclassIt()
    DetachMessage Me, Me.hWnd, WM_GETMINMAXINFO
    DetachMessage Me, Me.hWnd, WM_INITMENU
    DetachMessage Me, Me.hWnd, WM_SYSCOMMAND
    'DetachMessage Me, Me.hWnd, WM_DWMCOMPOSITIONCHANGED
    DetachMessage Me, Me.hWnd, WM_SETTINGCHANGE
    DetachMessage Me, Me.hWnd, WM_THEMECHANGED
    DetachMessage Me, Me.hWnd, WM_CTLCOLORSCROLLBAR
End Sub

Private Function IBSSubclass_WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByRef wParam As Long, ByRef lParam As Long, ByRef bConsume As Boolean) As Long
    On Error Resume Next
    
    Dim hSysMenu As Long
    Dim MII As MENUITEMINFO
    
    Select Case uMsg
        Case WM_GETMINMAXINFO
            Dim lpMMI As MINMAXINFO
            CopyMemory lpMMI, ByVal lParam, Len(lpMMI)
            lpMMI.ptMinTrackSize.X = FormWidth * (DPI / 96)
            lpMMI.ptMinTrackSize.Y = FormMinHeight * (DPI / 96)
            lpMMI.ptMaxTrackSize.X = lpMMI.ptMinTrackSize.X
            lpMMI.ptMaxTrackSize.Y = FormMaxHeight * (DPI / 96)
            CopyMemory ByVal lParam, lpMMI, Len(lpMMI)
            
            IBSSubclass_WindowProc = 1&
            Exit Function
        Case WM_INITMENU
            hSysMenu = GetSystemMenu(Me.hWnd, 0)
            With MII
                .cbSize = Len(MII)
                .fMask = MIIM_STATE
                .fState = MFS_ENABLED Or IIf(MainFormOnTop, MFS_CHECKED, 0)
            End With
            SetMenuItemInfo hSysMenu, 1000, 0, MII
            
            IBSSubclass_WindowProc = 1&
            Exit Function
        Case WM_SYSCOMMAND
            If wParam = 1000 Then '�׻� ���� ǥ��
                MainFormOnTop = Not MainFormOnTop
                SetWindowPos hWnd, IIf(MainFormOnTop, hWnd_TOPMOST, hWnd_NOTOPMOST), 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
                SaveSetting "DownloadBooster", "Options", "AlwaysOnTop", -(MainFormOnTop)
                
                IBSSubclass_WindowProc = 1&
                Exit Function
            ElseIf wParam = 1003 And (Not (frmMain.Height <= 6930 + PaddedBorderWidth * 15 * 2)) Then 'â ũ�� �ʱ�ȭ
                Me.Height = 8985 + PaddedBorderWidth * 15 * 2
            
                IBSSubclass_WindowProc = 1&
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
                        If TypeOf ctrl Is FrameW Or TypeOf ctrl Is CheckBoxW Or TypeOf ctrl Is CommandButtonW Or TypeOf ctrl Is Slider Then ctrl.Refresh
                    Next ctrl
                    trThreadCount.VisualStyles = Not trThreadCount.VisualStyles
                    trThreadCount.VisualStyles = Not trThreadCount.VisualStyles
                    
                    SetTextColors
            End Select
        Case WM_THEMECHANGED
            SetTextColors
        Case WM_CTLCOLORSCROLLBAR
            IBSSubclass_WindowProc = 0&
            Exit Function
    End Select
    
    IBSSubclass_WindowProc = CallOldWindowProc(hWnd, uMsg, wParam, lParam)
End Function

Private Sub mnuErrorInfo_Click()
    If lvBatchFiles.SelectedItem Is Nothing Then Exit Sub
    If lvBatchFiles.SelectedItem.ForeColor <> vbRed Then Exit Sub
    Dim StatusString$
    StatusString = lvBatchFiles.SelectedItem.ListSubItems(3).Text
    StatusString = Mid(StatusString, InStr(StatusString, "(") + 1)
    StatusString = Left$(StatusString, Len(StatusString) - 1)
    If Not IsNumeric(StatusString) Then
        'MsgBox t("���� ������ ǥ���� �� �����ϴ�.", "Unable to show the error information."), 16
        MsgBox t("�ٿ�ε尡 �����Ǿ����ϴ�.", "Download was aborted."), 64
        Exit Sub
    End If
    MsgBox t("���� �ڵ�", "Error code") & ": " & StatusString & vbCrLf & t("����", "Description") & ": " & IIf(Exists(ErrorCodeDescription, StatusString), ErrorCodeDescription(StatusString), t("������ �����ϴ�.", "Description is unavailable")), 64, t("���� ����", "Error information")
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
                    '���� ����
                    If CBool(InStr(Line, "Downloading ")) And CBool(InStr(Line, " format(s): ")) Then
                        ytdlTotalFormatCount = UBound(Split(Data(5), "+")) + 1
                    End If
                    lvLogTest.ListItems.Add(, , "����").ListSubItems.Add , , Line
                Case "[download]"
                    If UBound(Data) > 1 Then
                        If Data(1) = "Destination:" Then
                            ytdlFileName = Replace(Line, "[download] Destination: ", "")
                        End If
                    End If
                    lvLogTest.ListItems.Add(, , "�ٿ�ε�").ListSubItems.Add , , Line
                Case "[Merger]"
                    If UBound(Data) > 3 Then
                        If Data(1) = "Merging" And Data(3) = "into" Then
                            ytdlFileName = Replace(Line, "[Merger] Merging formats into ", "")
                            If Left$(ytdlFileName, 1) = """" And Right$(ytdlFileName, 1) = """" Then
                                ytdlFileName = Mid$(ytdlFileName, 2, Len(ytdlFileName) - 2)
                            End If
                        End If
                    End If
                    lvLogTest.ListItems.Add(, , "��ü").ListSubItems.Add , , Line
                Case Else
                    lvLogTest.ListItems.Add(, , Data(0)).ListSubItems.Add , , Line
            End Select
        
nextLine:
        Next k
    Next i
End Sub
#End If

Sub OnData(Data As String)
    'If Left$(Data, 6) = "DEBUG " Then Debug.Print Data
    Dim output$
    Dim idx%
    Dim progress%
    Dim DownloadedBytes As Double
    If Left$(Data, 7) = "STATUS " Then
        Select Case Replace(Right$(Data, Len(Data) - 7), " ", "")
            Case "CHECKREDIRECT"
                sbStatusBar.Panels(1).Text = t("������ ã�� ��...", "Finding server...")
            Case "CHECKFILE"
                sbStatusBar.Panels(1).Text = t("���뼺 Ȯ�� ��...", "Checking availability...")
            Case "DOWNLOADING"
                sbStatusBar.Panels(1).Text = t("�ٿ�ε� ��...", "Downloading...")
            Case "MERGING"
                sbStatusBar.Panels(1).Text = t("���� ���� ���� ��...", "Merging segments...")
                pbTotalProgressMarquee.Visible = -1
                pbTotalProgressMarquee.MarqueeAnimation = -1
                cmdStop.Enabled = 0
            Case "COMPLETE"
                sbStatusBar.Panels(1).Text = t("�Ϸ�", "Complete")
                sbStatusBar.Panels(2).Text = ""
                sbStatusBar.Panels(3).Text = ""
                sbStatusBar.Panels(4).Text = ""
                pbTotalProgressMarquee.MarqueeAnimation = 0
                pbTotalProgressMarquee.Visible = 0
                pbTotalProgress.Value = 100
                OnExit 0&
            Case "UNABLETOCONTINUE"
                ShowMessageBox t("�̾�ޱⰡ �Ұ����մϴ�. ó������ �ٽ� �ٿ�ε��մϴ�.", "Unable to resume. Starting over..."), App.Title, 48, False, 5000, MsgBoxMode:=1
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
        
        If trThreadCount.Value > 1 And idx = 1 And (CDbl(Split(output, ",")(2)) > 0 Or lblTotalBytes.Caption = "0 ����Ʈ") Then lblTotalSizeThread.Caption = ParseSize(CDbl(Split(output, ",")(2)), True)
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
            If fTotal.Caption <> t(" ��ü �ٿ�ε� ��Ȳ ", " Total Progress ") Then
                fTotal.Caption = t(" ��ü �ٿ�ε� ��Ȳ ", " Total Progress ")
                lblProgress = ""
            End If
            If pbTotalProgress.Value <> 0 Then pbTotalProgress.Value = 0
            If DownloadedBytes = -1 Then
                sbStatusBar.Panels(2).Text = ""
            ElseIf total <= 0 Then
                sbStatusBar.Panels(2).Text = ParseSize(DownloadedBytes)
            Else
                sbStatusBar.Panels(2).Text = t(ParseSize(total) & " �� " & ParseSize(DownloadedBytes), ParseSize(DownloadedBytes) & " of " & ParseSize(total))
            End If
            If DownloadedBytes <> -1 Then timElapsed.Enabled = -1
            If total <= 0 Then
                If lblTotalBytes.Caption <> t("�� �� ����", "Unknown") Then lblTotalBytes.Caption = t("�� �� ����", "Unknown")
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
                sbStatusBar.Panels(2).Text = t(ParseSize(strTotal) & " �� " & ParseSize(DownloadedBytes), ParseSize(DownloadedBytes) & " of " & ParseSize(strTotal))
            End If
            If strTotal = "NaN" Or strTotal = "-1" Then
                lblTotalBytes.Caption = t("�� �� ����", "Unknown")
            Else
                lblTotalBytes.Caption = ParseSize(total, True)
                TotalSize = total
            End If
            lblDownloadedBytes.Caption = ParseSize(DownloadedBytes, True)
            pbTotalProgress.Value = progress
            fTotal.Caption = t(" ��ü �ٿ�ε� ��Ȳ (" & progress & "%) ", " Total Progress (" & progress & "%) ")
            lblProgress = "(" & progress & "%)"
            If Not BatchStarted Then SetTitle progress & "% " & t("�ٿ�ε� ��", "Downloading")
        End If
        
        Dim Speed As Double
        SpeedCount = SpeedCount + 1
        If SpeedCount >= 10 Then
            Speed = (DownloadedBytes - PrevDownloadedBytes)
            lblSpeed.Caption = ParseSize(Speed, True, "/" & t("��", "sec"))
            sbStatusBar.Panels(3).Text = ParseSize(Speed, False, "/" & t("��", "sec"))
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
    ElseIf Left$(Data, 16) = "SETMODIFIEDDATE " Then
        On Error Resume Next
        SetFileDate DownloadPath, Right$(Data, Len(Data) - 16)
    End If
End Sub

Sub NextBatchDownload()
    If Not BatchStarted Then Exit Sub
    Dim i&
    
    If lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(3).Text = t("�Ϸ�", "Done") Then _
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
        
        If lvBatchFiles.ListItems.Count Then
            Dim Enable As Boolean
            For i = 1& To lvBatchFiles.ListItems.Count
                If lvBatchFiles.ListItems(i).Checked Then
                    Enable = True
                    Exit For
                End If
            Next i
            cmdStartBatch.Enabled = Enable
        Else
            cmdStartBatch.Enabled = 0
        End If
        
        If BatchErrorCount Then
            MsgBox t("�ϳ� �̻��� ������ �߻��߽��ϴ�. �ش� �׸��� �� �� ������ ���� ������ �� �� �ֽ��ϴ�.", "One or more errors have occurred. Double click the error item to see details."), 48
        ElseIf GetSetting("DownloadBooster", "Options", "PlaySound", 1) <> 0 And BatchErrorAllCount <= 0 Then
            PlayWave Trim$(GetSetting("DownloadBooster", "Options", "CompleteSoundPath", "")), FallbackSound:=vbInformation
        End If
        
        If lblState.Caption = t("�Ϸ��", "Done") Then
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
            Case 0, -1
                '���� ����
                GoTo nextln
            Case 999
                GoTo nextln
            Case 1
                If chkAutoRetry.Value <> 1 Then
                    If pbTotalProgressMarquee.Visible And (lblDownloadedBytes.Caption = "-" Or lblDownloadedBytes.Caption = "��� ��...") Then
                        MsgBox t("�ش� ���� �ּҿ� ������ �� �����ϴ�. �ּҰ� ��ȿ���� �ʰų� ������ �������� �ʽ��ϴ�.", "The server does not respond or the file URL is invalid."), 16
                    Else
                        MsgBox t("�������� ������ ������ϴ�. �ٿ�ε� ���߿� ��Ʈ��ũ ������ �߻����� �� �ֽ��ϴ�.", "Network error while downloading."), 16
                    End If
                End If
            Case 102
                MsgBox "�ּҳ� ���� �̸��� �������� �ʾҽ��ϴ�.", 16
            Case 3, 103
                MsgBox t("���� ��ΰ� �������� �ʽ��ϴ�.", "Save path doesn't exist."), 16
            Case 104
                MsgBox t("������ ���ϸ��� ��� ���Դϴ�. �ٸ� �̸��� �����Ͻʽÿ�.", "File name already exists."), 16
            Case 106
                MsgBox t("���� ������ �ٿ�ε� �ν�Ʈ�� �������� �ʽ��ϴ�. ������ 1�� ������ ���ʽÿ�.", "Download boosting not supported. Try changing the thread count to 1."), 16
            Case 107
                MsgBox t("������ ũ�⸦ �� �� ��� �ٿ�ε带 �ν�Ʈ�� �� �����ϴ�. ������ 1�� ������ ���ʽÿ�.", "Unable to boost download because the file size is not provided. Try changing the thread count to 1."), 16
            Case 108
                Dim statusMsg As String
                statusMsg = ""
                Dim ErrDesc As String
                Dim Icon As VbMsgBoxStyle
                Icon = vbCritical
                If Len(HttpStatusCode) > 0 And LangID = 1042 Then
                    Select Case HttpStatusCode
                        Case "400"
                            ErrDesc = "��û�� �߸��Ǿ����ϴ�."
                        Case "401"
                            ErrDesc = "�����Ϸ��� ���� ������ �ʿ��մϴ�."
                        Case "402"
                            ErrDesc = "�����Ϸ��� ������ �ʿ��մϴ�."
                        Case "403"
                            ErrDesc = "���� ������ �����ϴ�."
                        Case "404"
                            ErrDesc = "������ ������ �������� �ʽ��ϴ�."
                        Case "405"
                            ErrDesc = "������ �������� ����Ÿ�� �����ؾ� �մϴ�."
                        Case "406"
                            ErrDesc = "��û�� �޾Ƶ��� �� �����ϴ�."
                        Case "407"
                            ErrDesc = "���Ͻ� ������ �ʿ��մϴ�."
                        Case "408"
                            ErrDesc = "��û�� ���ð� �ȿ� ���������� �ʾҽ��ϴ�."
                        Case "409"
                            ErrDesc = "��û�� ������ �浹�߽��ϴ�."
                        Case "410"
                            If Month(Now) = 4 And Day(Now) = 1 Then
                                ErrDesc = "������ �־��µ� �������ϴ�."
                                Icon = vbInformation
                            Else
                                ErrDesc = "������ �� �̻� ������ �����ϴ�."
                            End If
                        Case "414"
                            ErrDesc = "�ּҰ� �ʹ� ��ϴ�."
                        Case "418"
                            ErrDesc = "������ �ڽ��� �������ڶ� Ŀ�Ǹ� ���� �� ���ٰ� �մϴ� ��   "
                            Icon = vbInformation
                        Case "451"
                            ErrDesc = "������ ������ ������ �ٿ�ε� ���� �� �����ϴ�."
                        Case "500"
                            ErrDesc = "���� ������ ������ �߻��߽��ϴ�."
                        Case "502"
                            ErrDesc = "����Ʈ���̰� �ҷ��Դϴ�."
                        Case "503"
                            ErrDesc = "������ �Ͻ������� ������ �� ���� �����Դϴ�."
                        Case "504"
                            ErrDesc = "����Ʈ���� �ð��� �ʰ��Ǿ����ϴ�."
                        Case "505"
                            ErrDesc = "HTTP ������ �������� �ʽ��ϴ�."
                        Case Else
                            ErrDesc = "���� �� �����̰ų� �������� �������� �ʰų� ���� ������ ���� �� �ֽ��ϴ�."
                            statusMsg = " HTTP ���� �ڵ�� ( " & HttpStatusCode & " ) �Դϴ�."
                    End Select
                End If
                MsgBox t("������ ��û�� �ź��߽��ϴ�. " & ErrDesc & statusMsg, "Server denied your request. The file may not exist or have insufficient permissions to access it."), Icon
            Case 109
                MsgBox t("����� ������ �̾�ޱⰡ �Ұ����մϴ�.", "Unable to resume due to technical issues."), 16
            Case 110
                MsgBox t("��ũ ������ �����մϴ�.", "Low disk space."), 16
            Case 111
                MsgBox t("��ũ ���� �����Դϴ�. �Ͻ����� �����̰ų� ��ũ�� �б� �����̰ų� ��ũ�� ���°� 452��ó�� ���� ���� �� �ֽ��ϴ�.", "Disk write error. It may be just a temporary error, or disk is in an unhealthy state or disk is write protected."), 16
            Case Else
                MsgBox t("���� ������ �߻��߽��ϴ�. ���μ��� ��ȯ ���� ( " & RetVal & " ) �Դϴ�.", "Internal error. Process returned ( " & RetVal & " )."), 16
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
            lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(3).Text = t("����", "Error") & " (" & RetVal & ")"
            lvBatchFiles.ListItems(CurrentBatchIdx).ForeColor = 255
            lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(1).ForeColor = 255
            lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(2).ForeColor = 255
            lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(3).ForeColor = 255
            If RetVal <> 999& Then BatchErrorCount = BatchErrorCount + 1
            BatchErrorAllCount = BatchErrorAllCount + 1
        Else
            lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(3).Text = t("�Ϸ�", "Done")
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
    
    lblTotalBytes.Caption = t("��� ��...", "Pending...")
    lblDownloadedBytes.Caption = t("��� ��...", "Pending...")
    If trThreadCount.Value > 1 Then
        lblTotalSizeThread.Caption = t("��� ��...", "Pending...")
        lblThreadCount2.Caption = trThreadCount.Value
    Else
        lblTotalSizeThread.Caption = "-"
        lblThreadCount2.Caption = "-"
    End If
    lblElapsed.Caption = "0" & t("��", " seconds")
    lblSpeed.Caption = "-"
    lblRemaining.Caption = "-"
    
    fTotal.Caption = t(" ��ü �ٿ�ε� ��Ȳ ", " Total Progress ")
    lblProgress = ""
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
    
    lblState.Caption = t("���� ��", "Working")
    lblLBState = lblState
    sbStatusBar.Panels(1).Text = t("���� ��...", "Starting...")
    
    If BatchStarted Then
'        Dim BatchCount%
'        BatchCount = 0
'        For i = 1 To lvBatchFiles.ListItems.Count
'            BatchCount = BatchCount + -(lvBatchFiles.ListItems(i).Checked)
'        Next i
        SetTitle t(lvBatchFiles.ListItems.Count & "�� �� " & CurrentBatchIdx & "��° �׸� �ٿ�ε� ��", "Downloading " & CurrentBatchIdx & " of " & lvBatchFiles.ListItems.Count)
    Else
        SetTitle t("�ٿ�ε� ��", "Downloading")
    End If
End Sub

Sub OnStop(Optional PlayBeep As Boolean = True)
    SP.ClosePipe
    
    IsDownloading = False
    cmdGo.Enabled = Not BatchStarted
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
    cmdDecreaseThreads.Enabled = (trThreadCount.Value > trThreadCount.Min)
    cmdIncreaseThreads.Enabled = (trThreadCount.Value < trThreadCount.Max)
    cmdDownloadOptions.Enabled = -1
    lblThreadCount.Enabled = -1
    
    SP.FinishChild 0, 0
    
    Dim i&
    For i = 1 To trThreadCount.Value
        pbProgressMarquee(i).MarqueeAnimation = 0
        pbProgressMarquee(i).Visible = 0
    Next i
    
    If pbTotalProgressMarquee.Visible Then
        pbTotalProgressMarquee.MarqueeAnimation = 0
        pbTotalProgressMarquee.Visible = 0
    End If
    
    If pbTotalProgress.Value < 100 Then pbTotalProgress.Value = 0
    
    If pbTotalProgress.Value < 100 Then
        lblState.Caption = t("������", "Stopped")
        lblLBState = lblState
        sbStatusBar.Panels(1).Text = t("�غ�", "Ready")
    
        fTotal.Caption = t(" ��ü �ٿ�ε� ��Ȳ ", " Total Progress ")
        lblProgress = ""
        For i = 1 To lblDownloader.UBound
            pbProgress(i).Value = 0
            lblPercentage(i).Caption = ""
        Next i
    Else
        lblState.Caption = t("�Ϸ��", "Done")
        lblLBState = lblState
        sbStatusBar.Panels(1).Text = t("�Ϸ�", "Done")
        sbStatusBar.Panels(2).Text = ""
        sbStatusBar.Panels(3).Text = ""
        sbStatusBar.Panels(4).Text = ""
    End If
    
    If Not BatchStarted Then
        timElapsed.Enabled = 0
        sbStatusBar.Panels(3).Text = ""
        sbStatusBar.Panels(4).Text = ""
        
        If lvBatchFiles.ListItems.Count Then
            Dim Enable As Boolean
            For i = 1& To lvBatchFiles.ListItems.Count
                If lvBatchFiles.ListItems(i).Checked Then
                    Enable = True
                    Exit For
                End If
            Next i
            cmdStartBatch.Enabled = Enable
        Else
            cmdStartBatch.Enabled = 0
        End If
        
        If PlayBeep And GetSetting("DownloadBooster", "Options", "PlaySound", 1) <> 0 Then
            PlayWave Trim$(GetSetting("DownloadBooster", "Options", "CompleteSoundPath", "")), FallbackSound:=vbInformation
            lblState.Caption = t("�Ϸ��", "Done")
            lblLBState = lblState
            sbStatusBar.Panels(1).Text = t("�Ϸ�", "Done")
            sbStatusBar.Panels(2).Text = ""
        End If
    End If
    
    If lblTotalBytes.Caption = t("��� ��...", "Pending...") Then lblTotalBytes.Caption = "-"
    If lblDownloadedBytes.Caption = t("��� ��...", "Pending...") Then lblDownloadedBytes.Caption = "-"
    If PlayBeep And lblDownloadedBytes.Caption <> "-" Then lblTotalBytes.Caption = lblDownloadedBytes.Caption
    If lblTotalSizeThread.Caption = t("��� ��...", "Pending...") Then lblTotalSizeThread.Caption = "-"
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
    If LenB(Replace(txtURL.Text, " ", "")) Then
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
        MsgBox URL & " - " & t("�ּҰ� �ùٸ��� �ʽ��ϴ�. 'http://' �Ǵ� 'https://'�� �����ؾ� �մϴ�.", "Invalid address. Must start with 'http://' or 'https://'."), 16
        AddBatchURLs = False
        Exit Function
    End If
    
    On Error GoTo justadd
    If GetSetting("DownloadBooster", "Options", "AllowDuplicatesInQueue", 0) <> 0 Then GoTo justadd
    Dim i&
    If lvBatchFiles.ListItems.Count Then
        For i = 1& To lvBatchFiles.ListItems.Count
            If lvBatchFiles.ListItems(i).ListSubItems(2).Text = Trim$(URL) Then
                AddBatchURLs = True
                Exit Function
            End If
        Next i
    End If
justadd:
    On Error GoTo 0
    
    If Headers = "-" Then Headers = SessionHeaderCache
    
    If LenB(Trim$(SavePath)) = 0 Then SavePath = txtFileName.Text
    SavePath = Trim$(SavePath)
    Do While Replace(SavePath, "\\", "\") <> SavePath
        SavePath = Replace(SavePath, "\\", "\")
    Loop

    Dim idx&
    Dim FileName$
    Dim ServerName$
    FileName = SavePath
    If FolderExists(FileName) Then
        If Not (Right$(FileName, 1) = "\") Then FileName = FileName & "\"
        ServerName = FilterFilename(ExcludeParameters(URLDecode(Split(URL, "/")(UBound(Split(URL, "/"))))))
        If LenB(Replace(ServerName, " ", "")) = 0 Then ServerName = "download_" & CStr(Rnd * 1E+15)
        FileName = FileName & ServerName
    Else
        ServerName = FilterFilename(ExcludeParameters(URLDecode(Split(URL, "/")(UBound(Split(URL, "/"))))))
        If LenB(Replace(ServerName, " ", "")) = 0 Then
            ServerName = "download_" & CStr(Rnd * 1E+15)
        Else
            ServerName = CStr(Rnd * 1E+15) & "_" & ServerName
        End If
        FileName = GetParentFolderName(txtFileName.Text) & "\"
        FileName = Replace(FileName, "\\", "\") & ServerName
    End If
    idx = lvBatchFiles.ListItems.Add(Text:=ServerName).Index
    lvBatchFiles.ListItems(idx).ListSubItems.Add , , FileName
    lvBatchFiles.ListItems(idx).ListSubItems.Add , , URL
    lvBatchFiles.ListItems(idx).ListSubItems.Add , , t("���", "Queued")
    lvBatchFiles.ListItems(idx).ListSubItems.Add , , "Y"
    lvBatchFiles.ListItems(idx).ListSubItems.Add , , Headers
#If HIDEYTDL Then
#Else
    lvBatchFiles.ListItems(idx).ListSubItems.Add , , "N"
    lvBatchFiles.ListItems(idx).ListSubItems.Add , , ""
    lvBatchFiles.ListItems(idx).ListSubItems.Add , , "N"
    lvBatchFiles.ListItems(idx).ListSubItems.Add , , ""
    lvBatchFiles.ListItems(idx).ListSubItems.Add , , ""
    lvBatchFiles.ListItems(idx).ListSubItems.Add , , ""
    lvBatchFiles.ListItems(idx).ListSubItems.Add , , ""
#End If
    lvBatchFiles.ListItems(idx).Checked = -1
    cmdStartBatch.Enabled = Not (IsDownloading Or cmdStop.Enabled Or BatchStarted)
    AddBatchURLs = True
End Function

Private Sub cmdAddToQueue_Click()
    If lvBatchFiles.ListItems.Count >= MAX_32BIT_SIGNED_INT Then
        MsgBox t("�ִ� �ϰ� �ٿ�ε� ������ �ʰ��߽��ϴ�.", "Maximum number of items exceeded."), vbExclamation
        Exit Sub
    End If
    If LenB(Replace(txtURL.Text, " ", "")) = 0 Then
        MsgBox t("���� �ּҸ� �Է��Ͻʽÿ�.", "Specify the file URL."), 64
        Exit Sub
    End If
    On Error GoTo justadd
    If GetSetting("DownloadBooster", "Options", "AllowDuplicatesInQueue", 0) <> 0 Then GoTo justadd
    Dim i&
    If lvBatchFiles.ListItems.Count Then
        For i = 1& To lvBatchFiles.ListItems.Count
            If lvBatchFiles.ListItems(i).ListSubItems(2).Text = Trim$(txtURL.Text) Then
                MsgBox t("�ش� �ּҴ� �̹� ��⿭�� �߰��Ǿ����ϴ�.", "That URL is already added"), 64
                Exit Sub
            End If
        Next i
    End If
justadd:
    AddBatchURLs txtURL.Text, , "-"
End Sub

Sub cmdBatch_Click()
    On Error Resume Next
    
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
    End If
    SetBackgroundPosition
End Sub

Private Sub cmdBrowse_Click()
    ShowFileDialog
End Sub

Private Sub cmdClear_Click()
    txtURL.Text = ""
End Sub

Private Sub cmdDecreaseThreads_Click()
    If trThreadCount.Value > trThreadCount.Min Then trThreadCount.Value = trThreadCount.Value - 1
End Sub

Private Sub cmdDelete_Click()
    If BatchStarted And CurrentBatchIdx = lvBatchFiles.SelectedItem.Index Then Exit Sub

    If BatchStarted And CurrentBatchIdx > lvBatchFiles.SelectedItem.Index Then CurrentBatchIdx = CurrentBatchIdx - 1
    lvBatchFiles.ListItems.Remove lvBatchFiles.SelectedItem.Index
    If lvBatchFiles.ListItems.Count < 1 Or cmdStop.Enabled Or BatchStarted Then
        cmdStartBatch.Enabled = 0
        Exit Sub
    End If
    
    Dim i&
    Dim Enable As Boolean
    For i = 1& To lvBatchFiles.ListItems.Count
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

Sub StartDownload(ByVal URL As String, ByVal FileName As String, Optional NeedsDecoding As Boolean = False)
    If BatchStarted Then
        If Not lvBatchFiles.ListItems(CurrentBatchIdx).Checked Then
            lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(3).Text = t("���", "Skip")
            lvBatchFiles.ListItems(CurrentBatchIdx).ForeColor = &H808080
            lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(1).ForeColor = &H808080
            lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(2).ForeColor = &H808080
            lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(3).ForeColor = &H808080
            NextBatchDownload
            Exit Sub
        End If
        
        If lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(3).Text = t("�Ϸ�", "Done") Then
            NextBatchDownload
            Exit Sub
        End If
    
        lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(3).Text = t("���� ��...", "Working...")
        lvBatchFiles.ListItems(CurrentBatchIdx).ForeColor = &HFF0000
        lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(1).ForeColor = &HFF0000
        lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(2).ForeColor = &HFF0000
        lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(3).ForeColor = &HFF0000
        
        On Error GoTo L1
        If lvBatchFiles.SelectedItem.Index = CurrentBatchIdx Then
            cmdDelete.Enabled = 0
            cmdDeleteDropdown.Enabled = 0
            cmdEdit.Enabled = 0
        ElseIf LenB(lvBatchFiles.SelectedItem.Text) And lvBatchFiles.SelectedItem.Selected Then
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
        If LenB(Trim$(SplittedPath(i))) And LenB(Replace(Trim$(SplittedPath(i)), ".", "")) = 0 Then
            MsgBox t("���� ��ΰ� ��ȿ���� �ʽ��ϴ�.", "Invalid save path."), 16
            OnExit 999
            Exit Sub
        End If
    Next i
    
    If (Not FolderExists(Trim$(FileName))) And ((Not FolderExists(GetParentFolderName(Trim$(FileName)))) Or Right$(FileName, 1) = "\") Then
        MsgBox t("���� ��ΰ� �������� �ʽ��ϴ�.", "Save path does not exist."), 16
        OnExit 999
        Exit Sub
    End If
    
    If LenB(Replace(FileName, " ", "")) = 0 Then
        FileName = Replace(CurDir() & "\", "\\", "\")
    End If
    Dim ServerName$
    Dim AutoName As Boolean
    AutoName = False
    If FolderExists(FileName) Then
        If Not (Right$(FileName, 1) = "\") Then FileName = FileName & "\"
        ServerName = FilterFilename(ExcludeParameters(URLDecode(Split(URL, "/")(UBound(Split(URL, "/"))), Not NeedsDecoding)))
        If LenB(Replace(ServerName, " ", "")) = 0 Then ServerName = "download_" & CStr(Rnd * 1E+15)
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
        PrevPartialDownload = FileExists(FileName & ".part.tmp")
        If PrevPartialDownload Then
            Dim ContinueMsgboxResult As VbMsgBoxResult
            ContinueMsgboxResult = MsgBox(t("������ �ٿ�ε� �޴ٰ� ������ �����Դϴ�. �ٿ�ε���� �������� �̾ �����ðڽ��ϱ�?" & vbCrLf & "��[�ƴϿ�]�� ���� ��� ó������ �ٽ� �ٿ�ε�˴ϴ�.", "This file was previously downloaded partially. Would you like to resume?" & vbCrLf & "  We will download from the start if you choose No."), vbQuestion + vbYesNoCancel)
            If ContinueMsgboxResult = vbYes Then
                ContinueDownload = 1
            ElseIf ContinueMsgboxResult = vbCancel Then
                OnExit 999
                Exit Sub
            End If
        End If
    End If
    
    Dim CurrentHeaderCache$
    If BatchStarted Then
        CurrentHeaderCache = lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(5).Text
    Else
        CurrentHeaderCache = Functions.SessionHeaderCache
    End If
    Dim SPResult As SP_RESULTS
    SPResult = RunNode(SP, DownloadScript, """" & _
        Replace(Replace(URL, " ", "%20"), """", "%22") & """ """ & _
        FileName & """ " & _
        trThreadCount.Value & " " & _
        "0 " & _
        cbWhenExist.ListIndex & " " & _
        ContinueDownload & " " & _
        GetSetting("DownloadBooster", "Options", "NoRedirectCheck", 0) & " " & _
        GetSetting("DownloadBooster", "Options", "ForceGet", 1) & " " & _
        GetSetting("DownloadBooster", "Options", "Ignore300", 0) & " " & _
        (-AutoName) & " " & _
        GetSetting("DownloadBooster", "Options", "ThreadRequestInterval", 100) & " " & _
        GetSetting("DownloadBooster", "Options", "UseServerModifiedDate", 1) & " " & _
        Replace(Col(Functions.HeaderCache, "-"), vbCrLf, "") & " " & _
        Replace(Col(CurrentHeaderCache, "-"), vbCrLf, "") _
    )
    'EnvironmentVariables:=StrPtr("UV_THREADPOOL_SIZE=" & (trThreadCount.Value + 1) & vbNullChar & vbNullChar)
    Select Case SPResult
        Case SP_SUCCESS
            'SP.ClosePipe
            Exit Sub
        Case SP_CREATEPIPEFAILED
            MsgBox t("�ٿ�ε� ���ۿ� �����߽��ϴ�. �ٿ�δ� ���μ����κ��� ������ �޾ƿ� �� �����ϴ�. ���丮 �������� �ùٸ� ���α׷��� �����ߴ��� Ȯ���Ͻʽÿ�.", "Failed to receieve data from the downloader process. Check if the directory settings are valid."), 16
        Case SP_CREATEPROCFAILED
            MsgBox t("�ٿ�ε� ���ۿ� �����߽��ϴ�. �ٿ�δ� ���μ����� ������ �� �����ϴ�. ���丮 �������� �ùٸ� ���α׷��� �����ߴ��� Ȯ���Ͻʽÿ�.", "Failed to create the downloader process. Check if the directory settings are valid."), 16
    End Select
    
    If Not BatchStarted Then cmdGo.Enabled = -1
    cmdStop.Enabled = 0
    cmdStop.Left = Me.Width + 1200
    cmdGo.Enabled = -1
    cmdGo.Visible = -1
    OnStop False
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
    
    If LenB(Replace(txtURL.Text, " ", "")) = 0 Then
        MsgBox t("���� �ּҸ� �Է��Ͻʽÿ�.", "Specify the file URL."), 64
        Exit Sub
    End If
    
    If Left$(txtURL.Text, 7) <> "http://" And Left$(txtURL.Text, 8) <> "https://" Then
        MsgBox t("�ּҰ� �ùٸ��� �ʽ��ϴ�. 'http://' �Ǵ� 'https://'�� �����ؾ� �մϴ�.", "Invalid address. Must start with 'http://' or 'https://'."), 16
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
        If LenB(Trim$(SplittedPath(i))) And LenB(Replace(Trim$(SplittedPath(i)), ".", "")) = 0 Then
            MsgBox t("���� ��ΰ� ��ȿ���� �ʽ��ϴ�.", "Invalid save path."), 16
            Exit Sub
        End If
    Next i
    
    If (Not FolderExists(Trim$(txtFileName.Text))) And ((Not FolderExists(GetParentFolderName(Trim$(txtFileName.Text)))) Or Right$(txtFileName.Text, 1) = "\") Then
        MsgBox t("���� ��ΰ� �������� �ʽ��ϴ�.", "Save path does not exist."), 16
        Exit Sub
    End If
    
    txtURL.Text = Trim$(txtURL.Text)

    Elapsed = 0
    If GetSetting("DownloadBooster", "Options", "LazyElapsed", "0") <> "1" Then timElapsed.Enabled = -1
    StartDownload txtURL.Text, txtFileName.Text, True
End Sub

Private Sub cmdIncreaseThreads_Click()
    If trThreadCount.Value < trThreadCount.Max Then trThreadCount.Value = trThreadCount.Value + 1
End Sub

Private Sub cmdOpen_Click()
    ShellExecute DownloadPath
End Sub

Private Sub cmdOpen_DropDown()
    cmdOpenFileDropdown_Click
End Sub

Private Sub cmdOpenBatch_Click()
    On Error Resume Next
    ShellExecute lvBatchFiles.SelectedItem.ListSubItems(1).Text
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
    If LenB(DownloadPath) Then pth = DownloadPath Else pth = txtFileName.Text
    OpenFolder pth
End Sub

Private Sub cmdOptions_Click()
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
    StopDownload NormalStop
End Sub

Private Sub cmdStopBatch_Click()
    StopDownload BatchStop
End Sub

Sub SetBackgroundPosition(Optional ByVal ForceRefresh As Boolean = False)
    On Error Resume Next
    Dim i%, j%, k%
    If imgBackground.Visible Then
        Dim ImageCentered As Boolean
        Dim imgpos As Integer
        ImageCentered = False
        imgpos = ImagePosition
        If ImagePosition > 3 And ImagePosition <= 6 Then
            ImageCentered = True
            imgpos = imgpos - 3
        End If
        Dim Width&, Height&
        Width = GetPictureWidth(imgBackground.Picture)
        Height = GetPictureHeight(imgBackground.Picture)
        Select Case imgpos
            Case 0 '���̱�
                If imgBackground.Stretch <> True Then imgBackground.Stretch = True
                imgBackground.Move 0, 0, Me.Width, Me.Height
            Case 1 '���̿� ���߱�
                If imgBackground.Stretch <> True Then imgBackground.Stretch = True
                imgBackground.Move IIf(ImageCentered, (Me.Width - imgBackground.Width) \ 2, 0), 0, Width / Height * Me.Height, Me.Height
            Case 2 '�ʺ� ���߱�
                If imgBackground.Stretch <> True Then imgBackground.Stretch = True
                imgBackground.Move 0, IIf(ImageCentered, ((Me.Height - sbStatusBar.Height - CaptionHeight * 15 - 15) - imgBackground.Height) \ 2, 0), Me.Width, Height / Width * Me.Width
            Case 3 '���� ũ��
                If imgBackground.Stretch = True Then imgBackground.Stretch = False
                imgBackground.Move IIf(ImageCentered, (Me.Width - imgBackground.Width) \ 2, 0), IIf(ImageCentered, ((Me.Height - sbStatusBar.Height - CaptionHeight * 15 - 15) - imgBackground.Height) \ 2, 0), Width, Height
            Case 7 '�ٵ��ǽ�
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
        If imgpos <> 7 And MaxLoadedTileBackgroundImage > 0 Then
            For i = 1 To MaxLoadedTileBackgroundImage
                Set imgBackgroundTile(i).Picture = Nothing
                Unload imgBackgroundTile(i)
                Set imgBackgroundTile(i) = Nothing
            Next i
        End If
        If ImagePosition < 2 Or ImagePosition = 4 Or ForceRefresh Or ImageCentered Then
dorefresh:
            On Error Resume Next
            fOptions.Refresh
            chkOpenAfterComplete.Refresh
            chkOpenFolder.Refresh
            chkContinueDownload.Refresh
            chkAutoRetry.Refresh
            If Not fTygemFrameTransparent.Visible Then
                Frame4.Refresh
                fTotal.Refresh
            End If
            trThreadCount.VisualStyles = Not trThreadCount.VisualStyles
            trThreadCount.VisualStyles = Not trThreadCount.VisualStyles
            fTabs.Refresh
            optTabDownload2.Refresh
            optTabThreads2.Refresh
            If (Not cmdBrowse.IsTygemButton) And cmdBrowse.VisualStyles Then
                Dim ctrl As Control
                For Each ctrl In Me.Controls
                    If TypeOf ctrl Is CommandButtonW Then ctrl.Refresh
                Next ctrl
            End If
            fDownloadInfo.Refresh
            fThreadInfo.Refresh
            pbProgressOuterContainer.Refresh
            pbProgressContainer.Refresh
            If fTygemFrameTransparent.Visible And fTygemFrameTransparent.Transparent Then fTygemFrameTransparent.Refresh
            fState.Refresh
            fProgress.Refresh
        End If
    Else
        GoTo dorefresh
    End If
End Sub

Sub SetBackgroundImage()
    On Error Resume Next
    Dim i%
    Dim BackgroundImagePath$
    BackgroundImagePath = Trim$(GetSetting("DownloadBooster", "Options", "BackgroundImagePath", ""))
    If GetSetting("DownloadBooster", "Options", "UseBackgroundImage", 0) = 1 And LenB(BackgroundImagePath) Then
        Set imgBackground.Picture = LoadPictureFromFile(BackgroundImagePath)
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
End Sub

Sub SetPattern()
    pgPattern.FillStyle = CInt(GetSetting("DownloadBooster", "Options", "FormFillStyle", 0)) + 1
    pgPattern.FillColor = CLng(GetSetting("DownloadBooster", "Options", "FormFillColor", 0))
End Sub

Sub SetFrameTexture()
    On Error Resume Next
    
    Select Case LCase(GetSetting("DownloadBooster", "Options", "LiveBadukMemoSkinFrameType", "solidcolor"))
        Case "texture"
            fTygemFrameTransparent.Transparent = False
            Set imgFrameTexture.Picture = LoadPictureFromFile(GetSetting("DownloadBooster", "Options", "LiveBadukMemoSkinFrameTexture", ""))
            imgFrameTexture.Visible = True
        Case "solidcolor"
            fTygemFrameTransparent.Transparent = False
            imgFrameTexture.Visible = False
        Case Else
            fTygemFrameTransparent.Transparent = True
            imgFrameTexture.Visible = True
            If DarkTransparent Is Nothing Then Set DarkTransparent = LoadPictureFromResource(101, RCData)
            Set imgFrameTexture.Picture = DarkTransparent
    End Select
End Sub

Sub LoadLiveBadukSkin()
    On Error Resume Next
    Dim i%
    Dim LBEnabled As Boolean
    Dim fSkin As Byte
    fSkin = CByte(GetSetting("DownloadBooster", "Options", "ProgressFrameSkin", 1))
    LBEnabled = (fSkin > 0 And DPI = 96)
    If LBEnabled Then
        fTygemFrameTransparent.Top = 1200
        fTygemFrameTransparent.Left = 120
        SetTygemFrameRgn
        fTygemFrameTransparent.Visible = -1
        fTygemFrameTransparent.BackColor = CLng(GetSetting("DownloadBooster", "Options", "LiveBadukMemoSkinFrameColor", 11194780))
        
        SetFrameTexture
        
        For i = imgFrame.LBound To imgFrame.UBound
            imgFrame(i).Visible = (i = fSkin)
        Next i

        fTotal.Visible = 0
        Frame4.Visible = 0
        'lblLBCaption.Visible = -1
        fDownloadInfo.Refresh
        pbProgressOuterContainer.Refresh
        pbProgressContainer.Refresh

        pbTotalProgress.Top = 1800 - 90
        pbTotalProgressMarquee.Top = 1800 - 90
        
        optTabDownload2.Width = 1130
        optTabThreads2.Width = 1485
        optTabDownload2.Caption = fTabDownload.Caption
        optTabThreads2.Caption = fTabThreads.Caption
        
        pbTotalProgressMarquee.Left = 360 + 735
        pbTotalProgressMarquee.Width = 6015 - 735 - 735
        pbTotalProgress.Left = 360 + 735
        pbTotalProgress.Width = 6015 - 735 - 735
        lblState.Visible = False
        
        fState.Visible = True
        fState.Top = 1755
        fState.Left = 360
        fState.Refresh
        
        fProgress.Visible = True
        fProgress.Top = 1755
        fProgress.Refresh
        
        LBFrameEnabled = True
        fOptions.BorderStyle = 0
        cmdOptions.Left = 7200
        cmdAbout.Left = 7200
        Label11.Visible = True
        lbOptionsHeader.Visible = True
        lbOptionsHeader3D.Visible = True
        
        lblURL.Font.Bold = True
        lblFilePath.Font.Bold = True
        lblThreadCountLabel.Font.Bold = True
        
        Dim EnableShadow As Boolean
        EnableShadow = (GetSetting("DownloadBooster", "Options", "LiveBadukMemoSkinEnableShadow", 1) <> 0)
        If EnableShadow Then
            Dim ShadowColor As Long
            ShadowColor = CLng(GetSetting("DownloadBooster", "Options", "LiveBadukMemoSkinShadowColor", 16777215))
            If ShadowColor < 0& Then ShadowColor = 16777215
            'lblLBCaptionShadow.ForeColor = ShadowColor
            lblLBCaptionShadow2.ForeColor = ShadowColor
            lblURLShadow.ForeColor = ShadowColor
            lblFilePathShadow.ForeColor = ShadowColor
            lblThreadCountLabelShadow.ForeColor = ShadowColor
        End If
        'lblLBCaptionShadow.Visible = EnableShadow
        lblLBCaptionShadow2.Visible = EnableShadow
        lblURLShadow.Visible = EnableShadow
        lblFilePathShadow.Visible = EnableShadow
        lblThreadCountLabelShadow.Visible = EnableShadow
        
        cbWhenExist.Width = 1305
        
        Dim ShowBorder As Boolean
        ShowBorder = (GetSetting("DownloadBooster", "Options", "LiveBadukMemoSkinEnableBorder", 1) <> 0)
        imgBorderTopLeft.Visible = ShowBorder
        imgBorderTopRight.Visible = ShowBorder
        imgBorderBottomLeft.Visible = ShowBorder
        imgBorderBottomRight.Visible = ShowBorder
        pgBorderLeft.Visible = ShowBorder
        pgBorderTop.Visible = ShowBorder
        pgBorderRight.Visible = ShowBorder
        pgBorderBottom.Visible = ShowBorder
        
        imgLBContentBackground.Visible = True
        Select Case LCase(GetSetting("DownloadBooster", "Options", "LiveBadukMemoSkinFrameBackgroundType", "transparent"))
            Case "texture"
                Set imgLBContentBackground.Picture = LoadPictureFromFile(GetSetting("DownloadBooster", "Options", "LiveBadukMemoSkinFrameBackground", ""))
            Case "solidcolor"
                Set imgLBContentBackground.Picture = GenerateSolidColor(CLng(GetSetting("DownloadBooster", "Options", "LiveBadukMemoSkinFrameBackgroundColor", 16777215)))
            Case Else
                If LightTransparent Is Nothing Then Set LightTransparent = LoadPictureFromResource(102, RCData)
                Set imgLBContentBackground.Picture = LightTransparent
        End Select
        
        fTabThreads.Visible = False
        fTabDownload.Visible = False
        
        lblLBCaption2.Left = FrameSkinLabelOffset(fSkin)
        lblLBCaptionShadow2.Left = FrameSkinLabelOffset(fSkin) + Screen.TwipsPerPixelX
    Else
        fTygemFrameTransparent.Visible = 0

        'lblLBCaption.Visible = 0
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
        lbOptionsHeader3D.Visible = False
        
        'lblLBCaptionShadow.Visible = 0
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
        
        fTabThreads.Visible = True
        fTabDownload.Visible = True
        
        fState.Visible = False
        fProgress.Visible = False
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
        lblFilename.ForeColor = ContentTextColor
        lblTotalBytes.ForeColor = ContentTextColor
        lblDownloadedBytes.ForeColor = ContentTextColor
        lblElapsed.ForeColor = ContentTextColor
        lblSpeed.ForeColor = ContentTextColor
        lblThreadCount2.ForeColor = ContentTextColor
        lblTotalSizeThread.ForeColor = ContentTextColor
        lblRemaining.ForeColor = ContentTextColor
        
        If GetSetting("DownloadBooster", "Options", "LiveBadukMemoSkinEnableTextColor", 1) <> 0 Then
            Dim LBTextColor As Long
            LBTextColor = CLng(GetSetting("DownloadBooster", "Options", "LiveBadukMemoSkinTextColor", 2111255))
            lblURL.ForeColor = LBTextColor
            lblFilePath.ForeColor = LBTextColor
            lblThreadCountLabel.ForeColor = LBTextColor
            lblLBCaption.ForeColor = LBTextColor
            lblLBCaption2.ForeColor = LBTextColor
        End If
    End If

    Dim ctrl As Control
    For Each ctrl In Me.Controls
        If TypeOf ctrl Is FrameW Or TypeOf ctrl Is CheckBoxW Or TypeOf ctrl Is CommandButtonW Or TypeOf ctrl Is Slider Then
            ctrl.Refresh
        End If
    Next ctrl
    
    SetFont Me
    
    If Not LBEnabled Then
        SetTextColors
        fTabDownload.Refresh
        fTabThreads.Refresh
    End If
End Sub

#If HIDEYTDL Then
#Else
Private Sub cmdYtdlTest_Click()
    StartYtdlDownload
End Sub
#End If

Sub SetTitle(Optional ByVal Title As String = "")
    If Title = "" Then
        Me.Caption = FormCaption
    Else
        Me.Caption = Title & " - " & FormCaption
    End If
End Sub

Private Sub Form_Load()
    On Error Resume Next
    
    '���� �ڵ� ����
    Set ErrorCodeDescription = New Collection
    ErrorCodeDescription.Add t("�������� ������ ������ϴ�. �ٿ�ε� �� ��Ʈ��ũ ������ �߻��߰ų� �ּҰ� ��ȿ���� �ʰų� ������ �������� �ʽ��ϴ�.", "Network error"), "1"
    ErrorCodeDescription.Add t("�ּҳ� ���� �̸��� �������� �ʾҽ��ϴ�.", "Address or file name unspecified"), "102"
    ErrorCodeDescription.Add t("���� ��ΰ� �������� �ʽ��ϴ�.", "Save path doesn't exist"), "103"
    ErrorCodeDescription.Add t("������ ���ϸ��� ��� ���Դϴ�. �ٸ� �̸��� �����Ͻʽÿ�.", "File name already exists"), "104"
    ErrorCodeDescription.Add t("���� ������ �ٿ�ε� �ν�Ʈ�� �������� �ʽ��ϴ�. ������ 1�� ������ ���ʽÿ�.", "Download boosting not supported. Try changing the thread count to 1."), "106"
    ErrorCodeDescription.Add t("������ ũ�⸦ �� �� ��� �ٿ�ε带 �ν�Ʈ�� �� �����ϴ�. ������ 1�� ������ ���ʽÿ�.", "Unable to boost download because the file size is not provided. Try changing the thread count to 1."), "107"
    ErrorCodeDescription.Add t("������ ��û�� �ź��߽��ϴ�. ���� �� �����̰ų� �������� �������� �ʰų� ���� ������ ���� �� �ֽ��ϴ�.", "Server has denied your request. The file may not exist or have insufficient permissions to access it."), "108"
    
    '�ִ� ������ ����
    MAX_THREAD_COUNT = CInt(GetSetting("DownloadBooster", "Options", "MaxThreadCount", 25))
    
    sbStatusBar.Panels(1).Text = t("�غ�", "Ready")
    
    'â ���� ����
    FormCaption = App.Title & " " & App.Major & "." & App.Minor
    If App.Revision > 0 Then FormCaption = FormCaption & "." & App.Revision
#If BETA Then
    FormCaption = FormCaption & " " & t("��Ÿ", "Beta") & " " & BetaVer
#End If
    SetTitle
    
    '��ũ�� ���� ����
    ScrollOneScreen = GetSetting("DownloadBooster", "Options", "ScrollOneScreen", 0) <> 0
    vsProgressScroll.LargeChange = IIf(ScrollOneScreen, 1, 10)
    
    MaxLoadedTileBackgroundImage = 0
    ImagePosition = GetSetting("DownloadBooster", "Options", "ImagePosition", 1)
    
    'â ��ġ �ҷ�����
    Dim Lft%
    Dim Top%
    Top = GetSetting("DownloadBooster", "UserData", "FormTop", "")
    Lft = GetSetting("DownloadBooster", "UserData", "FormLeft", "")
    If LenB(Top) And LenB(Lft) Then
        Me.Top = Top
        Me.Left = Lft
    End If
    
    '������ ���� â UI �ʱ� ����
    Dim i%, ThreadInfoLabelTop#, ThreadInfoProgressTop#, ThreadInfoDownloaderCaption$
    ThreadInfoDownloaderCaption = t("������", "Thread") & " "
    For i = 1 To MAX_THREAD_COUNT
        Load lblDownloader(i)
        Load lblPercentage(i)
        Load pbProgress(i)
        Load pbProgressMarquee(i)
        ThreadInfoProgressTop = 360# * CDbl(i - 1)
        ThreadInfoLabelTop = ThreadInfoProgressTop + 45#
        lblDownloader(i).Top = ThreadInfoLabelTop
        lblPercentage(i).Top = ThreadInfoLabelTop
        pbProgress(i).Top = ThreadInfoProgressTop
        pbProgress(i).ZOrder 1
        pbProgressMarquee(i).Top = ThreadInfoProgressTop
        pbProgressMarquee(i).ZOrder 0
        lblDownloader(i).Caption = ThreadInfoDownloaderCaption & i & ":"
    Next i
    
    '�����̴� ���� �� ����
    If MAX_THREAD_COUNT >= 250 Then
        trThreadCount.TickFrequency = 16
    ElseIf MAX_THREAD_COUNT >= 100 Then
        trThreadCount.TickFrequency = 8
    ElseIf MAX_THREAD_COUNT >= 50 Then
        trThreadCount.TickFrequency = 4
    ElseIf MAX_THREAD_COUNT <= 15 Then
        trThreadCount.TickFrequency = 1
    End If
    trThreadCount.Max = MAX_THREAD_COUNT
    
    pbProgressContainer.Height = 360# * CDbl(MAX_THREAD_COUNT)
    fDownloadInfo.Top = fThreadInfo.Top
    
    LoadLiveBadukSkin
    
    'â �ʺ� ����
    Me.Width = MAIN_FORM_WIDTH + PaddedBorderWidth * 15 * 2 * (DPI / 96)
    cmdStop.Left = Me.Width + 1200
    cmdStopBatch.Left = Me.Width + 1200
    
    '���õ� �� ���
    If GetSetting("DownloadBooster", "UserData", "LastTab", 1) = 1 Then
        fTabDownload_Click
    Else
        fTabThreads_Click
    End If
    
    '������ �� ���
    trThreadCount.Value = GetSetting("DownloadBooster", "UserData", "ThreadCount", GetSetting("DownloadBooster", "Options", "ThreadCount", 1))
    trThreadCount_Scroll
    
    '�ϰ� �ٿ�ε� �� ����
    lvBatchFiles.ColumnHeaders.Add , , t("���� �̸�", "File Name"), 2895
    lvBatchFiles.ColumnHeaders.Add , , t("��ü ���", "Full Path"), 0
    lvBatchFiles.ColumnHeaders.Add , , t("���� �ּ�", "File URL"), 4495
    lvBatchFiles.ColumnHeaders.Add , , t("����", "Status"), 1105, LvwColumnHeaderAlignmentCenter
    lvBatchFiles.ColumnHeaders.Add , , t("���� �̸� �ڵ� ����", "Autodetect File Name"), 0
    lvBatchFiles.ColumnHeaders.Add , , t("���ڵ��� ���", "Encoded Headers"), 0
#If HIDEYTDL Then
#Else
    lvBatchFiles.ColumnHeaders.Add , , "youtube-dl " & t("���", "used"), 0
    lvBatchFiles.ColumnHeaders.Add , , "youtube-dl: " & t("����", "format"), 0
    lvBatchFiles.ColumnHeaders.Add , , "youtube-dl: " & t("����� ����", "extract audio"), 0
    lvBatchFiles.ColumnHeaders.Add , , "youtube-dl: " & t("����� ����", "audio format"), 0
    lvBatchFiles.ColumnHeaders.Add , , "youtube-dl: " & t("����� ���� ����", "audio quality type"), 0
    lvBatchFiles.ColumnHeaders.Add , , "youtube-dl: CBR", 0
    lvBatchFiles.ColumnHeaders.Add , , "youtube-dl: VBR", 0
#End If
    
    'ȭ�� �̸� �� ��� ���
    txtFileName.Text = GetSetting("DownloadBooster", "UserData", "SavePath", CurDir())
    
    Me.Height = 6930 + PaddedBorderWidth * 15 * 2
    
    '�׻� ���� ǥ��
    If GetSetting("DownloadBooster", "Options", "AlwaysOnTop", 0) = 1 Then
        MainFormOnTop = True
        SetWindowPos hWnd, hWnd_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    Else
        MainFormOnTop = False
    End If
    
    '���� �޴� �׸� �߰�
    Dim hSysMenu As Long
    Dim MenuCount As Long
    hSysMenu = GetSystemMenu(Me.hWnd, 0)
    MenuCount = GetMenuItemCount(hSysMenu)
    Dim MII As MENUITEMINFO
    
    MII.cbSize = Len(MII)
    
    '�׻� ���� ǥ��
    With MII
        .fMask = MIIM_STATE Or MIIM_ID Or MIIM_TYPE
        .fType = MFT_STRING
        .fState = MFS_ENABLED
        .wID = 1000
        .dwTypeData = t("�׻� ���� ǥ��(&A)", "&Always on top")
        .cch = Len(.dwTypeData)
    End With
    InsertMenuItem hSysMenu, 0, 1, MII

    '���� ����
    With MII
        .fMask = MIIM_STATE Or MIIM_ID Or MIIM_TYPE
        .fType = MFT_STRING
        .fState = MFS_ENABLED
        .wID = 1003
        .dwTypeData = t("â ũ�� �ʱ�ȭ(&E)", "R&eset window size")
        .cch = Len(.dwTypeData)
    End With
    InsertMenuItem hSysMenu, 1, 1, MII

    '���м�
    With MII
        .cbSize = Len(MII)
        .fMask = MIIM_ID Or MIIM_TYPE
        .fType = MFT_SEPARATOR
        .wID = 2000
    End With
    InsertMenuItem hSysMenu, 2, 1, MII
    
    '�ϰ� ó�� â ǥ�� ���� ���
    If GetSetting("DownloadBooster", "UserData", "BatchExpanded", 1) <> 0 Then
        cmdBatch_Click
    Else
        FormWidth = (MAIN_FORM_WIDTH + PaddedBorderWidth * 15 * 2) / 15
        FormMinHeight = (6930 + PaddedBorderWidth * 15 * 2) / 15
        FormMaxHeight = (6930 + PaddedBorderWidth * 15 * 2) / 15
    End If
    
    '����� ���� �ҷ�����
    AddItemToComboBox cbWhenExist, t("�ǳʶٱ�", "Skip")
    AddItemToComboBox cbWhenExist, t("�����", "Overwrite")
    AddItemToComboBox cbWhenExist, t("�̸� ����", "Rename")
    cbWhenExist.ListIndex = GetSetting("DownloadBooster", "Options", "WhenFileExists", 0)
    chkOpenAfterComplete.Value = GetSetting("DownloadBooster", "Options", "OpenWhenComplete", 0)
    chkOpenFolder.Value = GetSetting("DownloadBooster", "Options", "OpenFolderWhenComplete", 0)
    chkContinueDownload.Value = GetSetting("DownloadBooster", "Options", "ContinueDownload", 0)
    chkAutoRetry.Value = GetSetting("DownloadBooster", "Options", "AutoRetry", 0)
    
    If GetSetting("DownloadBooster", "Options", "RememberURL", 0) <> 0 Then
        txtURL.Text = GetSetting("DownloadBooster", "UserData", "FileURL", "")
        txtURL.SelStart = 0
        txtURL.SelLength = Len(txtURL.Text)
    End If
    
    '�ɰ��� ���� ����
    SetupSplitButtons

    '���ö���¡ ����
    tr lblURL, "File &address:"
    tr lblFilePath, "Save &file to:"
    tr lblThreadCountLabel, "&Threads:"
    tr cmdClear, "Clear(&Y)"
    tr cmdBrowse, "&Browse..."
    tr fTotal, " Total Progress "
    tr fTabDownload, " Summary "
    tr fTabThreads, " Threads "
    tr cmdOptions, "More opt&ions..."
    tr cmdOpen, "&Open"
    tr cmdOpenFolder, "Op&en folder"
    tr cmdGo, "&Download"
    tr cmdStop, "Sto&p"
    tr cmdAddToQueue, "Add to &queue"
    tr cmdBatch, "Batc&h download"
    tr lblState, "Stopped"
    lblLBState = lblState
    tr cmdOpenBatch, "Open(&W)"
    tr cmdAdd, "Add U&RL..."
    tr cmdDelete, "Remo&ve"
    tr cmdStartBatch, "&Start"
    tr cmdStopBatch, "Stop(&Z)"
    tr Label8, "File name:"
    tr Label2, "Total:"
    tr Label3, "Recieved:"
    tr Label4, "Elapsed:"
    tr Label5, "Speed:"
    tr Label6, "Threads:"
    tr Label7, "Size per thread:"
    tr fOptions, " Settings "
    
    tr chkOpenAfterComplete, "Open when &complete"
    tr chkOpenFolder, "Open fo&lder when done"
    tr chkContinueDownload, "Always resume(&J)"
    tr chkAutoRetry, "Auto retry on error(&G)"
    
    tr Label1, "Exists(&K):"
    tr mnuAddItem, "&Add URL..."
    tr mnuClearBatch, "&Clear list"
    tr mnuClearBatch2, "&Clear list"
    tr mnuClearBatch3, "&Clear list"
    tr mnuDeleteItem, "&Remove"
    tr mnuOpenFolder, "Open &folder"
    tr cmdAbout, "Abo&ut application"
    tr Label10, "Remaining:"
    tr lblLBCaption, "Progress"
    
    tr mnuEdit, "&Edit..."
    tr mnuMoveUp, "Move &up"
    tr mnuMoveDown, "Move &down"
    tr mnuAddItem2, "&Add URL..."
    tr mnuOpenBatch, "&Open"
    tr mnuOpenFolder2, "Open &folder"
    
    tr cmdEdit, "Edit(&N)..."
    
    tr mnuProperties, "View p&roperties"
    tr mnuPropertiesBatch, "View p&roperties"

#If HIDEYTDL Then
    cmdDownloadOptions.Caption = t("��� ����(&S)...", "Edit header&s...")
#Else
    tr cmdDownloadOptions, "Download &settings..."
#End If
    
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
    '���ö���¡ ��
    
    lbOptionsHeader.X1 = Label11.Width + 60
    lbOptionsHeader3D.X1 = Label11.Width + 75
    
    'â ȭ���� ���� �ҷ�����
    If GetSetting("DownloadBooster", "Options", "DisableDWMWindow", DefaultDisableDWMWindow) = 1 Then DisableDWMWindow Me.hWnd
    SetPattern
    SetBackgroundImage
    SetBackgroundPosition
    SetTextColors
    SetFont Me
    
    '�̹��� ����Ʈ �ε�
    Dim imlPicture As IPictureDisp
    
    Set imlPicture = imgDropdown.ListImages(1).ExtractIcon()
    imgDropdown.ListImages.Add 1, Picture:=imlPicture
    imgDropdown.ListImages.Add 1, Picture:=imlPicture
    imgDropdown.ListImages.Add 5, Picture:=imlPicture
    
    Set imlPicture = imgPlay.ListImages(1).ExtractIcon()
    imgPlay.ListImages.Add 1, Picture:=imlPicture
    imgPlay.ListImages.Add 1, Picture:=imlPicture
    imgPlay.ListImages.Add 5, Picture:=imlPicture
    
    Set imlPicture = imgMinus.ListImages(1).ExtractIcon()
    imgMinus.ListImages.Add 1, Picture:=imlPicture
    imgMinus.ListImages.Add 1, Picture:=imlPicture
    imgMinus.ListImages.Add 5, Picture:=imlPicture
    
    Set imlPicture = imgOpenFile.ListImages(1).ExtractIcon()
    imgOpenFile.ListImages.Add 1, Picture:=imlPicture
    imgOpenFile.ListImages.Add 1, Picture:=imlPicture
    imgOpenFile.ListImages.Add 5, Picture:=imlPicture
    
    Set imlPicture = Nothing
    
    '����Ŭ����
    AttachMessage Me, Me.hWnd, WM_GETMINMAXINFO
    AttachMessage Me, Me.hWnd, WM_INITMENU
    AttachMessage Me, Me.hWnd, WM_SYSCOMMAND
    'AttachMessage Me, Me.hWnd, WM_DWMCOMPOSITIONCHANGED
    AttachMessage Me, Me.hWnd, WM_SETTINGCHANGE
    AttachMessage Me, Me.hWnd, WM_THEMECHANGED
    AttachMessage Me, Me.hWnd, WM_CTLCOLORSCROLLBAR
    
    '��ũ�� ǥ�� ����
    vsProgressScroll.Visible = (trThreadCount.Value > 10 And optTabThreads2.Value)
End Sub

Private Sub SetTygemFrameRgn()
    Dim fSkin As Byte
    fSkin = CByte(GetSetting("DownloadBooster", "Options", "ProgressFrameSkin", 1))
    
    Dim RC As RECT
    GetWindowRect fTygemFrameTransparent.hWnd, RC
    Dim Rgn&, Rgn1&, Rgn2&, Rgn3&, Rgn4&, Rgn5&, Rgn6&, Rgn7&, Rgn8&, Rgn9&
    Rgn = CreateRectRgn(0, 0, RC.Right - RC.Left, RC.Bottom - RC.Top)
    
    Select Case fSkin
        Case 1
            Rgn1 = CreateRectRgn(0, 0, 17, 1)
            Rgn2 = CreateRectRgn(0, 1, 15, 2)
            Rgn3 = CreateRectRgn(0, 2, 14, 3)
            Rgn4 = CreateRectRgn(0, 3, 13, 4)
            Rgn5 = CreateRectRgn(0, 4, 12, 7)
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
            Rgn1 = CreateRectRgn(0, 7, 12, 17)
            Rgn2 = CreateRectRgn(0, 17, 11, 18)
            CombineRgn Rgn, Rgn, Rgn1, RGN_DIFF
            CombineRgn Rgn, Rgn, Rgn2, RGN_DIFF
            DeleteObject Rgn1
            DeleteObject Rgn2
            Rgn1 = CreateRectRgn(0, 18, 6, 19)
            Rgn2 = CreateRectRgn(0, 19, 4, 20)
            Rgn3 = CreateRectRgn(0, 20, 3, 21)
            Rgn4 = CreateRectRgn(0, 21, 2, 22)
            Rgn5 = CreateRectRgn(0, 22, 1, 23)
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
            Rgn1 = CreateRectRgn(100, 0, 112, 1)
            Rgn2 = CreateRectRgn(102, 1, 112, 2)
            Rgn3 = CreateRectRgn(103, 2, 112, 3)
            Rgn4 = CreateRectRgn(104, 3, 112, 4)
            Rgn5 = CreateRectRgn(105, 4, 112, 17)
            Rgn6 = CreateRectRgn(106, 17, 112, 18)
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
            Rgn1 = CreateRectRgn(10, 30, RC.Right - RC.Left - 10, RC.Bottom - RC.Top - 13)
            CombineRgn Rgn, Rgn, Rgn1, RGN_DIFF
            DeleteObject Rgn1
        Case 2
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
    End Select
    
    SetWindowRgn fTygemFrameTransparent.hWnd, Rgn, True
    DeleteObject Rgn
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
    
    If fOptions.VisualStyles Then
        fTabDownload.ForeColor = FrameCaptionColor
        fTabThreads.ForeColor = FrameCaptionColor
    End If
End Sub

Private Sub SetupSplitButton(ByRef LeftButton As CommandButtonW, ByRef RightButton As CommandButtonW)
    LeftButton.GetTygemButton().SplitLeft = True
    RightButton.GetTygemButton().SplitRight = True
    LeftButton.SetRgn
    RightButton.SetRgn
    If ComCtlsSupportLevel() >= 2 Then
        If CurrentButtonSkin = 0 Then
            If Not LeftButton.SplitButton Then LeftButton.Width = LeftButton.Width + 255
            LeftButton.SplitButton = True
            RightButton.Visible = False
        Else
            If LeftButton.SplitButton Then LeftButton.Width = LeftButton.Width - 255
            LeftButton.SplitButton = False
            RightButton.Visible = True
        End If
    End If
End Sub

Sub SetupSplitButtons()
    SetupSplitButton cmdOpenBatch, cmdOpenDropdown
    SetupSplitButton cmdDelete, cmdDeleteDropdown
    SetupSplitButton cmdOpen, cmdOpenFileDropdown
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
    
    If imgBackground.Visible Or imgBackgroundTile(0).Visible Then SetBackgroundPosition
End Sub

Private Sub StopDownload(Optional ByVal StopMode As DownloadStopMode = NormalStop, Optional ByRef Cancel As Integer)
    Dim IsMarquee As Boolean
    IsMarquee = pbTotalProgressMarquee.Visible
    Dim ConfirmResult As VbMsgBoxResult
    If IsMarquee Or ResumeUnsupported Then
        ConfirmResult = ConfirmEx(t("�ٿ�ε带 �����Ͻðڽ��ϱ�? ���� ������ �̾�ޱⰡ �������� �����Ƿ� ó������ �ٽ� �ٿ�ε�޾ƾ� �մϴ�.", "Cancel download? Resuming is not supported for this file."), t("�ٿ�ε� ���", "Cancel download"), 48)
    Else
        ConfirmResult = MsgBox(t("�ٿ�ε带 �����Ͻðڽ��ϱ�? �̾�ޱ� ����� ���� �ߴ��� ������ ��� �ٿ�ε���� �� �ֽ��ϴ�.", "Cancel download? You can resume later."), vbQuestion + vbYesNo, t("�ٿ�ε� ���", "Cancel download"))
    End If
    If ConfirmResult = vbYes Then
        Dim CurrentProgress As Long
        CurrentProgress = pbTotalProgress.Value
        
        Select Case StopMode
            Case BatchStop
                lvBatchFiles.ListItems(CurrentBatchIdx).ListSubItems(3).Text = t("����", "Stopped")
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
            Case ExitApplication
                BatchStarted = False
                SP.FinishChild 0, 0
            Case NormalStop
                OnStop False
                cmdOpen.Enabled = 0
                cmdOpenFileDropdown.Enabled = 0
        End Select
        
        If IsMarquee Or (CurrentProgress > 0 And CurrentProgress < 100) Then
            Dim KillTemp As Boolean
            KillTemp = False
            If IsMarquee Or ResumeUnsupported Then
                KillTemp = True
            Else
                KillTemp = MsgBox(t("���߿� �̾ �ٿ�ε���� �� �ֵ��� �ٿ�ε��� ����Ÿ�� �����Ͻðڽ��ϱ�?", "Would you like to keep the partially downloaded data to resume later?"), vbQuestion + vbYesNo) <> vbYes
            End If
            If KillTemp Then
                On Error Resume Next
                Kill DownloadPath & ".part.tmp"
            End If
        End If
        
        If StopMode = BatchStop And BatchErrorCount Then
            MsgBox t("�ϳ� �̻��� ������ �߻��߽��ϴ�. �ش� �׸��� �� �� ������ ���� ������ �� �� �ֽ��ϴ�.", "One or more errors have occurred. Double click the error item to see details."), 48
        End If
    ElseIf StopMode = ExitApplication Then
        Cancel = 1
        Exit Sub
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    If cmdStop.Enabled Or BatchStarted Then
        StopDownload ExitApplication, Cancel
        If Cancel Then Exit Sub
    Else
        BatchStarted = False
        SP.FinishChild 0, 0
    End If
    
    Me.Hide
    
    SaveSetting "DownloadBooster", "UserData", "SavePath", Trim$(txtFileName.Text)
    SaveSetting "DownloadBooster", "UserData", "BatchExpanded", CInt(Me.Height > 6930 + PaddedBorderWidth * 15 * 2) * -1
    SaveSetting "DownloadBooster", "Options", "WhenFileExists", cbWhenExist.ListIndex
    If GetSetting("DownloadBooster", "Options", "RememberURL", 0) <> 0 Then SaveSetting "DownloadBooster", "UserData", "FileURL", Trim$(txtURL.Text)
    SaveSetting "DownloadBooster", "UserData", "FormTop", Me.Top
    SaveSetting "DownloadBooster", "UserData", "FormLeft", Me.Left
    If Me.Height >= 8220 Then SaveSetting "DownloadBooster", "UserData", "FormHeight", Me.Height - PaddedBorderWidth * 15 * 2
    SaveSetting "DownloadBooster", "UserData", "LastTab", (CInt(optTabThreads2.Value) * -1) + 1
    
    Unload frmBatchAdd
    Unload frmBrowse
    Unload frmOptions
    Unload frmExplorer
    Unload frmDummyForm
    Unload frmEditBatch
    IBSSubclass_UnsubclassIt
    GetSystemMenu Me.hWnd, 1&
    Unload frmMessageBox
    Unload frmInputBox
    Unload frmAbout
    Unload frmDownloadOptions
    'If Not InIDE Then ExitProcess 0&
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
    Dim ItemCount As Long
    ItemCount = lvBatchFiles.ListItems.Count
    On Error GoTo ErrLn
    If lvBatchFiles.SelectedItem.Selected Then
        If cmdDelete.Enabled Then
            Dim ErrorOccurred As Boolean, Openable As Boolean, SelectedIndex As Long
            SelectedIndex = lvBatchFiles.SelectedItem.Index
            ErrorOccurred = (lvBatchFiles.SelectedItem.ForeColor = vbRed)
            Openable = cmdOpenBatch.Enabled
            mnuOpenBatch.Visible = Openable
            mnuOpenFolder2.Visible = Openable
            mnuErrorInfo.Visible = ErrorOccurred
            mnuSepOpen.Visible = ErrorOccurred Or Openable
            mnuMoveUp.Enabled = (SelectedIndex <> 1) And (Not BatchStarted)
            mnuMoveDown.Enabled = (SelectedIndex <> ItemCount) And (Not BatchStarted)
            
            Dim DefaultMenu As Menu
            If cmdOpenBatch.Enabled Then
                Set DefaultMenu = mnuOpenBatch
            ElseIf mnuErrorInfo.Visible Then
                Set DefaultMenu = mnuErrorInfo
            Else
                Set DefaultMenu = mnuEdit
            End If
            Me.PopupMenu mnuListContext, DefaultMenu:=DefaultMenu
        End If
        Exit Sub
    End If
    
ErrLn:
    mnuClearBatch2.Enabled = (ItemCount > 0)
    Me.PopupMenu mnuListContext2
End Sub

Private Sub lvBatchFiles_ItemCheck(Item As LvwListItem, ByVal Checked As Boolean)
    If BatchStarted And Item.Index = CurrentBatchIdx Then
        Item.Checked = True
        Exit Sub
    Else
        Dim ForeColor&, StatusText$
        If Not Checked Then
            ForeColor = &H808080
            StatusText = t("���", "Skip")
        Else
            ForeColor = 0&
            StatusText = t("���", "Queued")
        End If
        Item.ListSubItems(3).Text = StatusText
        Item.ForeColor = ForeColor
        Item.ListSubItems(1).ForeColor = ForeColor
        Item.ListSubItems(2).ForeColor = ForeColor
        Item.ListSubItems(3).ForeColor = ForeColor
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
    cmdStartBatch.Enabled = Enable
End Sub

Private Sub lvBatchFiles_ItemDblClick(Item As LvwListItem, ByVal Button As Integer)
    On Error Resume Next
    If Not Item.Selected Then Exit Sub
    If cmdOpenBatch.Enabled And Item.ListSubItems(3).Text = t("�Ϸ�", "Done") Then
        cmdOpenBatch_Click
    ElseIf Item.ForeColor = vbRed Then
        mnuErrorInfo_Click
    ElseIf (Not BatchStarted) Or (BatchStarted And CurrentBatchIdx <> Item.Index) Then
        mnuEdit_Click
    End If
End Sub

Private Sub lvBatchFiles_ItemSelect(Item As LvwListItem, ByVal Selected As Boolean)
    If Selected Then
        Dim Enable As Boolean
        Enable = (Not (BatchStarted And Item.Index = CurrentBatchIdx))
        cmdDelete.Enabled = Enable
        cmdDeleteDropdown.Enabled = Enable
        cmdEdit.Enabled = Enable
        
        Enable = (Item.ListSubItems(3).Text = t("�Ϸ�", "Done"))
        cmdOpenBatch.Enabled = Enable
        cmdOpenDropdown.Enabled = Enable
    Else
        cmdDelete.Enabled = 0
        cmdDeleteDropdown.Enabled = 0
        cmdOpenBatch.Enabled = 0
        cmdOpenDropdown.Enabled = 0
        cmdEdit.Enabled = 0
    End If
End Sub

Private Sub lvBatchFiles_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 46 And lvBatchFiles.SelectedItem.Selected Then cmdDelete_Click
End Sub

Private Sub mnuAddItem_Click()
    cmdAdd_Click
End Sub

Private Sub mnuAddItem2_Click()
    mnuAddItem_Click
End Sub

Private Sub mnuClearBatch_Click()
    If lvBatchFiles.ListItems.Count Then
        If MsgBox(t("��⿭�� ��� �׸��� �����Ͻðڽ��ϱ�?", "Are you sure you want to clear the queue?"), vbQuestion + vbYesNo) <> vbYes Then Exit Sub
        
        Dim i&
        i = 1&
        Do While i <= lvBatchFiles.ListItems.Count
            If Not (BatchStarted And CurrentBatchIdx = i) Then
                lvBatchFiles.ListItems.Remove i
                If BatchStarted And CurrentBatchIdx > i Then
                    CurrentBatchIdx = CurrentBatchIdx - 1&
                End If
            ElseIf BatchStarted And CurrentBatchIdx = i Then
                i = i + 1&
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
    'frmEditBatch.Label2.Enabled = (lvBatchFiles.SelectedItem.ListSubItems(3).Text <> t("�Ϸ�", "Done"))
    'frmEditBatch.txtFilePath.Enabled = frmEditBatch.Label2.Enabled
    frmEditBatch.OriginalURL = lvBatchFiles.SelectedItem.ListSubItems(2).Text
    frmEditBatch.OriginalPath = lvBatchFiles.SelectedItem.ListSubItems(1).Text
    frmEditBatch.Show vbModal, Me
exitsub22:
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
#Else
    lvBatchFiles.ListItems(NewIdx).ListSubItems.Add , , lvBatchFiles.ListItems(DownIdx).ListSubItems(6).Text
    lvBatchFiles.ListItems(NewIdx).ListSubItems.Add , , lvBatchFiles.ListItems(DownIdx).ListSubItems(7).Text
    lvBatchFiles.ListItems(NewIdx).ListSubItems.Add , , lvBatchFiles.ListItems(DownIdx).ListSubItems(8).Text
    lvBatchFiles.ListItems(NewIdx).ListSubItems.Add , , lvBatchFiles.ListItems(DownIdx).ListSubItems(9).Text
    lvBatchFiles.ListItems(NewIdx).ListSubItems.Add , , lvBatchFiles.ListItems(DownIdx).ListSubItems(10).Text
    lvBatchFiles.ListItems(NewIdx).ListSubItems.Add , , lvBatchFiles.ListItems(DownIdx).ListSubItems(11).Text
    lvBatchFiles.ListItems(NewIdx).ListSubItems.Add , , lvBatchFiles.ListItems(DownIdx).ListSubItems(12).Text
#End If
    lvBatchFiles.ListItems(NewIdx).Checked = lvBatchFiles.ListItems(DownIdx).Checked
    lvBatchFiles.ListItems(NewIdx).ForeColor = lvBatchFiles.ListItems(DownIdx).ForeColor
    lvBatchFiles.ListItems(NewIdx).ListSubItems(1).ForeColor = lvBatchFiles.ListItems(DownIdx).ListSubItems(1).ForeColor
    lvBatchFiles.ListItems(NewIdx).ListSubItems(2).ForeColor = lvBatchFiles.ListItems(DownIdx).ListSubItems(2).ForeColor
    lvBatchFiles.ListItems(NewIdx).ListSubItems(3).ForeColor = lvBatchFiles.ListItems(DownIdx).ListSubItems(3).ForeColor
    lvBatchFiles.ListItems(NewIdx).ListSubItems(3).Text = lvBatchFiles.ListItems(DownIdx).ListSubItems(3).Text
    
    lvBatchFiles.ListItems.Remove DownIdx
    
exitsub:
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
#Else
    lvBatchFiles.ListItems(NewIdx).ListSubItems.Add , , lvBatchFiles.ListItems(UpIdx).ListSubItems(6).Text
    lvBatchFiles.ListItems(NewIdx).ListSubItems.Add , , lvBatchFiles.ListItems(UpIdx).ListSubItems(7).Text
    lvBatchFiles.ListItems(NewIdx).ListSubItems.Add , , lvBatchFiles.ListItems(UpIdx).ListSubItems(8).Text
    lvBatchFiles.ListItems(NewIdx).ListSubItems.Add , , lvBatchFiles.ListItems(UpIdx).ListSubItems(9).Text
    lvBatchFiles.ListItems(NewIdx).ListSubItems.Add , , lvBatchFiles.ListItems(UpIdx).ListSubItems(10).Text
    lvBatchFiles.ListItems(NewIdx).ListSubItems.Add , , lvBatchFiles.ListItems(UpIdx).ListSubItems(11).Text
    lvBatchFiles.ListItems(NewIdx).ListSubItems.Add , , lvBatchFiles.ListItems(UpIdx).ListSubItems(12).Text
#End If
    lvBatchFiles.ListItems(NewIdx).Checked = lvBatchFiles.ListItems(UpIdx).Checked
    lvBatchFiles.ListItems(NewIdx).ForeColor = lvBatchFiles.ListItems(UpIdx).ForeColor
    lvBatchFiles.ListItems(NewIdx).ListSubItems(1).ForeColor = lvBatchFiles.ListItems(UpIdx).ListSubItems(1).ForeColor
    lvBatchFiles.ListItems(NewIdx).ListSubItems(2).ForeColor = lvBatchFiles.ListItems(UpIdx).ListSubItems(2).ForeColor
    lvBatchFiles.ListItems(NewIdx).ListSubItems(3).ForeColor = lvBatchFiles.ListItems(UpIdx).ListSubItems(3).ForeColor
    lvBatchFiles.ListItems(NewIdx).ListSubItems(3).Text = lvBatchFiles.ListItems(UpIdx).ListSubItems(3).Text
    
    lvBatchFiles.ListItems.Remove UpIdx
    
exitsub:
End Sub

Private Sub mnuOpenBatch_Click()
    cmdOpenBatch_Click
End Sub

Private Sub mnuOpenFolder_Click()
    Dim pth$
    pth = lvBatchFiles.SelectedItem.ListSubItems(1).Text
    If LenB(pth) = 0 Then pth = txtFileName.Text
    ShellExecute pth
End Sub

Private Sub mnuOpenFolder2_Click()
    mnuOpenFolder_Click
End Sub

Private Sub mnuProperties_Click()
    On Error Resume Next
    ShellExecute DownloadPath, "properties"
End Sub

Private Sub mnuPropertiesBatch_Click()
    On Error Resume Next
    ShellExecute lvBatchFiles.SelectedItem.ListSubItems(1).Text, "properties"
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
    CancelDisplay = True
    SP.FinishChild 0
    OnStop
End Sub

Private Sub timElapsed_Timer()
    Elapsed = Elapsed + 1
    sbStatusBar.Panels(4).Text = FormatTime(Elapsed) & t(" ���", " elapsed")
    lblElapsed.Caption = Replace(sbStatusBar.Panels(4).Text, " " & t("���", "elapsed"), "")
End Sub

Private Sub trThreadCount_Change()
    trThreadCount_Scroll
    SaveSetting "DownloadBooster", "UserData", "ThreadCount", trThreadCount.Value
End Sub

Sub trThreadCount_Scroll()
    If trThreadCount.Value = 1 Then
        lblThreadCount.Caption = "(" & t("�Ϲ� �ٿ�ε�", "No threading") & ")"
    Else
        lblThreadCount.Caption = "(" & trThreadCount.Value & t("�� ������", " threads") & ")"
    End If
    Dim i%, bRowVisible As Boolean
    For i = 1 To lblDownloader.UBound
        bRowVisible = (i <= trThreadCount.Value)
        lblDownloader(i).Visible = bRowVisible
        pbProgress(i).Visible = bRowVisible
        lblPercentage(i).Visible = bRowVisible
    Next i
    
    If trThreadCount.Value - 10 > 0 Then
        If ScrollOneScreen Then
            vsProgressScroll.Max = Ceil(trThreadCount.Value / 10) - 1
        Else
            vsProgressScroll.Max = trThreadCount.Value - 10
        End If
        'vsProgressScroll.Enabled = -1
        vsProgressScroll.Visible = fThreadInfo.Visible
    Else
        If vsProgressScroll.Max <> 0 Then vsProgressScroll.Max = 0
        'If vsProgressScroll.Enabled Then vsProgressScroll.Enabled = 0
        
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
    
    cmdDecreaseThreads.Enabled = (trThreadCount.Value <> trThreadCount.Min)
    cmdIncreaseThreads.Enabled = (trThreadCount.Value <> trThreadCount.Max)
    
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
    If LBFrameEnabled Or imgBackground.Visible Then pbProgressContainer.Refresh
End Sub
