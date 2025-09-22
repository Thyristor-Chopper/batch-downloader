VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "정보"
   ClientHeight    =   5265
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   7650
   BeginProperty Font 
      Name            =   "굴림"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   7650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.PictureBox pbLicenses 
      BorderStyle     =   0  '없음
      Height          =   3255
      Index           =   2
      Left            =   2640
      ScaleHeight     =   3255
      ScaleWidth      =   4815
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1440
      Width           =   4815
      Begin VB.TextBox txtLicense 
         Height          =   3255
         Index           =   2
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  '양방향
         TabIndex        =   7
         Top             =   0
         Width           =   4815
      End
   End
   Begin VB.PictureBox pbLicenses 
      BorderStyle     =   0  '없음
      Height          =   3255
      Index           =   6
      Left            =   2640
      ScaleHeight     =   3255
      ScaleWidth      =   4815
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1440
      Width           =   4815
      Begin VB.TextBox txtLicense 
         Height          =   3255
         Index           =   6
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   18
         Top             =   0
         Width           =   4815
      End
   End
   Begin prjDownloadBooster.CommandButtonW cmdOK 
      Cancel          =   -1  'True
      Default         =   -1  'True
      Height          =   345
      Left            =   6120
      TabIndex        =   5
      Top             =   4800
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   609
      Caption         =   "확인"
   End
   Begin VB.PictureBox pbLicenses 
      BorderStyle     =   0  '없음
      Height          =   3255
      Index           =   3
      Left            =   2640
      ScaleHeight     =   3255
      ScaleWidth      =   4815
      TabIndex        =   16
      Top             =   1440
      Width           =   4815
      Begin VB.TextBox txtLicense 
         Height          =   3255
         Index           =   3
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  '양방향
         TabIndex        =   4
         Top             =   0
         Width           =   4815
      End
   End
   Begin VB.PictureBox pbLicenses 
      BorderStyle     =   0  '없음
      Height          =   3255
      Index           =   5
      Left            =   2640
      ScaleHeight     =   3255
      ScaleWidth      =   4815
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1440
      Width           =   4815
      Begin VB.TextBox txtLicense 
         Height          =   3255
         Index           =   5
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   10
         Top             =   0
         Width           =   4815
      End
   End
   Begin VB.PictureBox pbLicenses 
      BorderStyle     =   0  '없음
      Height          =   3255
      Index           =   1
      Left            =   2640
      ScaleHeight     =   3255
      ScaleWidth      =   4815
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1440
      Width           =   4815
      Begin VB.TextBox txtLicense 
         Height          =   3255
         Index           =   1
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   9
         Top             =   0
         Width           =   4815
      End
   End
   Begin VB.PictureBox pbLicenses 
      BorderStyle     =   0  '없음
      Height          =   3255
      Index           =   4
      Left            =   2640
      ScaleHeight     =   3255
      ScaleWidth      =   4815
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1440
      Width           =   4815
      Begin VB.TextBox txtLicense 
         Height          =   3255
         Index           =   4
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   8
         Top             =   0
         Width           =   4815
      End
   End
   Begin prjDownloadBooster.ImageList imgItems 
      Left            =   360
      Top             =   4320
      _ExtentX        =   1005
      _ExtentY        =   1005
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   16711935
      InitListImages  =   "frmAbout.frx":000C
   End
   Begin prjDownloadBooster.FrameW FrameW1 
      Height          =   3255
      Left            =   1080
      TabIndex        =   6
      Top             =   1440
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   5741
      BorderStyle     =   0
      Caption         =   "라이선스(&L)"
      Begin prjDownloadBooster.ListView lvItems 
         Height          =   3255
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   5741
         Icons           =   "imgItems"
         Arrange         =   2
         LabelEdit       =   2
         HideSelection   =   0   'False
         ShowInfoTips    =   -1  'True
         ShowLabelTips   =   -1  'True
         ShowColumnTips  =   -1  'True
         SnapToGrid      =   -1  'True
      End
   End
   Begin VB.PictureBox pbLicenses 
      BorderStyle     =   0  '없음
      Height          =   3255
      Index           =   7
      Left            =   2640
      ScaleHeight     =   3255
      ScaleWidth      =   4815
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1440
      Width           =   4815
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  '투명
      Caption         =   "버전"
      Height          =   225
      Left            =   1050
      TabIndex        =   1
      Top             =   600
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  '투명
      Caption         =   "응용 프로그램 제목"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1050
      TabIndex        =   0
      Top             =   240
      Width           =   3885
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  '투명
      Caption         =   "This product includes software developed by vbAccelerator."
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   1050
      TabIndex        =   2
      Top             =   960
      Width           =   6405
   End
   Begin VB.Image picIcon 
      Height          =   480
      Left            =   240
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    lvItems.SetFocus
End Sub

Private Sub Form_Load()
    On Error Resume Next
    InitForm Me
    
    Me.Caption = t(App.Title & " 정보", "About " & App.Title)
    Set picIcon.Picture = frmMain.Icon
    lblVersion.Caption = t("버전 ", "Version ") & App.Major & "." & App.Minor
    If App.Revision Then lblVersion.Caption = lblVersion.Caption & "." & App.Revision
    lblTitle.Caption = App.Title
    'lblDescription.Caption = "This product includes software developed by vbAccelerator. " & t("이 프로그램에는 외부 라이브러리가 사용됐으며 라이선스는 아래와 같습니다.", "This program includes external libraries. Check out the license of them below.")
    tr cmdOK, "OK"
    'tr FrameW1, "&License"
    
    imgItems.ListImages.Add Picture:=Train(1)
    imgItems.ListImages.Add Picture:=Train(2)
    
    lvItems.ListItems.Add , , "Krool's Comctl", 1
    lvItems.ListItems.Add , , "Node.js (v5.12.0)", 2
    lvItems.ListItems.Add , , "ShellPipe (v7)", 1
    lvItems.ListItems.Add , , "iconv-lite (v0.6.3)", 2
    lvItems.ListItems.Add , , "PNG with alpha", 1
    lvItems.ListItems.Add , , "vbAccelerator SSubTmr", 2
    lvItems.ListItems(1).Selected = True
    
    Dim i As Byte
    For i = txtLicense.LBound To txtLicense.UBound
        txtLicense(i).Text = LoadResText(200 + i, RCData)
    Next i
End Sub

Private Sub lvItems_ItemSelect(Item As LvwListItem, ByVal Selected As Boolean)
    On Error Resume Next
    If Not Selected Then Exit Sub 'If Item Is lvItems.SelectedItem Then Item.Selected = True: Exit Sub
    Static i As Byte
    For i = pbLicenses.LBound To pbLicenses.UBound
        pbLicenses(i).Visible = (i = Item.Index)
    Next i
End Sub
