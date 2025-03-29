VERSION 5.00
Begin VB.Form frmLiveBadukSkinProperties 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "라이브 바둑 쪽지 스킨 등록 정보"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6270
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "굴림"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLiveBadukSkinProperties.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   6270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin prjDownloadBooster.CommandButtonW cmdSelectTexture 
      Height          =   330
      Left            =   1800
      TabIndex        =   10
      Top             =   1410
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      Caption         =   "선택(&S)..."
   End
   Begin prjDownloadBooster.OptionButtonW optTexture 
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   1440
      Width           =   1335
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "텍스처(&E):"
   End
   Begin prjDownloadBooster.OptionButtonW optColor 
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   1080
      Width           =   1335
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "색(&C):"
   End
   Begin prjDownloadBooster.OptionButtonW optTransparent 
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   720
      Width           =   2655
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "반투명(&T)"
   End
   Begin prjDownloadBooster.CommandButtonW cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   1920
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "취소"
   End
   Begin prjDownloadBooster.CommandButtonW cmdOK 
      Default         =   -1  'True
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   1920
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "확인"
   End
   Begin VB.CommandButton cmdSelectShadow 
      Caption         =   "Command1"
      Height          =   180
      Left            =   -3720
      TabIndex        =   2
      Top             =   120
      Width           =   90
   End
   Begin VB.Label lblSelectFrameColor 
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   9
      Top             =   1080
      Width           =   855
   End
   Begin VB.Shape pgFrameColor 
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00404040&
      FillColor       =   &H00808080&
      Height          =   255
      Left            =   1800
      Shape           =   4  '둥근 사각형
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      Caption         =   "프레임 색(&F):"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "글자 그림자 색(&H):"
      Height          =   180
      Left            =   120
      TabIndex        =   1
      Top             =   165
      Width           =   1530
   End
   Begin VB.Shape pgShadow 
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00404040&
      FillColor       =   &H00808080&
      Height          =   255
      Left            =   1800
      Shape           =   4  '둥근 사각형
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblSelectShadow 
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   0
      Top             =   165
      Width           =   855
   End
End
Attribute VB_Name = "frmLiveBadukSkinProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    SaveSetting "DownloadBooster", "Options", "LiveBadukMemoSkinShadowColor", CLng(pgShadow.BackColor)
    SaveSetting "DownloadBooster", "Options", "LiveBadukMemoSkinFrameColor", CLng(pgFrameColor.BackColor)
    SaveSetting "DownloadBooster", "Options", "LiveBadukMemoSkinFrameType", IIf(optTransparent.Value, "transparent", IIf(optColor.Value, "solidcolor", "texture"))
    frmOptions.ColorChanged = True
    frmOptions.cmdApply.Enabled = -1
    Unload Me
End Sub

Private Sub cmdSelectShadow_Click()
    lblSelectShadow_Click
End Sub

Private Sub cmdSelectTexture_Click()
    Tags.BrowseTargetForm = 5
    Tags.BrowsePresetPath = GetSetting("DownloadBooster", "Options", "LiveBadukMemoSkinFrameTexture", "")
    frmExplorer.Show vbModal, Me
End Sub

Private Sub Form_Load()
    If GetSetting("DownloadBooster", "Options", "DisableDWMWindow", DefaultDisableDWMWindow) = 1 Then DisableDWMWindow Me.hWnd
    SetFormBackgroundColor Me
    SetFont Me
    SetWindowPos Me.hWnd, IIf(MainFormOnTop, hWnd_TOPMOST, hWnd_NOTOPMOST), 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    
    lblSelectShadow.Top = pgShadow.Top
    lblSelectShadow.Left = pgShadow.Left
    lblSelectShadow.Width = pgShadow.Width
    lblSelectShadow.Height = pgShadow.Height
    pgShadow.BackColor = GetSetting("DownloadBooster", "Options", "LiveBadukMemoSkinShadowColor", 16777215)
    lblSelectFrameColor.Top = pgFrameColor.Top
    lblSelectFrameColor.Left = pgFrameColor.Left
    lblSelectFrameColor.Width = pgFrameColor.Width
    lblSelectFrameColor.Height = pgFrameColor.Height
    pgFrameColor.BackColor = GetSetting("DownloadBooster", "Options", "LiveBadukMemoSkinFrameColor", 16777215)
    Select Case LCase(GetSetting("DownloadBooster", "Options", "LiveBadukMemoSkinFrameType", "transparent"))
        Case "solidcolor"
            optColor.Value = True
        Case "texture"
            optTexture.Value = True
        Case Else
            optTransparent.Value = True
    End Select
    
    tr Label20, "Text s&hadow color:"
    tr Label1, "&Frame color:"
    tr optTransparent, "Semi-&transparent"
    tr optColor, "&Color:"
    tr optTexture, "T&exture:"
    tr cmdSelectTexture, "&Select..."
    tr cmdOK, "OK"
    tr cmdCancel, "Cancel"
    Me.Caption = t(Me.Caption, "LiveBaduk Memo Skin Properties")
End Sub

Private Sub lblSelectFrameColor_Click()
    Dim Color As OLE_COLOR
    Color = ShowColorDialog(Me.hWnd, True, pgFrameColor.BackColor)
    If Color = -1 Then Exit Sub
    pgFrameColor.BackColor = Color
    optColor.Value = True
End Sub

Private Sub lblSelectShadow_Click()
    Dim Color As OLE_COLOR
    Color = ShowColorDialog(Me.hWnd, True, pgShadow.BackColor)
    If Color = -1 Then Exit Sub
    pgShadow.BackColor = Color
End Sub
