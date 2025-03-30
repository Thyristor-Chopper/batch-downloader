VERSION 5.00
Begin VB.Form frmLiveBadukSkinProperties 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "라이브 바둑 쪽지 스킨 등록 정보"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5415
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
   ScaleHeight     =   4905
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin prjDownloadBooster.CheckBoxW chkEnableBorders 
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   4080
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   450
      Caption         =   "쪽지 테두리 표시(&W)"
   End
   Begin prjDownloadBooster.FrameW fFrameBackground 
      Height          =   1335
      Left            =   120
      TabIndex        =   11
      Top             =   2640
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   2355
      Caption         =   " 프레임 배경 "
      Begin prjDownloadBooster.OptionButtonW optBackgroundColor 
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         Caption         =   "단색(&O):"
      End
      Begin prjDownloadBooster.CommandButtonW cmdSelectFrameTexture 
         Height          =   330
         Left            =   1920
         TabIndex        =   16
         Top             =   900
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         Caption         =   "선택(&L)..."
      End
      Begin prjDownloadBooster.OptionButtonW optFrameTexture 
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         Caption         =   "텍스처(&U):"
      End
      Begin prjDownloadBooster.OptionButtonW optFrameTransparent 
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   450
         Value           =   -1  'True
         Caption         =   "반투명(&R)"
      End
      Begin VB.Label lblFrameBackgroundColorSelect 
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
         Left            =   2520
         TabIndex        =   14
         Top             =   600
         Width           =   855
      End
      Begin VB.Shape pgFrameBackgroundColor 
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H00404040&
         FillColor       =   &H00808080&
         Height          =   255
         Left            =   1920
         Shape           =   4  '둥근 사각형
         Top             =   600
         Width           =   615
      End
   End
   Begin prjDownloadBooster.FrameW fText 
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   1720
      Caption         =   " 글자 "
      Begin prjDownloadBooster.CheckBoxW chkShadowColor 
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1695
         _ExtentX        =   0
         _ExtentY        =   0
         Value           =   1
         Caption         =   "그림자 색(&H):"
      End
      Begin prjDownloadBooster.CheckBoxW chkTextColor 
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1695
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "글자 색(&X):"
      End
      Begin VB.Label lblSelectTextColor 
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
         Left            =   2520
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin VB.Shape pgText 
         BackColor       =   &H00000000&
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H00404040&
         FillColor       =   &H00808080&
         Height          =   255
         Left            =   1920
         Shape           =   4  '둥근 사각형
         Top             =   240
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
         Left            =   2520
         TabIndex        =   4
         Top             =   600
         Width           =   855
      End
      Begin VB.Shape pgShadow 
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H00404040&
         FillColor       =   &H00808080&
         Height          =   255
         Left            =   1920
         Shape           =   4  '둥근 사각형
         Top             =   600
         Width           =   615
      End
   End
   Begin prjDownloadBooster.FrameW fFrameColor 
      Height          =   1335
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   2355
      Caption         =   " 프레임 테두리 "
      Begin prjDownloadBooster.CommandButtonW cmdSelectTexture 
         Height          =   330
         Left            =   1920
         TabIndex        =   10
         Top             =   930
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         Caption         =   "선택(&S)..."
      End
      Begin prjDownloadBooster.OptionButtonW optTexture 
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   1335
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "텍스처(&E):"
      End
      Begin prjDownloadBooster.OptionButtonW optColor 
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   1335
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "단색(&C):"
      End
      Begin prjDownloadBooster.OptionButtonW optTransparent 
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   2655
         _ExtentX        =   0
         _ExtentY        =   0
         Value           =   -1  'True
         Caption         =   "반투명(&T)"
      End
      Begin VB.Shape pgFrameColor 
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H00404040&
         FillColor       =   &H00808080&
         Height          =   255
         Left            =   1920
         Shape           =   4  '둥근 사각형
         Top             =   600
         Width           =   615
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
         Left            =   2160
         TabIndex        =   8
         Top             =   600
         Width           =   855
      End
   End
   Begin prjDownloadBooster.CommandButtonW cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   3120
      TabIndex        =   19
      Top             =   4440
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "취소"
   End
   Begin prjDownloadBooster.CommandButtonW cmdOK 
      Default         =   -1  'True
      Height          =   375
      Left            =   960
      TabIndex        =   18
      Top             =   4440
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "확인"
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
    SaveSetting "DownloadBooster", "Options", "LiveBadukMemoSkinTextColor", CLng(pgText.BackColor)
    SaveSetting "DownloadBooster", "Options", "LiveBadukMemoSkinEnableShadow", chkShadowColor.Value
    SaveSetting "DownloadBooster", "Options", "LiveBadukMemoSkinEnableTextColor", chkTextColor.Value
    SaveSetting "DownloadBooster", "Options", "LiveBadukMemoSkinEnableBorder", chkEnableBorders.Value
    SaveSetting "DownloadBooster", "Options", "LiveBadukMemoSkinFrameBackgroundType", IIf(optFrameTransparent.Value, "transparent", IIf(optBackgroundColor.Value, "solidcolor", "texture"))
    SaveSetting "DownloadBooster", "Options", "LiveBadukMemoSkinFrameBackgroundColor", CLng(pgFrameBackgroundColor.BackColor)
    frmOptions.ColorChanged = True
    frmOptions.cmdApply.Enabled = -1
    Unload Me
End Sub

Private Sub cmdSelectFrameTexture_Click()
    Tags.BrowseTargetForm = 6
    Tags.BrowsePresetPath = GetSetting("DownloadBooster", "Options", "LiveBadukMemoSkinFrameBackground", "")
    frmExplorer.Show vbModal, Me
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
    pgShadow.BackColor = CLng(GetSetting("DownloadBooster", "Options", "LiveBadukMemoSkinShadowColor", 16777215))
    lblSelectFrameColor.Top = pgFrameColor.Top
    lblSelectFrameColor.Left = pgFrameColor.Left
    lblSelectFrameColor.Width = pgFrameColor.Width
    lblSelectFrameColor.Height = pgFrameColor.Height
    pgFrameColor.BackColor = CLng(GetSetting("DownloadBooster", "Options", "LiveBadukMemoSkinFrameColor", 16777215))
    lblSelectTextColor.Top = pgText.Top
    lblSelectTextColor.Left = pgText.Left
    lblSelectTextColor.Width = pgText.Width
    lblSelectTextColor.Height = pgText.Height
    pgText.BackColor = CLng(GetSetting("DownloadBooster", "Options", "LiveBadukMemoSkinTextColor", 0))
    lblFrameBackgroundColorSelect.Top = pgFrameBackgroundColor.Top
    lblFrameBackgroundColorSelect.Left = pgFrameBackgroundColor.Left
    lblFrameBackgroundColorSelect.Width = pgFrameBackgroundColor.Width
    lblFrameBackgroundColorSelect.Height = pgFrameBackgroundColor.Height
    pgFrameBackgroundColor.BackColor = CLng(GetSetting("DownloadBooster", "Options", "LiveBadukMemoSkinFrameBackgroundColor", 16777215))
    chkShadowColor.Value = CInt(GetSetting("DownloadBooster", "Options", "LiveBadukMemoSkinEnableShadow", 1))
    chkTextColor.Value = CInt(GetSetting("DownloadBooster", "Options", "LiveBadukMemoSkinEnableTextColor", 0))
    Select Case LCase(GetSetting("DownloadBooster", "Options", "LiveBadukMemoSkinFrameType", "transparent"))
        Case "solidcolor"
            optColor.Value = True
        Case "texture"
            optTexture.Value = True
        Case Else
            optTransparent.Value = True
    End Select
    Select Case LCase(GetSetting("DownloadBooster", "Options", "LiveBadukMemoSkinFrameBackgroundType", "transparent"))
        Case "texture"
            optFrameTexture.Value = True
        Case "solidcolor"
            optBackgroundColor.Value = True
        Case Else
            optFrameTransparent.Value = True
    End Select
    chkEnableBorders.Value = CInt(GetSetting("DownloadBooster", "Options", "LiveBadukMemoSkinEnableBorder", 1))
    
    tr fText, " Text "
    tr chkTextColor, "Te&xt color:"
    tr chkShadowColor, "S&hadow color:"
    tr fFrameColor, " Frame "
    tr optTransparent, "Semi-&transparent"
    tr optColor, "&Color:"
    tr optTexture, "T&exture:"
    tr cmdSelectTexture, "&Select..."
    tr cmdOK, "OK"
    tr cmdCancel, "Cancel"
    Me.Caption = t(Me.Caption, "LiveBaduk Memo Skin Properties")
    tr fFrameBackground, " Frame Background "
    tr optFrameTransparent, "Semi-t&ransparent"
    tr optFrameTexture, "Text&ure:"
    tr chkEnableBorders, "Sho&w borders"
    tr cmdSelectFrameTexture, "Se&lect..."
    tr optBackgroundColor, "C&olor:"
End Sub

Private Sub lblFrameBackgroundColorSelect_Click()
    Dim Color As OLE_COLOR
    Color = ShowColorDialog(Me.hWnd, True, pgFrameBackgroundColor.BackColor)
    If Color = -1 Then Exit Sub
    pgFrameBackgroundColor.BackColor = Color
    optBackgroundColor.Value = True
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
    chkShadowColor.Value = 1
End Sub

Private Sub lblSelectTextColor_Click()
    Dim Color As OLE_COLOR
    Color = ShowColorDialog(Me.hWnd, True, pgText.BackColor)
    If Color = -1 Then Exit Sub
    pgText.BackColor = Color
    chkTextColor.Value = 1
End Sub
