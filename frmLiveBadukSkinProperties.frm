VERSION 5.00
Begin VB.Form frmLiveBadukSkinProperties 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "스킨 설정"
   ClientHeight    =   5295
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
   ScaleHeight     =   5295
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin prjDownloadBooster.CheckBoxW chkEnableBorders 
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   4440
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   450
      Caption         =   "쪽지 테두리 표시(&W)"
   End
   Begin prjDownloadBooster.FrameW fFrameBackground 
      Height          =   1335
      Left            =   120
      TabIndex        =   13
      Top             =   3000
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   2355
      Caption         =   " 프레임 배경 "
      Begin VB.OptionButton optBackgroundColor 
         Caption         =   "단색(&O):"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   1695
      End
      Begin prjDownloadBooster.CommandButtonW cmdSelectFrameTexture 
         Height          =   330
         Left            =   2160
         TabIndex        =   18
         Top             =   900
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         Caption         =   "선택(&L)..."
      End
      Begin VB.OptionButton optFrameTexture 
         Caption         =   "텍스처(&U):"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   1335
      End
      Begin VB.OptionButton optFrameTransparent 
         Caption         =   "반투명(&R)"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Value           =   -1  'True
         Width           =   2055
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
         Left            =   2160
         TabIndex        =   16
         Top             =   600
         Width           =   615
      End
      Begin VB.Shape pgFrameBackgroundColor 
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H00404040&
         FillColor       =   &H00808080&
         Height          =   255
         Left            =   2160
         Shape           =   4  '둥근 사각형
         Top             =   600
         Width           =   615
      End
   End
   Begin prjDownloadBooster.FrameW fText 
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   2355
      Caption         =   " 글자 "
      Begin prjDownloadBooster.CheckBoxW chkShadowColor 
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   450
         Value           =   1
         Caption         =   "라벨 그림자 색(&H):"
      End
      Begin prjDownloadBooster.CheckBoxW chkTextColor 
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   450
         Caption         =   "라벨 글자 색(&X):"
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
         Left            =   2160
         TabIndex        =   2
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
         Left            =   2160
         TabIndex        =   4
         Top             =   600
         Width           =   615
      End
      Begin VB.Label lblSelectContentColor 
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
         TabIndex        =   6
         Top             =   960
         Width           =   615
      End
      Begin VB.Shape pgContentColor 
         BackColor       =   &H00000000&
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H00404040&
         FillColor       =   &H00808080&
         Height          =   255
         Left            =   2160
         Shape           =   4  '둥근 사각형
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "현황 글자 색(&P):"
         Height          =   180
         Left            =   390
         TabIndex        =   5
         Top             =   1005
         Width           =   1350
      End
      Begin VB.Shape pgText 
         BackColor       =   &H00000000&
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H00404040&
         FillColor       =   &H00808080&
         Height          =   255
         Left            =   2160
         Shape           =   4  '둥근 사각형
         Top             =   240
         Width           =   615
      End
      Begin VB.Shape pgShadow 
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H00404040&
         FillColor       =   &H00808080&
         Height          =   255
         Left            =   2160
         Shape           =   4  '둥근 사각형
         Top             =   600
         Width           =   615
      End
   End
   Begin prjDownloadBooster.FrameW fFrameColor 
      Height          =   1335
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   2355
      Caption         =   " 프레임 테두리 "
      Begin prjDownloadBooster.CommandButtonW cmdSelectTexture 
         Height          =   330
         Left            =   2160
         TabIndex        =   12
         Top             =   930
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         Caption         =   "선택(&S)..."
      End
      Begin VB.OptionButton optTexture 
         Caption         =   "텍스처(&E):"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   1335
      End
      Begin VB.OptionButton optColor 
         Caption         =   "단색(&C):"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   1335
      End
      Begin VB.OptionButton optTransparent 
         Caption         =   "반투명(&T)"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Value           =   -1  'True
         Width           =   2655
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
         TabIndex        =   10
         Top             =   600
         Width           =   615
      End
      Begin VB.Shape pgFrameColor 
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H00404040&
         FillColor       =   &H00808080&
         Height          =   255
         Left            =   2160
         Shape           =   4  '둥근 사각형
         Top             =   600
         Width           =   615
      End
   End
   Begin prjDownloadBooster.CommandButtonW cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   3120
      TabIndex        =   21
      Top             =   4800
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "취소"
   End
   Begin prjDownloadBooster.CommandButtonW cmdOK 
      Default         =   -1  'True
      Height          =   375
      Left            =   960
      TabIndex        =   20
      Top             =   4800
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
    SaveSetting "DownloadBooster", "Options", "LiveBadukMemoSkinContentTextColor", CLng(pgContentColor.BackColor)
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
    InitForm Me
    
    pgShadow.BackColor = CLng(GetSetting("DownloadBooster", "Options", "LiveBadukMemoSkinShadowColor", 16777215))
    pgFrameColor.BackColor = CLng(GetSetting("DownloadBooster", "Options", "LiveBadukMemoSkinFrameColor", 16777215))
    pgText.BackColor = CLng(GetSetting("DownloadBooster", "Options", "LiveBadukMemoSkinTextColor", 0))
    pgFrameBackgroundColor.BackColor = CLng(GetSetting("DownloadBooster", "Options", "LiveBadukMemoSkinFrameBackgroundColor", 16777215))
    pgContentColor.BackColor = CLng(GetSetting("DownloadBooster", "Options", "LiveBadukMemoSkinContentTextColor", 0))
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
    tr chkTextColor, "Label te&xt color:"
    tr chkShadowColor, "Label s&hadow color:"
    tr fFrameColor, " Frame "
    tr optTransparent, "Semi-&transparent"
    tr optColor, "&Color:"
    tr optTexture, "T&exture:"
    tr cmdSelectTexture, "&Select..."
    tr cmdOK, "OK"
    tr cmdCancel, "Cancel"
    tr Me, "Skin Settings"
    tr fFrameBackground, " Frame Background "
    tr optFrameTransparent, "Semi-t&ransparent"
    tr optFrameTexture, "Text&ure:"
    tr chkEnableBorders, "Sho&w borders"
    tr cmdSelectFrameTexture, "Se&lect..."
    tr optBackgroundColor, "C&olor:"
    tr Label1, "&Progress text color:"
End Sub

Private Sub lblFrameBackgroundColorSelect_Click()
    Dim Color As OLE_COLOR
    Color = ShowColorDialog(Me.hWnd, True, pgFrameBackgroundColor.BackColor)
    If Color = -1 Then Exit Sub
    pgFrameBackgroundColor.BackColor = Color
    optBackgroundColor.Value = True
End Sub

Private Sub lblSelectContentColor_Click()
    Dim Color As OLE_COLOR
    Color = ShowColorDialog(Me.hWnd, True, pgContentColor.BackColor, True)
    If Color = -1 Then Exit Sub
    pgContentColor.BackColor = Color
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
    Color = ShowColorDialog(Me.hWnd, True, pgShadow.BackColor, True)
    If Color = -1 Then Exit Sub
    pgShadow.BackColor = Color
    chkShadowColor.Value = 1
End Sub

Private Sub lblSelectTextColor_Click()
    Dim Color As OLE_COLOR
    Color = ShowColorDialog(Me.hWnd, True, pgText.BackColor, True)
    If Color = -1 Then Exit Sub
    pgText.BackColor = Color
    chkTextColor.Value = 1
End Sub
