VERSION 5.00
Begin VB.Form frmLiveBadukSkinProperties 
   BorderStyle     =   3  'ũ�� ���� ��ȭ ����
   Caption         =   "��Ų ����"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5310
   BeginProperty Font 
      Name            =   "����"
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
   ScaleHeight     =   6000
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '������ ���
   Begin prjDownloadBooster.CheckBoxW chkEnableBorders 
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   5160
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   450
      Caption         =   "���� �׵θ� ǥ��(&W)"
   End
   Begin prjDownloadBooster.FrameW fFrameBackground 
      Height          =   1335
      Left            =   120
      TabIndex        =   13
      Top             =   3720
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   2355
      Caption         =   "������ ���"
      Begin VB.OptionButton optBackgroundColor 
         Caption         =   "�ܻ�(&O):"
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
         Caption         =   "����(&L)..."
      End
      Begin VB.OptionButton optFrameTexture 
         Caption         =   "�ؽ�ó(&U):"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   1335
      End
      Begin VB.OptionButton optFrameTransparent 
         Caption         =   "������(&R)"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Value           =   -1  'True
         Width           =   2055
      End
      Begin VB.Label lblFrameBackgroundColorSelect 
         BackStyle       =   0  '����
         BeginProperty Font 
            Name            =   "����"
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
         BackStyle       =   1  '�������� ����
         BorderColor     =   &H00404040&
         FillColor       =   &H00808080&
         Height          =   255
         Left            =   2160
         Shape           =   4  '�ձ� �簢��
         Top             =   600
         Width           =   615
      End
   End
   Begin prjDownloadBooster.FrameW fText 
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   3625
      Caption         =   "����"
      Begin prjDownloadBooster.CheckBoxW chkEnableFontSize 
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   1320
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   450
         Value           =   1
         Caption         =   "�� ���� ũ��(&S):"
      End
      Begin prjDownloadBooster.CheckBoxW chkBold 
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1680
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   450
         Value           =   1
         Caption         =   "�� ����(&B)"
      End
      Begin prjDownloadBooster.SpinBox sbFontSize 
         Height          =   255
         Left            =   2160
         TabIndex        =   22
         Top             =   1320
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   450
         Min             =   6
         Max             =   14
         Value           =   10
         ThousandsSeparator=   0   'False
         AllowOnlyNumbers=   -1  'True
         TextAlignment   =   1
      End
      Begin prjDownloadBooster.CheckBoxW chkShadowColor 
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   450
         Value           =   1
         Caption         =   "�� �׸��� ��(&H):"
      End
      Begin prjDownloadBooster.CheckBoxW chkTextColor 
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   450
         Caption         =   "�� ���� ��(&X):"
      End
      Begin VB.Label lblSelectTextColor 
         BackStyle       =   0  '����
         BeginProperty Font 
            Name            =   "����"
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
         BackStyle       =   0  '����
         BeginProperty Font 
            Name            =   "����"
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
         BackStyle       =   0  '����
         BeginProperty Font 
            Name            =   "����"
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
         BackStyle       =   1  '�������� ����
         BorderColor     =   &H00404040&
         FillColor       =   &H00808080&
         Height          =   255
         Left            =   2160
         Shape           =   4  '�ձ� �簢��
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "��Ȳ ���� ��(&P):"
         Height          =   180
         Left            =   390
         TabIndex        =   5
         Top             =   1005
         Width           =   1350
      End
      Begin VB.Shape pgText 
         BackColor       =   &H00000000&
         BackStyle       =   1  '�������� ����
         BorderColor     =   &H00404040&
         FillColor       =   &H00808080&
         Height          =   255
         Left            =   2160
         Shape           =   4  '�ձ� �簢��
         Top             =   240
         Width           =   615
      End
      Begin VB.Shape pgShadow 
         BackStyle       =   1  '�������� ����
         BorderColor     =   &H00404040&
         FillColor       =   &H00808080&
         Height          =   255
         Left            =   2160
         Shape           =   4  '�ձ� �簢��
         Top             =   600
         Width           =   615
      End
   End
   Begin prjDownloadBooster.FrameW fFrameColor 
      Height          =   1335
      Left            =   120
      TabIndex        =   7
      Top             =   2280
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   2355
      Caption         =   "������ �׵θ�"
      Begin prjDownloadBooster.CommandButtonW cmdSelectTexture 
         Height          =   330
         Left            =   2160
         TabIndex        =   12
         Top             =   930
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         Caption         =   "����(&S)..."
      End
      Begin VB.OptionButton optTexture 
         Caption         =   "�ؽ�ó(&E):"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   1335
      End
      Begin VB.OptionButton optColor 
         Caption         =   "�ܻ�(&C):"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   1335
      End
      Begin VB.OptionButton optTransparent 
         Caption         =   "������(&T)"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Value           =   -1  'True
         Width           =   2655
      End
      Begin VB.Label lblSelectFrameColor 
         BackStyle       =   0  '����
         BeginProperty Font 
            Name            =   "����"
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
         BackStyle       =   1  '�������� ����
         BorderColor     =   &H00404040&
         FillColor       =   &H00808080&
         Height          =   255
         Left            =   2160
         Shape           =   4  '�ձ� �簢��
         Top             =   600
         Width           =   615
      End
   End
   Begin prjDownloadBooster.CommandButtonW cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   3000
      TabIndex        =   21
      Top             =   5520
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "���"
   End
   Begin prjDownloadBooster.CommandButtonW cmdOK 
      Default         =   -1  'True
      Height          =   375
      Left            =   960
      TabIndex        =   20
      Top             =   5520
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "Ȯ��"
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
    frmOptions.LiveBadukMemoSkinShadowColor = CLng(pgShadow.BackColor)
    frmOptions.LiveBadukMemoSkinFrameColor = CLng(pgFrameColor.BackColor)
    frmOptions.LiveBadukMemoSkinFrameType = IIf(optTransparent.Value, "transparent", IIf(optColor.Value, "solidcolor", "texture"))
    frmOptions.LiveBadukMemoSkinTextColor = CLng(pgText.BackColor)
    frmOptions.LiveBadukMemoSkinEnableShadow = chkShadowColor.Value
    frmOptions.LiveBadukMemoSkinEnableTextColor = chkTextColor.Value
    frmOptions.LiveBadukMemoSkinEnableBorder = chkEnableBorders.Value
    frmOptions.LiveBadukMemoSkinFrameBackgroundType = IIf(optFrameTransparent.Value, "transparent", IIf(optBackgroundColor.Value, "solidcolor", "texture"))
    frmOptions.LiveBadukMemoSkinFrameBackgroundColor = CLng(pgFrameBackgroundColor.BackColor)
    frmOptions.LiveBadukMemoSkinContentTextColor = CLng(pgContentColor.BackColor)
    frmOptions.LiveBadukMemoSkinEnableLabelFontSize = chkEnableFontSize.Value
    frmOptions.LiveBadukMemoSkinLabelFontSize = sbFontSize.Value
    frmOptions.LiveBadukMemoSkinLabelFontBold = chkBold.Value
    
    frmOptions.ColorChanged = True
    frmOptions.FontChanged = True
    frmOptions.ProgressSkinChanged = True
    frmOptions.cmdApply.Enabled = -1
    Unload Me
End Sub

Private Sub cmdSelectFrameTexture_Click()
    ShowFileDialog 6, frmOptions.LiveBadukMemoSkinFrameBackground
End Sub

Private Sub cmdSelectTexture_Click()
    ShowFileDialog 5, frmOptions.LiveBadukMemoSkinFrameTexture
End Sub

Private Sub Form_Load()
    InitForm Me
    
    pgShadow.BackColor = frmOptions.LiveBadukMemoSkinShadowColor
    pgFrameColor.BackColor = frmOptions.LiveBadukMemoSkinFrameColor
    pgText.BackColor = frmOptions.LiveBadukMemoSkinTextColor
    pgFrameBackgroundColor.BackColor = frmOptions.LiveBadukMemoSkinFrameBackgroundColor
    pgContentColor.BackColor = frmOptions.LiveBadukMemoSkinContentTextColor
    chkShadowColor.Value = frmOptions.LiveBadukMemoSkinEnableShadow
    chkTextColor.Value = frmOptions.LiveBadukMemoSkinEnableTextColor
    Select Case frmOptions.LiveBadukMemoSkinFrameType
        Case "solidcolor"
            optColor.Value = True
        Case "texture"
            optTexture.Value = True
        Case Else
            optTransparent.Value = True
    End Select
    Select Case frmOptions.LiveBadukMemoSkinFrameBackgroundType
        Case "texture"
            optFrameTexture.Value = True
        Case "solidcolor"
            optBackgroundColor.Value = True
        Case Else
            optFrameTransparent.Value = True
    End Select
    chkEnableBorders.Value = frmOptions.LiveBadukMemoSkinEnableBorder
    
    chkEnableFontSize.Value = frmOptions.LiveBadukMemoSkinEnableLabelFontSize
    sbFontSize.Value = frmOptions.LiveBadukMemoSkinLabelFontSize
    chkBold.Value = frmOptions.LiveBadukMemoSkinLabelFontBold
    
    tr fText, "Text"
    tr chkTextColor, "Label te&xt color:"
    tr chkShadowColor, "Label s&hadow color:"
    tr fFrameColor, "Frame"
    tr optTransparent, "Semi-&transparent"
    tr optColor, "&Color:"
    tr optTexture, "T&exture:"
    tr cmdSelectTexture, "&Select..."
    tr cmdOK, "OK"
    tr cmdCancel, "Cancel"
    tr Me, "Skin Settings"
    tr fFrameBackground, "Frame Background"
    tr optFrameTransparent, "Semi-t&ransparent"
    tr optFrameTexture, "Text&ure:"
    tr chkEnableBorders, "Sho&w borders"
    tr cmdSelectFrameTexture, "Se&lect..."
    tr optBackgroundColor, "C&olor:"
    tr Label1, "&Progress text color:"
    tr chkEnableFontSize, "Label text &size:"
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

Private Sub sbFontSize_Change()
    chkEnableFontSize.Value = 1
End Sub

