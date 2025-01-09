VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "스킨 설정 "
   ClientHeight    =   4905
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6135
   BeginProperty Font 
      Name            =   "굴림"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.PictureBox pbPanel 
      Height          =   1185
      Index           =   2
      Left            =   3600
      ScaleHeight     =   1125
      ScaleWidth      =   3675
      TabIndex        =   5
      Top             =   2400
      Visible         =   0   'False
      Width           =   3735
      Begin VB.CheckBox chkRememberURL 
         Caption         =   "파일 주소 기억(&M)"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   600
         Width           =   2055
      End
      Begin VB.CheckBox chkNoCleanup 
         Caption         =   "조각 파일 유지(&N)"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   2130
      End
      Begin VB.Frame Frame2 
         Caption         =   " 다운로드 설정 "
         Height          =   855
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   3495
      End
   End
   Begin VB.PictureBox pbPanel 
      BorderStyle     =   0  '없음
      Height          =   3825
      Index           =   1
      Left            =   165
      ScaleHeight     =   3825
      ScaleWidth      =   5790
      TabIndex        =   4
      Top             =   435
      Width           =   5790
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  '없음
         Height          =   735
         Left            =   240
         ScaleHeight     =   735
         ScaleWidth      =   1575
         TabIndex        =   18
         Top             =   1680
         Width           =   1575
         Begin VB.OptionButton optUserFore 
            Caption         =   "사용자 지정(&E)"
            Height          =   255
            Left            =   0
            TabIndex        =   20
            Top             =   450
            Width           =   1575
         End
         Begin VB.OptionButton optSystemFore 
            Caption         =   "시스템 색상(&Y)"
            Height          =   255
            Left            =   0
            TabIndex        =   19
            Top             =   0
            Width           =   1815
         End
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  '없음
         Height          =   735
         Left            =   240
         ScaleHeight     =   735
         ScaleWidth      =   1575
         TabIndex        =   15
         Top             =   360
         Width           =   1575
         Begin VB.OptionButton optSystemColor 
            Caption         =   "시스템 색상(&S)"
            Height          =   255
            Left            =   0
            TabIndex        =   17
            Top             =   0
            Width           =   1815
         End
         Begin VB.OptionButton optUserColor 
            Caption         =   "사용자 지정(&U)"
            Height          =   255
            Left            =   0
            TabIndex        =   16
            Top             =   450
            Width           =   1575
         End
      End
      Begin VB.CheckBox chkNoDWMWindow 
         Caption         =   "윈도우 7 모양으로 바꾸기(&I)"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   3000
         Width           =   2655
      End
      Begin VB.Frame Frame3 
         Caption         =   " 스타일 "
         Height          =   615
         Left            =   120
         TabIndex        =   9
         Top             =   2760
         Width           =   3015
      End
      Begin VB.Frame Frame1 
         Caption         =   " 배경색 "
         Height          =   1215
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   3375
         Begin VB.Label lblSelectColor 
            BackStyle       =   0  '투명
            Height          =   495
            Left            =   1800
            TabIndex        =   12
            Top             =   600
            Width           =   1455
         End
         Begin VB.Shape pgColor 
            BackStyle       =   1  '투명하지 않음
            BorderColor     =   &H00404040&
            FillColor       =   &H00808080&
            FillStyle       =   2  '수평선
            Height          =   375
            Left            =   1800
            Shape           =   4  '둥근 사각형
            Top             =   630
            Width           =   1455
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   " 글자색 "
         Height          =   1215
         Left            =   120
         TabIndex        =   13
         Top             =   1440
         Width           =   3375
         Begin VB.Label lblSelectFore 
            BackStyle       =   0  '투명
            Height          =   495
            Left            =   1800
            TabIndex        =   14
            Top             =   600
            Width           =   1455
         End
         Begin VB.Shape pgFore 
            BackStyle       =   1  '투명하지 않음
            BorderColor     =   &H00404040&
            FillColor       =   &H00808080&
            FillStyle       =   2  '수평선
            Height          =   375
            Left            =   1800
            Shape           =   4  '둥근 사각형
            Top             =   630
            Width           =   1455
         End
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "적용(&A)"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4560
      TabIndex        =   3
      Top             =   4440
      Width           =   1455
   End
   Begin prjDownloadBooster.TabStrip tsTabStrip 
      Height          =   4215
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   7435
      TabFixedWidth   =   53
      InitTabs        =   "frmOptions.frx":0442
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "취소"
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "확인"
      Default         =   -1  'True
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   4440
      Width           =   1455
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub chkNoCleanup_Click()
    cmdApply.Enabled = -1
End Sub

Private Sub chkNoDWMWindow_Click()
    cmdApply.Enabled = -1
End Sub

Private Sub chkRememberURL_Click()
    cmdApply.Enabled = -1
End Sub

Private Sub cmdApply_Click()
    SaveSetting "DownloadBooster", "Options", "NoCleanup", chkNoCleanup.Value
    If WinVer >= 6.1 Then SaveSetting "DownloadBooster", "Options", "DisableDWMWindow", chkNoDWMWindow.Value
    SaveSetting "DownloadBooster", "Options", "RememberURL", chkRememberURL.Value
    If chkNoDWMWindow.Value Then
        DisableDWMWindow Me.hWnd
        DisableDWMWindow frmMain.hWnd
    Else
        EnableDWMWindow Me.hWnd
        EnableDWMWindow frmMain.hWnd
    End If
    If optSystemColor.Value Then
        SaveSetting "DownloadBooster", "Options", "BackColor", "-1"
        pgColor.BackColor = &H8000000F
    ElseIf optUserColor.Value Then
        SaveSetting "DownloadBooster", "Options", "BackColor", CLng(pgColor.BackColor)
    End If
    If optSystemFore.Value Then
        SaveSetting "DownloadBooster", "Options", "ForeColor", "-1"
        pgFore.BackColor = &H80000012
    ElseIf optUserFore.Value Then
        SaveSetting "DownloadBooster", "Options", "ForeColor", CLng(pgFore.BackColor)
    End If
    SetFormBackgroundColor Me
    SetFormBackgroundColor frmMain
    cmdApply.Enabled = 0
End Sub

Private Sub Form_Load()
    If GetSetting("DownloadBooster", "Options", "DisableDWMWindow", DefaultDisableDWMWindow) = 1 Then DisableDWMWindow Me.hWnd
    SetFormBackgroundColor Me
    
    pbPanel(2).Top = pbPanel(1).Top
    pbPanel(2).Left = pbPanel(1).Left
    pbPanel(2).Width = pbPanel(1).Width
    pbPanel(2).Height = pbPanel(1).Height
    pbPanel(2).BorderStyle = 0
    If WinVer < 6.1 Then chkNoDWMWindow.Enabled = 0
    chkNoCleanup.Value = GetSetting("DownloadBooster", "Options", "NoCleanup", 0)
    chkNoDWMWindow.Value = GetSetting("DownloadBooster", "Options", "DisableDWMWindow", DefaultDisableDWMWindow)
    If WinVer < 6.1 Then chkNoDWMWindow.Value = 0
    chkRememberURL.Value = GetSetting("DownloadBooster", "Options", "RememberURL", 0)
    If WinVer < 6.2 Then chkNoDWMWindow.Caption = "Aero 효과 사용 안 함(&I)"
    
    Dim clrBackColor As Long
    clrBackColor = GetSetting("DownloadBooster", "Options", "BackColor", DefaultBackColor)
    If clrBackColor < 0 Or clrBackColor > 16777215 Then
        optSystemColor.Value = True
        pgColor.BackColor = &H8000000F
    Else
        optUserColor.Value = True
        pgColor.BackColor = clrBackColor
    End If
    
    Dim clrForeColor As Long
    clrForeColor = GetSetting("DownloadBooster", "Options", "ForeColor", -1)
    If clrForeColor < 0 Or clrForeColor > 16777215 Then
        optSystemFore.Value = True
        pgFore.BackColor = &H80000012
    Else
        optUserFore.Value = True
        pgFore.BackColor = clrForeColor
    End If
    
    cmdApply.Enabled = 0
End Sub

Private Sub lblSelectColor_Click()
    Dim Color As OLE_COLOR
    Color = ShowColorDialog(Me.hWnd, True, pgColor.BackColor)
    If Color = -1 Then Exit Sub
    pgColor.BackColor = Color
    cmdApply.Enabled = -1
    optUserColor.Value = True
End Sub

Private Sub lblSelectFore_Click()
    Dim Color As OLE_COLOR
    Color = ShowColorDialog(Me.hWnd, True, pgFore.BackColor)
    If Color = -1 Then Exit Sub
    pgFore.BackColor = Color
    cmdApply.Enabled = -1
    optUserFore.Value = True
End Sub

Private Sub OKButton_Click()
    cmdApply_Click
    Unload Me
End Sub

Private Sub optSystemColor_Click()
    cmdApply.Enabled = -1
End Sub

Private Sub optSystemFore_Click()
    cmdApply.Enabled = -1
End Sub

Private Sub optUserColor_Click()
    cmdApply.Enabled = -1
End Sub

Private Sub optUserFore_Click()
    cmdApply.Enabled = -1
End Sub

Private Sub tsTabStrip_TabClick(ByVal TabItem As TbsTab)
    pbPanel(1).Visible = 0
    pbPanel(2).Visible = 0
    pbPanel(TabItem.Index).Visible = -1
End Sub
