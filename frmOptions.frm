VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'ũ�� ���� ��ȭ ����
   Caption         =   "�ɼ�"
   ClientHeight    =   5490
   ClientLeft      =   2760
   ClientTop       =   3855
   ClientWidth     =   12405
   BeginProperty Font 
      Name            =   "����"
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
   ScaleHeight     =   5490
   ScaleWidth      =   12405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '������ ���
   Begin prjDownloadBooster.TygemButton tygOK 
      Height          =   360
      Left            =   2040
      TabIndex        =   52
      Top             =   5040
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   635
      Caption         =   "Ȯ��"
   End
   Begin prjDownloadBooster.TygemButton tygCancel 
      Height          =   360
      Left            =   3480
      TabIndex        =   51
      Top             =   5040
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   635
      Caption         =   "���"
   End
   Begin prjDownloadBooster.TygemButton tygApply 
      Height          =   360
      Left            =   4920
      TabIndex        =   50
      Top             =   5040
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   635
      Enabled         =   0   'False
      Caption         =   "����"
   End
   Begin VB.PictureBox pbPanel 
      AutoRedraw      =   -1  'True
      Height          =   2055
      Index           =   3
      Left            =   6360
      ScaleHeight     =   1995
      ScaleWidth      =   5955
      TabIndex        =   35
      Top             =   120
      Width           =   6015
      Begin prjDownloadBooster.FrameW FrameW2 
         Height          =   1335
         Left            =   120
         TabIndex        =   36
         Top             =   600
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   2355
         Caption         =   " ��� ���� "
         Transparent     =   -1  'True
         Begin prjDownloadBooster.TextBoxW txtYtdlPath 
            Height          =   255
            Left            =   2040
            TabIndex        =   42
            Top             =   960
            Visible         =   0   'False
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   450
         End
         Begin prjDownloadBooster.TextBoxW txtNodePath 
            Height          =   255
            Left            =   2040
            TabIndex        =   38
            Top             =   240
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   450
         End
         Begin prjDownloadBooster.TextBoxW txtScriptPath 
            Height          =   255
            Left            =   2040
            TabIndex        =   39
            Top             =   600
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   450
         End
         Begin VB.Label Label7 
            BackStyle       =   0  '����
            Caption         =   "&youtube-dl/yt-dlp:"
            Height          =   255
            Left            =   120
            TabIndex        =   43
            Top             =   990
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.Label Label5 
            BackStyle       =   0  '����
            Caption         =   "�ٿ�ε� ��ũ��Ʈ(&D):"
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   630
            Width           =   1935
         End
         Begin VB.Label Label4 
            BackStyle       =   0  '����
            Caption         =   "&Node.js:"
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   270
            Width           =   1455
         End
      End
      Begin VB.Label Label6 
         BackStyle       =   0  '����
         Caption         =   "�⺻���� ����Ϸ��� �ʵ带 ����νʽÿ�. �� �ɼ��� ��� ����ڸ�     ���� ���̸� �Ϲ������� ������ �ʿ䰡 �����ϴ�."
         Height          =   480
         Left            =   120
         TabIndex        =   41
         Top             =   120
         Width           =   5775
      End
   End
   Begin VB.PictureBox pbPanel 
      AutoRedraw      =   -1  'True
      Height          =   2865
      Index           =   1
      Left            =   6360
      ScaleHeight     =   2805
      ScaleWidth      =   5595
      TabIndex        =   5
      Top             =   2280
      Width           =   5655
      Begin prjDownloadBooster.FrameW Frame5 
         Height          =   675
         Left            =   120
         TabIndex        =   21
         Top             =   1935
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   1191
         Caption         =   " �������̽� "
         Transparent     =   -1  'True
         Begin prjDownloadBooster.ComboBoxW cbLanguage 
            Height          =   300
            Left            =   1080
            TabIndex        =   23
            Top             =   240
            Width           =   1935
            _ExtentX        =   0
            _ExtentY        =   0
            Style           =   2
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '����
            Caption         =   "���(&L):"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Tag             =   "nocolorchange"
            Top             =   285
            Width           =   975
         End
      End
      Begin prjDownloadBooster.CheckBoxW chkRememberURL 
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1080
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   450
         Caption         =   "���� �ּ� ���(&M)"
         Transparent     =   -1  'True
      End
      Begin prjDownloadBooster.CheckBoxW chkNoCleanup 
         Height          =   255
         Left            =   2640
         TabIndex        =   8
         Top             =   600
         Width           =   2250
         _ExtentX        =   3969
         _ExtentY        =   450
         Caption         =   "���� ���� ����(&N)"
         Transparent     =   -1  'True
      End
      Begin prjDownloadBooster.FrameW Frame2 
         Height          =   1710
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   3016
         Caption         =   " �ٿ�ε� ���� "
         Begin prjDownloadBooster.ComboBoxW cbWhenExist 
            Height          =   300
            Left            =   2055
            TabIndex        =   31
            Top             =   1320
            Width           =   1455
            _ExtentX        =   0
            _ExtentY        =   0
            Style           =   2
         End
         Begin prjDownloadBooster.CheckBoxW chkAutoRetry 
            Height          =   255
            Left            =   2520
            TabIndex        =   29
            Top             =   720
            Width           =   2625
            _ExtentX        =   4630
            _ExtentY        =   450
            Caption         =   "��Ʈ��ũ ���� �� ��õ�(&U)"
            Transparent     =   -1  'True
         End
         Begin prjDownloadBooster.CheckBoxW chkAlwaysResume 
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   720
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   450
            Caption         =   "�׻� �̾�ޱ�(&A)"
            Transparent     =   -1  'True
         End
         Begin prjDownloadBooster.CheckBoxW chkBeepWhenComplete 
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   480
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   450
            Caption         =   "�Ϸ� �� ��ȣ�� ���(&B)"
            Transparent     =   -1  'True
         End
         Begin prjDownloadBooster.CheckBoxW chkOpenDirWhenComplete 
            Height          =   255
            Left            =   2520
            TabIndex        =   26
            Top             =   240
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   450
            Caption         =   "�Ϸ� �� ���� ����(&P)"
            Transparent     =   -1  'True
         End
         Begin prjDownloadBooster.CheckBoxW chkOpenWhenComplete 
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   450
            Caption         =   "�Ϸ� �� ���� ����(&O)"
            Transparent     =   -1  'True
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '����
            Caption         =   "�ߺ� ���ϸ� ó��(&D):"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Tag             =   "nocolorchange"
            Top             =   1365
            Width           =   1935
         End
      End
   End
   Begin VB.PictureBox pbPanel 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  '����
      Height          =   4425
      Index           =   2
      Left            =   165
      ScaleHeight     =   4425
      ScaleWidth      =   6030
      TabIndex        =   4
      Top             =   450
      Visible         =   0   'False
      Width           =   6030
      Begin VB.PictureBox pbBackground 
         BorderStyle     =   0  '����
         Height          =   1335
         Left            =   120
         ScaleHeight     =   1335
         ScaleWidth      =   4395
         TabIndex        =   44
         Top             =   120
         Width           =   4395
         Begin prjDownloadBooster.TygemButton tygSample 
            Height          =   285
            Left            =   2340
            TabIndex        =   54
            Top             =   780
            Visible         =   0   'False
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   503
            Caption         =   "����"
         End
         Begin prjDownloadBooster.CommandButtonW cmdSample 
            Height          =   285
            Left            =   2340
            TabIndex        =   45
            TabStop         =   0   'False
            Top             =   780
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   503
            Caption         =   "����"
            Transparent     =   -1  'True
         End
         Begin VB.Shape Shape4 
            BackColor       =   &H000000FF&
            BackStyle       =   1  '�������� ����
            Height          =   135
            Left            =   3960
            Shape           =   3  '����
            Top             =   120
            Width           =   135
         End
         Begin VB.Line lnForePreview 
            BorderWidth     =   2
            X1              =   300
            X2              =   2100
            Y1              =   540
            Y2              =   540
         End
         Begin VB.Image imgPreview 
            Height          =   375
            Left            =   420
            Stretch         =   -1  'True
            Top             =   540
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000009&
            BorderWidth     =   2
            X1              =   300
            X2              =   2100
            Y1              =   180
            Y2              =   180
         End
         Begin VB.Shape pgBackPreview 
            BackColor       =   &H8000000F&
            BackStyle       =   1  '�������� ����
            Height          =   885
            Left            =   120
            Top             =   330
            Width           =   4095
         End
         Begin VB.Shape Shape3 
            BackColor       =   &H80000002&
            BackStyle       =   1  '�������� ����
            BorderStyle     =   0  '����
            Height          =   495
            Left            =   75
            Top             =   285
            Width           =   4200
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H80000002&
            BackStyle       =   1  '�������� ����
            Height          =   615
            Left            =   60
            Shape           =   4  '�ձ� �簢��
            Top             =   60
            Width           =   4215
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H80000002&
            BackStyle       =   1  '�������� ����
            Height          =   975
            Left            =   60
            Top             =   300
            Width           =   4215
         End
      End
      Begin prjDownloadBooster.FrameW FrameW1 
         Height          =   975
         Left            =   3000
         TabIndex        =   32
         Top             =   2280
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   1720
         Caption         =   " ��� �׸� "
         Transparent     =   -1  'True
         Begin prjDownloadBooster.TygemButton tygChooseBackground 
            Height          =   330
            Left            =   2040
            TabIndex        =   53
            Top             =   210
            Visible         =   0   'False
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   582
            Caption         =   "..."
         End
         Begin prjDownloadBooster.ComboBoxW cbImagePosition 
            Height          =   300
            Left            =   960
            TabIndex        =   47
            Top             =   600
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   529
            Style           =   2
            Text            =   "ComboBoxW1"
         End
         Begin prjDownloadBooster.CommandButtonW cmdChooseBackground 
            Height          =   330
            Left            =   2040
            TabIndex        =   34
            Top             =   210
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   582
            Caption         =   "..."
         End
         Begin prjDownloadBooster.CheckBoxW chkEnableBackgroundImage 
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   240
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   450
            Caption         =   "��� �׸� ���(&B)"
            Transparent     =   -1  'True
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '����
            Caption         =   "��ġ(&P):"
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   645
            Width           =   840
         End
      End
      Begin prjDownloadBooster.FrameW Frame6 
         Height          =   975
         Left            =   3000
         TabIndex        =   24
         Top             =   3360
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   1720
         Caption         =   " ��Ų "
         Transparent     =   -1  'True
         Begin prjDownloadBooster.ComboBoxW cbSkin 
            Height          =   300
            Left            =   360
            TabIndex        =   49
            Top             =   480
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   529
            Style           =   2
            Text            =   "ComboBoxW1"
         End
         Begin VB.Label Label8 
            BackStyle       =   0  '����
            Caption         =   "��Ų ����(&K):"
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.PictureBox pbOptionContainer 
         BorderStyle     =   0  '����
         Height          =   615
         Index           =   2
         Left            =   240
         ScaleHeight     =   615
         ScaleWidth      =   1680
         TabIndex        =   18
         Top             =   3600
         Width           =   1680
         Begin prjDownloadBooster.OptionButtonW optUserFore 
            Height          =   255
            Left            =   0
            TabIndex        =   20
            Top             =   330
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   450
            Caption         =   "����� ����(&T):"
            Transparent     =   -1  'True
         End
         Begin prjDownloadBooster.OptionButtonW optSystemFore 
            Height          =   255
            Left            =   0
            TabIndex        =   19
            Top             =   0
            Width           =   1815
            _ExtentX        =   0
            _ExtentY        =   0
            Caption         =   "�ý��� ����(&Y)"
            Transparent     =   -1  'True
         End
      End
      Begin VB.PictureBox pbOptionContainer 
         BorderStyle     =   0  '����
         Height          =   615
         Index           =   1
         Left            =   240
         ScaleHeight     =   615
         ScaleWidth      =   1680
         TabIndex        =   15
         Top             =   2520
         Width           =   1680
         Begin prjDownloadBooster.OptionButtonW optSystemColor 
            Height          =   255
            Left            =   0
            TabIndex        =   17
            Top             =   0
            Width           =   1815
            _ExtentX        =   0
            _ExtentY        =   0
            Caption         =   "�ý��� ����(&S)"
            Transparent     =   -1  'True
         End
         Begin prjDownloadBooster.OptionButtonW optUserColor 
            Height          =   255
            Left            =   0
            TabIndex        =   16
            Top             =   330
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   450
            Caption         =   "����� ����(&U):"
            Transparent     =   -1  'True
         End
      End
      Begin prjDownloadBooster.CheckBoxW chkNoDWMWindow 
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1800
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   450
         Caption         =   "������ 7 ������� �ٲٱ�(&I)"
         Transparent     =   -1  'True
      End
      Begin prjDownloadBooster.FrameW Frame3 
         Height          =   615
         Left            =   120
         TabIndex        =   9
         Top             =   1560
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   1085
         Caption         =   " â ��� "
      End
      Begin prjDownloadBooster.FrameW Frame1 
         Height          =   975
         Left            =   120
         TabIndex        =   6
         Top             =   2280
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   1720
         Caption         =   " ���� "
         Begin VB.Label lblSelectColor 
            BackStyle       =   0  '����
            Height          =   255
            Left            =   1800
            TabIndex        =   12
            Top             =   240
            Width           =   1455
         End
         Begin VB.Shape pgColor 
            BackStyle       =   1  '�������� ����
            BorderColor     =   &H00404040&
            FillColor       =   &H00808080&
            Height          =   255
            Left            =   1920
            Shape           =   4  '�ձ� �簢��
            Top             =   585
            Width           =   495
         End
      End
      Begin prjDownloadBooster.FrameW Frame4 
         Height          =   975
         Left            =   120
         TabIndex        =   13
         Top             =   3360
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   1720
         Caption         =   " ���ڻ� "
         Begin VB.Label lblSelectFore 
            BackStyle       =   0  '����
            Height          =   255
            Left            =   1800
            TabIndex        =   14
            Top             =   240
            Width           =   1455
         End
         Begin VB.Shape pgFore 
            BackStyle       =   1  '�������� ����
            BorderColor     =   &H00404040&
            FillColor       =   &H00808080&
            Height          =   255
            Left            =   1920
            Shape           =   4  '�ձ� �簢��
            Top             =   585
            Width           =   495
         End
      End
   End
   Begin prjDownloadBooster.CommandButtonW cmdApply 
      Height          =   360
      Left            =   4920
      TabIndex        =   3
      Top             =   5040
      Width           =   1320
      _ExtentX        =   0
      _ExtentY        =   0
      Enabled         =   0   'False
      Caption         =   "����(&A)"
   End
   Begin prjDownloadBooster.TabStrip tsTabStrip 
      Height          =   4815
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   8493
      MultiRow        =   0   'False
      TabFixedWidth   =   53
      TabScrollWheel  =   0   'False
      Transparent     =   -1  'True
      InitTabs        =   "frmOptions.frx":000C
   End
   Begin prjDownloadBooster.CommandButtonW CancelButton 
      Cancel          =   -1  'True
      Height          =   360
      Left            =   3480
      TabIndex        =   1
      Top             =   5040
      Width           =   1320
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "���"
   End
   Begin prjDownloadBooster.CommandButtonW OKButton 
      Default         =   -1  'True
      Height          =   360
      Left            =   2040
      TabIndex        =   0
      Top             =   5040
      Width           =   1320
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "Ȯ��"
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"
Dim Loaded As Boolean
Dim ColorChanged As Boolean
Public ImageChanged As Boolean
Dim VisualStyleChanged As Boolean
Dim SkinChanged As Boolean

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub cbImagePosition_Click()
    If Loaded Then
        cmdApply.Enabled = -1
        tygApply.Enabled = -1
    End If
End Sub

Private Sub cbLanguage_Click()
    If Loaded Then
        Alert t("�� �����Ϸ��� ���α׷��� ������ؾ� �մϴ�.", "To change the language you must restart the application."), App.Title, Me, 64
        cmdApply.Enabled = -1
        tygApply.Enabled = -1
    End If
End Sub

Private Sub cbSkin_Click()
    cmdSample.VisualStyles = (cbSkin.ListIndex = 0 Or cbSkin.ListIndex = 2)
    tygSample.Visible = (cbSkin.ListIndex = 2)
    If Loaded Then
        cmdApply.Enabled = -1
        tygApply.Enabled = -1
        SkinChanged = True
        VisualStyleChanged = True
    End If
End Sub

Private Sub cbWhenExist_Click()
    If Loaded Then
        cmdApply.Enabled = -1
        tygApply.Enabled = -1
    End If
End Sub

Private Sub chkAlwaysResume_Click()
    If Loaded Then
        cmdApply.Enabled = -1
        tygApply.Enabled = -1
    End If
End Sub

Private Sub chkAutoRetry_Click()
    If Loaded Then
        cmdApply.Enabled = -1
        tygApply.Enabled = -1
    End If
End Sub

Private Sub chkBeepWhenComplete_Click()
    If Loaded Then
        cmdApply.Enabled = -1
        tygApply.Enabled = -1
    End If
End Sub

Private Sub chkEnableBackgroundImage_Click()
    If Loaded Then
        cmdApply.Enabled = -1
        tygApply.Enabled = -1
        ImageChanged = True
    End If
    
    If chkEnableBackgroundImage.Value = 0 Then
        cmdChooseBackground.Enabled = 0
        tygChooseBackground.Enabled = 0
        imgPreview.Visible = 0
        cmdSample.Refresh
    Else
        cmdChooseBackground.Enabled = -1
        tygChooseBackground.Enabled = -1
        
        On Error Resume Next
        If LCase(Right$(GetSetting("DownloadBooster", "Options", "BackgroundImagePath", ""), 4)) = ".png" Then
            Set imgPreview.Picture = LoadPngIntoPictureWithAlpha(GetSetting("DownloadBooster", "Options", "BackgroundImagePath", ""))
        Else
            imgPreview.Picture = LoadPicture(GetSetting("DownloadBooster", "Options", "BackgroundImagePath", ""))
        End If
        
        imgPreview.Visible = -1
        cmdSample.Refresh
    End If
End Sub

Private Sub chkNoCleanup_Click()
    If Loaded Then
        cmdApply.Enabled = -1
        tygApply.Enabled = -1
    End If
End Sub

Private Sub chkNoDWMWindow_Click()
    If Loaded Then
        cmdApply.Enabled = -1
        tygApply.Enabled = -1
    End If
End Sub

'Private Sub chkNoTheming_Click()
'    If Loaded Then
'        cmdApply.Enabled = -1
'        tygApply.Enabled = -1
'        VisualStyleChanged = True
'    End If
'    cmdSample.VisualStyles = (Not CBool(chkNoTheming.Value * (-1)))
'End Sub

Private Sub chkOpenDirWhenComplete_Click()
    If Loaded Then
        cmdApply.Enabled = -1
        tygApply.Enabled = -1
    End If
End Sub

Private Sub chkOpenWhenComplete_Click()
    If Loaded Then
        cmdApply.Enabled = -1
        tygApply.Enabled = -1
    End If
End Sub

Private Sub chkRememberURL_Click()
    If Loaded Then
        cmdApply.Enabled = -1
        tygApply.Enabled = -1
    End If
End Sub

Private Sub cmdApply_Click()
    If WinVer >= 6.1 And chkNoDWMWindow.Enabled Then SaveSetting "DownloadBooster", "Options", "DisableDWMWindow", chkNoDWMWindow.Value
    Dim i%
    
    SaveSetting "DownloadBooster", "Options", "NoCleanup", chkNoCleanup.Value
    SaveSetting "DownloadBooster", "Options", "RememberURL", chkRememberURL.Value
    
    SaveSetting "DownloadBooster", "Options", "OpenWhenComplete", chkOpenWhenComplete.Value
    SaveSetting "DownloadBooster", "Options", "OpenFolderWhenComplete", chkOpenDirWhenComplete.Value
    SaveSetting "DownloadBooster", "Options", "PlaySound", chkBeepWhenComplete.Value
    SaveSetting "DownloadBooster", "Options", "ContinueDownload", chkAlwaysResume.Value
    SaveSetting "DownloadBooster", "Options", "AutoRetry", chkAutoRetry.Value
    SaveSetting "DownloadBooster", "Options", "WhenFileExists", cbWhenExist.ListIndex
    frmMain.cbWhenExist.ListIndex = cbWhenExist.ListIndex
    
    frmMain.chkOpenAfterComplete.Value = chkOpenWhenComplete.Value
    frmMain.chkOpenFolder.Value = chkOpenDirWhenComplete.Value
    frmMain.chkPlaySound.Value = chkBeepWhenComplete.Value
    frmMain.chkContinueDownload.Value = chkAlwaysResume.Value
    frmMain.chkAutoRetry.Value = chkAutoRetry.Value
    
    If chkNoDWMWindow.Enabled = True Then
        If chkNoDWMWindow.Value Then
            DisableDWMWindow Me.hWnd
            DisableDWMWindow frmMain.hWnd
        Else
            EnableDWMWindow Me.hWnd
            EnableDWMWindow frmMain.hWnd
        End If
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
        frmMain.pgSettingsBackground.Visible = 0
        frmMain.chkOpenAfterComplete.Tag = ""
        frmMain.chkOpenFolder.Tag = ""
        frmMain.chkPlaySound.Tag = ""
        frmMain.chkContinueDownload.Tag = ""
        frmMain.chkAutoRetry.Tag = ""
        frmMain.chkOpenAfterComplete.Transparent = -1
        frmMain.chkOpenFolder.Transparent = -1
        frmMain.chkPlaySound.Transparent = -1
        frmMain.chkContinueDownload.Transparent = -1
        frmMain.chkAutoRetry.Transparent = -1
    ElseIf optUserFore.Value Then
        SaveSetting "DownloadBooster", "Options", "ForeColor", CLng(pgFore.BackColor)
        frmMain.pgSettingsBackground.Visible = -1
        frmMain.chkOpenAfterComplete.Tag = "nobackcolorchange"
        frmMain.chkOpenFolder.Tag = "nobackcolorchange"
        frmMain.chkPlaySound.Tag = "nobackcolorchange"
        frmMain.chkContinueDownload.Tag = "nobackcolorchange"
        frmMain.chkAutoRetry.Tag = "nobackcolorchange"
        frmMain.chkOpenAfterComplete.Transparent = 0
        frmMain.chkOpenFolder.Transparent = 0
        frmMain.chkPlaySound.Transparent = 0
        frmMain.chkContinueDownload.Transparent = 0
        frmMain.chkAutoRetry.Transparent = 0
    End If
    SaveSetting "DownloadBooster", "Options", "DisableVisualStyle", CBool(cbSkin.ListIndex = 1) * (-1)
    SaveSetting "DownloadBooster", "Options", "EnableLiveBadukMemoSkin", CBool(cbSkin.ListIndex = 2) * (-1)
    If ColorChanged Or VisualStyleChanged Or SkinChanged Then
        SetFormBackgroundColor Me
        SetFormBackgroundColor frmMain
        frmMain.LoadLiveBadukSkin
    End If
    If VisualStyleChanged Then
        On Error Resume Next
        Dim ctrl As Control
        For Each ctrl In Me.Controls
            If TypeName(ctrl) = "PictureBox" Then
                ctrl.AutoRedraw = True
                tsTabStrip.DrawBackground ctrl.hWnd, ctrl.hDC
            End If
        Next ctrl
        For Each ctrl In Me.Controls
            If TypeName(ctrl) = "FrameW" Then
                ctrl.Transparent = True 'Not CBool(chkNoTheming.Value)
            End If
            ctrl.Refresh
        Next ctrl
        On Error GoTo 0
    End If
    If cbLanguage.ListIndex = 0 Then
        SaveSetting "DownloadBooster", "Options", "Language", "0"
    ElseIf cbLanguage.ListIndex = 1 Then
        SaveSetting "DownloadBooster", "Options", "Language", 1042
    Else
        SaveSetting "DownloadBooster", "Options", "Language", 1033
    End If
    SaveSetting "DownloadBooster", "Options", "ImagePosition", cbImagePosition.ListIndex
    frmMain.ImagePosition = cbImagePosition.ListIndex
    frmMain.SetBackgroundPosition True
    If ImageChanged Then
        SaveSetting "DownloadBooster", "Options", "UseBackgroundImage", chkEnableBackgroundImage.Value
        frmMain.SetBackgroundImage
    End If
    
    On Error Resume Next
    If GetSetting("DownloadBooster", "Options", "ForeColor", -1) <> -1 Or GetSetting("DownloadBooster", "Options", "UseBackgroundImage", 0) = 1 Then
        For i = frmMain.pgOverlay.LBound To frmMain.pgOverlay.UBound
            frmMain.pgOverlay(i).Visible = -1
            frmMain.lblOverlay(i).Visible = -1
        Next i
        frmMain.optTabDownload2.Transparent = 0
        frmMain.optTabDownload2.BackColor = frmMain.pgOverlay(0).BackColor
        frmMain.optTabDownload2.Refresh
        frmMain.optTabThreads2.Transparent = 0
        frmMain.optTabThreads2.BackColor = frmMain.pgOverlay(0).BackColor
        frmMain.optTabThreads2.Refresh
        frmMain.fTabDownload.Transparent = 0
        frmMain.fTabDownload.BackColor = frmMain.pgOverlay(0).BackColor
        frmMain.fTabDownload.Refresh
        frmMain.fTabThreads.Transparent = 0
        frmMain.fTabThreads.BackColor = frmMain.pgOverlay(0).BackColor
        frmMain.fTabThreads.Refresh
    Else
        For i = frmMain.pgOverlay.LBound To frmMain.pgOverlay.UBound
            frmMain.pgOverlay(i).Visible = 0
            frmMain.lblOverlay(i).Visible = 0
        Next i
        frmMain.optTabDownload2.Transparent = -1
        frmMain.optTabDownload2.Refresh
        frmMain.optTabThreads2.Transparent = -1
        frmMain.optTabThreads2.Refresh
        frmMain.fTabDownload.Transparent = -1
        frmMain.fTabDownload.Refresh
        frmMain.fTabThreads.Transparent = -1
        frmMain.fTabThreads.Refresh
    End If
    On Error GoTo 0
    Dim NoDisable As Boolean
    NoDisable = False
    If Trim$(txtNodePath.Text) <> "" Then
        If FileExists(Trim$(txtNodePath.Text)) Then
            SaveSetting "DownloadBooster", "Options", "NodePath", Trim$(txtNodePath.Text)
        Else
            Alert t("Node.js ��ΰ� �������� �ʽ��ϴ�.", "Node.js path does not exist."), App.Title, Me, 16
            NoDisable = True
        End If
    Else
        SaveSetting "DownloadBooster", "Options", "NodePath", ""
    End If
    If Trim$(txtScriptPath.Text) <> "" Then
        If FileExists(Trim$(txtScriptPath.Text)) Then
            SaveSetting "DownloadBooster", "Options", "ScriptPath", Trim$(txtScriptPath.Text)
        Else
            Alert t("�ٿ�ε� ��ũ��Ʈ ��ΰ� �������� �ʽ��ϴ�.", "Download script path does not exist."), App.Title, Me, 16
            NoDisable = True
        End If
    Else
        SaveSetting "DownloadBooster", "Options", "ScriptPath", ""
    End If
    If Trim$(txtYtdlPath.Text) <> "" Then
        If FileExists(Trim$(txtYtdlPath.Text)) Then
            SaveSetting "DownloadBooster", "Options", "YtdlPath", Trim$(txtYtdlPath.Text)
        Else
            Alert t("Youtube-dl ��ΰ� �������� �ʽ��ϴ�.", "Youtube-dl path does not exist."), App.Title, Me, 16
            NoDisable = True
        End If
    Else
        SaveSetting "DownloadBooster", "Options", "YtdlPath", ""
    End If
    
    ColorChanged = False
    ImageChanged = False
    VisualStyleChanged = False
    SkinChanged = False
    If Not NoDisable Then
        cmdApply.Enabled = 0
        tygApply.Enabled = 0
    End If
End Sub

Private Sub cmdChooseBackground_Click()
    frmCustomBackground.Show vbModal, Me
End Sub

Private Sub cmdSysInfo_Click()
    Call StartSysInfo
End Sub

Private Sub Form_Load()
    VisualStyleChanged = False
    Loaded = False
    ImageChanged = False
    ColorChanged = False
    SkinChanged = False
    If GetSetting("DownloadBooster", "Options", "DisableDWMWindow", DefaultDisableDWMWindow) = 1 Then DisableDWMWindow Me.hWnd
    SetFormBackgroundColor Me
    SetFont Me
    
    Me.Width = 6495 '6840
    Me.Height = 5970
    
    lblSelectColor.Top = pgColor.Top
    lblSelectColor.Left = pgColor.Left
    lblSelectColor.Width = pgColor.Width
    lblSelectColor.Height = pgColor.Height
    
    lblSelectFore.Top = pgFore.Top
    lblSelectFore.Left = pgFore.Left
    lblSelectFore.Width = pgFore.Width
    lblSelectFore.Height = pgFore.Height
    
    imgPreview.Width = pgBackPreview.Width - 30
    imgPreview.Height = pgBackPreview.Height - 30
    imgPreview.Top = pgBackPreview.Top + 15
    imgPreview.Left = pgBackPreview.Left + 15
    
    Dim i%
    For i = 1 To pbPanel.Count
        If i <> 1 Then
            pbPanel(i).Visible = 0
            pbPanel(i).Enabled = 0
        End If
        If i <> 2 Then
            pbPanel(i).Top = pbPanel(2).Top
            pbPanel(i).Left = pbPanel(2).Left
            pbPanel(i).Width = pbPanel(2).Width
            pbPanel(i).Height = pbPanel(2).Height
            pbPanel(i).BorderStyle = 0
            pbPanel(i).AutoRedraw = True
        End If
    Next i
    chkNoCleanup.Value = GetSetting("DownloadBooster", "Options", "NoCleanup", 0)
    chkNoDWMWindow.Value = GetSetting("DownloadBooster", "Options", "DisableDWMWindow", DefaultDisableDWMWindow)
    If (Not IsDWMEnabled()) Then
        chkNoDWMWindow.Enabled = False
        chkNoDWMWindow.Value = 1
    End If
    If WinVer < 6.1 Then chkNoDWMWindow.Value = 0
    chkRememberURL.Value = GetSetting("DownloadBooster", "Options", "RememberURL", 0)
    chkNoDWMWindow.Caption = t(chkNoDWMWindow.Caption, "Use W&indows 7 style")
    If WinVer < 6# Then
        chkNoDWMWindow.Caption = t("DWM â ��Ȱ��ȭ(&I)", "Disable DWM w&indow")
        chkNoDWMWindow.Value = 1
    ElseIf WinVer < 6.2 Then
        chkNoDWMWindow.Caption = t("Aero â ��� �� ��(&I)", "Disable Aero w&indow")
    End If
    
    chkEnableBackgroundImage.Value = GetSetting("DownloadBooster", "Options", "UseBackgroundImage", 0)
    If chkEnableBackgroundImage.Value = 0 Then
        cmdChooseBackground.Enabled = 0
        tygChooseBackground.Enabled = 0
    End If
    
    Dim clrBackColor As Long
    clrBackColor = GetSetting("DownloadBooster", "Options", "BackColor", DefaultBackColor)
    If clrBackColor < 0 Or clrBackColor > 16777215 Then
        optSystemColor.Value = True
        pgColor.BackColor = &H8000000F
    Else
        optUserColor.Value = True
        pgColor.BackColor = clrBackColor
    End If
    pgBackPreview.BackColor = pgColor.BackColor
    
    Dim clrForeColor As Long
    clrForeColor = GetSetting("DownloadBooster", "Options", "ForeColor", -1)
    If clrForeColor < 0 Or clrForeColor > 16777215 Then
        optSystemFore.Value = True
        pgFore.BackColor = &H80000012
    Else
        optUserFore.Value = True
        pgFore.BackColor = clrForeColor
    End If
    lnForePreview.BorderColor = pgFore.BackColor
    
    cmdApply.Enabled = 0
    tygApply.Enabled = 0
    'On Error Resume Next
    Dim ctrl As Control
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "PictureBox" Then
            ctrl.AutoRedraw = True
            tsTabStrip.DrawBackground ctrl.hWnd, ctrl.hDC
        End If
    Next ctrl
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "FrameW" Then
            ctrl.Transparent = True
            ctrl.Refresh
        End If
    Next ctrl
    
    chkOpenWhenComplete.Value = frmMain.chkOpenAfterComplete.Value
    chkOpenDirWhenComplete.Value = frmMain.chkOpenFolder.Value
    chkBeepWhenComplete.Value = frmMain.chkPlaySound.Value
    chkAlwaysResume.Value = frmMain.chkContinueDownload.Value
    chkAutoRetry.Value = frmMain.chkAutoRetry.Value
    
    cbSkin.Clear
    cbSkin.AddItem t("�ý��� ��Ÿ��", "System style")
    cbSkin.AddItem t("���� ��Ÿ��", "Classic style")
    cbSkin.AddItem t("Ÿ���� �ٵ� ��Ÿ��", "LiveBaduk style")
    If CInt(GetSetting("DownloadBooster", "Options", "EnableLiveBadukMemoSkin", 0)) Then
        cbSkin.ListIndex = 2
    ElseIf Abs(CInt(GetSetting("DownloadBooster", "Options", "DisableVisualStyle", 0))) Then
        cbSkin.ListIndex = 1
    Else
        cbSkin.ListIndex = 0
    End If
    
    'chkNoTheming.Value = Abs(CInt(GetSetting("DownloadBooster", "Options", "DisableVisualStyle", 0)))
    cmdSample.VisualStyles = (Not CBool(CInt(GetSetting("DownloadBooster", "Options", "DisableVisualStyle", 0))))
    tygSample.Visible = Abs(CInt(GetSetting("DownloadBooster", "Options", "EnableLiveBadukMemoSkin", 0))) * (-1)
    
    cbLanguage.Clear
    cbLanguage.AddItem t("�ڵ�", "Auto")
    cbLanguage.AddItem "�ѱ���"
    cbLanguage.AddItem "English"
    Dim LangSet As String
    LangSet = GetSetting("DownloadBooster", "Options", "Language", "0")
    If LangSet = "0" Then
        cbLanguage.ListIndex = 0
    ElseIf LangSet = "1042" Then
        cbLanguage.ListIndex = 1
    Else
        cbLanguage.ListIndex = 2
    End If
    
    cbWhenExist.Clear
    cbWhenExist.AddItem t("�ߴ�", "Abort")
    cbWhenExist.AddItem t("�����", "Overwrite")
    cbWhenExist.AddItem t("�̸� ����", "Rename")
    cbWhenExist.ListIndex = GetSetting("DownloadBooster", "Options", "WhenFileExists", 0)
    
    cbImagePosition.Clear
    cbImagePosition.AddItem t("���̱�", "Stretch")
    cbImagePosition.AddItem t("���̿� ���߱�", "Fit to height")
    cbImagePosition.AddItem t("�ʺ� ���߱�", "Fit to width")
    cbImagePosition.AddItem t("���� ũ�� ����", "True size")
    cbImagePosition.ListIndex = GetSetting("DownloadBooster", "Options", "ImagePosition", 1)
    
    txtNodePath.Text = GetSetting("DownloadBooster", "Options", "NodePath", "")
    txtScriptPath.Text = GetSetting("DownloadBooster", "Options", "ScriptPath", "")
    txtYtdlPath.Text = GetSetting("DownloadBooster", "Options", "YtdlPath", "")
    
    chkNoDWMWindow_Click
    
    On Error Resume Next
    If chkEnableBackgroundImage.Value And GetSetting("DownloadBooster", "Options", "BackgroundImagePath", "") <> "" Then
        If LCase(Right$(GetSetting("DownloadBooster", "Options", "BackgroundImagePath", ""), 4)) = ".png" Then
            Set imgPreview.Picture = LoadPngIntoPictureWithAlpha(GetSetting("DownloadBooster", "Options", "BackgroundImagePath", ""))
        Else
            imgPreview.Picture = LoadPicture(GetSetting("DownloadBooster", "Options", "BackgroundImagePath", ""))
        End If
        
        imgPreview.Visible = -1
    End If
    
    tsTabStrip.Tabs(1).Caption = t(tsTabStrip.Tabs(1).Caption, " General ")
    tsTabStrip.Tabs(2).Caption = t(tsTabStrip.Tabs(2).Caption, " Appearance ")
    tsTabStrip.Tabs(3).Caption = t(tsTabStrip.Tabs(3).Caption, "  Paths  ")
    Frame1.Caption = t(Frame1.Caption, " Background color ")
    Frame4.Caption = t(Frame4.Caption, " Text color ")
    Frame3.Caption = t(Frame3.Caption, " Window appearance ")
    Frame2.Caption = t(Frame2.Caption, " Download options ")
    Frame5.Caption = t(Frame5.Caption, " Interface ")
    chkNoCleanup.Caption = t(chkNoCleanup.Caption, "Preserve segme&nts")
    chkRememberURL.Caption = t(chkRememberURL.Caption, "Re&member URL")
    optSystemColor.Caption = t(optSystemColor.Caption, "&System color")
    optUserColor.Caption = t(optUserColor.Caption, "C&ustom color:")
    optSystemFore.Caption = t(optSystemFore.Caption, "S&ystem color")
    optUserFore.Caption = t(optUserFore.Caption, "Cus&tom color:")
    Label1.Caption = t(Label1.Caption, "&Language:")
    OKButton.Caption = t(OKButton.Caption, "OK")
    CancelButton.Caption = t(CancelButton.Caption, "Cancel")
    cmdApply.Caption = t(cmdApply.Caption, "&Apply")
    Me.Caption = t(Me.Caption, "Options")
    Frame6.Caption = t(Frame6.Caption, " Skin ")
    chkOpenWhenComplete.Caption = t(chkOpenWhenComplete.Caption, "&Open file when complete")
    chkOpenDirWhenComplete.Caption = t(chkOpenDirWhenComplete.Caption, "O&pen folder when complete")
    chkBeepWhenComplete.Caption = t(chkBeepWhenComplete.Caption, "&Beep when complete")
    chkAlwaysResume.Caption = t(chkAlwaysResume.Caption, "&Always resume")
    chkAutoRetry.Caption = t(chkAutoRetry.Caption, "A&uto retry on network error")
    Label3.Caption = t(Label3.Caption, "If filename alrea&dy exists:")
    FrameW1.Caption = t(FrameW1.Caption, " Background image ")
    chkEnableBackgroundImage.Caption = t(chkEnableBackgroundImage.Caption, "Use &background image")
    'cmdChooseBackground.Caption = t(cmdChooseBackground.Caption, "C&hoose image...")
    'chkNoTheming.Caption = t(chkNoTheming.Caption, "&Use classic theme")
    Label6.Caption = t(Label6.Caption, "Leave the field blank to use defaults. This option is for advanced users and there is no need to change for normal use.")
    FrameW2.Caption = t(FrameW2.Caption, " Directory settings ")
    Label5.Caption = t(Label5.Caption, "&Download script:")
    cmdSample.Caption = t(cmdSample.Caption, "Button")
    tygSample.Caption = cmdSample.Caption
    Label2.Caption = t(Label2.Caption, "&Position:")
    Label8.Caption = t(Label8.Caption, "Select s&kin:")
    tygOK.Caption = t(tygOK.Caption, "OK")
    tygCancel.Caption = t(tygCancel.Caption, "Cancel")
    tygApply.Caption = t(tygApply.Caption, "Apply")
    
    Loaded = True
End Sub

Private Sub lblSelectColor_Click()
    Dim Color As OLE_COLOR
    Color = ShowColorDialog(Me.hWnd, True, pgColor.BackColor)
    If Color = -1 Then Exit Sub
    pgColor.BackColor = Color
    cmdApply.Enabled = -1
    tygApply.Enabled = -1
    optUserColor.Value = True
    ColorChanged = True
    pgBackPreview.BackColor = pgColor.BackColor
    cmdSample.Refresh
End Sub

Private Sub lblSelectFore_Click()
    Dim Color As OLE_COLOR
    Color = ShowColorDialog(Me.hWnd, True, pgFore.BackColor, True)
    If Color = -1 Then Exit Sub
    pgFore.BackColor = Color
    cmdApply.Enabled = -1
    tygApply.Enabled = -1
    optUserFore.Value = True
    ColorChanged = True
    lnForePreview.BorderColor = pgFore.BackColor
End Sub

Private Sub OKButton_Click()
    cmdApply_Click
    Unload Me
End Sub

Private Sub optSystemColor_Click()
    If Loaded Then
        cmdApply.Enabled = -1
        tygApply.Enabled = -1
        ColorChanged = True
    End If
    pgBackPreview.BackColor = &H8000000F
    cmdSample.Refresh
End Sub

Private Sub optSystemFore_Click()
    If Loaded Then
        cmdApply.Enabled = -1
        tygApply.Enabled = -1
        ColorChanged = True
    End If
    lnForePreview.BorderColor = &H80000012
End Sub

Private Sub optUserColor_Click()
    If Loaded Then
        cmdApply.Enabled = -1
        tygApply.Enabled = -1
        ColorChanged = True
    End If
    pgBackPreview.BackColor = pgColor.BackColor
    cmdSample.Refresh
End Sub

Private Sub optUserFore_Click()
    If Loaded Then
        cmdApply.Enabled = -1
        tygApply.Enabled = -1
        ColorChanged = True
    End If
    lnForePreview.BorderColor = pgFore.BackColor
End Sub

Private Sub tsTabStrip_TabClick(ByVal TabItem As TbsTab)
    Dim i%
    For i = 1 To pbPanel.Count
        If i = TabItem.Index Then
            pbPanel(i).Visible = -1
            pbPanel(i).Enabled = -1
        Else
            pbPanel(i).Visible = 0
            pbPanel(i).Enabled = 0
        End If
    Next i
End Sub

Sub StartSysInfo()
    On Error GoTo SysInfoErr
  
    Dim RC As Long
    Dim SysInfoPath As String
    
    ' �ý��� ���� ���α׷��� ��ο� �̸��� ������Ʈ������ ���� �ɴϴ�...
    SysInfoPath = GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, "")
    If SysInfoPath = "" Then
        SysInfoPath = GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, "")
        If SysInfoPath <> "" Then
            ' �˷��� 32��Ʈ ���� ������ ���� ���θ� Ȯ���մϴ�.
            If Dir(SysInfoPath & "\MSINFO32.EXE") <> "" Then
                SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
                
            ' ���� - ������ ã�� �� �����ϴ�...
            Else
                GoTo SysInfoErr
            End If
        ' ���� - ������Ʈ�� �׸��� ã�� �� �����ϴ�...
        Else
            GoTo SysInfoErr
        End If
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    Alert t("������ �ý��� ������ ����� �� �����ϴ�.", "System Information is unavailable."), App.Title, Me, 48
End Sub

Private Sub txtNodePath_Change()
    If Loaded Then
        cmdApply.Enabled = -1
        tygApply.Enabled = -1
    End If
End Sub

Private Sub txtScriptPath_Change()
    If Loaded Then
        cmdApply.Enabled = -1
        tygApply.Enabled = -1
    End If
End Sub

Private Sub txtYtdlPath_Change()
    If Loaded Then
        cmdApply.Enabled = -1
        tygApply.Enabled = -1
    End If
End Sub

Private Sub tygApply_Click()
    cmdApply_Click
End Sub

Private Sub tygCancel_Click()
    CancelButton_Click
End Sub

Private Sub tygChooseBackground_Click()
    cmdChooseBackground_Click
End Sub

Private Sub tygOK_Click()
    OKButton_Click
End Sub
