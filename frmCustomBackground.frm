VERSION 5.00
Begin VB.Form frmCustomBackground 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "배경 그림 선택"
   ClientHeight    =   4965
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6495
   BeginProperty Font 
      Name            =   "굴림"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCustomBackground.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin prjDownloadBooster.CheckBoxW chkHidden 
      Height          =   255
      Left            =   4920
      TabIndex        =   10
      Top             =   2760
      Width           =   1575
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "숨김 표시(&H)"
   End
   Begin VB.Timer timDelayer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5640
      Top             =   3600
   End
   Begin VB.TextBox txtFileName 
      Height          =   270
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   2175
   End
   Begin VB.PictureBox picPreviewFrame 
      Height          =   1335
      Left            =   120
      ScaleHeight     =   1275
      ScaleWidth      =   3675
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3480
      Width           =   3735
      Begin VB.Image imgPreview 
         Height          =   1275
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   3675
      End
   End
   Begin VB.ComboBox selFileType 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      Style           =   2  '드롭다운 목록
      TabIndex        =   4
      Top             =   2760
      Width           =   2175
   End
   Begin VB.DriveListBox selDrive 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2520
      TabIndex        =   9
      Top             =   2760
      Width           =   2175
   End
   Begin VB.DirListBox lvDir 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1770
      Left            =   2520
      TabIndex        =   7
      Top             =   720
      Width           =   2175
   End
   Begin VB.FileListBox lvFiles 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1710
      Left            =   120
      Pattern         =   "*.JPG; *.GIF; *.BMP; *.DIB; *.WMF; *.EMF"
      System          =   -1  'True
      TabIndex        =   2
      Top             =   720
      Width           =   2175
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "취소"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   340
      Left            =   4920
      TabIndex        =   14
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "열기(&O)"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   340
      Left            =   4920
      TabIndex        =   13
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblDirectory 
      BackStyle       =   0  '투명
      Height          =   255
      Left            =   2520
      TabIndex        =   6
      Top             =   375
      Width           =   2175
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '투명
      Caption         =   "미리보기:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '투명
      Caption         =   "파일 형식(&F):"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '투명
      Caption         =   "드라이브(&R):"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   8
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '투명
      Caption         =   "폴더 선택(&D):"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   5
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      Caption         =   "그림 선택(&P):"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmCustomBackground"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
    Unload Me
End Sub

Function GetFilename(Path As String) As String
    GetFilename = Right(Path, Len(Path) - InStrRev(Path, "\"))
End Function

Private Sub chkHidden_Click()
    lvFiles.Hidden = chkHidden.Value
End Sub

Private Sub Form_Load()
    If GetSetting("DownloadBooster", "Options", "DisableDWMWindow", DefaultDisableDWMWindow) = 1 Then DisableDWMWindow Me.hWnd
    SetFormBackgroundColor Me
    SetFont Me
    SetWindowPos Me.hWnd, IIf(MainFormOnTop, hWnd_TOPMOST, hWnd_NOTOPMOST), 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    
    selFileType.Clear
    selFileType.AddItem t("모든 그림", "All pictures") & " (*.JPG; *.GIF; *.BMP; *.DIB; *.PNG; *.WMF; *.EMF; *.ICO; *.CUR)"
    selFileType.AddItem "JPEG (*.JPG)"
    selFileType.AddItem "GIF (*.GIF)"
    selFileType.AddItem t("비트맵", "Bitmap") & " (*.BMP; *.DIB)"
    selFileType.AddItem "PNG (*.PNG)"
    selFileType.AddItem t("그래픽", "Graphics") & " (*.WMF; *.EMF)"
    selFileType.AddItem t("아이콘", "Icon") & " (*.ICO)"
    selFileType.AddItem t("커서", "Cursor") & " (*.CUR)"
    selFileType.ListIndex = 0
    
    lvDir_Change
    
    'On Error Resume Next
    Dim Path$
    Dim posn%
    Path = GetSetting("DownloadBooster", "Options", "BackgroundImagePath", "")
    If Path <> "" Then
        posn = InStrRev(Path, "\")
        If posn > 0 Then
            Dim i%
            For i = 0 To selDrive.ListCount - 1
                If LCase(Left$(selDrive.List(i), 1)) = LCase(Left$(Path, 1)) Then
                    selDrive.ListIndex = i
                    Exit For
                End If
            Next i
            lvDir.Path = Left$(Path, posn)
            lvFiles.Path = Left$(Path, posn)
            For i = 0 To lvFiles.ListCount - 1
                If GetFilename(Path) = lvFiles.List(i) Then
                    lvFiles.ListIndex = i
                    Exit For
                End If
            Next i
        End If
    End If
    
    Label1.Caption = t(Label1.Caption, "&Select picture:")
    Label4.Caption = t(Label4.Caption, "File &type:")
    Label3.Caption = t(Label3.Caption, "Dri&ve:")
    Label2.Caption = t(Label2.Caption, "&Directory:")
    chkHidden.Caption = t(chkHidden.Caption, "Show &hidden")
    OKButton.Caption = t(OKButton.Caption, "OK")
    CancelButton.Caption = t(CancelButton.Caption, "Cancel")
    Label5.Caption = t(Label5.Caption, "Preview:")
    Me.Caption = t(Me.Caption, "Select background image")
End Sub

Private Sub lvDir_Change()
    lvFiles.Path = lvDir.Path
    lblDirectory.Caption = Right$(lvDir.Path, 12)
    If lblDirectory.Caption <> lvDir.Path Then lblDirectory.Caption = "..." & lblDirectory.Caption
End Sub

Private Sub lvFiles_Click()
    On Error Resume Next
    If LCase(Right$(lvFiles.List(lvFiles.ListIndex), 4)) = ".png" Then
        Set imgPreview.Picture = LoadPngIntoPictureWithAlpha(lvFiles.Path & "\" & lvFiles.List(lvFiles.ListIndex))
    Else
        imgPreview.Picture = LoadPicture(lvFiles.Path & "\" & lvFiles.List(lvFiles.ListIndex))
    End If
    If Not timDelayer.Enabled Then txtFileName.Text = lvFiles.List(lvFiles.ListIndex)
End Sub

Private Sub lvFiles_DblClick()
    OKButton_Click
End Sub

Private Sub OKButton_Click()
    If lvFiles.ListIndex < 0 Then Exit Sub
    On Error GoTo e
    If LCase(Right$(lvFiles.List(lvFiles.ListIndex), 4)) = ".png" Then
        LoadPngIntoPictureWithAlpha lvFiles.Path & "\" & lvFiles.List(lvFiles.ListIndex)
    Else
        LoadPicture lvFiles.Path & "\" & lvFiles.List(lvFiles.ListIndex)
    End If
    SaveSetting "DownloadBooster", "Options", "BackgroundImagePath", lvFiles.Path & "\" & lvFiles.List(lvFiles.ListIndex)
    frmOptions.ImageChanged = True
    frmOptions.cmdApply.Enabled = True
    If LCase(Right$(lvFiles.List(lvFiles.ListIndex), 4)) = ".png" Then
        Set frmOptions.imgPreview.Picture = LoadPngIntoPictureWithAlpha(lvFiles.Path & "\" & lvFiles.List(lvFiles.ListIndex))
    Else
        frmOptions.imgPreview.Picture = LoadPicture(lvFiles.Path & "\" & lvFiles.List(lvFiles.ListIndex))
    End If
    frmOptions.cmdSample.Refresh
    frmOptions.RedrawPreview
    Unload Me
    Exit Sub
    
e:
    Alert t("그림이 손상되었거나 올바르지 않습니다.", "The selected picture is corrupt or invalid."), App.Title, 16
End Sub

Private Sub selDrive_Change()
    On Error GoTo e
    lvDir.Path = selDrive.Drive
    Exit Sub
    
e:
    Alert t("시스템에 부착된 장치를 사용할 수 없습니다.", "There is no disk in the selected drive"), App.Title, 16
End Sub

Private Sub selFileType_Click()
    lvFiles.Pattern = Replace(Mid$(selFileType.Text, InStr(1, selFileType.Text, "(") + 1, Len(selFileType.Text) - InStr(1, selFileType.Text, "(") - 1), " ", "")
End Sub

Private Sub timDelayer_Timer()
    timDelayer.Enabled = 0
End Sub

Private Sub txtFileName_Change()
    If Replace(txtFileName.Text, " ", "") = "" Then Exit Sub
    timDelayer.Enabled = -1
    
    On Error Resume Next
    Dim i%
    For i = 0 To lvFiles.ListCount - 1
        If LCase(Left$(lvFiles.List(i), Len(txtFileName.Text))) = LCase(txtFileName.Text) Then
            lvFiles.ListIndex = i
            Exit For
        End If
    Next i
End Sub
