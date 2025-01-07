VERSION 5.00
Begin VB.Form frmBrowse 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "다운로드 경로 선택"
   ClientHeight    =   3255
   ClientLeft      =   2760
   ClientTop       =   3870
   ClientWidth     =   6345
   BeginProperty Font 
      Name            =   "굴림"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBrowse.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   6345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.Timer timDelayer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5400
      Top             =   1560
   End
   Begin VB.CheckBox chkHidden 
      Caption         =   "숨김 표시(&H)"
      Height          =   255
      Left            =   4920
      TabIndex        =   6
      Top             =   2880
      Width           =   1350
   End
   Begin VB.TextBox txtFileName 
      Height          =   270
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   2175
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
      TabIndex        =   11
      Top             =   2880
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
      TabIndex        =   3
      Top             =   2880
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
      TabIndex        =   0
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
      TabIndex        =   5
      Top             =   510
      Width           =   1335
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "확인"
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
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblDirectory 
      BackStyle       =   0  '투명
      Height          =   255
      Left            =   2520
      TabIndex        =   12
      Top             =   415
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '투명
      Caption         =   "파일 형식(&T):"
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
      TabIndex        =   10
      Top             =   2640
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
      TabIndex        =   9
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '투명
      Caption         =   "폴더(&D):"
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
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      Caption         =   "파일 이름(&F):"
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
      TabIndex        =   7
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmBrowse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub chkHidden_Click()
    lvFiles.Hidden = chkHidden.Value
End Sub

Private Sub Form_Load()
    selFileType.AddItem "모든 파일 (*.*)"
    selFileType.ListIndex = 0
    
    On Error Resume Next
    Dim Path$
    Path = lvDir.Path
    
    Dim fmpth As String
    fmpth = Trim$(frmMain.txtFileName.Text)
    If FolderExists(fmpth) Then
        Path = fmpth
    Else
        Path = fso.GetParentFolderName(fmpth)
        txtFileName.Text = Split(fmpth, "\")(UBound(Split(fmpth, "\")))
    End If
    
    Dim i%
    For i = 0 To selDrive.ListCount - 1
        If LCase(Left$(selDrive.List(i), 1)) = LCase(Left$(Path, 1)) Then
            selDrive.ListIndex = i
            Exit For
        End If
    Next i
    
    lvDir.Path = Path
    lvDir_Change
End Sub

Private Sub lvDir_Change()
    lvFiles.Path = lvDir.Path
    lblDirectory.Caption = Right$(lvDir.Path, 12)
    If lblDirectory.Caption <> lvDir.Path Then lblDirectory.Caption = "..." & lblDirectory.Caption
    SaveSetting "DownloadBooster", "UserData", "LastSaveDir", lvDir.Path
End Sub

Private Sub lvFiles_Click()
    If frmMain.cbWhenExist.ListIndex = 0 Then Exit Sub
    If Not timDelayer.Enabled Then txtFileName.Text = lvFiles.List(lvFiles.ListIndex)
End Sub

Private Sub lvFiles_DblClick()
    If frmMain.cbWhenExist.ListIndex <> 0 Then _
        OKButton_Click
End Sub

Private Sub OKButton_Click()
    txtFileName.Text = Trim$(txtFileName.Text)
    
    If _
        InStr(1, txtFileName.Text, "\") > 0 Or _
        InStr(1, txtFileName.Text, "/") > 0 Or _
        InStr(1, txtFileName.Text, """") > 0 Or _
        InStr(1, txtFileName.Text, "*") > 0 Or _
        InStr(1, txtFileName.Text, "?") > 0 Or _
        InStr(1, txtFileName.Text, "<") > 0 Or _
        InStr(1, txtFileName.Text, ">") > 0 Or _
        InStr(1, txtFileName.Text, "|") > 0 Or _
        UCase(txtFileName.Text) = "CON" Or _
        UCase(txtFileName.Text) = "AUX" Or _
        UCase(txtFileName.Text) = "PRN" Or _
        UCase(txtFileName.Text) = "NUL" Or _
        UCase(txtFileName.Text) = "COM1" Or _
        UCase(txtFileName.Text) = "COM2" Or _
        UCase(txtFileName.Text) = "COM3" Or _
        UCase(txtFileName.Text) = "COM4" Or _
        UCase(txtFileName.Text) = "LPT1" Or _
        UCase(txtFileName.Text) = "LPT2" Or _
        UCase(txtFileName.Text) = "LPT3" Or _
        UCase(txtFileName.Text) = "LPT4" _
    Then
        MsgBox "파일 이름이 올바르지 않습니다.", 48
        Exit Sub
    End If

    Dim Data$, Path$
    
    If Right$(lvFiles.Path, 1) = "\" Then
        Path = lvFiles.Path & txtFileName.Text
    Else
        Path = lvFiles.Path & "\" & txtFileName.Text
    End If
    On Error Resume Next
    If FileExists(Path) Then
        If frmMain.cbWhenExist.ListIndex = 0 Then
            MsgBox "파일 이름이 이미 존재합니다. 다른 이름을 선택하십시오.", 16
            Exit Sub
        ElseIf frmMain.cbWhenExist.ListIndex = 1 Then
            If MsgBox("파일 이름이 이미 존재합니다. 덮어쓰시겠습니까?", 48 + vbYesNo) <> vbYes Then
                Exit Sub
            End If
        End If
    End If
    
    On Error GoTo e
    If Right$(Path, 2) = "\\" Then Path = Left$(Path, Len(Path) - 1)
    frmMain.txtFileName = Path
    
    Unload Me
    Exit Sub
    
e:
    MsgBox "문제가 발생했습니다!", 16
    Exit Sub
End Sub

Private Sub selDrive_Change()
    On Error GoTo e
    lvDir.Path = selDrive.Drive
    Exit Sub
    
e:
    MsgBox "선택한 드라이브 안에 디스크가 없습니다.", 16
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
