VERSION 5.00
Begin VB.Form frmBrowse 
   BorderStyle     =   3  'ũ�� ���� ��ȭ ����
   Caption         =   "�ٿ�ε� ��� ����"
   ClientHeight    =   3255
   ClientLeft      =   2760
   ClientTop       =   3870
   ClientWidth     =   6345
   BeginProperty Font 
      Name            =   "����"
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
   StartUpPosition =   1  '������ ���
   Begin VB.Timer timDelayer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5400
      Top             =   1560
   End
   Begin prjDownloadBooster.CheckBoxW chkHidden 
      Height          =   255
      Left            =   4920
      TabIndex        =   10
      Top             =   2880
      Width           =   1455
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "���� ǥ��(&H)"
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
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      Style           =   2  '��Ӵٿ� ���
      TabIndex        =   4
      Top             =   2880
      Width           =   2175
   End
   Begin VB.DriveListBox selDrive 
      BeginProperty Font 
         Name            =   "����"
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
      Top             =   2880
      Width           =   2175
   End
   Begin VB.DirListBox lvDir 
      BeginProperty Font 
         Name            =   "����"
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
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1890
      Left            =   120
      System          =   -1  'True
      TabIndex        =   2
      Top             =   720
      Width           =   2175
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   340
      Left            =   4920
      TabIndex        =   12
      Top             =   510
      Width           =   1335
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "Ȯ��"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   340
      Left            =   4920
      TabIndex        =   11
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblDirectory 
      Height          =   255
      Left            =   2520
      TabIndex        =   6
      Top             =   415
      Width           =   2175
   End
   Begin VB.Label Label4 
      Caption         =   "���� ����(&T):"
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
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "����̺�(&V):"
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
      Left            =   2520
      TabIndex        =   8
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "����(&D):"
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
      Left            =   2520
      TabIndex        =   5
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "���� �̸�(&F):"
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
      Left            =   120
      TabIndex        =   0
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
    InitForm Me
    
    AddItemToComboBox selFileType, t("��� ����", "All files") & " (*.*)"
    selFileType.ListIndex = 0
    
    On Error Resume Next
    
    Dim Path$
    Path = lvDir.Path
    
    Dim fmpth As String
    If Tags.BrowsePresetPath = "" Then
        fmpth = Trim$(frmMain.txtFileName.Text)
    Else
        fmpth = Tags.BrowsePresetPath
    End If
    If FolderExists(fmpth) Then
        Path = fmpth
    Else
        Path = GetParentFolderName(fmpth)
        txtFileName.Text = Split(fmpth, "\")(UBound(Split(fmpth, "\")))
    End If
    
'    If Trim$(txtFileName.Text) = "" Then
'        txtFileName.Text = lvFiles.Pattern
'    End If
    
    txtFileName.SelStart = 0
    txtFileName.SelLength = Len(txtFileName.Text)
    
    Dim i%
    For i = 0 To selDrive.ListCount - 1
        If LCase(Left$(selDrive.List(i), 1)) = LCase(Left$(Path, 1)) Then
            selDrive.ListIndex = i
            Exit For
        End If
    Next i
    
    lvDir.Path = Path
    lvDir_Change
    
    Label1.Caption = t(Label1.Caption, "&File name:")
    Label4.Caption = t(Label4.Caption, "File &type:")
    Label3.Caption = t(Label3.Caption, "Dri&ve:")
    Label2.Caption = t(Label2.Caption, "&Directory:")
    chkHidden.Caption = t(chkHidden.Caption, "Show &hidden")
    OKButton.Caption = t(OKButton.Caption, "OK")
    CancelButton.Caption = t(CancelButton.Caption, "Cancel")
    Me.Caption = t(Me.Caption, "Select download path")
    
    txtFileName.Visible = (Tags.BrowseTargetForm <> 2)
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
    If frmMain.cbWhenExist.ListIndex <> 0 And Tags.BrowseTargetForm <> 2 Then _
        OKButton_Click
End Sub

Private Sub OKButton_Click()
    txtFileName.Text = Trim$(txtFileName.Text)
    
    On Error Resume Next
    If InStr(1, txtFileName.Text, "*") > 0 Or InStr(1, txtFileName.Text, "?") > 0 Then
        lvFiles.Pattern = txtFileName.Text
        txtFileName.SelStart = 0
        txtFileName.SelLength = Len(txtFileName.Text)
        MessageBeep 0
        Exit Sub
    End If
    On Error GoTo 0
    
    Dim IsColonPresent As Boolean
    If Len(txtFileName.Text) > 3 And Mid$(txtFileName.Text, 2, 2) = ":\" Then
        IsColonPresent = InStr(1, Right$(txtFileName.Text, Len(txtFileName.Text) - 3), "|") > 0
    Else
        IsColonPresent = Includes(txtFileName.Text, ":")
    End If
    
    If _
        InStr(1, txtFileName.Text, "\") > 0 Or _
        InStr(1, txtFileName.Text, "/") > 0 Or _
        InStr(1, txtFileName.Text, """") > 0 Or _
        InStr(1, txtFileName.Text, "*") > 0 Or _
        InStr(1, txtFileName.Text, "?") > 0 Or _
        InStr(1, txtFileName.Text, "<") > 0 Or _
        InStr(1, txtFileName.Text, ">") > 0 Or _
        InStr(1, txtFileName.Text, "|") > 0 Or _
        IsColonPresent Or _
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
        UCase(txtFileName.Text) = "LPT4" Or _
        (LenB(txtFileName.Text) And Replace(txtFileName.Text, ".", "") = "") Or _
        Right$(txtFileName.Text, 1) = "." _
    Then
        MsgBox t("���� �̸��� �ùٸ��� �ʽ��ϴ�.", "Invalid file name."), 48
        Exit Sub
    End If

    Dim Data$, Path$
    
    If Tags.BrowseTargetForm = 2 Then
        Path = lvFiles.Path
    Else
        If Right$(lvFiles.Path, 1) = "\" Then
            Path = lvFiles.Path & txtFileName.Text
        Else
            Path = lvFiles.Path & "\" & txtFileName.Text
        End If
    End If
    On Error Resume Next
    If FileExists(Path) Then
        If frmMain.cbWhenExist.ListIndex = 0 Then
            MsgBox t("���� �̸��� �̹� �����մϴ�. �ٸ� �̸��� �����Ͻʽÿ�.", "File name already exists."), 16
            Exit Sub
        ElseIf frmMain.cbWhenExist.ListIndex = 1 Then
            If MsgBox(t("���� �̸��� �̹� �����մϴ�. ����ðڽ��ϱ�?", "File name already exists. Overwrite?"), 48 + vbYesNo) <> vbYes Then
                Exit Sub
            End If
        End If
    End If
    
    On Error GoTo e
    If Right$(Path, 2) = "\\" Then Path = Left$(Path, Len(Path) - 1)
    Select Case Tags.BrowseTargetForm
        Case 1
            frmEditBatch.txtFilePath.Text = Path
        Case 2
            frmBatchAdd.txtSavePath.Text = Path
        Case Else
            frmMain.txtFileName.Text = Path
    End Select
    
    Unload Me
    Exit Sub
    
e:
    MsgBox t("������ �߻��߽��ϴ�!", "Error!"), 16
    Exit Sub
End Sub

Private Sub selDrive_Change()
    On Error GoTo e
    lvDir.Path = selDrive.Drive
    Exit Sub
    
e:
    MsgBox t("����̺긦 �� �� �����ϴ�. ������ ����̺� �ȿ� ��ũ�� ���ų� ����̺갡 ��� �ֽ��ϴ�.", "The drive is inaccessible. There is no disk in the selected drive or the drive is locked."), 16
End Sub

Private Sub selFileType_Change()
    selFileType_Click
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
