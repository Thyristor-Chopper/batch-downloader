VERSION 5.00
Begin VB.Form frmExplorer 
   Caption         =   "�ٿ�ε� ��� ����"
   ClientHeight    =   8325
   ClientLeft      =   2775
   ClientTop       =   3885
   ClientWidth     =   9750
   BeginProperty Font 
      Name            =   "����"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmExplorer.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   9750
   StartUpPosition =   1  '������ ���
   Begin prjDownloadBooster.CommandButtonW cmdPreview 
      Height          =   345
      Left            =   8160
      TabIndex        =   9
      Top             =   5280
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      Caption         =   "�̸� ���(&P)"
   End
   Begin prjDownloadBooster.CommandButtonW CancelButton 
      Cancel          =   -1  'True
      Height          =   340
      Left            =   8160
      TabIndex        =   5
      Top             =   4920
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   609
      Caption         =   "���"
   End
   Begin prjDownloadBooster.CommandButtonW OKButton 
      Default         =   -1  'True
      Height          =   340
      Left            =   8160
      TabIndex        =   4
      Top             =   4560
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   609
      Caption         =   "Ȯ��"
   End
   Begin prjDownloadBooster.CheckBoxW chkShowFiles 
      Height          =   255
      Left            =   6240
      TabIndex        =   8
      Top             =   5340
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   450
      Caption         =   "���� ǥ��(&S)"
   End
   Begin prjDownloadBooster.ImageList imgPlaces 
      Left            =   8640
      Top             =   7200
      _ExtentX        =   1005
      _ExtentY        =   1005
      ImageWidth      =   32
      ImageHeight     =   32
      ColorDepth      =   4
      MaskColor       =   16711935
      InitListImages  =   "frmExplorer.frx":000C
   End
   Begin VB.PictureBox pbPlacesBarContainer 
      BackColor       =   &H8000000C&
      Height          =   5175
      Left            =   120
      ScaleHeight     =   5115
      ScaleWidth      =   1410
      TabIndex        =   18
      TabStop         =   0   'False
      Tag             =   "nobackcolorchange"
      Top             =   480
      Width           =   1470
      Begin prjDownloadBooster.ToolBar tbPlaces 
         Height          =   5175
         Left            =   0
         Tag             =   "nobackcolorchange novisualstylechange"
         Top             =   0
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   9128
         VisualStyles    =   0   'False
         ImageList       =   "imgPlaces"
         BackColor       =   -2147483636
         Style           =   1
         Orientation     =   1
         Divider         =   0   'False
         AllowCustomize  =   0   'False
         ButtonHeight    =   35
         ButtonWidth     =   94
         MinButtonWidth  =   94
         MaxButtonWidth  =   94
         InitButtons     =   "frmExplorer.frx":111C
      End
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
      Height          =   300
      Left            =   120
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   6360
      Visible         =   0   'False
      Width           =   2175
   End
   Begin prjDownloadBooster.ImageCombo cbFolderList 
      Height          =   315
      Left            =   1680
      TabIndex        =   11
      Top             =   120
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   582
      ImageList       =   "imgFolderSmall"
      Style           =   2
      MaxDropDownItems=   20
   End
   Begin prjDownloadBooster.ImageList imgFolderSmall 
      Left            =   8640
      Top             =   6480
      _ExtentX        =   1005
      _ExtentY        =   1005
      ImageWidth      =   16
      ImageHeight     =   16
      ColorDepth      =   32
      MaskColor       =   16711935
      InitListImages  =   "frmExplorer.frx":172C
   End
   Begin prjDownloadBooster.ImageList imgFolder 
      Left            =   8640
      Top             =   5760
      _ExtentX        =   1005
      _ExtentY        =   1005
      ImageWidth      =   32
      ImageHeight     =   32
      ColorDepth      =   32
      MaskColor       =   16711935
      InitListImages  =   "frmExplorer.frx":409C
   End
   Begin VB.PictureBox picPreviewFrame 
      Enabled         =   0   'False
      Height          =   2415
      Left            =   2880
      ScaleHeight     =   2355
      ScaleWidth      =   4140
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   5760
      Width           =   4200
      Begin VB.Image imgPreview 
         Height          =   2355
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   4140
      End
   End
   Begin prjDownloadBooster.CommandButtonW cmdViews 
      Height          =   330
      Left            =   9090
      TabIndex        =   15
      Top             =   120
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   582
      Caption         =   "v"
   End
   Begin prjDownloadBooster.ListView lvFiles 
      Height          =   3960
      Left            =   1680
      TabIndex        =   12
      Top             =   480
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   6985
      VisualTheme     =   1
      Icons           =   "imgFolder"
      SmallIcons      =   "imgFolderSmall"
      Arrange         =   2
      AllowColumnReorder=   -1  'True
      HideSelection   =   0   'False
      ShowInfoTips    =   -1  'True
      ShowLabelTips   =   -1  'True
      ShowColumnTips  =   -1  'True
      HighlightColumnHeaders=   -1  'True
      UseColumnChevron=   -1  'True
      AutoSelectFirstItem=   0   'False
   End
   Begin prjDownloadBooster.CheckBoxW chkHidden 
      Height          =   255
      Left            =   1680
      TabIndex        =   6
      Top             =   5340
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   450
      Caption         =   "���� ���� ǥ��(&H)"
   End
   Begin VB.TextBox txtFileName 
      Height          =   270
      Left            =   2880
      TabIndex        =   1
      Top             =   4590
      Width           =   5175
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
      Left            =   2880
      Style           =   2  '��Ӵٿ� ���
      TabIndex        =   3
      Top             =   4950
      Width           =   5175
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
      Left            =   120
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   6720
      Visible         =   0   'False
      Width           =   2175
   End
   Begin prjDownloadBooster.ToolBar tbToolBar 
      Height          =   330
      Left            =   8400
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      ImageList       =   "imgFolderSmall"
      Style           =   1
      Divider         =   0   'False
      ShowTips        =   -1  'True
      Wrappable       =   0   'False
      AllowCustomize  =   0   'False
      ButtonWidth     =   23
      InitButtons     =   "frmExplorer.frx":695C
   End
   Begin prjDownloadBooster.CheckBoxW chkUnixHidden 
      Height          =   255
      Left            =   3960
      TabIndex        =   7
      Top             =   5340
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   450
      Caption         =   "������ ���� ǥ��(&U)"
   End
   Begin VB.Label Label2 
      Alignment       =   1  '������ ����
      BackStyle       =   0  '����
      Caption         =   "ã�� ��ġ(&I):"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   180
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '����
      Caption         =   "�̸�����:"
      Height          =   255
      Left            =   1680
      TabIndex        =   17
      Top             =   5760
      Width           =   1215
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
      Left            =   1680
      TabIndex        =   2
      Top             =   5010
      Width           =   1215
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
      Left            =   1680
      TabIndex        =   0
      Top             =   4620
      Width           =   1215
   End
   Begin VB.Menu mnuFolderFloor 
      Caption         =   "����"
      Visible         =   0   'False
      Begin VB.Menu mnuNewFolder 
         Caption         =   "�� ����(&N)"
      End
      Begin VB.Menu mnuCmd 
         Caption         =   "��� ������Ʈ(&M)"
      End
      Begin VB.Menu mnuView 
         Caption         =   "����(&V)"
         Begin VB.Menu mnuIconSize 
            Caption         =   "ū ������(&L)"
            Index           =   0
         End
         Begin VB.Menu mnuIconSize 
            Caption         =   "���� ������(&S)"
            Index           =   1
         End
         Begin VB.Menu mnuIconSize 
            Caption         =   "������(&I)"
            Index           =   2
         End
         Begin VB.Menu mnuIconSize 
            Caption         =   "�ڼ���(&D)"
            Index           =   3
         End
         Begin VB.Menu mnuIconSize 
            Caption         =   "������ ����(&T)"
            Index           =   4
         End
      End
      Begin VB.Menu mnuSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "���� ��ħ(&E)"
      End
      Begin VB.Menu mnuSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFolderProperties 
         Caption         =   "�Ӽ�(&R)"
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "����(&F)"
      Visible         =   0   'False
      Begin VB.Menu mnuSelect 
         Caption         =   "����(&L)"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "����(&O)"
      End
      Begin VB.Menu mnuExplore 
         Caption         =   "Ž��(&X)"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "����(&D)"
      End
      Begin VB.Menu mnuRename 
         Caption         =   "�̸� �ٲٱ�(&M)"
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProperties 
         Caption         =   "�Ӽ�(&R)"
      End
   End
End
Attribute VB_Name = "frmExplorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Pattern$
Dim IsMyComputer As Boolean
'Dim mnuTop&, mnuBottom&, mnuViewID&
Dim Loaded As Boolean
Dim ListedOn As String
Dim ExtToIcon As Collection
Dim ExtToSmallIcon As Collection
Dim FirstListed As Boolean
Dim LoadFinished As Boolean

Implements IBSSubclass

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub cbFolderList_Click()
    Dim i%
    Dim Path$
    
    For i = 1 To tbPlaces.Buttons.Count
        tbPlaces.Buttons(i).Value = TbrButtonValueUnpressed
    Next i
    
    Select Case cbFolderList.SelectedItem.Index
        Case 1
            lvDir.Path = GetSpecialFolder(CSIDL_RECENT)
            tbPlaces.Buttons(1).Value = TbrButtonValuePressed
        Case 2
            lvDir.Path = GetSpecialFolder(CSIDL_DESKTOP)
            tbPlaces.Buttons(2).Value = TbrButtonValuePressed
        Case 3
            lvDir.Path = GetSpecialFolder(CSIDL_PERSONAL)
            tbPlaces.Buttons(3).Value = TbrButtonValuePressed
        Case 4
            ShowMyComputer
            Exit Sub
    End Select
    
    If IsMyComputer Then ListedOn = ""
    
    If cbFolderList.SelectedItem.Indentation = 2 Then
        On Error Resume Next
        'selDrive.ListIndex = cbFolderList.SelectedItem.Index - 5
        Dim DriveLetter$
        DriveLetter = LCase(Left$(cbFolderList.SelectedItem.Text, 1))
        For i = 0 To selDrive.ListCount
            If LCase(Left$(selDrive.List(i), 1)) = DriveLetter Then
                selDrive.ListIndex = i
retrydrive:
                lvDir.Path = DriveLetter & ":\"
                
                If Err Then
                    If MsgBox(t("����̺긦 �� �� �����ϴ�. ������ ����̺� �ȿ� ��ũ�� ���ų� ����̺갡 ��� �ֽ��ϴ�.", "The drive is inaccessible. There is no disk in the selected drive or the drive is locked."), vbRetryCancel + vbCritical) = vbRetry Then
                        GoTo retrydrive
                    Else
                        GoTo exitsub
                    End If
                End If
                
                If Loaded Then ListFiles
                Exit Sub
            End If
        Next i
        Exit Sub
    ElseIf cbFolderList.SelectedItem.Indentation > 2 Then
        Path = UCase(Left$(cbFolderList.ComboItems(cbFolderList.SelectedItem.Index - cbFolderList.SelectedItem.Indentation + 2).Text, 1)) & ":\"
        For i = cbFolderList.SelectedItem.Index - cbFolderList.SelectedItem.Indentation + 3 To cbFolderList.SelectedItem.Index
            Path = Path & cbFolderList.ComboItems(i).Text & "\"
        Next i
        lvDir.Path = Path
    ElseIf cbFolderList.SelectedItem.Indentation = 1 And cbFolderList.SelectedItem.Index > 4 Then
        Path = GetSpecialFolder(CSIDL_DESKTOP)
        If Right$(Path, 1) <> "\" Then Path = Path & "\"
        lvDir.Path = Path & cbFolderList.SelectedItem.Text
    End If
    
    If Loaded Then ListFiles
exitsub:
End Sub

Private Sub chkHidden_Click()
    ListedOn = ""
    If Loaded Then ListFiles
    SaveSetting "DownloadBooster", "UserData", "ShowHidden", chkHidden.Value
End Sub

Private Sub ListFiles()
    If LenB(ListedOn) And ListedOn = lvDir.Path Then Exit Sub
    ListedOn = lvDir.Path
    LoadFinished = False

    Dim li As LvwListItem
    Dim i%, k#
    On Error Resume Next
    If Not lvFiles.SelectedItem Is Nothing Then
        lvFiles.SelectedItem.Selected = False
        Set lvFiles.SelectedItem = Nothing
    End If
    lvFiles.ListItems.Clear
    If IsMyComputer Then
        IsMyComputer = False
        lvFiles.ColumnHeaders(2).Text = t("ũ��", "Size")
        lvFiles.ColumnHeaders(2).Alignment = LvwColumnHeaderAlignmentRight
        lvFiles.ColumnHeaders(2).Width = 1455
        lvFiles.ColumnHeaders(3).Text = t("����", "Type")
        lvFiles.ColumnHeaders(3).Alignment = LvwColumnHeaderAlignmentLeft
        lvFiles.ColumnHeaders(3).Width = 1800
        lvFiles.ColumnHeaders(4).Text = t("������ ��¥", "Modified")
        lvFiles.ColumnHeaders(4).Alignment = LvwColumnHeaderAlignmentLeft
        lvFiles.ColumnHeaders(4).Width = 2250
    End If
    lvFiles.GroupView = False
    
    If imgFolder.ListImages.Count > 10 Then
        For i = 11 To imgFolder.ListImages.Count
            imgFolder.ListImages.Remove 11
        Next i
    End If
    If imgFolderSmall.ListImages.Count > 18 Then
        For i = 19 To imgFolderSmall.ListImages.Count
            imgFolderSmall.ListImages.Remove 19
        Next i
    End If
    
    If ExtToIcon.Count Then
        For i = 1 To ExtToIcon.Count
            ExtToIcon.Remove 1
        Next i
    End If
    If ExtToSmallIcon.Count Then
        For i = 1 To ExtToSmallIcon.Count
            ExtToSmallIcon.Remove 1
        Next i
    End If
    
    Dim Path$, Name$
    Path = lvDir.Path
    Dim IsDesktop As Boolean
    IsDesktop = (Path = GetSpecialFolder(CSIDL_DESKTOP))
    
    Dim totalcnt As Double
    If Right$(Path, 1) <> "\" Then Path = Path & "\"
    On Error Resume Next
    tbToolBar.Buttons(2).Enabled = False
    
    If Len(lvDir.Path) > 3 Then
        tbToolBar.Buttons(2).Enabled = True
        If lvDir.Path <> GetSpecialFolder(CSIDL_DESKTOP) And lvDir.Path <> GetSpecialFolder(CSIDL_RECENT) Then
            Set li = lvFiles.ListItems.Add(, , "..", 1, 1)
            li.ListSubItems.Add , , "-"
            li.ListSubItems.Add , , t("���� ����", "Parent Folder")
            li.ListSubItems.Add , , "-"
            totalcnt = 1
        End If
    End If
    
    Dim Attributes As VbFileAttribute
    Attributes = (vbHidden * chkHidden.Value) Or vbReadOnly Or vbArchive
    
    Name = Dir$(Path, vbDirectory Or Attributes)
    Dim FullPath As String
    Do While LenB(Name)
        If Name <> "." And Name <> ".." Then
            FullPath = Path & Name
            If (GetAttr(FullPath) And vbDirectory) = vbDirectory Then
                If ((Not chkUnixHidden.Value) And Left$(Name, 1) = ".") Or InStr(FullPath, "?") Then GoTo nextdir
                
                Set li = lvFiles.ListItems.Add(, , Name, 1, 1)
                li.ListSubItems.Add Text:="-"
                li.ListSubItems.Add Text:=t("���� ����", "File Folder")
                li.ListSubItems.Add Text:="-"
                If Name <> ".." Then li.ListSubItems(3).Text = FormatModified(FileDateTime(FullPath))
                
                TotalCntProc totalcnt
            End If
        End If
nextdir:
        Name = Dir$
    Loop
    
    Dim PatternMatched As Boolean
    Dim PatternsSplit() As String
    Dim CurrentPattern$
    Dim ExtName$
    Dim Icon%, SmallIcon%
    PatternsSplit = Split(Pattern, ";")
    Dim cnt As Double
    Dim ext As String
    Dim UseFileAttr As Boolean
    Dim ShellIcon As IPicture, ShellSmallIcon As IPicture
    Dim IconFlags As Long
    If Tags.BrowseTargetForm = 2 And chkShowFiles.Value = 0 Then GoTo afterfileload
    Name = Dir$(Path, Attributes And (Not vbDirectory))
    Do While LenB(Name)
        FullPath = Path & Name
        PatternMatched = False
        For i = LBound(PatternsSplit) To UBound(PatternsSplit)
            CurrentPattern = Trim$(PatternsSplit(i))
            If CurrentPattern = "*.*" Then CurrentPattern = "*"
            PatternMatched = (LCase(Name) Like LCase(CurrentPattern))
            If PatternMatched Then Exit For
        Next i
        If (Not PatternMatched) Or InStr(FullPath, "?") Or (chkUnixHidden.Value = 0 And Left$(Name, 1) = ".") Then GoTo NextItem

        ext = UCase(GetExtensionName(Name))
        If ext = "LNK" Then
            If FolderExists(RemoveQuotes(GetShortcutTarget(FullPath))) Then
                Icon = 1
                SmallIcon = 1
                GoTo aftericonproc
            End If
        End If
        
        Icon = 2
        SmallIcon = 2
        UseFileAttr = Not (ext = "EXE" Or ext = "LNK" Or ext = "PIF" Or ext = "ICO")
        ShellGetFileInfo FullPath, UseFileAttr, ShellIcon, ShellSmallIcon, ExtName
        If Not UseFileAttr Then
            GoTo addicon
        ElseIf Exists(ExtToIcon, ext) Then
            Icon = ExtToIcon(ext)
            SmallIcon = ExtToSmallIcon(ext)
        Else
addicon:
            If cnt < 250 Then
                If ShellIcon Is Nothing Or ShellSmallIcon Is Nothing Then GoTo aftericonproc
                Icon = imgFolder.ListImages.Add(, , ShellIcon).Index
                SmallIcon = imgFolderSmall.ListImages.Add(, , ShellSmallIcon).Index
                If UseFileAttr Then
                    ExtToIcon.Add Icon, ext
                    ExtToSmallIcon.Add SmallIcon, ext
                End If
                cnt = cnt + 1
            End If
        End If
aftericonproc:
        Set li = lvFiles.ListItems.Add(, , Name, Icon, SmallIcon)
        li.ListSubItems.Add Text:=ParseSize(FileLen(FullPath))
        If LenB(Trim$(ExtName)) = 0 Then ExtName = ext & " " & t("����", "File")
        li.ListSubItems.Add Text:=ExtName
        li.ListSubItems.Add Text:=FormatModified(FileDateTime(FullPath))
        
        If Not FirstListed Then
            If Tags.BrowseTargetForm >= 3 And Tags.BrowseTargetForm <= 6 Then
                If LCase(Name) = LCase(GetFilename(Tags.BrowsePresetPath)) Then
                    li.Selected = True
                    li.EnsureVisible
                End If
            End If
        End If
        
        TotalCntProc totalcnt
NextItem:
        Name = Dir$
    Loop
afterfileload:
    
    If IsDesktop Then
        Set li = lvFiles.ListItems.Add(1, , t("�� ��ǻ��", "My Computer"), 9, 14)
        li.ListSubItems.Add , , "-"
        li.ListSubItems.Add , , t("�ý��� ����", "System Folder")
        li.ListSubItems.Add , , "-"
        
        Set li = lvFiles.ListItems.Add(1, , t("�� ����", "My Documents"), 10, 13)
        li.ListSubItems.Add , , "-"
        li.ListSubItems.Add , , t("�ý��� ����", "System Folder")
        li.ListSubItems.Add , , "-"
    End If
    
    tbToolBar.Buttons(3).Enabled = True
    FirstListed = True
    LoadFinished = True
    cbFolderList.Enabled = True
    'tbPlaces.Enabled = True
    tbToolBar.Enabled = True
    chkHidden.Enabled = True
    chkUnixHidden.Enabled = True
    chkShowFiles.Enabled = True
    selFileType.Enabled = True
    OKButton.Enabled = True
    CancelButton.Enabled = True
    Label1.Enabled = True
    Label4.Enabled = True
    txtFileName.Enabled = True
    Label2.Enabled = True
End Sub

Private Sub TotalCntProc(ByRef totalcnt As Double)
    If totalcnt >= 250 Then
        If totalcnt = 250 Then
            cbFolderList.Enabled = 0
            'tbPlaces.Enabled = 0
            tbToolBar.Enabled = 0
            chkHidden.Enabled = 0
            chkUnixHidden.Enabled = 0
            chkShowFiles.Enabled = 0
            selFileType.Enabled = 0
            OKButton.Enabled = 0
            CancelButton.Enabled = 0
            Label1.Enabled = 0
            Label4.Enabled = 0
            txtFileName.Enabled = 0
            Label2.Enabled = 0
        End If
        If totalcnt Mod 100 = 0 Then DoEvents
    End If
    totalcnt = totalcnt + 1
End Sub

Private Sub chkShowFiles_Click()
    ListedOn = ""
    If Loaded Then ListFiles
    SaveSetting "DownloadBooster", "UserData", "ShowFiles", chkShowFiles.Value
End Sub

Private Sub chkUnixHidden_Click()
    ListedOn = ""
    If Loaded Then ListFiles
    SaveSetting "DownloadBooster", "UserData", "ShowUnixHidden", chkUnixHidden.Value
End Sub

Private Sub cmdPreview_Click()
    If lvFiles.SelectedItem Is Nothing Then Exit Sub
    If Not lvFiles.SelectedItem.Selected Then Exit Sub
    
    Dim Path$
    Path = lvDir.Path
    If Right$(Path, 1) <> "\" Then Path = Path & "\"
    PlayWave Path & lvFiles.SelectedItem.Text
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    txtFileName.SetFocus
    Loaded = True
    If Not LoadFinished Then Exit Sub
    ListFiles
End Sub

Private Sub Form_Load()
    On Error Resume Next
    InitForm Me
    LoadFinished = True
    
    Set ExtToIcon = New Collection
    Set ExtToSmallIcon = New Collection
    
    lvFiles.ColumnHeaders.Add , , t("�̸�", "Name"), 2295
    lvFiles.ColumnHeaders.Add(, , t("ũ��", "Size"), 1455).Alignment = LvwColumnHeaderAlignmentRight
    lvFiles.ColumnHeaders.Add , , t("����", "Type"), 1800
    lvFiles.ColumnHeaders.Add , , t("������ ��¥", "Modified"), 2250
    
    'lvFiles.ColumnHeaders(1).SortArrow = LvwColumnHeaderSortArrowUp
    
    Select Case Tags.BrowseTargetForm
        Case 3, 5, 6
            AddItemToComboBox selFileType, t("��� �׸�", "All pictures") & " (*.JPG; *.JPEG; *.JPE; *.JFIF; *.GIF; *.BMP; *.DIB; *.RLE; *.PNG; *.TIF; *.TIFF; *.WMF; *.EMF; *.ICO; *.CUR)"
            AddItemToComboBox selFileType, "JPEG (*.JPG; *.JPEG; *.JPE; *.JFIF)"
            AddItemToComboBox selFileType, "GIF (*.GIF)"
            AddItemToComboBox selFileType, t("��Ʈ��", "Bitmap") & " (*.BMP; *.DIB; *.RLE)"
            AddItemToComboBox selFileType, "PNG (*.PNG)"
            AddItemToComboBox selFileType, "TIFF (*.TIF; *.TIFF)"
            AddItemToComboBox selFileType, t("��Ÿ����", "Metafile") & " (*.WMF; *.EMF)"
            AddItemToComboBox selFileType, t("������", "Icon") & " (*.ICO)"
            AddItemToComboBox selFileType, t("Ŀ��", "Cursor") & " (*.CUR)"
        Case 4
            AddItemToComboBox selFileType, t("�Ҹ�", "Sound") & " (*.WAV)"
        Case Else
            AddItemToComboBox selFileType, t("��� ����", "All files") & " (*.*)"
    End Select
    selFileType.ListIndex = 0
    
    Dim Path$
    Path = lvDir.Path
    
    Dim fmpth As String
    If Tags.BrowseTargetForm = 3 Then
        fmpth = GetSetting("DownloadBooster", "Options", "BackgroundImagePath", "")
setpreview:
        Set imgPreview.Picture = LoadPictureFromFile(fmpth)
    ElseIf Tags.BrowseTargetForm = 5 Then
        fmpth = GetSetting("DownloadBooster", "Options", "LiveBadukMemoSkinFrameTexture", "")
        GoTo setpreview
    ElseIf Tags.BrowseTargetForm = 6 Then
        fmpth = GetSetting("DownloadBooster", "Options", "LiveBadukMemoSkinFrameBackground", "")
        GoTo setpreview
    Else
        If Tags.BrowsePresetPath = "" Then
            fmpth = Trim$(frmMain.txtFileName.Text)
        Else
            fmpth = Tags.BrowsePresetPath
        End If
    End If
    If FolderExists(fmpth) Then
        Path = fmpth
    ElseIf FolderExists(GetParentFolderName(fmpth)) Then
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
    
    tr Label1, "&File name:"
    tr OKButton, "OK"
    Select Case Tags.BrowseTargetForm
        Case 2
            Label1.Caption = t("���� �̸�(&F):", "&Folder name:")
            txtFileName.Width = 6735
            OKButton.Top = CancelButton.Top
            OKButton.Left = CancelButton.Left - 120 - OKButton.Width
            OKButton.Caption = t("���� ����(&E)", "S&elect Folder")
        Case 0
            If txtFileName.Text = "" Then
                OKButton.Caption = t("���� ����(&E)", "S&elect Folder")
            Else
                OKButton.Caption = t("����(&E)", "Sav&e")
            End If
        Case 1
            OKButton.Caption = t("����(&E)", "Sav&e")
        Case 3, 4
            OKButton.Caption = t("����(&O)", "&Open")
    End Select
    tr Label4, "File &type:"
    Label4.Visible = Tags.BrowseTargetForm <> 2
    selFileType.Visible = Tags.BrowseTargetForm <> 2
    tr Label2, "F&ind in:"
    tr chkHidden, "Show &hidden files"
    tr chkUnixHidden, "Show &Unix hidden"
    tr CancelButton, "Cancel"
    tr Me, "Select download path"
    If Tags.BrowseTargetForm = 3 Then Me.Caption = t("��� ���� ����", "Choose background image")
    If Tags.BrowseTargetForm = 4 Then Me.Caption = t("ȿ���� ����", "Choose sound")
    If Tags.BrowseTargetForm = 5 Or Tags.BrowseTargetForm = 6 Then Me.Caption = t("�ؽ�ó ����", "Choose texture")
    chkShowFiles.Visible = (Tags.BrowseTargetForm = 2)
    tr chkShowFiles, "&Show files"
    tr Label5, "Preview:"
    tr cmdPreview, "&Preview"
    
    tbPlaces.Buttons(1).Caption = t("�� �ֱ� ����", "Recent")
    tbPlaces.Buttons(2).Caption = t("���� ȭ��", "Desktop")
    tbPlaces.Buttons(3).Caption = t("�� ����", "My Documents")
    tbPlaces.Buttons(4).Caption = t("�� ��ǻ��", "My Computer")
    tbPlaces.Buttons(5).Caption = t("���ã��", "Favorites")
    tbPlaces.Buttons(6).Caption = t("�� ������", "My Profile")
    
    tbToolBar.Buttons(1).ToolTipText = t("�ڷ�", "Back")
    tbToolBar.Buttons(2).ToolTipText = t("�� ���� ����", "Up one level")
    tbToolBar.Buttons(3).ToolTipText = t("�� ���� �����", "New folder")
    tbToolBar.Buttons(4).ToolTipText = t("����", "Icon size")
    
    tr mnuNewFolder, "&New folder"
    tr mnuCmd, "Open Co&mmand Prompt"
    tr mnuView, "&View"
    tr mnuIconSize(0), "&Large icons"
    tr mnuIconSize(1), "&Small icons"
    tr mnuIconSize(2), "L&ist"
    tr mnuIconSize(3), "&Details"
    tr mnuIconSize(4), "&Tiles"
    tr mnuRefresh, "&Refresh"
    tr mnuFolderProperties, "P&roperties"
    tr mnuSelect, "Se&lect"
    tr mnuOpen, "&Open"
    tr mnuExplore, "E&xplore"
    tr mnuDelete, "&Delete"
    tr mnuRename, "Rena&me"
    tr mnuProperties, "P&roperties"
    
    lvDir.Path = Path
    selFileType_Click
    
    If Tags.BrowseTargetForm = 3 Or Tags.BrowseTargetForm = 5 Or Tags.BrowseTargetForm = 6 Then
        'Me.Height = 8280
    Else
        Me.Height = 6165
    End If
    
    lvFiles.View = GetSetting("DownloadBooster", "UserData", "FileListView", LvwViewConstants.LvwViewReport)
    
    lvFiles.Groups.Add , , t("�� ��ǻ�Ϳ� ����� ����", "Files Stored on This Computer")
    lvFiles.Groups.Add , , t("�ϵ� ��ũ ����̺�", "Hard Disk Drives")
    lvFiles.Groups.Add , , t("�̵��� ����Ұ� �ִ� ��ġ", "Drives with Removable Storage")
    lvFiles.Groups.Add , , t("��Ʈ��ũ ����̺�", "Network Drives")
    lvFiles.Groups.Add , , t("��Ÿ", "Others")
    
    chkHidden.Value = GetSetting("DownloadBooster", "UserData", "ShowHidden", 0)
    chkUnixHidden.Value = GetSetting("DownloadBooster", "UserData", "ShowUnixHidden", 1)
    chkShowFiles.Value = GetSetting("DownloadBooster", "UserData", "ShowFiles", 0)
    
    cmdPreview.Visible = (Tags.BrowseTargetForm = 4)
    
    On Error Resume Next
    'Me.Icon = frmMain.Icon
    Me.Width = GetSetting("DownloadBooster", "UserData", "ComdlgWidth", 10245) + PaddedBorderWidth * 15 * 2
    Me.Height = GetSetting("DownloadBooster", "UserData", "ComdlgHeight", 6165) + IIf(Tags.BrowseTargetForm = 3 Or Tags.BrowseTargetForm = 5 Or Tags.BrowseTargetForm = 6, 8835 - 6165, 0) + PaddedBorderWidth * 15 * 2
    
    AttachMessage Me, Me.hWnd, WM_GETMINMAXINFO
    AttachMessage Me, Me.hWnd, WM_SETTINGCHANGE
    
    mnuIconSize(lvFiles.View).Checked = True
    
    If WinVer >= 6# And Build >= 5048 Then lvFiles.FullRowSelect = True
    
    If Tags.BrowseTargetForm <> 3 And Tags.BrowseTargetForm <> 5 And Tags.BrowseTargetForm <> 6 Then
        Label5.Visible = 0
        picPreviewFrame.Visible = 0
    End If
    
    lvDir_Change
End Sub

'Sub CheckViewRadio()
'    mnuFolderFloor.Visible = -1
'    mnuViewID = GetSubMenu(GetSubMenu(GetMenu(Me.hWnd), 0), 2)
'    mnuTop = GetMenuItemID(mnuViewID, 0)
'    mnuBottom = GetMenuItemID(mnuViewID, 4)
'    Dim CurrentView As LvwViewConstants
'    CurrentView = lvFiles.View
'    Select Case CurrentView
'        Case LvwViewIcon
'            CheckMenuRadioItem mnuViewID, mnuTop, mnuBottom, mnuTop, MF_BYCOMMAND
'        Case LvwViewSmallIcon
'            CheckMenuRadioItem mnuViewID, mnuTop, mnuBottom, mnuTop + 1, MF_BYCOMMAND
'        Case LvwViewList
'            CheckMenuRadioItem mnuViewID, mnuTop, mnuBottom, mnuTop + 2, MF_BYCOMMAND
'        Case LvwViewReport
'            CheckMenuRadioItem mnuViewID, mnuTop, mnuBottom, mnuTop + 3, MF_BYCOMMAND
'        Case LvwViewTile
'            CheckMenuRadioItem mnuViewID, mnuTop, mnuBottom, mnuTop + 4, MF_BYCOMMAND
'    End Select
'    mnuFolderFloor.Visible = 0
'End Sub

Private Sub ShowMyComputer()
    IsMyComputer = True
    LoadFinished = True
    ListedOn = ""
    
    lvFiles.ColumnHeaders(2).Text = t("����", "Type")
    lvFiles.ColumnHeaders(2).Alignment = LvwColumnHeaderAlignmentLeft
    lvFiles.ColumnHeaders(2).Width = 1800
    lvFiles.ColumnHeaders(3).Text = t("��ü ũ��", "Total Space")
    lvFiles.ColumnHeaders(3).Alignment = LvwColumnHeaderAlignmentRight
    lvFiles.ColumnHeaders(3).Width = 1455
    lvFiles.ColumnHeaders(4).Text = t("��� ����", "Free Space")
    lvFiles.ColumnHeaders(4).Alignment = LvwColumnHeaderAlignmentRight
    lvFiles.ColumnHeaders(4).Width = 1455
    
    lvFiles.GroupView = True
    Dim Icon%
    lvFiles.ListItems.Clear
    Dim k%
    Dim Item As LvwListItem
    Dim DriveType$
    Dim Group As LvwGroup
    Dim TotalSpace As Double
    Dim FreeSpace As Double
    On Error Resume Next
    For k = 0 To selDrive.ListCount - 1
        Select Case GetDriveType(Left$(selDrive.List(k), 2))
            Case DRIVE_FIXED, DRIVE_UNKNOWN, DRIVE_NO_ROOT_DIR
                Icon = 6
                DriveType = t("���� ��ũ", "Local Disk")
                Set Group = lvFiles.Groups(2)
            Case DRIVE_REMOVABLE
                If LCase(Left$(selDrive.List(k), 1)) < "c" Then
                    Icon = 3
                    DriveType = t("����", "Diskette")
                Else
                    Icon = 4
                    DriveType = t("�̵��� ��ũ", "Removable Disk")
                End If
                Set Group = lvFiles.Groups(3)
            Case DRIVE_CDROM
                Icon = 5
                DriveType = t("CD/DVD ����̺�", "CD/DVD Drive")
                Set Group = lvFiles.Groups(3)
            Case DRIVE_REMOTE
                Icon = 7
                DriveType = t("��Ʈ��ũ ����̺�", "Network Drive")
                Set Group = lvFiles.Groups(4)
            Case DRIVE_RAMDISK
                Icon = 8
                DriveType = t("RAM ��ũ", "RAM Disk")
                Set Group = lvFiles.Groups(5)
        End Select
        Set Item = lvFiles.ListItems.Add(, , selDrive.List(k), Icon, Icon)
        GetDiskSpace UCase(Left$(selDrive.List(k), 2)), TotalSpace, FreeSpace
        Item.ListSubItems.Add , , DriveType
        Item.ListSubItems.Add , , ParseSize(TotalSpace)
        Item.ListSubItems.Add , , ParseSize(FreeSpace)
        Set Item.Group = Group
    Next k
    
    For k = 1 To tbPlaces.Buttons.Count
        tbPlaces.Buttons(k).Value = TbrButtonValueUnpressed
    Next k
    tbPlaces.Buttons(4).Value = TbrButtonValuePressed
    cbFolderList.ComboItems(4).Selected = True
    tbToolBar.Buttons(2).Enabled = True
    tbToolBar.Buttons(3).Enabled = False
End Sub

Sub Form_Resize()
    cbFolderList.Width = Me.Width - PaddedBorderWidth * 15 * 2 - (9870 - 6495)
    tbToolBar.Left = Me.Width - PaddedBorderWidth * 15 * 2 - 1215 - 120 - 120
    cmdViews.Left = Me.Width - PaddedBorderWidth * 15 * 2 - 495 - 120 - 30 - 120
    lvFiles.Width = Me.Width - PaddedBorderWidth * 15 * 2 - (9870 - 7935)
    CancelButton.Left = Me.Width - PaddedBorderWidth * 15 * 2 - CancelButton.Width - 120 - 120 - 15
    If Tags.BrowseTargetForm = 3 Or Tags.BrowseTargetForm = 5 Or Tags.BrowseTargetForm = 6 Then
        lvFiles.Height = Me.Height - PaddedBorderWidth * 15 * 2 - (8835 - 3960)
    Else
        lvFiles.Height = Me.Height - PaddedBorderWidth * 15 * 2 - (6165 - 3960)
    End If
    CancelButton.Top = lvFiles.Top + lvFiles.Height + 120 + OKButton.Height + 45
    txtFileName.Width = Me.Width - PaddedBorderWidth * 15 * 2 - (9870 - 5175)
    selFileType.Width = Me.Width - PaddedBorderWidth * 15 * 2 - (9870 - 5175)
    Label1.Top = lvFiles.Height + lvFiles.Top + 195
    Label4.Top = lvFiles.Height + lvFiles.Top + 195 + 360 + 15
    txtFileName.Top = Label1.Top - 45
    selFileType.Top = Label4.Top - 45
    If Tags.BrowseTargetForm = 2 Then
        OKButton.Left = CancelButton.Left - 120 - OKButton.Width
        OKButton.Top = CancelButton.Top
        txtFileName.Width = txtFileName.Width + 120 + OKButton.Width - 15
    Else
        OKButton.Left = CancelButton.Left
        OKButton.Top = CancelButton.Top - OKButton.Height - 60
    End If
    chkHidden.Top = selFileType.Top + selFileType.Height + 60
    chkUnixHidden.Top = selFileType.Top + selFileType.Height + 60
    chkShowFiles.Top = selFileType.Top + selFileType.Height + 60
    pbPlacesBarContainer.Height = chkHidden.Top + chkHidden.Height - pbPlacesBarContainer.Top
    tbPlaces.Height = pbPlacesBarContainer.Height
    Label5.Top = chkHidden.Top + chkHidden.Height + 180
    picPreviewFrame.Top = Label5.Top
    cmdPreview.Left = CancelButton.Left - 120 - CancelButton.Width
    cmdPreview.Top = CancelButton.Top + CancelButton.Height + 30
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not LoadFinished Then
        Cancel = 1
        Exit Sub
    End If
    If Me.WindowState = 0 Then
        SaveSetting "DownloadBooster", "UserData", "ComdlgWidth", Me.Width - PaddedBorderWidth * 15 * 2
        SaveSetting "DownloadBooster", "UserData", "ComdlgHeight", Me.Height - PaddedBorderWidth * 15 * 2 - IIf(Tags.BrowseTargetForm = 3 Or Tags.BrowseTargetForm = 5 Or Tags.BrowseTargetForm = 6, 8835 - 6165, 0)
    End If
    
    IBSSubclass_UnsubclassIt
    
    On Error Resume Next
    imgFolder.ListImages.Clear
    imgFolderSmall.ListImages.Clear
End Sub

Private Function IBSSubclass_MsgResponse(ByVal hWnd As Long, ByVal uMsg As Long) As EMsgResponse
    IBSSubclass_MsgResponse = emrConsume
End Function

Private Sub IBSSubclass_UnsubclassIt()
    DetachMessage Me, Me.hWnd, WM_GETMINMAXINFO
    DetachMessage Me, Me.hWnd, WM_SETTINGCHANGE
End Sub

Private Function IBSSubclass_WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByRef wParam As Long, ByRef lParam As Long, ByRef bConsume As Boolean) As Long
    On Error Resume Next
 
    Select Case uMsg
        Case WM_GETMINMAXINFO
            Dim lpMMI As MINMAXINFO
            CopyMemory lpMMI, ByVal lParam, Len(lpMMI)
            lpMMI.ptMinTrackSize.X = (10245 + PaddedBorderWidth * 15 * 2) / 15 * (DPI / 96)
            lpMMI.ptMinTrackSize.Y = (IIf(Tags.BrowseTargetForm = 3 Or Tags.BrowseTargetForm = 5 Or Tags.BrowseTargetForm = 6, 8835, 6165) + PaddedBorderWidth * 15 * 2) / 15 * (DPI / 96)
            lpMMI.ptMaxTrackSize.X = (Screen.Width + 1200) * (DPI / 96)
            lpMMI.ptMaxTrackSize.Y = (Screen.Height + 1200) * (DPI / 96)
            CopyMemory ByVal lParam, lpMMI, Len(lpMMI)
            
            IBSSubclass_WindowProc = 1&
            Exit Function
        Case WM_SETTINGCHANGE
            Select Case GetStrFromPtr(lParam)
                Case "WindowMetrics"
                    UpdateBorderWidth
                    Form_Resize
            End Select
    End Select
    
    IBSSubclass_WindowProc = CallOldWindowProc(hWnd, uMsg, wParam, lParam)
End Function

Private Sub lvDir_Change()
    Dim i%
    Dim k%
    Dim indentLevel%
    Dim Item As ImcComboItem
    Dim Path$
    Dim ItemCount%
    Dim ItemSelectPos
    Path = lvDir.Path
    If Right$(lvDir.Path, 1) <> "\" Then Path = Path & "\"
    Dim PathSplit() As String
    PathSplit = Split(Path, "\")
    ItemCount = 0
    ItemSelectPos = 1
    
    cbFolderList.ComboItems.Clear
    cbFolderList.ComboItems.Add , , t("�� �ֱ� ����", "Recent"), 11
    cbFolderList.ComboItems.Add , , t("���� ȭ��", "Desktop"), 12
    cbFolderList.ComboItems.Add , , t("�� ����", "My Documents"), 13, , 1
    cbFolderList.ComboItems.Add , , t("�� ��ǻ��", "My Computer"), 14, , 1
    ItemCount = 4
    
    Dim Icon%
    For k = 0 To selDrive.ListCount - 1
        Select Case GetDriveType(Left$(selDrive.List(k), 2))
            Case DRIVE_FIXED, DRIVE_UNKNOWN, DRIVE_NO_ROOT_DIR
                Icon = 6
            Case DRIVE_REMOVABLE
                Icon = 3
            Case DRIVE_CDROM
                Icon = 5
            Case DRIVE_REMOTE
                Icon = 7
            Case DRIVE_RAMDISK
                Icon = 8
        End Select
        cbFolderList.ComboItems.Add , , selDrive.List(k), Icon, , 2
        ItemCount = ItemCount + 1
        If LCase(Left$(selDrive.List(k), 1)) = LCase(Left$(PathSplit(LBound(PathSplit)), 1)) Then
            indentLevel = 3
            For i = LBound(PathSplit) + 1 To UBound(PathSplit) - 1
                cbFolderList.ComboItems.Add , , PathSplit(i), 10, , indentLevel
                indentLevel = indentLevel + 1
                ItemCount = ItemCount + 1
            Next i
            ItemSelectPos = (k + i + 4)
'            For i = 0 To lvDir.ListCount - 1
'                cbFolderList.ComboItems.Add , , Replace(lvDir.List(i), Path, "", , 1), 1, , indentLevel
'                ItemCount = ItemCount + 1
'            Next i
        End If
    Next k
    
    Dim Name$
    Path = GetSpecialFolder(CSIDL_DESKTOP)
    If Right$(Path, 1) <> "\" Then Path = Path & "\"
    Name = Dir(Path, vbDirectory)
    On Error Resume Next
    Dim FullPath As String
    Do While LenB(Name)
        If Name <> "." And Name <> ".." Then
            FullPath = Path & Name
            If (GetAttr(FullPath) And vbDirectory) = vbDirectory Then
                Set Item = cbFolderList.ComboItems.Add(, , Name, 1, , 1)
                If (FullPath) = lvDir.Path Then ItemSelectPos = Item.Index
                ItemCount = ItemCount + 1
            End If
        End If
        Name = Dir
    Loop
    
    If lvDir.Path = GetSpecialFolder(CSIDL_RECENT) Then
        cbFolderList.ComboItems(1).Selected = True
    ElseIf lvDir.Path = GetSpecialFolder(CSIDL_DESKTOP) Then
        cbFolderList.ComboItems(2).Selected = True
    ElseIf lvDir.Path = GetSpecialFolder(CSIDL_PERSONAL) Then
        cbFolderList.ComboItems(3).Selected = True
    Else
        cbFolderList.ComboItems(ItemSelectPos).Selected = True
    End If
    
    'ListFiles
    SaveSetting "DownloadBooster", "UserData", "LastSaveDir", lvDir.Path
    
    For i = 1 To tbPlaces.Buttons.Count
        tbPlaces.Buttons(i).Value = TbrButtonValueUnpressed
    Next i
    
    Path = lvDir.Path
    Select Case Path
        Case GetSpecialFolder(CSIDL_RECENT)
            tbPlaces.Buttons(1).Value = TbrButtonValuePressed
        Case GetSpecialFolder(CSIDL_DESKTOP)
            tbPlaces.Buttons(2).Value = TbrButtonValuePressed
        Case GetSpecialFolder(CSIDL_PERSONAL)
            tbPlaces.Buttons(3).Value = TbrButtonValuePressed
        Case GetSpecialFolder(CSIDL_FAVORITES)
            tbPlaces.Buttons(5).Value = TbrButtonValuePressed
        Case Environ$("USERPROFILE")
            tbPlaces.Buttons(6).Value = TbrButtonValuePressed
    End Select
    
    If ItemCount < cbFolderList.ComboItems.Count Then
        For i = cbFolderList.ComboItems.Count To (ItemCount + 1) Step -1
            cbFolderList.ComboItems.Remove i
        Next i
    End If
End Sub

Private Sub lvFiles_AfterLabelEdit(Cancel As Boolean, NewString As String)
    On Error Resume Next
    If lvFiles.SelectedItem Is Nothing Then Exit Sub
    If Not lvFiles.SelectedItem.Selected Then Exit Sub
    Dim Path$
    Path = lvDir.Path
    If Right$(Path, 1) <> "\" Then Path = Path & "\"
    NewString = Trim$(NewString)
    If FileExists(Path & NewString) Then
        MsgBox t("���� �̸��� �̹� �����մϴ�.", "File name already exists."), 16
        Cancel = True
        Exit Sub
    End If
    If _
        InStr(1, NewString, "\") > 0 Or _
        InStr(1, NewString, "/") > 0 Or _
        InStr(1, NewString, """") > 0 Or _
        InStr(1, NewString, "*") > 0 Or _
        InStr(1, NewString, "?") > 0 Or _
        InStr(1, NewString, "<") > 0 Or _
        InStr(1, NewString, ">") > 0 Or _
        InStr(1, NewString, "|") > 0 Or _
        UCase(NewString) = "CON" Or _
        UCase(NewString) = "AUX" Or _
        UCase(NewString) = "PRN" Or _
        UCase(NewString) = "NUL" Or _
        UCase(NewString) = "COM1" Or _
        UCase(NewString) = "COM2" Or _
        UCase(NewString) = "COM3" Or _
        UCase(NewString) = "COM4" Or _
        UCase(NewString) = "LPT1" Or _
        UCase(NewString) = "LPT2" Or _
        UCase(NewString) = "LPT3" Or _
        UCase(NewString) = "LPT4" Or _
        LenB(Replace(NewString, ".", "")) = 0 Or _
        Right$(NewString, 1) = "." _
    Then
        MsgBox t("���� �̸��� �ùٸ��� �ʽ��ϴ�.", "Invalid file name."), 16
        Cancel = True
        Exit Sub
    End If
    If NewString = "" Then
        Cancel = True
        Exit Sub
    End If
    
    Dim FullPath$
    FullPath = Path & lvFiles.SelectedItem.Text
    On Error GoTo renamefail
    Name FullPath As (Path & NewString)
    Exit Sub
renamefail:
    Cancel = True
    MsgBox t("���� �̸��� �ٲٴ� �� �����߽��ϴ�.", "Failed to rename the selected file."), 16
End Sub

Private Sub lvFiles_BeforeLabelEdit(Cancel As Boolean)
    On Error Resume Next
    If lvFiles.SelectedItem Is Nothing Then Exit Sub
    If Not lvFiles.SelectedItem.Selected Then Exit Sub
    If lvFiles.SelectedItem.Text = ".." Then Cancel = True
    If lvFiles.SelectedItem.IconIndex > 2 And lvFiles.SelectedItem.IconIndex <= 10 Then Cancel = True
    
    Dim FullPath$
    Dim Path$
    Path = lvDir.Path
    If Right$(Path, 1) <> "\" Then Path = Path & "\"
    FullPath = Path & lvFiles.SelectedItem.Text
    If (Not FileExists(FullPath)) And (Not FolderExists(FullPath)) Then Cancel = True
End Sub

Private Sub lvFiles_ContextMenu(ByVal X As Single, ByVal Y As Single)
    On Error Resume Next
    Dim Item As LvwListItem
    Set Item = lvFiles.SelectedItem
    If Not Item Is Nothing Then
        If Item.Selected Then
            mnuRename.Enabled = ((Item.IconIndex <= 2 Or Item.IconIndex > 10) And Item.Text <> "..")
            mnuDelete.Enabled = (Not IsMyComputer) And Item.Text <> ".."
            mnuExplore.Visible = IsMyComputer Or Item.IconIndex = 1
            mnuOpen.Enabled = (IsMyComputer Or Item.IconIndex <= 2 Or Item.IconIndex > 10)
            mnuProperties.Enabled = (((Item.IconIndex <= 2 Or Item.IconIndex > 10) And Item.Text <> "..") Or IsMyComputer)
            If Tags.BrowseTargetForm = 2 Then
                mnuSelect.Enabled = (Item.IconIndex = 1 Or IsMyComputer) And LoadFinished
            Else
                mnuSelect.Enabled = LoadFinished
            End If
            If mnuSelect.Enabled Then
                Me.PopupMenu mnuFile, DefaultMenu:=mnuSelect
            Else
                Me.PopupMenu mnuFile
            End If
        Else
            GoTo folderfloor
        End If
    Else
folderfloor:
        mnuView.Visible = True
        mnuNewFolder.Enabled = tbToolBar.Buttons(3).Enabled And LoadFinished
        mnuFolderProperties.Enabled = Not IsMyComputer
        mnuCmd.Enabled = tbToolBar.Buttons(3).Enabled
        mnuRefresh.Enabled = LoadFinished
        Me.PopupMenu mnuFolderFloor
    End If
End Sub

Private Sub lvFiles_ItemDblClick(Item As LvwListItem, ByVal Button As Integer)
    If Item Is Nothing Then Exit Sub
    If Not Item.Selected Then Exit Sub
    
    Dim FullPath$
    If Right$(lvDir.Path, 1) = "\" Then
        FullPath = lvDir.Path & Item.Text
    Else
        FullPath = lvDir.Path & "\" & Item.Text
    End If
    
    If (Item.IconIndex <= 2 Or Item.IconIndex > 10) And UCase(GetExtensionName(Item.Text)) = "LNK" And (Not FolderExists(FullPath)) Then
        Dim LnkPath As String
        LnkPath = GetShortcutTarget(FullPath)
        If Left$(LnkPath, 1) = """" And Right$(LnkPath, 1) = """" Then _
            LnkPath = Mid$(LnkPath, 2, Len(LnkPath) - 2)
        If FolderExists(LnkPath) Then
            If LoadFinished Then
                lvDir.Path = LnkPath
                If Tags.BrowseTargetForm = 2 Then txtFileName.Text = ""
            End If
        ElseIf (frmMain.cbWhenExist.ListIndex <> 0 And Tags.BrowseTargetForm <> 2) Or Tags.BrowseTargetForm = 3 Or Tags.BrowseTargetForm = 4 Or Tags.BrowseTargetForm = 5 Or Tags.BrowseTargetForm = 6 Then
            OKButton_Click
        End If
    ElseIf Item.IconIndex > 2 And Item.IconIndex <= 10 Then
        If Item.Text = t("�� ��ǻ��", "My Computer") Then
            ShowMyComputer
            Exit Sub
        ElseIf Item.Text = t("�� ����", "My Documents") Then
            lvDir.Path = GetSpecialFolder(CSIDL_PERSONAL)
            Exit Sub
        End If
    
        On Error Resume Next
retrydrive:
        ListedOn = ""
        lvDir.Path = UCase(Left$(Item.Text, 2)) & "\"
        If Err Then
            If MsgBox(t("������ ����̺� �ȿ� ��ũ�� ���ų� ����̺갡 ��� �ֽ��ϴ�.", "There is no disk in the selected drive or the drive is locked."), vbRetryCancel + vbCritical) = vbRetry Then
                GoTo retrydrive
            End If
        End If
    ElseIf Item.IconIndex = 1 Then
        If LoadFinished Then
            On Error GoTo folderinaccessible
            lvDir.Path = FullPath
            If Tags.BrowseTargetForm = 2 Then txtFileName.Text = ""
            Exit Sub
folderinaccessible:
            MsgBox t("������ �������� �ʰų� ���� ������ �����ϴ�.", "The folder does not exist or there are no permission to access it."), 16
        End If
    ElseIf (frmMain.cbWhenExist.ListIndex <> 0 And Tags.BrowseTargetForm <> 2) Or Tags.BrowseTargetForm = 3 Or Tags.BrowseTargetForm = 4 Or Tags.BrowseTargetForm = 5 Or Tags.BrowseTargetForm = 6 Then
        OKButton_Click
    End If
End Sub

Private Sub lvFiles_ItemSelect(Item As LvwListItem, ByVal Selected As Boolean)
    cmdPreview.Enabled = Selected
    
    If Item.IconIndex = 1 And Tags.BrowseTargetForm = 2 Then
        'If Item.Text <> ".." Then
            txtFileName.Text = Item.Text
        'End If
    End If
    If (frmMain.cbWhenExist.ListIndex = 0 And Tags.BrowseTargetForm <> 3 And Tags.BrowseTargetForm <> 4 And Tags.BrowseTargetForm <> 5 And Tags.BrowseTargetForm <> 6) Or (Not Selected) Then Exit Sub
    If Item.IconIndex = 1 Or (Item.IconIndex > 2 And Item.IconIndex <= 10) Then Exit Sub
    If Tags.BrowseTargetForm <> 2 Then txtFileName.Text = Item.Text
    
    If (Tags.BrowseTargetForm = 3 Or Tags.BrowseTargetForm = 5 Or Tags.BrowseTargetForm = 6) And Item.IconIndex <> 1 And (Not IsMyComputer) Then
        On Error Resume Next
        Dim Path$
        Path = lvDir.Path
        If Right$(Path, 1) <> "\" Then Path = Path & "\"
        Set imgPreview.Picture = LoadPictureFromFile(Path & Item.Text)
    End If
End Sub

Private Sub lvFiles_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 116 Then
        If IsMyComputer Then
            ShowMyComputer
        ElseIf LoadFinished Then
            ListedOn = ""
            ListFiles
        End If
    ElseIf KeyCode = 113 And (Not lvFiles.SelectedItem Is Nothing) Then
        If lvFiles.SelectedItem.Selected Then lvFiles.StartLabelEdit
    ElseIf KeyCode = 8 Then
        If tbToolBar.Buttons(2).Enabled And Len(lvDir.Path) > 3 And LoadFinished Then _
            lvDir.Path = GetParentFolderName(lvDir.Path)
    ElseIf KeyCode = 46 And (Not lvFiles.SelectedItem Is Nothing) Then
        If lvFiles.SelectedItem.Selected And (lvFiles.SelectedItem.IconIndex = 2 Or lvFiles.SelectedItem.IconIndex > 10) Then mnuDelete_Click
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        If IsMyComputer Then
            ShowMyComputer
        ElseIf LoadFinished Then
            ListedOn = ""
            ListFiles
        End If
    End If
End Sub

Private Sub mnuCmd_Click()
    ShellExecute "cmd.exe", WorkingDirectory:=lvDir.Path
End Sub

Private Sub mnuDelete_Click()
    On Error Resume Next
    If lvFiles.SelectedItem Is Nothing Then Exit Sub
    If Not lvFiles.SelectedItem.Selected Then Exit Sub
    If IsMyComputer Then Exit Sub

    Dim FullPath$
    If Right$(lvDir.Path, 1) = "\" Then
        FullPath = lvDir.Path & lvFiles.SelectedItem.Text
    Else
        FullPath = lvDir.Path & "\" & lvFiles.SelectedItem.Text
    End If
    
    If ConfirmEx("'" & lvFiles.SelectedItem.Text & "' " & t("�׸��� ���������� �����Ͻðڽ��ϱ�?", " - delete item permanently?"), App.Title, 48) = vbYes Then
        On Error GoTo deletefail
        Dim IsDirectory As Boolean
        IsDirectory = (lvFiles.SelectedItem.IconIndex = 1)
        If IsDirectory Then RmDir FullPath Else Kill FullPath
        lvFiles.ListItems.Remove lvFiles.SelectedItem.Index
        Exit Sub
deletefail:
        Dim ErrorMessage As String
        If IsDirectory Then
            ErrorMessage = t("������ ��� ���� �ʰų� ���� ������ �����ϴ�.", "Directory is not empty or there is no delete permission.")
        Else
            ErrorMessage = t("�׸��� ����� �� �����߽��ϴ�.", "Failed to delete the specified item.")
        End If
        MsgBox ErrorMessage, 16
    End If
End Sub

Private Sub mnuExplore_Click()
    On Error Resume Next
    If lvFiles.SelectedItem Is Nothing Then Exit Sub
    If Not lvFiles.SelectedItem.Selected Then Exit Sub

    Dim FullPath$
    
    If IsMyComputer Then
        FullPath = Left$(lvFiles.SelectedItem.Text, 1) & ":\"
        GoTo isfolder
    End If
    
    If Right$(lvDir.Path, 1) = "\" Then
        FullPath = lvDir.Path & lvFiles.SelectedItem.Text
    Else
        FullPath = lvDir.Path & "\" & lvFiles.SelectedItem.Text
    End If
    
    If lvFiles.SelectedItem.IconIndex = 1 And UCase(GetExtensionName(lvFiles.SelectedItem.Text)) = "LNK" And (Not FolderExists(FullPath)) Then
        Dim LnkPath As String
        LnkPath = RemoveQuotes(GetShortcutTarget(FullPath))
        If FolderExists(LnkPath) Then
            FullPath = LnkPath
            GoTo isfolder
        End If
        Exit Sub
    End If
    
isfolder:
    ShellExecute FullPath, "explore"
End Sub

Private Sub mnuFolderProperties_Click()
    ShellExecute lvDir.Path, "properties"
End Sub

Private Sub mnuIconSize_Click(Index As Integer)
    lvFiles.View = Index
    SaveSetting "DownloadBooster", "UserData", "FileListView", Index
    Static i As Byte
    For i = mnuIconSize.LBound To mnuIconSize.UBound
        mnuIconSize(i).Checked = False
    Next i
    mnuIconSize(Index).Checked = True
End Sub

Private Sub mnuNewFolder_Click()
    CreateNewFolder
End Sub

Private Sub mnuOpen_Click()
    On Error Resume Next
    If lvFiles.SelectedItem Is Nothing Then Exit Sub
    If Not lvFiles.SelectedItem.Selected Then Exit Sub
    
    Dim FullPath$
    
    If IsMyComputer Then
        FullPath = Left$(lvFiles.SelectedItem.Text, 1) & ":\"
        GoTo exec
    End If

    If Right$(lvDir.Path, 1) = "\" Then
        FullPath = lvDir.Path & lvFiles.SelectedItem.Text
    Else
        FullPath = lvDir.Path & "\" & lvFiles.SelectedItem.Text
    End If
    
    If (lvFiles.SelectedItem.IconIndex <= 2 Or lvFiles.SelectedItem.IconIndex > 10) And UCase(GetExtensionName(lvFiles.SelectedItem.Text)) = "LNK" And (Not FolderExists(FullPath)) Then
        FullPath = RemoveQuotes(GetShortcutTarget(FullPath))
    End If
    
exec:
    ShellExecute FullPath
End Sub

Private Sub mnuProperties_Click()
    On Error Resume Next
    If lvFiles.SelectedItem Is Nothing Then Exit Sub
    If Not lvFiles.SelectedItem.Selected Then Exit Sub
    
    If IsMyComputer Then
        ShellExecute Left$(lvFiles.SelectedItem.Text, 1) & ":\", "properties"
        Exit Sub
    End If

    Dim FullPath$
    If Right$(lvDir.Path, 1) = "\" Then
        FullPath = lvDir.Path & lvFiles.SelectedItem.Text
    Else
        FullPath = lvDir.Path & "\" & lvFiles.SelectedItem.Text
    End If
    
    ShellExecute FullPath, "properties"
End Sub

Private Sub mnuRefresh_Click()
    If IsMyComputer Then
        ShowMyComputer
    Else
        If Not LoadFinished Then Exit Sub
        ListedOn = ""
        ListFiles
    End If
End Sub

Private Sub mnuRename_Click()
    On Error Resume Next
    If Not lvFiles.SelectedItem Is Nothing Then
        If lvFiles.SelectedItem.Selected Then
            If IsMyComputer Then Exit Sub
            If (lvFiles.SelectedItem.IconIndex <= 2 Or lvFiles.SelectedItem.IconIndex > 10) And lvFiles.SelectedItem.Text <> ".." Then lvFiles.StartLabelEdit
        End If
    End If
End Sub

Private Sub mnuSelect_Click()
    On Error Resume Next
    If Not lvFiles.SelectedItem Is Nothing Then
        If lvFiles.SelectedItem.Selected Then
            lvFiles_ItemDblClick lvFiles.SelectedItem, 1
        End If
    End If
End Sub

Private Sub OKButton_Click()
    txtFileName.Text = Trim$(txtFileName.Text)
    
    'mft ���� ����
    If (Len(lvDir.Path) = 3 And LCase(txtFileName.Text) = "$mft") Or (Len(txtFileName.Text) > 3 And Mid$(txtFileName.Text, 2, 2) = ":\" And (LCase(Right(txtFileName.Text, Len(txtFileName.Text) - 3)) = "$mft" Or StartsWith(LCase(Right(txtFileName.Text, Len(txtFileName.Text) - 3)), "$mft\"))) Then
        If MsgBox(t("��� ��ũ���� �߻���ų �����̽ʴϱ�? ^^;", "You'd better not do that!"), t(vbQuestion + vbYesNo, vbExclamation)) = vbYes Then
            MsgBox "�׷��� �� �˴ϴ�!", 16
        End If
        Exit Sub
    End If
    
    On Error Resume Next
    If InStr(1, txtFileName.Text, "*") > 0 Or InStr(1, txtFileName.Text, "?") > 0 Then
        If Not LoadFinished Then Exit Sub
        Pattern = txtFileName.Text
        txtFileName.SelStart = 0
        txtFileName.SelLength = Len(txtFileName.Text)
        ListedOn = ""
        ListFiles
        Exit Sub
    End If

    Dim Data$, Path$
    
    If Not lvFiles.SelectedItem Is Nothing Then
        If lvFiles.SelectedItem.Selected Then
            Dim FullPath$
            If Right$(lvDir.Path, 1) = "\" Then
                FullPath = lvDir.Path & lvFiles.SelectedItem.Text
            Else
                FullPath = lvDir.Path & "\" & lvFiles.SelectedItem.Text
            End If
            
            Dim FullPath2$
            If Right$(lvDir.Path, 1) = "\" Then
                FullPath2 = lvDir.Path & txtFileName.Text
            Else
                FullPath2 = lvDir.Path & "\" & txtFileName.Text
            End If
        
            If lvFiles.SelectedItem.IconIndex = 1 And UCase(GetExtensionName(lvFiles.SelectedItem.Text)) = "LNK" And (Not FolderExists(FullPath)) Then
                Dim LnkPath As String
                LnkPath = RemoveQuotes(GetShortcutTarget(FullPath))
                If FolderExists(LnkPath) Then FullPath = LnkPath
            End If
            
            If lvFiles.SelectedItem.IconIndex = 1 And FolderExists(FullPath) And (txtFileName.Text = "" Or ((Not FolderExists(txtFileName)) And (Not FolderExists(FullPath2)))) Then
                If LoadFinished Then lvDir.Path = FullPath
                Exit Sub
            End If
        End If
    End If
    
    If Tags.BrowseTargetForm = 3 Or Tags.BrowseTargetForm = 4 Or Tags.BrowseTargetForm = 5 Or Tags.BrowseTargetForm = 6 Then
        If FolderExists(txtFileName.Text) Then
            If LoadFinished Then
                txtFileName.SelStart = 0
                txtFileName.SelLength = Len(txtFileName.Text)
                lvDir.Path = txtFileName.Text
            End If
            Exit Sub
        ElseIf FolderExists(lvDir.Path & IIf(EndsWith(lvDir.Path, "\"), "", "\") & txtFileName.Text) Then
            If LoadFinished Then
                txtFileName.SelStart = 0
                txtFileName.SelLength = Len(txtFileName.Text)
                lvDir.Path = lvDir.Path & IIf(EndsWith(lvDir.Path, "\"), "", "\") & txtFileName.Text
            End If
            Exit Sub
        End If
        
        If FolderExists(GetParentFolderName(txtFileName.Text)) Then
            If Not LoadFinished Then Exit Sub
            lvDir.Path = GetParentFolderName(txtFileName.Text)
            txtFileName.Text = GetFilename(txtFileName.Text)
        End If
    ElseIf FolderExists(txtFileName.Text) Then
        If txtFileName.Text = "." Or txtFileName.Text = ".." Then
            If LoadFinished Then
                lvDir.Path = txtFileName.Text
                txtFileName.Text = ""
'                txtFileName.SelStart = 0
'                txtFileName.SelLength = Len(txtFileName.Text)
            End If
            Exit Sub
        End If
        If Not LoadFinished Then Exit Sub
        lvDir.Path = txtFileName.Text
        txtFileName.Text = ""
    ElseIf FolderExists(GetParentFolderName(txtFileName.Text)) Then
        If Not LoadFinished Then Exit Sub
        lvDir.Path = GetParentFolderName(txtFileName.Text)
        txtFileName.Text = GetFilename(txtFileName.Text)
        If txtFileName.Text = "." Or txtFileName.Text = ".." Then
            Exit Sub
        End If
    ElseIf Tags.BrowseTargetForm = 2 Then
        Path = lvDir.Path
        If Right$(lvDir.Path, 1) <> "\" Then Path = Path & "\"
        If FolderExists(Path & txtFileName.Text) Then
            If Not LoadFinished Then Exit Sub
            lvDir.Path = Path & txtFileName.Text
            If LenB(txtFileName.Text) Then
                txtFileName.Text = ""
                Exit Sub
            End If
        End If
    ElseIf InStr(1, txtFileName.Text, "\") Then
        MsgBox t("�Է��� ������ ��ΰ� �������� �ʽ��ϴ�.", "The specified folder path does not exist."), 48
        Exit Sub
    End If
    On Error GoTo 0
    
    If Tags.BrowseTargetForm >= 3 And Tags.BrowseTargetForm <= 6 Then
        Path = lvDir.Path
        If Right$(lvDir.Path, 1) <> "\" Then Path = Path & "\"
        If Not FileExists(Path & txtFileName.Text) Then
            MsgBox txtFileName.Text & vbCrLf & t("������ �����ϴ�.", "File does not exist.") & vbCrLf & t("���� �̸��� �ùٸ��� �Է��ߴ��� Ȯ���Ͻʽÿ�.", "Check if you specified a valid file name."), 48
            Exit Sub
        End If
    End If
    
    Select Case Tags.BrowseTargetForm: Case 3, 5, 6
        Dim PicturePath As String
        PicturePath = Path & txtFileName.Text
        If LoadPictureFromFile(PicturePath) Is Nothing Then GoTo imgerr
        Select Case Tags.BrowseTargetForm
            Case 3
                frmOptions.ChangedBackgroundPath = PicturePath
                frmOptions.LoadBackgroundList
            Case 5
                SaveSetting "DownloadBooster", "Options", "LiveBadukMemoSkinFrameTexture", PicturePath
                frmLiveBadukSkinProperties.optTexture.Value = True
            Case 6
                SaveSetting "DownloadBooster", "Options", "LiveBadukMemoSkinFrameBackground", PicturePath
                frmLiveBadukSkinProperties.optFrameTexture.Value = True
        End Select
        Unload Me
        Exit Sub
imgerr:
        MsgBox t("�׸��� �ջ�Ǿ��ų� �ùٸ��� �ʽ��ϴ�.", "The selected picture is corrupt or invalid."), 16
        Exit Sub
    End Select
    
    If Tags.BrowseTargetForm = 4 And LenB(txtFileName.Text) = 0 Then Exit Sub
    
    Dim IsColonPresent As Boolean
    If Len(txtFileName.Text) > 3 And Mid$(txtFileName.Text, 2, 2) = ":\" Then
        IsColonPresent = InStr(1, Right$(txtFileName.Text, Len(txtFileName.Text) - 3), "|")
    Else
        IsColonPresent = InStr(txtFileName.Text, ":")
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
        (LenB(txtFileName.Text) And LenB(Replace(txtFileName.Text, ".", "")) = 0) Or _
        Right$(txtFileName.Text, 1) = "." _
    Then
        If Tags.BrowseTargetForm = 2 Then
            MsgBox t("���� ��ΰ� �������� �ʽ��ϴ�.", "Invalid folder path."), 48
        Else
            MsgBox t("���� �̸��� �ùٸ��� �ʽ��ϴ�.", "Invalid file name."), 48
        End If
        Exit Sub
    End If

    If Tags.BrowseTargetForm = 2 Then
        Path = lvDir.Path
    Else
        If Right$(lvDir.Path, 1) = "\" Then
            Path = lvDir.Path & txtFileName.Text
        Else
            Path = lvDir.Path & "\" & txtFileName.Text
        End If
    End If
    On Error Resume Next
    If FileExists(Path) And Tags.BrowseTargetForm <> 4 Then
        If frmMain.cbWhenExist.ListIndex = 0 Then
            MsgBox t("���� �̸��� �̹� �����մϴ�. �ٸ� �̸��� �����Ͻʽÿ�.", "File name already exists."), 16
            Exit Sub
        ElseIf frmMain.cbWhenExist.ListIndex = 1 Then
            If MsgBox(Path & t("��(��) �̹� �ֽ��ϴ�." & vbCrLf & "����ðڽ��ϱ�?", " already exists." & vbCrLf & "Overwrite?"), 48 + vbYesNo) <> vbYes Then
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
        Case 4
            Tags.BrowseTargetTextbox.Text = Path
        Case Else
            frmMain.txtFileName.Text = Path
    End Select
    
    Unload Me
    Exit Sub
    
e:
    MsgBox t("������ �߻��߽��ϴ�!", "Error!"), 16
    Exit Sub
End Sub

Private Sub selFileType_Change()
    selFileType_Click
End Sub

Private Sub selFileType_Click()
    Pattern = Replace(Mid$(selFileType.Text, InStr(1, selFileType.Text, "(") + 1, Len(selFileType.Text) - InStr(1, selFileType.Text, "(") - 1), " ", "")
    If Not LoadFinished Then Exit Sub
    ListedOn = ""
    If Loaded Then ListFiles
End Sub

Private Sub tbPlaces_ButtonClick(ByVal Button As TbrButton)
    If Not LoadFinished Then Exit Sub
    
    Dim i%
    For i = 1 To tbPlaces.Buttons.Count
        tbPlaces.Buttons(i).Value = TbrButtonValueUnpressed
    Next i
    
    If IsMyComputer Then ListedOn = ""
    
    Select Case Button.Index
        Case 1
            lvDir.Path = GetSpecialFolder(CSIDL_RECENT)
            cbFolderList.ComboItems(1).Selected = True
        Case 2
            lvDir.Path = GetSpecialFolder(CSIDL_DESKTOP)
            cbFolderList.ComboItems(2).Selected = True
        Case 3
            lvDir.Path = GetSpecialFolder(CSIDL_PERSONAL)
        Case 4
            ShowMyComputer
            Exit Sub
        Case 5
            lvDir.Path = GetSpecialFolder(CSIDL_FAVORITES)
        Case 6
            lvDir.Path = Environ$("USERPROFILE")
    End Select
    
    If IsMyComputer Then
        lvDir_Change
    End If
    
    'ListFiles
    
    Dim Path$
    Path = lvDir.Path
    Select Case Path
        Case GetSpecialFolder(CSIDL_RECENT)
            tbPlaces.Buttons(1).Value = TbrButtonValuePressed
        Case GetSpecialFolder(CSIDL_DESKTOP)
            tbPlaces.Buttons(2).Value = TbrButtonValuePressed
        Case GetSpecialFolder(CSIDL_PERSONAL)
            tbPlaces.Buttons(3).Value = TbrButtonValuePressed
        Case GetSpecialFolder(CSIDL_FAVORITES)
            tbPlaces.Buttons(5).Value = TbrButtonValuePressed
        Case Environ$("USERPROFILE")
            tbPlaces.Buttons(6).Value = TbrButtonValuePressed
    End Select
End Sub

Private Sub CreateNewFolder()
    If Not tbToolBar.Buttons(3).Enabled Then Exit Sub

    Dim DirName$
    Dim FullPath$
    Do
        DirName = CStr(Fix(Rnd * 100000000))
        If Right$(lvDir.Path, 1) = "\" Then
            FullPath = lvDir.Path & DirName
        Else
            FullPath = lvDir.Path & "\" & DirName
        End If
    Loop While FileExists(FullPath) Or FolderExists(FullPath)
    On Error Resume Next
    MkDir FullPath
    If Not FolderExists(FullPath) Then
        MsgBox t("������ ����� �� �����߽��ϴ�.", "Failed to create a folder here."), 16
        Exit Sub
    End If
    Dim Item As LvwListItem
    Set Item = lvFiles.ListItems.Add(, , DirName, 1, 1)
    Item.ListSubItems.Add , , "-"
    Item.ListSubItems.Add , , t("���� ����", "File Folder")
    Item.ListSubItems.Add , , FileDateTime(FullPath)
    Item.EnsureVisible
    Item.Selected = True
    lvFiles.StartLabelEdit
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As TbrButton)
    Select Case Button.Index
        Case 2
            If Not LoadFinished Then Exit Sub
            If IsMyComputer Then
                lvDir.Path = GetSpecialFolder(CSIDL_DESKTOP)
            ElseIf Len(lvDir.Path) > 3 Then
                lvDir.Path = GetParentFolderName(lvDir.Path)
            End If
        Case 3
            CreateNewFolder
    End Select
End Sub

Private Sub tbToolBar_ButtonDropDown(ByVal Button As TbrButton)
    If Button.Index = 4 Then
        Me.PopupMenu mnuView, , cmdViews.Left, cmdViews.Top + cmdViews.Height
    End If
End Sub

Private Sub txtFileName_Change()
    If Tags.BrowseTargetForm = 0 Then
        If txtFileName.Text = "" Then
            OKButton.Caption = t("���� ����(&E)", "S&elect Folder")
        Else
            OKButton.Caption = t("����(&E)", "Sav&e")
        End If
    End If
End Sub
