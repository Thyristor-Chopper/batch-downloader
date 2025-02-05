VERSION 5.00
Begin VB.Form frmEditHeader 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "헤더 편집"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8640
   BeginProperty Font 
      Name            =   "굴림"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEditHeader.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   8640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin prjDownloadBooster.LinkLabel lblDescription 
      Height          =   495
      Left            =   720
      TabIndex        =   8
      Top             =   180
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   873
      Caption         =   "frmEditHeader.frx":000C
      Transparent     =   -1  'True
   End
   Begin prjDownloadBooster.CommandButtonW CancelButton 
      Cancel          =   -1  'True
      Height          =   330
      Left            =   7320
      TabIndex        =   7
      Top             =   5280
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      Caption         =   "취소"
   End
   Begin prjDownloadBooster.CommandButtonW OKButton 
      Default         =   -1  'True
      Height          =   330
      Left            =   6000
      TabIndex        =   6
      Top             =   5280
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      Caption         =   "확인"
   End
   Begin prjDownloadBooster.CommandButtonW cmdEditHeaderName 
      Height          =   330
      Left            =   3360
      TabIndex        =   0
      Top             =   5280
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      Enabled         =   0   'False
      Caption         =   "이름변경(&R)"
   End
   Begin prjDownloadBooster.TextBoxW txtEdit 
      Height          =   255
      Left            =   2760
      TabIndex        =   1
      Top             =   960
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   450
   End
   Begin prjDownloadBooster.CommandButtonW cmdDeleteHeader 
      Height          =   330
      Left            =   2040
      TabIndex        =   2
      Top             =   5280
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      Enabled         =   0   'False
      Caption         =   "삭제(&D)"
   End
   Begin prjDownloadBooster.CommandButtonW cmdEditHeaderValue 
      Height          =   330
      Left            =   4680
      TabIndex        =   3
      Top             =   5280
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      Enabled         =   0   'False
      Caption         =   "편집(&E)"
   End
   Begin prjDownloadBooster.CommandButtonW cmdAddHeader 
      Height          =   330
      Left            =   720
      TabIndex        =   4
      Top             =   5280
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      Caption         =   "추가(&A)"
   End
   Begin prjDownloadBooster.ListView lvHeaders 
      Height          =   4455
      Left            =   720
      TabIndex        =   5
      Top             =   720
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   7858
      VisualTheme     =   1
      View            =   3
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HideSelection   =   0   'False
      ShowLabelTips   =   -1  'True
      HighlightColumnHeaders=   -1  'True
      AutoSelectFirstItem=   0   'False
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmEditHeader.frx":00E8
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmEditHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SelectedListItem As LvwListItem
Dim MouseY As Integer
Dim OKClicked As Boolean

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    If GetSetting("DownloadBooster", "Options", "DisableDWMWindow", DefaultDisableDWMWindow) = 1 Then DisableDWMWindow Me.hWnd
    SetFormBackgroundColor Me
    SetFont Me
    SetWindowPos Me.hWnd, IIf(MainFormOnTop, HWND_TOPMOST, HWND_NOTOPMOST), 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE

'    On Error Resume Next
'    Me.Icon = frmMain.Icon
'    On Error GoTo 0
    
    Me.Caption = t(Me.Caption, "Edit Headers")
    cmdAddHeader.Caption = t(cmdAddHeader.Caption, "&Add")
    cmdDeleteHeader.Caption = t(cmdDeleteHeader.Caption, "&Delete")
    cmdEditHeaderName.Caption = t(cmdEditHeaderName.Caption, "&Rename")
    cmdEditHeaderValue.Caption = t(cmdEditHeaderValue.Caption, "&Edit")
    lblDescription.Caption = t(lblDescription.Caption, "Headers here are only applied in this session. Go to <A>Options</A> to change them permanently.")
    OKButton.Caption = t(OKButton.Caption, "OK")
    CancelButton.Caption = t(CancelButton.Caption, "Cancel")
    
    lvHeaders.ColumnHeaders.Add , , t("이름", "Name"), 2655
    lvHeaders.ColumnHeaders.Add , , t("값", "Value"), 4815
    
    Dim Header
    For Each Header In SessionHeaders.Keys
        lvHeaders.ListItems.Add(, , Header).ListSubItems.Add , , SessionHeaders(CStr(Header))
    Next Header
    
    OKClicked = False
    
    On Error Resume Next
    Me.Icon = frmMain.imgWrench.ListImages(1).Picture
    On Error GoTo 0
End Sub

Private Sub cmdAddHeader_Click()
    lvHeaders.SetFocus
    Set lvHeaders.SelectedItem = lvHeaders.ListItems.Add(, , "")
    lvHeaders.SelectedItem.ListSubItems.Add , , ""
    lvHeaders.StartLabelEdit
End Sub

Private Sub cmdDeleteHeader_Click()
    If Not lvHeaders.SelectedItem Is Nothing Then
        If lvHeaders.SelectedItem.Selected Then
            lvHeaders.ListItems.Remove lvHeaders.SelectedItem.Index
        End If
    End If
End Sub

Private Sub cmdEditHeaderName_Click()
    On Error Resume Next
    lvHeaders.SetFocus
    lvHeaders.StartLabelEdit
End Sub

Private Sub cmdEditHeaderValue_Click()
    On Error GoTo exitsub
    If Not lvHeaders.SelectedItem Is Nothing Then
        Set SelectedListItem = lvHeaders.SelectedItem
        With txtEdit
            .Top = (lvHeaders.Top + MouseY) - Fix((txtEdit.Height) / 2)
            .Left = lvHeaders.Left + lvHeaders.ColumnHeaders(1).Width + 30
            .Width = lvHeaders.ColumnHeaders(2).Width
            .Text = SelectedListItem.ListSubItems(1).Text
            .Visible = True
            .SetFocus
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
        OKButton.Enabled = 0
    End If
exitsub:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not OKClicked Then
        If Confirm(t("변경된 내용을 취소하고 닫으시겠습니까?", "Do you want to discard changes and close?"), App.Title, Me) <> vbYes Then Cancel = 1
    End If
End Sub

Private Sub lblDescription_LinkActivate(ByVal Link As LlbLink, ByVal Reason As LlbLinkActivateReasonConstants)
    Load frmOptions
    frmOptions.tsTabStrip.Tabs(2).Selected = -1
    frmOptions.Show vbModal, Me
End Sub

Private Sub OKButton_Click()
    SessionHeaders.RemoveAll
    
    If lvHeaders.ListItems.Count > 0 Then
        Dim RawHeaders$
        RawHeaders = ""
        Dim i%
        For i = 1 To lvHeaders.ListItems.Count
            If Trim$(lvHeaders.ListItems(i).Text) <> "" Then
                SessionHeaders.Add CStr(Trim$(lvHeaders.ListItems(i).Text)), CStr(lvHeaders.ListItems(i).ListSubItems(1).Text)
                RawHeaders = RawHeaders & LCase(Trim$(lvHeaders.ListItems(i).Text)) & ": " & lvHeaders.ListItems(i).ListSubItems(1).Text & vbLf
            End If
        Next i
        If Right$(RawHeaders, 1) = vbLf Then RawHeaders = Left$(RawHeaders, Len(RawHeaders) - 1)
        SessionHeaderCache = btoa(RawHeaders)
    Else
        SessionHeaderCache = ""
    End If

    OKClicked = True
    Unload Me
End Sub

Private Sub txtEdit_LostFocus()
    On Error Resume Next
    SelectedListItem.ListSubItems(1).Text = txtEdit.Text
    txtEdit.Visible = False
    Set SelectedListItem = Nothing
    OKButton.Enabled = -1
End Sub
 
Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Or KeyAscii = 10 Then
        SelectedListItem.ListSubItems(1).Text = txtEdit.Text
        txtEdit.Visible = False
        Set SelectedListItem = Nothing
        OKButton.Enabled = -1
        lvHeaders.SetFocus
    End If
End Sub

Private Sub lvHeaders_AfterLabelEdit(Cancel As Boolean, NewString As String)
    NewString = Trim$(NewString)
    If NewString = "" Then
invalidname:
        Cancel = True
        Alert t("헤더 이름이 잘못되었습니다.", "Invalid header name."), App.Title, Me, 16
        Exit Sub
    End If
    
    Dim i%
    For i = 1 To Len(NewString)
        Select Case Mid$(NewString, i, 1)
            Case "a" To "z", "A" To "Z", "0" To "9", "-", "_"
            Case Else
                GoTo invalidname
        End Select
    Next i
    
    For i = 1 To lvHeaders.ListItems.Count
        If LCase(lvHeaders.ListItems(i).Text) = LCase(NewString) Then
            Cancel = True
            Alert t("해당 이름이 이미 존재합니다.", "Duplicate header name."), App.Title, Me, 16
            Exit Sub
            Exit For
        End If
    Next i
End Sub

Private Sub lvHeaders_ItemDblClick(ByVal Item As LvwListItem, ByVal Button As Integer)
    If Item.Selected Then _
        cmdEditHeaderValue_Click
End Sub

Private Sub lvHeaders_ItemSelect(ByVal Item As LvwListItem, ByVal Selected As Boolean)
    On Error GoTo justdisable
    If Selected Then
        cmdDeleteHeader.Enabled = -1
        cmdEditHeaderName.Enabled = -1
        cmdEditHeaderValue.Enabled = -1
        Exit Sub
    End If
justdisable:
    cmdDeleteHeader.Enabled = 0
    cmdEditHeaderName.Enabled = 0
    cmdEditHeaderValue.Enabled = 0
End Sub

Private Sub lvHeaders_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseY = Y
End Sub

