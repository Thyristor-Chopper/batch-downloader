VERSION 5.00
Begin VB.Form frmBatchAdd 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "일괄 다운로드"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5985
   BeginProperty Font 
      Name            =   "굴림"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBatchAdd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin prjDownloadBooster.TextBoxW txtURLs 
      Height          =   2535
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   4215
      _ExtentX        =   0
      _ExtentY        =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3
   End
   Begin prjDownloadBooster.CommandButtonW cmdCancel 
      Cancel          =   -1  'True
      Height          =   340
      Left            =   4560
      TabIndex        =   3
      Top             =   510
      Width           =   1335
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "취소"
   End
   Begin prjDownloadBooster.CommandButtonW cmdOK 
      Height          =   340
      Left            =   4560
      TabIndex        =   2
      Top             =   120
      Width           =   1335
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "확인"
   End
   Begin VB.Label Label1 
      Caption         =   "각 줄에 파일 주소를 입력하십시오(&L)."
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmBatchAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PrevKeyCode As Integer

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim URLs() As String
    URLs = Split(txtURLs.Text, vbCrLf)
    For i = 0 To UBound(URLs)
        If Replace(URLs(i), " ", "") <> "" Then
            frmMain.AddBatchURLs URLs(i)
        End If
    Next i
    Unload Me
End Sub

Private Sub Form_Load()
    If GetSetting("DownloadBooster", "Options", "DisableDWMWindow", DefaultDisableDWMWindow) = 1 Then DisableDWMWindow Me.hWnd
    SetFormBackgroundColor Me
    SetFont Me
    
    Me.Caption = t(Me.Caption, "Batch download")
    cmdOK.Caption = t(cmdOK.Caption, "OK")
    cmdCancel.Caption = t(cmdCancel.Caption, "Cancel")
    Label1.Caption = t(Label1.Caption, "Enter one UR&L per line:")
End Sub

Private Sub txtURLs_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = 13 Or KeyCode = 10) Then
        If (PrevKeyCode = 13 Or PrevKeyCode = 10) Then
            cmdOK_Click
        End If
        PrevKeyCode = KeyCode
    Else
        PrevKeyCode = 0
    End If
End Sub

Private Sub txtURLs_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    PrevKeyCode = 0
End Sub
