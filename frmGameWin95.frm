VERSION 5.00
Begin VB.Form frmGameWin95 
   BackColor       =   &H00F8EFE5&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "숫자야구"
   ClientHeight    =   2565
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4815
   BeginProperty Font 
      Name            =   "굴림"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGameWin95.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin prjDownloadBooster.CommandButtonW cmdGiveUp 
      Caption         =   "기권(&G)"
      Height          =   375
      Left            =   3480
      TabIndex        =   14
      Top             =   1080
      Width           =   1215
   End
   Begin prjDownloadBooster.CommandButtonW cmdReset 
      Caption         =   "초기화(&R)"
      Height          =   375
      Left            =   3480
      TabIndex        =   13
      Top             =   600
      Width           =   1215
   End
   Begin prjDownloadBooster.CommandButtonW cmdQuit 
      Caption         =   "그만(&C)"
      Height          =   375
      Left            =   3480
      TabIndex        =   12
      Top             =   120
      Width           =   1215
   End
   Begin prjDownloadBooster.CommandButtonW cmdGo 
      Caption         =   "결과 확인"
      Height          =   320
      Left            =   2160
      TabIndex        =   11
      Top             =   480
      Width           =   1095
   End
   Begin prjDownloadBooster.TextBoxW txtZ 
      Alignment       =   2  '가운데 맞춤
      Height          =   270
      Left            =   1440
      TabIndex        =   9
      Text            =   "0"
      Top             =   480
      Width           =   375
   End
   Begin prjDownloadBooster.TextBoxW txtY 
      Alignment       =   2  '가운데 맞춤
      Height          =   270
      Left            =   840
      TabIndex        =   8
      Text            =   "0"
      Top             =   480
      Width           =   375
   End
   Begin prjDownloadBooster.TextBoxW txtX 
      Alignment       =   2  '가운데 맞춤
      Height          =   270
      Left            =   240
      TabIndex        =   7
      Text            =   "0"
      Top             =   480
      Width           =   375
   End
   Begin VB.Frame fStatus 
      BackColor       =   &H00F8EFE5&
      Caption         =   "상태"
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   3015
      Begin VB.Label lblRemaining 
         Alignment       =   1  '오른쪽 맞춤
         BackStyle       =   0  '투명
         Caption         =   "10회 남음"
         Height          =   255
         Left            =   1920
         TabIndex        =   6
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblAnswer 
         BackStyle       =   0  '투명
         Caption         =   "정답:  -   -   -"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label lblStrike 
         BackStyle       =   0  '투명
         Caption         =   "0"
         Height          =   255
         Left            =   2640
         TabIndex        =   4
         Top             =   240
         Width           =   135
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '투명
         Caption         =   "스트라이크:"
         Height          =   255
         Left            =   1560
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblBall 
         BackStyle       =   0  '투명
         Caption         =   "0"
         Height          =   255
         Left            =   480
         TabIndex        =   2
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "볼:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '투명
      Caption         =   "숫자 추측:"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "frmGameWin95"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x%, y%, z%
Dim Remaining%

Private Sub cmdGiveUp_Click()
    cmdGo.Enabled = 0
    lblAnswer.Caption = "정답:  " & x & "   " & y & "   " & z
End Sub

Private Sub cmdGo_Click()
    If Not IsNumeric(txtX.Text) Or Not IsNumeric(txtY.Text) Or Not IsNumeric(txtZ.Text) Then Exit Sub
    lblStrike.Caption = 0
    If txtX.Text = x Then lblStrike.Caption = CInt(lblStrike.Caption) + 1
    If txtY.Text = y Then lblStrike.Caption = CInt(lblStrike.Caption) + 1
    If txtZ.Text = z Then lblStrike.Caption = CInt(lblStrike.Caption) + 1
    
    If lblStrike.Caption = 3 Then
        cmdGiveUp_Click
        Exit Sub
    End If
    
    lblBall.Caption = 0
    If txtX.Text <> x And (txtY.Text = x Or txtZ.Text = x) Then lblBall.Caption = CInt(lblBall.Caption) + 1
    If txtY.Text <> y And (txtX.Text = y Or txtZ.Text = y) Then lblBall.Caption = CInt(lblBall.Caption) + 1
    If txtZ.Text <> z And (txtX.Text = z Or txtY.Text = z) Then lblBall.Caption = CInt(lblBall.Caption) + 1
    
    Remaining = Remaining - 1
    lblRemaining.Caption = Remaining & "회 남음"
    If Remaining < 1 Then cmdGiveUp_Click
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub cmdReset_Click()
    Form_Load
End Sub

Private Sub Form_Load()
    Randomize
    x = Int((Rnd * 8) + 1)
    y = Int((Rnd * 8) + 1)
    z = Int((Rnd * 8) + 1)
    
    txtX.Text = 0
    txtY.Text = 0
    txtZ.Text = 0
    lblBall.Caption = 0
    lblStrike.Caption = 0
    
    lblAnswer.Caption = "정답:  -   -   -"
    
    Remaining = 10
    lblRemaining.Caption = Remaining & "회 남음"
    cmdGo.Enabled = -1
End Sub
