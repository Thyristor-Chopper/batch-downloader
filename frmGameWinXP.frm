VERSION 5.00
Begin VB.Form frmGameWinXP 
   BackColor       =   &H00F8EFE5&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "카지노"
   ClientHeight    =   3405
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5640
   BeginProperty Font 
      Name            =   "굴림"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGameWinXP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin prjDownloadBooster.TextBoxW txtDeal 
      Height          =   270
      Left            =   4200
      TabIndex        =   19
      Top             =   480
      Width           =   975
      _ExtentX        =   0
      _ExtentY        =   0
      Text            =   "frmGameWinXP.frx":000C
   End
   Begin VB.Timer timTimer 
      Enabled         =   0   'False
      Interval        =   60
      Left            =   1560
      Top             =   2880
   End
   Begin prjDownloadBooster.CommandButtonW cmdGo 
      Height          =   285
      Left            =   2280
      TabIndex        =   9
      Top             =   3000
      Width           =   1455
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "시작! (&S)"
   End
   Begin prjDownloadBooster.CommandButtonW cmdQuit 
      Height          =   285
      Left            =   3960
      TabIndex        =   8
      Top             =   3000
      Width           =   1455
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "그만(&Q)"
   End
   Begin prjDownloadBooster.TextBoxW txtMyNumber 
      Height          =   270
      Index           =   6
      Left            =   2640
      TabIndex        =   5
      Top             =   600
      Width           =   255
      _ExtentX        =   0
      _ExtentY        =   0
      Text            =   "frmGameWinXP.frx":0032
      Alignment       =   2
   End
   Begin prjDownloadBooster.TextBoxW txtMyNumber 
      Height          =   270
      Index           =   5
      Left            =   2160
      TabIndex        =   4
      Top             =   600
      Width           =   255
      _ExtentX        =   0
      _ExtentY        =   0
      Text            =   "frmGameWinXP.frx":0054
      Alignment       =   2
   End
   Begin prjDownloadBooster.TextBoxW txtMyNumber 
      Height          =   270
      Index           =   4
      Left            =   1680
      TabIndex        =   3
      Top             =   600
      Width           =   255
      _ExtentX        =   0
      _ExtentY        =   0
      Text            =   "frmGameWinXP.frx":0076
      Alignment       =   2
   End
   Begin prjDownloadBooster.TextBoxW txtMyNumber 
      Height          =   270
      Index           =   3
      Left            =   1200
      TabIndex        =   2
      Top             =   600
      Width           =   255
      _ExtentX        =   0
      _ExtentY        =   0
      Text            =   "frmGameWinXP.frx":0098
      Alignment       =   2
   End
   Begin prjDownloadBooster.TextBoxW txtMyNumber 
      Height          =   270
      Index           =   2
      Left            =   720
      TabIndex        =   1
      Top             =   600
      Width           =   255
      _ExtentX        =   0
      _ExtentY        =   0
      Text            =   "frmGameWinXP.frx":00BA
      Alignment       =   2
   End
   Begin prjDownloadBooster.TextBoxW txtMyNumber 
      Height          =   270
      Index           =   1
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   255
      _ExtentX        =   0
      _ExtentY        =   0
      Text            =   "frmGameWinXP.frx":00DC
      Alignment       =   2
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '투명
      Caption         =   "점수 걸기:"
      Height          =   255
      Left            =   4200
      TabIndex        =   18
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '투명
      Caption         =   "(0~9)"
      Height          =   255
      Left            =   3240
      TabIndex        =   17
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      Caption         =   "사용할 복권의 숫자를 입력하십시오. "
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   240
      Width           =   4455
   End
   Begin VB.Label lblCasino 
      BackStyle       =   0  '투명
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   4800
      TabIndex        =   15
      Top             =   1680
      Width           =   135
   End
   Begin VB.Line Line7 
      BorderStyle     =   5  '대시-점-점
      X1              =   4440
      X2              =   5280
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00FF80FF&
      BackStyle       =   1  '투명하지 않음
      Height          =   975
      Left            =   4560
      Shape           =   4  '둥근 사각형
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label lblCasino 
      BackStyle       =   0  '투명
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   3960
      TabIndex        =   14
      Top             =   1680
      Width           =   135
   End
   Begin VB.Line Line6 
      BorderStyle     =   5  '대시-점-점
      X1              =   3600
      X2              =   4440
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FFFF80&
      BackStyle       =   1  '투명하지 않음
      Height          =   975
      Left            =   3720
      Shape           =   4  '둥근 사각형
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label lblCasino 
      BackStyle       =   0  '투명
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   3120
      TabIndex        =   13
      Top             =   1680
      Width           =   135
   End
   Begin VB.Line Line5 
      BorderStyle     =   5  '대시-점-점
      X1              =   2760
      X2              =   3600
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H0080FF80&
      BackStyle       =   1  '투명하지 않음
      Height          =   975
      Left            =   2880
      Shape           =   4  '둥근 사각형
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label lblCasino 
      BackStyle       =   0  '투명
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   2280
      TabIndex        =   12
      Top             =   1680
      Width           =   135
   End
   Begin VB.Line Line4 
      BorderStyle     =   5  '대시-점-점
      X1              =   1920
      X2              =   2760
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  '투명하지 않음
      Height          =   975
      Left            =   2040
      Shape           =   4  '둥근 사각형
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label lblCasino 
      BackStyle       =   0  '투명
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1440
      TabIndex        =   11
      Top             =   1680
      Width           =   135
   End
   Begin VB.Line Line3 
      BorderStyle     =   5  '대시-점-점
      X1              =   1080
      X2              =   1920
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  '투명하지 않음
      Height          =   975
      Left            =   1200
      Shape           =   4  '둥근 사각형
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label lblCasino 
      BackStyle       =   0  '투명
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   600
      TabIndex        =   10
      Top             =   1680
      Width           =   135
   End
   Begin VB.Line Line2 
      BorderStyle     =   5  '대시-점-점
      X1              =   240
      X2              =   1080
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H008080FF&
      BackStyle       =   1  '투명하지 않음
      Height          =   975
      Left            =   360
      Shape           =   4  '둥근 사각형
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label lblScore 
      BackStyle       =   0  '투명
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   840
      TabIndex        =   7
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '투명
      Caption         =   "점수:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   3000
      Width           =   495
   End
End
Attribute VB_Name = "frmGameWinXP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Timer%

Private Sub cmdGo_Click()
    Dim i%
    For i = 1 To 6
        If Not IsNumeric(txtMyNumber(i).Text) Then Exit Sub
        If txtMyNumber(i).Text < 0 Or txtMyNumber(i).Text > 9 Then Exit Sub
    Next i
    
    If Not IsNumeric(txtDeal.Text) Then
        MsgBox "점수가 숫자여야 합니다.", 16, "카지노"
        Exit Sub
    ElseIf CInt(txtDeal.Text) < 1 Then
        MsgBox "좀 더 적극적으로 해 보시오.", 48, "카지노"
        Exit Sub
    ElseIf CInt(txtDeal.Text) > 100 Then
        MsgBox "적당이 하는 게 좋습니다.", 48, "카지노"
        Exit Sub
    End If
    
    Timer = 0
    
    Randomize
    timTimer.Enabled = -1
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub timTimer_Timer()
    Dim i%, ret%
    
    If Timer >= 60 Then
        timTimer.Enabled = 0
        ret = 1
        For i = 1 To 6
            If lblCasino(i).Caption <> txtMyNumber(i).Text Then
                ret = 0
                Exit Sub
            End If
        Next i
        If ret Then
            lblScore.Caption = CInt(lblScore.Caption) + CInt(txtDeal.Text)
        Else
            lblScore.Caption = CInt(lblScore.Caption) - CInt(txtDeal.Text)
        End If
        If lblScore.Caption < 0 Then lblScore.Caption = 0
    End If
    
    For i = Int(Timer / 10) + 1 To 6
        lblCasino(i).Caption = Int((Rnd * 8) + 1)
    Next i
    
    Timer = Timer + 1
End Sub
