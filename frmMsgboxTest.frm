VERSION 5.00
Begin VB.Form frmMsgboxTest 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows ±âº»°ª
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   3720
      Top             =   1560
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   495
      Left            =   960
      TabIndex        =   3
      Top             =   2280
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   960
      TabIndex        =   2
      Top             =   1680
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   960
      TabIndex        =   1
      Top             =   1080
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmMsgboxTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Alert "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
End Sub

Private Sub Command2_Click()
    Command2.Caption = Confirm("xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx", App.Title, Me)
End Sub

Private Sub Command3_Click()
    Command3.Caption = ConfirmEx("xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx", App.Title, Me)
End Sub

Private Sub Command4_Click()
    Command4.Caption = ConfirmCancel("xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx", App.Title, Me)
End Sub

Private Sub Form_Load()
    Label1.Caption = "ok:" & vbOK & ",cancel:" & vbCancel & ",yes:" & vbYes & ",no:" & vbNo
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = 0
    Command1_Click
End Sub
