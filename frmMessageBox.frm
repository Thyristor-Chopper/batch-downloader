VERSION 5.00
Begin VB.Form frmMessageBox 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "메시지 상자"
   ClientHeight    =   1365
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   28440
   BeginProperty Font 
      Name            =   "굴림"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMessageBox.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   28440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.CommandButton cmdTryAgain 
      Caption         =   "다시 시도(&T)"
      Height          =   315
      Left            =   15240
      TabIndex        =   11
      Top             =   840
      Width           =   1455
   End
   Begin VB.Timer timeout 
      Enabled         =   0   'False
      Left            =   360
      Top             =   960
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "취소"
      Height          =   320
      Left            =   5880
      TabIndex        =   3
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton cmdNo 
      Caption         =   "아니요(&N)"
      Height          =   320
      Left            =   4320
      TabIndex        =   1
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton cmdYes 
      Caption         =   "예(&Y)"
      Height          =   320
      Left            =   2760
      TabIndex        =   0
      Top             =   840
      Width           =   1455
   End
   Begin prjDownloadBooster.OptionButtonW optNo 
      Height          =   255
      Left            =   1080
      TabIndex        =   9
      Top             =   1320
      Width           =   1575
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "아니요(&N)"
   End
   Begin prjDownloadBooster.OptionButtonW optYes 
      Height          =   255
      Left            =   1080
      TabIndex        =   8
      Top             =   960
      Width           =   1575
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "예(&Y)"
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "확인"
      Height          =   315
      Left            =   7440
      TabIndex        =   2
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton cmdRetry 
      Caption         =   "다시 시도(&R)"
      Height          =   315
      Left            =   9000
      TabIndex        =   5
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton cmdAbort 
      Caption         =   "중단(&A)"
      Height          =   315
      Left            =   10560
      TabIndex        =   4
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton cmdIgnore 
      Caption         =   "무시(&I)"
      Height          =   315
      Left            =   12120
      TabIndex        =   6
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton cmdContinue 
      Caption         =   "계속(&C)"
      Height          =   315
      Left            =   13680
      TabIndex        =   12
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "도움말"
      Height          =   315
      Left            =   16800
      TabIndex        =   7
      Top             =   840
      Width           =   1455
   End
   Begin VB.Image imgError 
      Height          =   360
      Left            =   75
      Picture         =   "frmMessageBox.frx":000C
      Top             =   90
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgExclamation 
      Height          =   360
      Left            =   75
      Picture         =   "frmMessageBox.frx":00C1
      Top             =   90
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgQuestion 
      Height          =   360
      Left            =   75
      Picture         =   "frmMessageBox.frx":01DC
      Top             =   90
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgInformation 
      Height          =   360
      Left            =   75
      Picture         =   "frmMessageBox.frx":029A
      Top             =   90
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblContent 
      BackColor       =   &H00F8EFE5&
      BackStyle       =   0  '투명
      Caption         =   "내용"
      Height          =   495
      Left            =   960
      TabIndex        =   10
      Top             =   360
      Width           =   27255
   End
   Begin VB.Image imgTrain 
      Height          =   480
      Index           =   4
      Left            =   4440
      Picture         =   "frmMessageBox.frx":0351
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgTrain 
      Height          =   480
      Index           =   3
      Left            =   3840
      Picture         =   "frmMessageBox.frx":0615
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgTrain 
      Height          =   480
      Index           =   2
      Left            =   3240
      Picture         =   "frmMessageBox.frx":08D4
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgTrain 
      Height          =   480
      Index           =   1
      Left            =   2640
      Picture         =   "frmMessageBox.frx":0C91
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgTrain 
      Height          =   480
      Index           =   0
      Left            =   2040
      Picture         =   "frmMessageBox.frx":0F82
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "frmMessageBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public MsgBoxMode As Byte
Public MsgBoxResult As VbMsgBoxResult
Public ResultID As String
Public MessageBoxObject As frmMessageBox

Private Sub cmdAbort_Click()
    MsgBoxResult = vbAbort
    Unload Me
End Sub

Private Sub cmdContinue_Click()
    MsgBoxResult = vbContinue
    Unload Me
End Sub

Private Sub cmdIgnore_Click()
    MsgBoxResult = vbIgnore
    Unload Me
End Sub

Private Sub cmdNo_Click()
    MsgBoxResult = vbNo
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If MsgBoxMode = vbYesNoEx Then
        If optYes.Value = True Then
            MsgBoxResult = vbYes
        Else
            MsgBoxResult = vbNo
        End If
    Else
        MsgBoxResult = vbOK
    End If
    Unload Me
End Sub

Private Sub cmdRetry_Click()
    MsgBoxResult = vbRetry
    Unload Me
End Sub

Private Sub cmdTryAgain_Click()
    MsgBoxResult = vbTryAgain
    Unload Me
End Sub

Private Sub cmdYes_Click()
    MsgBoxResult = vbYes
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    MsgBoxResult = vbCancel
    Unload Me
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    Select Case MsgBoxMode
        Case vbOKOnly
            cmdOK.SetFocus
        Case vbYesNo
            cmdYes.SetFocus
        Case vbYesNoEx
            optNo.SetFocus
        Case vbYesNoCancel
            cmdCancel.SetFocus
        Case vbAbortRetryIgnore
            cmdAbort.SetFocus
        Case vbRetryCancel
            cmdRetry.SetFocus
        Case vbOKCancel
            cmdOK.SetFocus
        Case vbCancelTryContinue
            cmdCancel.SetFocus
    End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Select Case KeyCode
        Case 78 'N
            If optNo.Visible Then
                optNo.Value = True
                optNo_Click
                optNo.SetFocus
            End If
            If cmdNo.Visible Then cmdNo_Click
        Case 89 'Y
            If optYes.Visible Then
                optYes.Value = True
                optYes_Click
                optYes.SetFocus
            End If
            If cmdYes.Visible Then cmdYes_Click
        Case 82 'R
            If cmdRetry.Visible Then cmdRetry_Click
        Case 65 'A
            If cmdAbort.Visible Then cmdAbort_Click
        Case 73 'I
            If cmdIgnore.Visible Then cmdIgnore_Click
        Case 67 'C
            If cmdContinue.Visible Then cmdContinue_Click
        Case 84 'T
            If cmdTryAgain.Visible Then cmdTryAgain_Click
    End Select
End Sub

Private Sub Form_Load()
    InitForm Me
End Sub

Sub Init()
    Dim SystemMenu As Long
    SystemMenu = GetSystemMenu(Me.hWnd, 0&)
    DeleteMenu SystemMenu, 0&, MF_BYCOMMAND
    If MsgBoxMode = vbYesNo Or MsgBoxMode = vbAbortRetryIgnore Then
        DeleteMenu SystemMenu, SC_CLOSE, MF_BYCOMMAND
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> vbFormCode Then
        If MsgBoxMode = vbYesNo Or MsgBoxMode = vbAbortRetryIgnore Then
            Cancel = 1
            Exit Sub
        Else
            MsgBoxResult = vbCancel
        End If
    End If
    GetSystemMenu Me.hWnd, 1&
    If MsgBoxMode <> vbOKOnly Then
        If Functions.MsgBoxResults Is Nothing Then Set Functions.MsgBoxResults = New Collection
        If Exists(Functions.MsgBoxResults, ResultID) Then Functions.MsgBoxResults.Remove ResultID
        Functions.MsgBoxResults.Add MsgBoxResult, ResultID
    End If
    
    If Not MessageBoxObject Is Nothing Then
        Unload MessageBoxObject
        Set MessageBoxObject = Nothing
    End If
End Sub

Private Sub timeout_Timer()
    cmdOK_Click
End Sub

Private Sub optNo_Click()
    cmdOK.Enabled = True
End Sub

Private Sub optYes_Click()
    cmdOK.Enabled = True
End Sub
