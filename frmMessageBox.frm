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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   28440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
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
      TabIndex        =   10
      Top             =   1320
      Width           =   1575
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "아니요(&N)"
   End
   Begin prjDownloadBooster.OptionButtonW optYes 
      Height          =   255
      Left            =   1080
      TabIndex        =   9
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
   Begin VB.CommandButton cmdFail 
      Caption         =   "실패(&F)"
      Height          =   315
      Left            =   13680
      TabIndex        =   7
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "도움말"
      Height          =   315
      Left            =   15240
      TabIndex        =   8
      Top             =   840
      Width           =   1455
   End
   Begin VB.Image imgMBIconError 
      Height          =   630
      Index           =   1
      Left            =   60
      Picture         =   "frmMessageBox.frx":000C
      Stretch         =   -1  'True
      Top             =   60
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Image imgMBIconError 
      Height          =   630
      Index           =   0
      Left            =   60
      Picture         =   "frmMessageBox.frx":0384
      Stretch         =   -1  'True
      Top             =   60
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Image imgMBIconWarning 
      Height          =   630
      Index           =   1
      Left            =   60
      Picture         =   "frmMessageBox.frx":06E3
      Stretch         =   -1  'True
      Top             =   60
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Image imgMBIconWarning 
      Height          =   630
      Index           =   0
      Left            =   60
      Picture         =   "frmMessageBox.frx":0B4E
      Stretch         =   -1  'True
      Top             =   60
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Image imgMBIconQuestion 
      Height          =   630
      Index           =   1
      Left            =   60
      Picture         =   "frmMessageBox.frx":0F9C
      Stretch         =   -1  'True
      Top             =   60
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Image imgMBIconQuestion 
      Height          =   630
      Index           =   0
      Left            =   60
      Picture         =   "frmMessageBox.frx":1322
      Stretch         =   -1  'True
      Top             =   60
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Image imgMBIconInfo 
      Height          =   630
      Index           =   1
      Left            =   60
      Picture         =   "frmMessageBox.frx":1693
      Stretch         =   -1  'True
      Top             =   60
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Image imgMBIconInfo 
      Height          =   630
      Index           =   0
      Left            =   60
      Picture         =   "frmMessageBox.frx":1A09
      Stretch         =   -1  'True
      Top             =   60
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label lblContent 
      BackColor       =   &H00F8EFE5&
      BackStyle       =   0  '투명
      Caption         =   "내용"
      Height          =   495
      Left            =   960
      TabIndex        =   11
      Top             =   360
      Width           =   27255
   End
End
Attribute VB_Name = "frmMessageBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ButtonPressed As Boolean
Public MsgBoxMode As Byte
Public MsgBoxResult As VbMsgBoxResult
Public ResultID As String

Private Sub cmdAbort_Click()
    MsgBoxResult = vbAbort
    ButtonPressed = -1
    Unload Me
End Sub

Private Sub cmdAbort_KeyDown(KeyCode As Integer, Shift As Integer)
    OnKeyDown KeyCode
End Sub

Private Sub cmdCancel_KeyDown(KeyCode As Integer, Shift As Integer)
    OnKeyDown KeyCode
End Sub

Private Sub cmdIgnore_Click()
    MsgBoxResult = vbIgnore
    ButtonPressed = -1
    Unload Me
End Sub

Private Sub cmdIgnore_KeyDown(KeyCode As Integer, Shift As Integer)
    OnKeyDown KeyCode
End Sub

Private Sub cmdNo_Click()
    MsgBoxResult = vbNo
    ButtonPressed = -1
    Unload Me
End Sub

Private Sub cmdNo_KeyDown(KeyCode As Integer, Shift As Integer)
    OnKeyDown KeyCode
End Sub

Private Sub cmdOK_Click()
    If optYes.Visible Then
        If optYes.Value = True Then
            MsgBoxResult = vbYes
        Else
            MsgBoxResult = vbNo
        End If
    Else
        MsgBoxResult = vbOK
    End If
    ButtonPressed = -1
    Unload Me
End Sub

Private Sub cmdOK_KeyDown(KeyCode As Integer, Shift As Integer)
    OnKeyDown KeyCode
End Sub

Private Sub cmdRetry_Click()
    MsgBoxResult = vbRetry
    ButtonPressed = -1
    Unload Me
End Sub

Private Sub cmdRetry_KeyDown(KeyCode As Integer, Shift As Integer)
    OnKeyDown KeyCode
End Sub

Private Sub cmdYes_Click()
    MsgBoxResult = vbYes
    ButtonPressed = -1
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    MsgBoxResult = vbCancel
    ButtonPressed = -1
    Unload Me
End Sub

Private Sub cmdYes_KeyDown(KeyCode As Integer, Shift As Integer)
    OnKeyDown KeyCode
End Sub

Sub PreFocusDefaultButton()
    On Error Resume Next
    Select Case MsgBoxMode
        Case 1
            cmdOK.SetFocus
        Case 2
            cmdYes.SetFocus
        Case 3
            optNo.SetFocus
        Case 4
            cmdCancel.SetFocus
        Case 5
            cmdAbort.SetFocus
        Case 6
            cmdRetry.SetFocus
        Case 7
            cmdOK.SetFocus
    End Select
End Sub

Private Sub Form_Activate()
    PreFocusDefaultButton
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    OnKeyDown KeyCode
End Sub

Private Sub Form_Load()
    If GetSetting("DownloadBooster", "Options", "DisableDWMWindow", DefaultDisableDWMWindow) = 1 Then DisableDWMWindow Me.hWnd
    SetFormBackgroundColor Me
    SetFont Me
    SetWindowPos Me.hWnd, IIf(MainFormOnTop, HWND_TOPMOST, HWND_NOTOPMOST), 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    ButtonPressed = 0
    Init
End Sub

Sub Init()
    If MsgBoxMode = 2 Or MsgBoxMode = 5 Then
        Dim SystemMenu As Long
        SystemMenu = GetSystemMenu(Me.hWnd, 0)
        DeleteMenu SystemMenu, SC_CLOSE, MF_BYCOMMAND
        DeleteMenu SystemMenu, 0, MF_BYCOMMAND
    End If
End Sub

Private Sub Form_Paint()
    PreFocusDefaultButton
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If cmdYes.Visible And cmdNo.Visible And (Not cmdCancel.Visible) And (Not ButtonPressed) Then
        Cancel = 1
        Exit Sub
    End If
    If Not ButtonPressed Then MsgBoxResult = vbCancel
    GetSystemMenu Me.hWnd, 1
    If MsgBoxMode <> 1 Then
        If Functions.MsgBoxResults Is Nothing Then Set Functions.MsgBoxResults = New Collection
        If Exists(Functions.MsgBoxResults, ResultID) Then Functions.MsgBoxResults.Remove ResultID
        Functions.MsgBoxResults.Add MsgBoxResult, ResultID
    End If
End Sub

Private Sub optNo_KeyDown(KeyCode As Integer, Shift As Integer)
    OnKeyDown KeyCode
End Sub

Private Sub optYes_KeyDown(KeyCode As Integer, Shift As Integer)
    OnKeyDown KeyCode
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

Sub OnKeyDown(KeyCode As Integer)
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
    End Select
End Sub
