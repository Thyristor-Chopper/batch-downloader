VERSION 5.00
Begin VB.Form frmGame 
   BackColor       =   &H00F8EFE5&
   BorderStyle     =   3  'ũ�� ���� ��ȭ ����
   Caption         =   "���߷� ����"
   ClientHeight    =   2160
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "����"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGame.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '������ ���
   Begin prjDownloadBooster.CommandButtonW Btn 
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Timer Timer2 
      Interval        =   5000
      Left            =   3960
      Top             =   1080
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   3960
      Top             =   480
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "0ȸ ���� / 0ȸ ����"
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label lblDesc 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "���� ������ ���ʽÿ�."
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   600
      Width           =   3015
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Success%, Fail%
Dim Pressed As Boolean

Private Sub Btn_Click()
    If Timer1.Enabled Then
        Success = Success + 1
        Pressed = -1
    Else
        Fail = Fail + 1
    End If
    
    lblStatus.Caption = Success & "ȸ ���� / " & Fail & "ȸ ����"
End Sub

Private Sub Form_Load()
    Success = 0
    Fail = 0
    Pressed = 0
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = 0
    lblDesc.Caption = "���� ������ ���ʽÿ�."
    If Not Pressed Then Fail = Fail + 1
    lblStatus.Caption = Success & "ȸ ���� / " & Fail & "ȸ ����"
End Sub

Private Sub Timer2_Timer()
    lblDesc.Caption = "����!!"
    Timer1.Enabled = -1
    Timer2.Interval = CInt((Rnd * 10) + 1) * 1000
    Timer2.Enabled = 0
    Timer2.Enabled = -1
End Sub
