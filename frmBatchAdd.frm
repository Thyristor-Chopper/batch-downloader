VERSION 5.00
Begin VB.Form frmBatchAdd 
   BorderStyle     =   3  'ũ�� ���� ��ȭ ����
   Caption         =   "�ϰ� �ٿ�ε�"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6240
   BeginProperty Font 
      Name            =   "����"
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
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '������ ���
   Begin VB.TextBox txtURLs 
      Height          =   2535
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  '�����
      TabIndex        =   0
      Top             =   360
      Width           =   4215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "���"
      Height          =   375
      Left            =   4560
      TabIndex        =   3
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Ȯ��"
      Height          =   375
      Left            =   4560
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "�� �ٿ� ���� �ּҸ� �Է��Ͻʽÿ�."
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmBatchAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Sub txtURLs_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = 13 Or KeyCode = 10) And Right$(txtURLs.Text, 2) = vbCrLf Then
        cmdOK_Click
    End If
End Sub
