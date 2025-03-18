VERSION 5.00
Begin VB.Form frmButtonTest 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows ±âº»°ª
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   480
      Width           =   2175
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   120
      Width           =   2775
   End
   Begin prjDownloadBooster.CommandButtonEx CommandButtonEx3 
      Default         =   -1  'True
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   2160
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±¼¸²"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin prjDownloadBooster.CommandButtonEx CommandButtonEx2 
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   1560
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      SplitButton     =   -1  'True
      IsTygemButton   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±¼¸²"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin prjDownloadBooster.CommandButtonEx CommandButtonEx1 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   960
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±¼¸²"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmButtonTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButtonEx1_Click()
    MsgBox "cancel"
End Sub

Private Sub CommandButtonEx1_Dropdown()
    MsgBox 1
End Sub

Private Sub CommandButtonEx2_Dropdown()
    MsgBox 2
End Sub

Private Sub CommandButtonEx3_Click()
    MsgBox "default"
End Sub

Private Sub CommandButtonEx3_Dropdown()
    MsgBox 3
End Sub
