VERSION 5.00
Begin VB.Form frmDummyForm 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "±¼¸²"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows ±âº»°ª
   Begin prjDownloadBooster.ImageList imgFiles 
      Left            =   0
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      ImageWidth      =   16
      ImageHeight     =   16
      ColorDepth      =   4
      MaskColor       =   16711935
      InitListImages  =   "frmDummyForm.frx":0000
   End
   Begin VB.Image imgTrain 
      Height          =   480
      Index           =   5
      Left            =   2400
      Picture         =   "frmDummyForm.frx":01A8
      Top             =   0
      Width           =   480
   End
   Begin VB.Image imgTrain 
      Height          =   480
      Index           =   4
      Left            =   1800
      Picture         =   "frmDummyForm.frx":0467
      Top             =   0
      Width           =   480
   End
   Begin VB.Image imgTrain 
      Height          =   480
      Index           =   3
      Left            =   1200
      Picture         =   "frmDummyForm.frx":072B
      Top             =   0
      Width           =   480
   End
   Begin VB.Image imgTrain 
      Height          =   480
      Index           =   2
      Left            =   600
      Picture         =   "frmDummyForm.frx":0AE8
      Top             =   0
      Width           =   480
   End
   Begin VB.Image imgTrain 
      Height          =   480
      Index           =   1
      Left            =   0
      Picture         =   "frmDummyForm.frx":0DD9
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "frmDummyForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
