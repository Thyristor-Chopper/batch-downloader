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
   Begin VB.PictureBox pbDummy 
      AutoRedraw      =   -1  'True
      Height          =   135
      Left            =   720
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   0
      Top             =   600
      Width           =   135
   End
   Begin prjDownloadBooster.ImageList imgFiles 
      Left            =   0
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      ImageWidth      =   16
      ImageHeight     =   16
      ColorDepth      =   4
      MaskColor       =   16711935
      InitListImages  =   "frmDummyForm.frx":0000
   End
End
Attribute VB_Name = "frmDummyForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
