VERSION 5.00
Begin VB.Form frmDummyForm 
   BorderStyle     =   0  '����
   ClientHeight    =   3090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "����"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDummyForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows �⺻��
   Visible         =   0   'False
   Begin prjDownloadBooster.ImageList imgCancel 
      Left            =   600
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      InitListImages  =   "frmDummyForm.frx":212A
   End
   Begin prjDownloadBooster.ImageList imgOK 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      InitListImages  =   "frmDummyForm.frx":22D2
   End
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
      InitListImages  =   "frmDummyForm.frx":247A
   End
End
Attribute VB_Name = "frmDummyForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
