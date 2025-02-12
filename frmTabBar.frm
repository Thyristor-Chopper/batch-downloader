VERSION 5.00
Begin VB.Form frmTabBar 
   BorderStyle     =   0  '¾øÀ½
   Caption         =   "ÅÇ"
   ClientHeight    =   360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2415
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
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   360
   ScaleWidth      =   2415
   ShowInTaskbar   =   0   'False
   Begin prjDownloadBooster.TabStrip tsFormTabs 
      Height          =   345
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   2385
      _ExtentX        =   4207
      _ExtentY        =   609
      MultiRow        =   0   'False
      TabWidthStyle   =   1
      TabMinWidth     =   13
      Separators      =   0   'False
      TabScrollWheel  =   0   'False
      InitTabs        =   "frmTabBar.frx":0000
   End
   Begin prjDownloadBooster.TabStrip tsDummy 
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   315
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   450
      Enabled         =   0   'False
      InitTabs        =   "frmTabBar.frx":008C
   End
End
Attribute VB_Name = "frmTabBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
    If GetSetting("DownloadBooster", "Options", "DisableDWMWindow", DefaultDisableDWMWindow) = 1 Then DisableDWMWindow Me.hWnd
    SetFormBackgroundColor Me
    SetFont Me
    SetWindowPos Me.hWnd, IIf(MainFormOnTop, HWND_TOPMOST, HWND_NOTOPMOST), 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    
    tsFormTabs.Width = Me.Width - 60
    tsDummy.Width = Me.Width
End Sub

Private Sub tsFormTabs_TabClick(ByVal TabItem As TbsTab)
    If TabItem.Index = tsFormTabs.Tabs.Count And TabItem.Caption = " + " Then
        frmMDI.NewSession
    Else
        Dim MainForm As frmMain
        For Each MainForm In frmMDI.Forms
            If MainForm.FormID = CLng(TabItem.Key) Then
                MainForm.Show
            Else
                MainForm.Hide
            End If
        Next MainForm
    End If
End Sub
