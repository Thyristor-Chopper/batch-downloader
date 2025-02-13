VERSION 5.00
Begin VB.MDIForm frmMDI 
   Appearance      =   0  '���
   BackColor       =   &H8000000C&
   Caption         =   "�ٿ�ε� �ν���"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8340
   Icon            =   "frmMDI.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows �⺻��
   Begin prjDownloadBooster.StatusBar sbStatusBar 
      Align           =   2  '�Ʒ� ����
      Height          =   330
      Left            =   0
      Top             =   2760
      Width           =   8340
      _ExtentX        =   14711
      _ExtentY        =   582
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      InitPanels      =   "frmMDI.frx":212A
   End
End
Attribute VB_Name = "frmMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public FormCount As Long
Public Forms As Collection

Private Sub MDIForm_Load()
    If GetSetting("DownloadBooster", "Options", "DisableDWMWindow", DefaultDisableDWMWindow) = 1 Then DisableDWMWindow Me.hWnd
    SetWindowPos Me.hWnd, IIf(MainFormOnTop, HWND_TOPMOST, HWND_NOTOPMOST), 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    
    Set Forms = New Collection
    FormCount = 0
    
    frmTabBar.Show
    frmTabBar.Width = Me.Width - 120
    frmTabBar.Top = 0
    frmTabBar.Left = 0
    
    Dim Lft%
    Dim Top%
    Top = GetSetting("DownloadBooster", "UserData", "FormTop", -1)
    Lft = GetSetting("DownloadBooster", "UserData", "FormLeft", -1)
    If Top >= 0 And Lft >= 0 Then
        Me.Top = Top
        Me.Left = Lft
    End If
    Me.Width = 9450 + PaddedBorderWidth * 15 * 2 + 120
    SetWindowSizeLimit Me.hWnd, Me.Width, Me.Width, 8220 + PaddedBorderWidth * 15 * 2 + frmTabBar.Height, Screen.Height + 1200
    
    Me.Caption = t(Me.Caption, "Download Booster") & " v" & App.Major & "." & App.Minor & "." & App.Revision
    frmTabBar.tsFormTabs.Tabs.Clear
    
    frmMain.Top = frmTabBar.Height
    frmMain.Left = 0
    frmMain.Show
    frmMain.FormID = 0
    FormCount = 0
    frmTabBar.tsFormTabs.Tabs.Add 0, "0", " " & t("�ϰ� ó��", "Batch") & " "
    frmTabBar.tsFormTabs.Tabs.Add , , " + "
    Me.Height = frmMain.Height + frmTabBar.Height
    
    frmMain.lblURL.Visible = 0
    frmMain.lblFilePath.Visible = 0
    frmMain.txtURL.Visible = 0
    frmMain.txtURL.Text = ""
    frmMain.txtFileName.Visible = 0
    frmMain.cmdClear.Visible = 0
    frmMain.tygReset.Visible = 0
    frmMain.cmdBrowse.Visible = 0
    frmMain.tygBrowse.Visible = 0
    
    NewSession
End Sub

Private Sub MDIForm_Resize()
    frmTabBar.Width = Me.Width - 120
    frmMain.Height = Me.Height - frmTabBar.Height - 525 - 525 - 330 + 540
End Sub

Sub NewSession()
    Dim NewMainForm As frmMain
    Set NewMainForm = New frmMain
    FormCount = FormCount + 1
    NewMainForm.FormID = FormCount
    frmTabBar.tsFormTabs.Tabs.Add(frmTabBar.tsFormTabs.Tabs.Count, CStr(FormCount), " " & t("����", "Session") & " " & FormCount & " ").Selected = True
    NewMainForm.Top = frmTabBar.Height
    NewMainForm.Left = 0
    NewMainForm.Height = 6000
    NewMainForm.Show
    Forms.Add NewMainForm, CStr(FormCount)
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    Dim MainForm As frmMain
    For Each MainForm In Forms
        If Not MainForm Is Nothing Then
            Unload MainForm
        End If
    Next MainForm
    Unload frmMain
End Sub
