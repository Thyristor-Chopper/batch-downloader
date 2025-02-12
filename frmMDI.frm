VERSION 5.00
Begin VB.MDIForm frmMDI 
   Appearance      =   0  '평면
   BackColor       =   &H8000000C&
   Caption         =   "다운로드 부스터"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "frmMDI.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows 기본값
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
    
    Me.Caption = t(Me.Caption, "Download Booster") & " v" & App.Major & "." & App.Minor & "." & App.Revision
    frmTabBar.tsFormTabs.Tabs.Clear
    
    frmMain.Top = frmTabBar.Height
    frmMain.Left = 0
    frmMain.Show
    frmMain.FormID = 1
    FormCount = 1
    frmTabBar.tsFormTabs.Tabs.Add 0, "1", " " & t("세션", "Session") & " 1 "
    frmTabBar.tsFormTabs.Tabs.Add , , " + "
    Me.Height = frmMain.Height + frmTabBar.Height
    Me.Width = frmMain.Width + 120
    
    Forms.Add frmMain, "1"
End Sub

Private Sub MDIForm_Resize()
    frmTabBar.Width = Me.Width - 120
    Dim MainForm As frmMain
    For Each MainForm In Forms
        If Not MainForm Is Nothing Then
            MainForm.Height = Me.Height - frmTabBar.Height - 525
        End If
    Next MainForm
End Sub

Sub NewSession()
    Dim NewMainForm As frmMain
    Set NewMainForm = New frmMain
    FormCount = FormCount + 1
    NewMainForm.FormID = FormCount
    frmTabBar.tsFormTabs.Tabs.Add(frmTabBar.tsFormTabs.Tabs.Count, CStr(FormCount), " " & t("세션", "Session") & " " & FormCount & " ").Selected = True
    NewMainForm.Top = frmTabBar.Height
    NewMainForm.Left = 0
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
End Sub
