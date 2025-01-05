Attribute VB_Name = "Startup"
Option Explicit
Private Const NULL_PTR As Long = 0
Private Const PTR_SIZE As Long = 4
Private Declare Function FindWindow Lib "user32" Alias "FindWindowW" (ByVal lpClassName As Long, ByVal lpWindowName As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long

Sub Main()
    Call InitVisualStylesFixes
    frmMain.Show vbModeless
End Sub
