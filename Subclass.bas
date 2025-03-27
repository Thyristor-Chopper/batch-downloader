Attribute VB_Name = "Subclass"
' [ 참고 자료 ]
'- https://www.vbforums.com/showthread.php?213415-Visual-Basic-API-FAQs&p=1263307#post1263307
'- https://cafe.daum.net/0pds/37XW/36
'- http://www.jasinskionline.com/windowsapi/ref/i/insertmenuitem.html

Option Explicit

Private Const GWL_WNDPROC = (-4)
Public Const WM_MOVE = &H3&
Public Const WM_SETCURSOR = &H20&
Public Const WM_NCPAINT = &H85&
Public Const WM_COMMAND = &H111&
Public Const WM_SIZING = &H214
Public Const WM_GETMINMAXINFO = &H24
Public Const WM_SYSCOMMAND = &H112
Public Const WM_INITMENU = &H116
Public Const WM_SETTINGCHANGE = &H1A
Public Const WM_DWMCOMPOSITIONCHANGED = &H31E
Public Const WM_THEMECHANGED = &H31A
Public Const WM_DPICHANGED = &H2E0
Public Const WM_CTLCOLORSCROLLBAR = &H137&
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Private Const WMSZ_LEFT = 1
Private Const WMSZ_RIGHT = 2
Private Const WMSZ_TOP = 3
Private Const WMSZ_TOPLEFT = 4
Private Const WMSZ_TOPRIGHT = 5
Private Const WMSZ_BOTTOM = 6
Private Const WMSZ_BOTTOMLEFT = 7
Private Const WMSZ_BOTTOMRIGHT = 8

Private mPrevProc_Explorer As Long
Private mPrevProc_Options As Long
'Private mPrevProc_Bluemetal As Long

Public MainFormOnTop As Boolean
 
'Sub Hook_Bluemetal(hWnd As Long)
'    If mPrevProc_Bluemetal = 0& Then _
'        mPrevProc_Bluemetal = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WndProc_Bluemetal)
'End Sub
 
'Sub Unhook_Bluemetal(hWnd As Long)
'    SetWindowLong hWnd, GWL_WNDPROC, mPrevProc_Bluemetal
'    mPrevProc_Bluemetal = 0&
'End Sub

'Function WndProc_Bluemetal(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'    On Error Resume Next
'
'    Select Case uMsg
'        Case WM_NCPAINT, WM_MOVE
'            Dim rc As RECT
'            GetWindowRect Bluemetal.hWnd, rc
'            SetWindowPos Bluemetal.pbTopLeft.hWnd, 0, rc.Left, rc.Top, Bluemetal.pbTopLeft.Width / 15, Bluemetal.pbTopLeft.Height / 15, SWP_FRAMECHANGED
'            SetWindowPos Bluemetal.pbTopMiddle.hWnd, 0, rc.Left + Bluemetal.pbTopLeft.Width / 15, rc.Top, Bluemetal.pbTopMiddle.Width / 15, Bluemetal.pbTopMiddle.Height / 15, SWP_FRAMECHANGED
'            SetWindowPos Bluemetal.pbTopRight.hWnd, 0, rc.Right - Bluemetal.pbTopRight.Width / 15, rc.Top, Bluemetal.pbTopRight.Width / 15, Bluemetal.pbTopRight.Height / 15, SWP_FRAMECHANGED
'            SetWindowPos Bluemetal.pbLeft.hWnd, 0, rc.Left, rc.Top + Bluemetal.pbTopLeft.Height / 15, Bluemetal.pbLeft.Width / 15, Bluemetal.pbLeft.Height / 15, SWP_FRAMECHANGED
'            SetWindowPos Bluemetal.pbBottomLeft.hWnd, 0, rc.Left, rc.Top + Bluemetal.pbTopLeft.Height / 15 + Bluemetal.pbLeft.Height / 15, Bluemetal.pbBottomLeft.Width / 15, Bluemetal.pbBottomLeft.Height / 15, SWP_FRAMECHANGED
'            SetWindowPos Bluemetal.pbBottomMiddle.hWnd, 0, rc.Left + Bluemetal.pbBottomLeft.Width / 15, rc.Top + Bluemetal.pbTopLeft.Height / 15 + Bluemetal.pbLeft.Height / 15, Bluemetal.pbBottomMiddle.Width / 15, Bluemetal.pbBottomMiddle.Height / 15, SWP_FRAMECHANGED
'            SetWindowPos Bluemetal.pbBottomRight.hWnd, 0, rc.Left + Bluemetal.pbBottomLeft.Width / 15 + Bluemetal.pbBottomMiddle.Width / 15, rc.Top + Bluemetal.pbTopLeft.Height / 15 + Bluemetal.pbLeft.Height / 15, Bluemetal.pbBottomRight.Width / 15, Bluemetal.pbBottomRight.Height / 15, SWP_FRAMECHANGED
'            SetWindowPos Bluemetal.pbRight.hWnd, 0, rc.Right - Bluemetal.pbRight.Width / 15, rc.Top + Bluemetal.pbTopRight.Height / 15, Bluemetal.pbRight.Width / 15, Bluemetal.pbRight.Height / 15, SWP_FRAMECHANGED
'    End Select
'
'    If mPrevProc_Bluemetal <> 0& Then
'        WndProc_Bluemetal = CallWindowProc(mPrevProc_Bluemetal, hWnd, uMsg, wParam, lParam)
'    Else
'        WndProc_Bluemetal = DefWindowProc(hWnd, uMsg, wParam, lParam)
'    End If
'End Function


