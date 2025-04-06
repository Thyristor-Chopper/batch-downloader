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
Public Const hWnd_TOPMOST = -1
Public Const hWnd_NOTOPMOST = -2
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

Public MainFormOnTop As Boolean

