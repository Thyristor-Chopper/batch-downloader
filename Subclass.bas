Attribute VB_Name = "Subclass"
'https://www.vbforums.com/showthread.php?213415-Visual-Basic-API-FAQs&p=1263307#post1263307
'https://m.cafe.daum.net/0pds/37XW/36도 참고함
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Const GWL_WNDPROC = (-4)
Private Const WM_SIZING = &H214
Public Const WM_GETMINMAXINFO = &H24
Private Const WMSZ_LEFT = 1
Private Const WMSZ_RIGHT = 2
Private Const WMSZ_TOP = 3
Private Const WMSZ_TOPLEFT = 4
Private Const WMSZ_TOPRIGHT = 5
Private Const WMSZ_BOTTOM = 6
Private Const WMSZ_BOTTOMLEFT = 7
Private Const WMSZ_BOTTOMRIGHT = 8

Private Const MIN_WIDTH = 200  'The minimum width in pixels
Private Const MIN_HEIGHT = 200 'The minimum height in pixels
Private Const MAX_WIDTH = 500  'The maximum width in pixels
Private Const MAX_HEIGHT = 500 'The maximum height in pixels

Public MinWidth As Collection
Public MinHeight As Collection
Public MaxWidth As Collection
Public MaxHeight As Collection

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type MINMAXINFO
    ptReserved As POINTAPI
    ptMaxSize As POINTAPI
    ptMaxPosition As POINTAPI
    ptMinTrackSize As POINTAPI
    ptMaxTrackSize As POINTAPI
End Type
 
Private mPrevProc As Long
 
Sub SetWindowSizeLimit(hWnd As Long, minW As Integer, maxW As Integer, minH As Integer, maxH As Integer)
    If Exists(MinWidth, hWnd) Then MinWidth.Remove CStr(hWnd)
    If Exists(MinHeight, hWnd) Then MinHeight.Remove CStr(hWnd)
    If Exists(MaxWidth, hWnd) Then MaxWidth.Remove CStr(hWnd)
    If Exists(MaxHeight, hWnd) Then MaxHeight.Remove CStr(hWnd)
    MinWidth.Add minW / 15, CStr(hWnd)
    MinHeight.Add minH / 15, CStr(hWnd)
    MaxWidth.Add maxW / 15, CStr(hWnd)
    MaxHeight.Add maxH / 15, CStr(hWnd)
    If mPrevProc <= 0& Then _
        mPrevProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf NewWndProc)
End Sub
 
Sub Unhook(hWnd As Long)
    Call SetWindowLong(hWnd, GWL_WNDPROC, mPrevProc)
    mPrevProc = 0&
End Sub
 
Function NewWndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error Resume Next
 
    If uMsg = WM_GETMINMAXINFO Then
        Dim lpMMI As MINMAXINFO
        CopyMemory lpMMI, ByVal lParam, Len(lpMMI)
        lpMMI.ptMinTrackSize.x = MinWidth(CStr(hWnd))
        lpMMI.ptMinTrackSize.y = MinHeight(CStr(hWnd))
        lpMMI.ptMaxTrackSize.x = MaxWidth(CStr(hWnd))
        lpMMI.ptMaxTrackSize.y = MaxHeight(CStr(hWnd))
        CopyMemory ByVal lParam, lpMMI, Len(lpMMI)
        
        NewWndProc = 1&
        Exit Function
    End If
    
 
    If mPrevProc > 0& Then
        NewWndProc = CallWindowProc(mPrevProc, hWnd, uMsg, wParam, lParam)
    Else
        NewWndProc = DefWindowProc(hWnd, uMsg, wParam, lParam)
    End If
End Function

