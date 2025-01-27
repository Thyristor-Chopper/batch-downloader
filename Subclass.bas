Attribute VB_Name = "Subclass"
' [ 참고 자료 ]
'- https://www.vbforums.com/showthread.php?213415-Visual-Basic-API-FAQs&p=1263307#post1263307
'- https://cafe.daum.net/0pds/37XW/36
'- http://www.jasinskionline.com/windowsapi/ref/i/insertmenuitem.html

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Const GWL_WNDPROC = (-4)
Private Const WM_SIZING = &H214
Public Const WM_GETMINMAXINFO = &H24
Public Const WM_SYSCOMMAND = &H112
Public Const WM_INITMENU = &H116
Const WM_DWMCOMPOSITIONCHANGED As Long = &H31E
Const DWM_EC_DISABLECOMPOSITION As Long = 0
Const DWM_EC_ENABLECOMPOSITION As Long = 1
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

Private Const MIN_WIDTH = 200  'The minimum width in pixels
Private Const MIN_HEIGHT = 200 'The minimum height in pixels
Private Const MAX_WIDTH = 500  'The maximum width in pixels
Private Const MAX_HEIGHT = 500 'The maximum height in pixels

Public MinWidth As Collection
Public MinHeight As Collection
Public MaxWidth As Collection
Public MaxHeight As Collection

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type MINMAXINFO
    ptReserved As POINTAPI
    ptMaxSize As POINTAPI
    ptMaxPosition As POINTAPI
    ptMinTrackSize As POINTAPI
    ptMaxTrackSize As POINTAPI
End Type
 
Private mPrevProc As Long
Private mPrevProc2 As Long

Public MainFormOnTop As Boolean
 
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
 
Sub SetWindowSizeLimit2(hWnd As Long, minW As Integer, maxW As Integer, minH As Integer, maxH As Integer)
    If Exists(MinWidth, hWnd) Then MinWidth.Remove CStr(hWnd)
    If Exists(MinHeight, hWnd) Then MinHeight.Remove CStr(hWnd)
    If Exists(MaxWidth, hWnd) Then MaxWidth.Remove CStr(hWnd)
    If Exists(MaxHeight, hWnd) Then MaxHeight.Remove CStr(hWnd)
    MinWidth.Add minW / 15, CStr(hWnd)
    MinHeight.Add minH / 15, CStr(hWnd)
    MaxWidth.Add maxW / 15, CStr(hWnd)
    MaxHeight.Add maxH / 15, CStr(hWnd)
    If mPrevProc2 <= 0& Then _
        mPrevProc2 = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf NewWndProc2)
End Sub
 
Sub Unhook(hWnd As Long)
    Call SetWindowLong(hWnd, GWL_WNDPROC, mPrevProc)
    mPrevProc = 0&
End Sub
 
Sub Unhook2(hWnd As Long)
    Call SetWindowLong(hWnd, GWL_WNDPROC, mPrevProc2)
    mPrevProc2 = 0&
End Sub
 
Function NewWndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error Resume Next
    
    Dim hSysMenu As Long
    Dim mii As MENUITEMINFO
 
    Select Case uMsg
        Case WM_GETMINMAXINFO
            Dim lpMMI As MINMAXINFO
            CopyMemory lpMMI, ByVal lParam, Len(lpMMI)
            lpMMI.ptMinTrackSize.X = MinWidth(CStr(hWnd))
            lpMMI.ptMinTrackSize.Y = MinHeight(CStr(hWnd))
            lpMMI.ptMaxTrackSize.X = MaxWidth(CStr(hWnd))
            lpMMI.ptMaxTrackSize.Y = MaxHeight(CStr(hWnd))
            CopyMemory ByVal lParam, lpMMI, Len(lpMMI)
            
            NewWndProc = 1&
            Exit Function
        Case WM_INITMENU
            hSysMenu = GetSystemMenu(hWnd, 0)
            With mii
                .cbSize = Len(mii)
                .fMask = MIIM_STATE
                .fState = MFS_ENABLED Or IIf(MainFormOnTop, MFS_CHECKED, 0)
            End With
            SetMenuItemInfo hSysMenu, 1000, 0, mii
            
            NewWndProc = 1&
            Exit Function
        Case WM_SYSCOMMAND
            If wParam = 1000 Then '항상 위에 표시
                MainFormOnTop = Not MainFormOnTop
                SetWindowPos hWnd, IIf(MainFormOnTop, HWND_TOPMOST, HWND_NOTOPMOST), 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
                SaveSetting "DownloadBooster", "Options", "AlwaysOnTop", Abs(CInt(MainFormOnTop))
                
                NewWndProc = 1&
                Exit Function
            ElseIf wParam = 1001 And (Not (frmMain.Height <= 6930 + PaddedBorderWidth * 15 * 2)) Then '일괄처리 접기
                frmMain.cmdBatch_Click
                
                NewWndProc = 1&
                Exit Function
            ElseIf wParam = 1002 And (frmMain.Height <= 6930 + PaddedBorderWidth * 15 * 2) Then '일괄처리 펼치기
                frmMain.cmdBatch_Click
                
                NewWndProc = 1&
                Exit Function
            ElseIf wParam = 1003 And (Not (frmMain.Height <= 6930 + PaddedBorderWidth * 15 * 2)) Then
                frmMain.Height = 8985 + PaddedBorderWidth * 15 * 2
            
                NewWndProc = 1&
                Exit Function
            End If
        Case WM_DWMCOMPOSITIONCHANGED
            frmMain.OnDWMChange
            
            NewWndProc = 1&
            Exit Function
    End Select
    
    If mPrevProc > 0& Then
        NewWndProc = CallWindowProc(mPrevProc, hWnd, uMsg, wParam, lParam)
    Else
        NewWndProc = DefWindowProc(hWnd, uMsg, wParam, lParam)
    End If
End Function
 
Function NewWndProc2(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error Resume Next
 
    If uMsg = WM_GETMINMAXINFO Then
        Dim lpMMI As MINMAXINFO
        CopyMemory lpMMI, ByVal lParam, Len(lpMMI)
        lpMMI.ptMinTrackSize.X = MinWidth(CStr(hWnd))
        lpMMI.ptMinTrackSize.Y = MinHeight(CStr(hWnd))
        lpMMI.ptMaxTrackSize.X = MaxWidth(CStr(hWnd))
        lpMMI.ptMaxTrackSize.Y = MaxHeight(CStr(hWnd))
        CopyMemory ByVal lParam, lpMMI, Len(lpMMI)
        
        NewWndProc2 = 1&
        Exit Function
    End If
    
    If mPrevProc2 > 0& Then
        NewWndProc2 = CallWindowProc(mPrevProc2, hWnd, uMsg, wParam, lParam)
    Else
        NewWndProc2 = DefWindowProc(hWnd, uMsg, wParam, lParam)
    End If
End Function

