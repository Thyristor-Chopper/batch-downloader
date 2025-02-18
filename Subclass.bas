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
Public Const WM_SETTINGCHANGE As Long = &H1A
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
 
Private mPrevProc_Main As Long
Private mPrevProc2 As Long
Private mPrevProc3 As Long
Private mPrevProc_Options As Long
'Private mPrevProc_Bluemetal As Long

Public MainFormOnTop As Boolean
 
Sub SetWindowSizeLimit(hWnd As Long, minW As Integer, maxW As Integer, minH As Integer, maxH As Integer)
    If Not IsRunning Then Exit Sub
    
    If Exists(MinWidth, hWnd) Then MinWidth.Remove CStr(hWnd)
    If Exists(MinHeight, hWnd) Then MinHeight.Remove CStr(hWnd)
    If Exists(MaxWidth, hWnd) Then MaxWidth.Remove CStr(hWnd)
    If Exists(MaxHeight, hWnd) Then MaxHeight.Remove CStr(hWnd)
    MinWidth.Add minW / 15, CStr(hWnd)
    MinHeight.Add minH / 15, CStr(hWnd)
    MaxWidth.Add maxW / 15, CStr(hWnd)
    MaxHeight.Add maxH / 15, CStr(hWnd)
    If mPrevProc_Main <= 0& Then _
        mPrevProc_Main = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WndProc_Main)
End Sub
 
Sub SetWindowSizeLimit2(hWnd As Long, minW As Integer, maxW As Integer, minH As Integer, maxH As Integer)
    If Not IsRunning Then Exit Sub
    
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
 
Sub SetWindowSizeLimit3(hWnd As Long, minW As Integer, maxW As Integer, minH As Integer, maxH As Integer)
    If Not IsRunning Then Exit Sub
    
    If Exists(MinWidth, hWnd) Then MinWidth.Remove CStr(hWnd)
    If Exists(MinHeight, hWnd) Then MinHeight.Remove CStr(hWnd)
    If Exists(MaxWidth, hWnd) Then MaxWidth.Remove CStr(hWnd)
    If Exists(MaxHeight, hWnd) Then MaxHeight.Remove CStr(hWnd)
    MinWidth.Add minW / 15, CStr(hWnd)
    MinHeight.Add minH / 15, CStr(hWnd)
    MaxWidth.Add maxW / 15, CStr(hWnd)
    MaxHeight.Add maxH / 15, CStr(hWnd)
    If mPrevProc3 <= 0& Then _
        mPrevProc3 = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf NewWndProc3)
End Sub
 
Sub Hook_Options(hWnd As Long)
    If Not IsRunning Then Exit Sub
    
    If mPrevProc_Options <= 0& Then _
        mPrevProc_Options = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WndProc_Options)
End Sub
 
'Sub Hook_Bluemetal(hWnd As Long)
'    If Not IsRunning Then Exit Sub
'
'    If mPrevProc_Bluemetal <= 0& Then _
'        mPrevProc_Bluemetal = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WndProc_Bluemetal)
'End Sub
 
Sub Unhook_Main(hWnd As Long)
    If Not IsRunning Then Exit Sub
    
    SetWindowLong hWnd, GWL_WNDPROC, mPrevProc_Main
    mPrevProc_Main = 0&
End Sub
 
Sub Unhook2(hWnd As Long)
    If Not IsRunning Then Exit Sub
    
    SetWindowLong hWnd, GWL_WNDPROC, mPrevProc2
    mPrevProc2 = 0&
End Sub
 
Sub Unhook3(hWnd As Long)
    If Not IsRunning Then Exit Sub
    
    SetWindowLong hWnd, GWL_WNDPROC, mPrevProc3
    mPrevProc3 = 0&
End Sub
 
Sub Unhook_Options(hWnd As Long)
    If Not IsRunning Then Exit Sub
    
    SetWindowLong hWnd, GWL_WNDPROC, mPrevProc_Options
    mPrevProc_Options = 0&
End Sub
 
'Sub Unhook_Bluemetal(hWnd As Long)
'    If Not IsRunning Then Exit Sub
'
'    SetWindowLong hWnd, GWL_WNDPROC, mPrevProc_Bluemetal
'    mPrevProc_Bluemetal = 0&
'End Sub
 
Function WndProc_Main(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error Resume Next
    
    Dim hSysMenu As Long
    Dim MII As MENUITEMINFO
 
    Select Case uMsg
        Case WM_GETMINMAXINFO
            Dim lpMMI As MINMAXINFO
            CopyMemory lpMMI, ByVal lParam, Len(lpMMI)
            lpMMI.ptMinTrackSize.X = MinWidth(CStr(hWnd))
            lpMMI.ptMinTrackSize.Y = MinHeight(CStr(hWnd))
            lpMMI.ptMaxTrackSize.X = MaxWidth(CStr(hWnd))
            lpMMI.ptMaxTrackSize.Y = MaxHeight(CStr(hWnd))
            CopyMemory ByVal lParam, lpMMI, Len(lpMMI)
            
            WndProc_Main = 1&
            Exit Function
        Case WM_INITMENU
            hSysMenu = GetSystemMenu(hWnd, 0)
            With MII
                .cbSize = Len(MII)
                .fMask = MIIM_STATE
                .fState = MFS_ENABLED Or IIf(MainFormOnTop, MFS_CHECKED, 0)
            End With
            SetMenuItemInfo hSysMenu, 1000, 0, MII
            
            WndProc_Main = 1&
            Exit Function
        Case WM_SYSCOMMAND
            If wParam = 1000 Then '항상 위에 표시
                MainFormOnTop = Not MainFormOnTop
                SetWindowPos hWnd, IIf(MainFormOnTop, HWND_TOPMOST, HWND_NOTOPMOST), 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
                SaveSetting "DownloadBooster", "Options", "AlwaysOnTop", Abs(CInt(MainFormOnTop))
                
                WndProc_Main = 1&
                Exit Function
            ElseIf wParam = 1001 And (Not (frmMain.Height <= 6930 + PaddedBorderWidth * 15 * 2)) Then '일괄처리 접기
                frmMain.cmdBatch_Click
                
                WndProc_Main = 1&
                Exit Function
            ElseIf wParam = 1002 And (frmMain.Height <= 6930 + PaddedBorderWidth * 15 * 2) Then '일괄처리 펼치기
                frmMain.cmdBatch_Click
                
                WndProc_Main = 1&
                Exit Function
            ElseIf wParam = 1003 And (Not (frmMain.Height <= 6930 + PaddedBorderWidth * 15 * 2)) Then
                frmMain.Height = 8985 + PaddedBorderWidth * 15 * 2
            
                WndProc_Main = 1&
                Exit Function
            End If
        Case WM_DWMCOMPOSITIONCHANGED
            frmMain.OnDWMChange
        Case WM_SETTINGCHANGE
            Select Case GetStrFromPtr(lParam)
                Case "WindowMetrics"
                    UpdateBorderWidth
                    
                    If Exists(MinWidth, hWnd) Then MinWidth.Remove CStr(hWnd)
                    If Exists(MaxWidth, hWnd) Then MaxWidth.Remove CStr(hWnd)
                    If Exists(MinHeight, hWnd) Then MinHeight.Remove CStr(hWnd)
                    MinWidth.Add (9450 + PaddedBorderWidth * 15 * 2) / 15, CStr(hWnd)
                    MaxWidth.Add (9450 + PaddedBorderWidth * 15 * 2) / 15, CStr(hWnd)
                    MinHeight.Add (8220 + PaddedBorderWidth * 15 * 2) / 15, CStr(hWnd)
                    
                    frmMain.Width = 9450 + PaddedBorderWidth * 15 * 2
                    
                    On Error Resume Next
                    Dim ctrl As Control
                    For Each ctrl In frmMain.Controls
                        If TypeName(ctrl) = "FrameW" Or TypeName(ctrl) = "CheckBoxW" Or TypeName(ctrl) = "OptionButtonW" Or TypeName(ctrl) = "CommandButtonW" Or TypeName(ctrl) = "Slider" Then ctrl.Refresh
                    Next ctrl
                    Dim PrevTrackerVisualStyles As Boolean
                    PrevTrackerVisualStyles = frmMain.trThreadCount.VisualStyles
                    frmMain.trThreadCount.VisualStyles = False
                    frmMain.trThreadCount.VisualStyles = True
                    frmMain.trThreadCount.VisualStyles = PrevTrackerVisualStyles
            End Select
    End Select
    
    If mPrevProc_Main > 0& Then
        WndProc_Main = CallWindowProc(mPrevProc_Main, hWnd, uMsg, wParam, lParam)
    Else
        WndProc_Main = DefWindowProc(hWnd, uMsg, wParam, lParam)
    End If
End Function
 
Function NewWndProc2(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error Resume Next
 
    Select Case uMsg
        Case WM_GETMINMAXINFO
            Dim lpMMI As MINMAXINFO
            CopyMemory lpMMI, ByVal lParam, Len(lpMMI)
            lpMMI.ptMinTrackSize.X = MinWidth(CStr(hWnd))
            lpMMI.ptMinTrackSize.Y = MinHeight(CStr(hWnd))
            lpMMI.ptMaxTrackSize.X = MaxWidth(CStr(hWnd))
            lpMMI.ptMaxTrackSize.Y = MaxHeight(CStr(hWnd))
            CopyMemory ByVal lParam, lpMMI, Len(lpMMI)
            
            NewWndProc2 = 1&
            Exit Function
        Case WM_SETTINGCHANGE
            Select Case GetStrFromPtr(lParam)
                Case "WindowMetrics"
                    UpdateBorderWidth
                    
                    If Exists(MinWidth, hWnd) Then MinWidth.Remove CStr(hWnd)
                    If Exists(MinHeight, hWnd) Then MinHeight.Remove CStr(hWnd)
                    MinWidth.Add (5145 + PaddedBorderWidth * 15 * 2) / 15, CStr(hWnd)
                    MinHeight.Add (2310 + PaddedBorderWidth * 15 * 2) / 15, CStr(hWnd)
                    
                    frmBatchAdd.Form_Resize
            End Select
    End Select
    
    If mPrevProc2 > 0& Then
        NewWndProc2 = CallWindowProc(mPrevProc2, hWnd, uMsg, wParam, lParam)
    Else
        NewWndProc2 = DefWindowProc(hWnd, uMsg, wParam, lParam)
    End If
End Function

Function NewWndProc3(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error Resume Next
 
    Select Case uMsg
        Case WM_GETMINMAXINFO
            Dim lpMMI As MINMAXINFO
            CopyMemory lpMMI, ByVal lParam, Len(lpMMI)
            lpMMI.ptMinTrackSize.X = MinWidth(CStr(hWnd))
            lpMMI.ptMinTrackSize.Y = MinHeight(CStr(hWnd))
            lpMMI.ptMaxTrackSize.X = MaxWidth(CStr(hWnd))
            lpMMI.ptMaxTrackSize.Y = MaxHeight(CStr(hWnd))
            CopyMemory ByVal lParam, lpMMI, Len(lpMMI)
            
            NewWndProc3 = 1&
            Exit Function
        Case WM_SETTINGCHANGE
            Select Case GetStrFromPtr(lParam)
                Case "WindowMetrics"
                    UpdateBorderWidth
                    
                    If Exists(MinWidth, hWnd) Then MinWidth.Remove CStr(hWnd)
                    If Exists(MinHeight, hWnd) Then MinHeight.Remove CStr(hWnd)
                    MinWidth.Add (10245 + PaddedBorderWidth * 15 * 2) / 15, CStr(hWnd)
                    MinHeight.Add (IIf(Tags.BrowseTargetForm = 3, 8835, 6165) + PaddedBorderWidth * 15 * 2) / 15, CStr(hWnd)
                    
                    frmExplorer.Form_Resize
            End Select
    End Select
    
    If mPrevProc3 > 0& Then
        NewWndProc3 = CallWindowProc(mPrevProc3, hWnd, uMsg, wParam, lParam)
    Else
        NewWndProc3 = DefWindowProc(hWnd, uMsg, wParam, lParam)
    End If
End Function

Function WndProc_Options(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error Resume Next
 
    Select Case uMsg
        Case WM_SETTINGCHANGE
            Select Case GetStrFromPtr(lParam)
                Case "WindowMetrics"
                    UpdateBorderWidth
                    frmOptions.SetPreviewPosition
                    frmOptions.DrawTabBackground
            End Select
    End Select
    
    If mPrevProc_Options > 0& Then
        WndProc_Options = CallWindowProc(mPrevProc_Options, hWnd, uMsg, wParam, lParam)
    Else
        WndProc_Options = DefWindowProc(hWnd, uMsg, wParam, lParam)
    End If
End Function

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
'    If mPrevProc_Bluemetal > 0& Then
'        WndProc_Bluemetal = CallWindowProc(mPrevProc_Bluemetal, hWnd, uMsg, wParam, lParam)
'    Else
'        WndProc_Bluemetal = DefWindowProc(hWnd, uMsg, wParam, lParam)
'    End If
'End Function


