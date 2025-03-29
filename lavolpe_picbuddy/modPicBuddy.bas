Attribute VB_Name = "modPicBuddy"
Option Explicit


' WARNING //// WARNING //// WARNING
' Subclassing is in play. This means you must not END your project while in design mode
' -- do not execute an END statement in your code
' -- do not click the blue 'stop' button on the VB toolbar
' -- do not click the 'end' button in a debug message box
' Save your work often. While subclassing, any errors you generate in your code
'   can cause a crash of your application !!!
' Recommendations....
' 1) While tweaking/modifying your project, do not call AttachBuddy. Rem them out
'   -- if AttachBuddy is not called, AttachChildControl, DetachChildControl, DetachBuddy have no effect
' 2) If you do call AttachBuddy, do so only to test visual effects
' 3) Once you are done tweaking your project, then call AttachBuddy as needed (i.e., un-rem the calls)

' Usercontrol designed to fake transparency for pictureboxes only.
' When a picturebox is assigned via AttachBuddy, these things apply
' 1) Any picture property in the picturebox will be lost
' 2) The .AutoRedraw property is set to false
' 3) You can reset it to True, but each time the picturebox is updated, it will be reset to False
'   -- Monitor the picturebox's Change event. It will fire each time the picturebox is updated by this control
' 4) Option buttons and checkboxes can be rendered transparently
' 5) To remove subclassing for the picturebox, call DeatchBuddy
' If you load checkboxes, option buttons during runtime, and want them
'   rendered transparent in the picturebox....
'   Call AttachChildControl and pass the newly loaded control
' When unloading checkboxes, option buttons that have been added using AttachChildControl...
'   Call DetachChildControl before unloading that object

' This usercontrol cannot handle specific actions generated in code.
' 1) Changing .BackColor property of picturebox's container
'   Fix: After changing .BackColor call .Refresh (i.e., Me.BackColor = vbWhite: Me.Refresh)
' 2) Changing a checkbox or option button text at runtime
'   Fix: After changing the text, refresh (i.e., Check1.Caption = "Remove": Check1.Refresh)
' 3) Changing the .Alignment property at runtime. VB destroys the control & creates a new one
'   Fix: After changing alignment, re-add control (i.e., Check1.Alignment = vbCenter: AttachChildControl Check1)
'   Note: you should call DetachChildControl before changing alignment. But is not absolutely required


Private Declare Function GetClientRect Lib "user32.dll" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function ScreenToClient Lib "user32.dll" (ByVal hWnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function ClientToScreen Lib "user32.dll" (ByVal hWnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetStockObject Lib "gdi32.dll" (ByVal nIndex As Long) As Long
Private Declare Function SetBkMode Lib "gdi32.dll" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function GetProp Lib "user32.dll" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32.dll" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function SetWindowOrgEx Lib "gdi32.dll" (ByVal hDC As Long, ByVal nX As Long, ByVal nY As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function InvalidateRect Lib "user32.dll" (ByVal hWnd As Long, ByRef lpRect As RECT, ByVal bErase As Long) As Long
Private Const NULL_BRUSH As Long = 5
Private Const NEWTRANSPARENT As Long = 3
Private Const WM_CTLCOLORSTATIC As Long = &H138
Private Const WM_PAINT As Long = &HF&
Private Const WM_ERASEBKGND As Long = &H14
Private Const WM_ENABLE As Long = &HA
Private Const WM_PRINTCLIENT As Long = &H318
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type POINTAPI
    X As Long
    Y As Long
End Type


Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function comctl32DllGetVersion Lib "comctl32" Alias "DllGetVersion" (pdvi As DLLVERSIONINFO) As Long
Private Declare Function IsAppThemed Lib "uxtheme.dll" () As Long
Private Declare Function IsThemeActive Lib "uxtheme.dll" () As Long
Private Declare Function GetVersion Lib "kernel32" () As Long
Private Const WM_THEMECHANGD As Long = &H31A&
Private Type DLLVERSIONINFO
    cbSize As Long
    dwMajor As Long
    dwMinor As Long
    dwBuildNumber As Long
    dwPlatformID As Long
End Type

Private m_Themed As Boolean

Public Function ValidateThemeEmployed() As Boolean
    ' Purpose :: Determine if application is themed

    Dim lVersionInfo As Long, hMod As Long, fa As Long
    Dim tdVI As DLLVERSIONINFO
    
    lVersionInfo = GetVersion()
    Select Case (lVersionInfo And &HFF)
        Case 6&
            m_Themed = True
        Case 5&
            m_Themed = ((lVersionInfo And &HFF00&) \ &H100 > 0&)
            ' ^^ if minor=zero then Win2K
        Case Else
            m_Themed = False
    End Select
    If m_Themed Then
        m_Themed = False
        If IsThemeActive() Then
            If IsAppThemed() Then               ' app not themed
                hMod = LoadLibrary("comctl32.dll")
                If hMod Then
                    fa = GetProcAddress(hMod, "DllGetVersion")
                    If fa Then
                       tdVI.cbSize = Len(tdVI)
                       fa = comctl32DllGetVersion(tdVI)
                       If fa = 0& Then
                           ' currently has common controls v6 or better loaded?
                          m_Themed = (tdVI.dwMajor > 5&)
                       End If
                    End If
                    FreeLibrary hMod
                End If
            End If
        End If
    End If
    ValidateThemeEmployed = m_Themed
    
End Function

Public Function picBuddyWindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    ' never call this from your project
    
    Dim lProc As Long
    lProc = GetProp(hWnd, "WndProc")
    Select Case uMsg
        Case WM_CTLCOLORSTATIC ' message sent to containers
            If GetProp(lParam, "WndProc") Then
                Dim vPT As POINTAPI
                CallWindowProc lProc, hWnd, uMsg, wParam, lParam
                If GetProp(lParam, "EraseBkg") = 1& Or m_Themed = True Then
                    SetProp lParam, "EraseBkg", 0
                    ClientToScreen lParam, vPT
                    ScreenToClient hWnd, vPT
                    SetWindowOrgEx wParam, vPT.X, vPT.Y, vPT
                    SendMessage hWnd, WM_PAINT, wParam, ByVal 0&
                    SetWindowOrgEx wParam, vPT.X, vPT.Y, vPT
                End If
                lProc = 0&                                      ' don't pass this message
                SetBkMode wParam, NEWTRANSPARENT
                picBuddyWindowProc = GetStockObject(NULL_BRUSH)
            End If
        
        Case WM_ERASEBKGND                                       ' only handle ones sent to children
            If m_Themed = False Then
                If GetProp(hWnd, "ChildBtn") Then SetProp hWnd, "EraseBkg", 1&
            End If
            
        Case WM_THEMECHANGD
            If GetProp(hWnd, "ChildBtn") = 0& Then ValidateThemeEmployed
        
        Case WM_ENABLE                                       ' only handle ones sent to children
            If GetProp(hWnd, "ChildBtn") Then
                Dim wRect As RECT, iPt As POINTAPI
                picBuddyWindowProc = CallWindowProc(lProc, hWnd, uMsg, wParam, lParam)
                GetClientRect hWnd, wRect
                InvalidateRect hWnd, wRect, True
                lProc = 0&                                      ' don't pass this message
            End If
    End Select                                                  ' forward message along if appropriate
    If lProc Then picBuddyWindowProc = CallWindowProc(lProc, hWnd, uMsg, wParam, lParam)

End Function
