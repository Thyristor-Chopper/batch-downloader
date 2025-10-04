Attribute VB_Name = "Subclass"
' [ 참고 자료 ]
'- https://www.vbforums.com/showthread.php?213415-Visual-Basic-API-FAQs&p=1263307#post1263307
'- https://cafe.daum.net/0pds/37XW/36
'- http://www.jasinskionline.com/windowsapi/ref/i/insertmenuitem.html

Option Explicit

Dim IsWindowActive As Byte

Public TitleHeight&, BorderSize&
Public Enable As Boolean

Public Const GWL_EXSTYLE As Long = -20&
Public Const RGN_DIFF As Long = 4&

Public Const WM_NCLBUTTONDOWN = &HA1

Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Sub ReleaseCapture Lib "user32" ()
Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As Long) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal uFlags As Long, ByVal X As Long, ByVal Y As Long, ByVal nReserved As Long, ByVal hWnd As Long, lprcRect As Any) As Long
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long

Type POINTAPI
    X As Long
    Y As Long
End Type

Public Const TPM_LEFTALIGN As Long = &H0&
Public Const TPM_RETURNCMD As Long = &H100&

Private Const GWL_WNDPROC = (-4)
Public Const WM_ACTIVATEAPP As Long = &H1C
Public Const WA_INACTIVE As Long = 0
Public Const WA_ACTIVE As Long = 1
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
Public Const SWP_NOZORDER = &H4
Public Const SWP_FRAMECHANGED As Long = &H20&
Private Const WMSZ_LEFT = 1
Private Const WMSZ_RIGHT = 2
Private Const WMSZ_TOP = 3
Private Const WMSZ_TOPLEFT = 4
Private Const WMSZ_TOPRIGHT = 5
Private Const WMSZ_BOTTOM = 6
Private Const WMSZ_BOTTOMLEFT = 7
Private Const WMSZ_BOTTOMRIGHT = 8

Public Const HTERROR As Long = -2       ' Error
Public Const HTTRANSPARENT As Long = -1 ' Transparent
Public Const HTNOWHERE As Long = 0      ' Outside window
Public Const HTCLIENT As Long = 1       ' Client area
Public Const HTCAPTION As Long = 2      ' Title bar
Public Const HTSYSMENU As Long = 3      ' System menu
Public Const HTGROWBOX As Long = 4      ' Size box (old)
Public Const HTSIZE As Long = HTGROWBOX
Public Const HTMENU As Long = 5         ' Menu
Public Const HTHSCROLL As Long = 6      ' Horizontal scroll bar
Public Const HTVSCROLL As Long = 7      ' Vertical scroll bar
Public Const HTMINBUTTON As Long = 8    ' Minimize button
Public Const HTMAXBUTTON As Long = 9    ' Maximize button
Public Const HTLEFT As Long = 10        ' Left border
Public Const HTRIGHT As Long = 11       ' Right border
Public Const HTTOP As Long = 12         ' Top border
Public Const HTTOPLEFT As Long = 13     ' Top-left corner
Public Const HTTOPRIGHT As Long = 14    ' Top-right corner
Public Const HTBOTTOM As Long = 15      ' Bottom border
Public Const HTBOTTOMLEFT As Long = 16  ' Bottom-left corner
Public Const HTBOTTOMRIGHT As Long = 17 ' Bottom-right corner
Public Const HTBORDER As Long = 18      ' Border (obsolete)
Public Const HTREDUCE As Long = HTMINBUTTON
Public Const HTZOOM As Long = HTMAXBUTTON
Public Const HTSIZEFIRST As Long = HTLEFT
Public Const HTSIZELAST As Long = HTBOTTOMRIGHT
Public Const HTOBJECT As Long = 19      ' Object in window
Public Const HTCLOSE As Long = 20       ' Close button
Public Const HTHELP As Long = 21        ' Help button

Public Const WM_NCCALCSIZE As Long = &H83
Public Const WM_NCHITTEST As Long = &H84

Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private mPrevProc_Bluemetal As Long

Public MainFormOnTop As Boolean
 
Sub Hook_Bluemetal(hWnd As Long)
    If mPrevProc_Bluemetal = 0& Then _
        mPrevProc_Bluemetal = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WndProc_Bluemetal)
End Sub
 
Sub Unhook_Bluemetal(hWnd As Long)
    SetWindowLong hWnd, GWL_WNDPROC, mPrevProc_Bluemetal
    mPrevProc_Bluemetal = 0&
End Sub

Function WndProc_Bluemetal(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error Resume Next
    Dim rc As RECT

    Select Case uMsg
        Case WM_NCPAINT, WM_MOVE
            If Enable Then
                GetWindowRect Bluemetal.hWnd, rc
                SetWindowPos Bluemetal.pbTopLeft.hWnd, 0, rc.Left, rc.Top, Bluemetal.pbTopLeft.Width / 15, Bluemetal.pbTopLeft.Height / 15, SWP_FRAMECHANGED
                SetWindowPos Bluemetal.pbTopMiddle.hWnd, 0, rc.Left + Bluemetal.pbTopLeft.Width / 15, rc.Top, Bluemetal.pbTopMiddle.Width / 15, Bluemetal.pbTopMiddle.Height / 15, SWP_FRAMECHANGED
                SetWindowPos Bluemetal.pbTopRight.hWnd, 0, rc.Right - Bluemetal.pbTopRight.Width / 15, rc.Top, Bluemetal.pbTopRight.Width / 15, Bluemetal.pbTopRight.Height / 15, SWP_FRAMECHANGED
                SetWindowPos Bluemetal.pbLeft.hWnd, 0, rc.Left, rc.Top + Bluemetal.pbTopLeft.Height / 15, Bluemetal.pbLeft.Width / 15, Bluemetal.pbLeft.Height / 15, SWP_FRAMECHANGED
                SetWindowPos Bluemetal.pbBottomLeft.hWnd, 0, rc.Left, rc.Top + Bluemetal.pbTopLeft.Height / 15 + Bluemetal.pbLeft.Height / 15, Bluemetal.pbBottomLeft.Width / 15, Bluemetal.pbBottomLeft.Height / 15, SWP_FRAMECHANGED
                SetWindowPos Bluemetal.pbBottomMiddle.hWnd, 0, rc.Left + Bluemetal.pbBottomLeft.Width / 15, rc.Top + Bluemetal.pbTopLeft.Height / 15 + Bluemetal.pbLeft.Height / 15, Bluemetal.pbBottomMiddle.Width / 15, Bluemetal.pbBottomMiddle.Height / 15, SWP_FRAMECHANGED
                SetWindowPos Bluemetal.pbBottomRight.hWnd, 0, rc.Left + Bluemetal.pbBottomLeft.Width / 15 + Bluemetal.pbBottomMiddle.Width / 15, rc.Top + Bluemetal.pbTopLeft.Height / 15 + Bluemetal.pbLeft.Height / 15, Bluemetal.pbBottomRight.Width / 15, Bluemetal.pbBottomRight.Height / 15, SWP_FRAMECHANGED
                SetWindowPos Bluemetal.pbRight.hWnd, 0, rc.Right - Bluemetal.pbRight.Width / 15, rc.Top + Bluemetal.pbTopRight.Height / 15, Bluemetal.pbRight.Width / 15, Bluemetal.pbRight.Height / 15, SWP_FRAMECHANGED
                
                WndProc_Bluemetal = 0&
                Exit Function
            End If
        Case WM_NCCALCSIZE
            If wParam <> 0 And Enable Then
                CopyMemory rc, ByVal lParam, Len(rc)
        
                rc.Top = rc.Top + TitleHeight
                rc.Left = rc.Left + BorderSize
                rc.Right = rc.Right - BorderSize
                rc.Bottom = rc.Bottom - BorderSize + 1
        
                CopyMemory ByVal lParam, rc, Len(rc)
        
                WndProc_Bluemetal = 0
                Exit Function
            End If
        Case WM_NCHITTEST
            If Enable Then
                Dim X As Long, Y As Long
                X = (lParam And &HFFFF&)
                Y = ((lParam \ &H10000) And &HFFFF&)
                If X And &H8000& Then X = X Or &HFFFF0000
                If Y And &H8000& Then Y = Y Or &HFFFF0000
            
                Dim rcWin As RECT
                GetWindowRect hWnd, rcWin
            
                Dim hit As Long: hit = HTCLIENT
            
                If X >= rcWin.Left And X < rcWin.Left + BorderSize Then
                    hit = HTLEFT
                ElseIf X < rcWin.Right And X >= rcWin.Right - BorderSize Then
                    hit = HTRIGHT
                End If
            
                If Y >= rcWin.Top And Y < rcWin.Top + BorderSize Then
                    If hit = HTLEFT Then
                        hit = HTTOPLEFT
                    ElseIf hit = HTRIGHT Then
                        hit = HTTOPRIGHT
                    Else
                        hit = HTTOP
                    End If
                ElseIf Y < rcWin.Bottom And Y >= rcWin.Bottom - BorderSize + 1 Then
                    If hit = HTLEFT Then
                        hit = HTBOTTOMLEFT
                    ElseIf hit = HTRIGHT Then
                        hit = HTBOTTOMRIGHT
                    Else
                        hit = HTBOTTOM
                    End If
                End If
            
                If Y >= rcWin.Top + BorderSize And Y < rcWin.Top + TitleHeight Then
                    If hit = HTCLIENT Then hit = HTCAPTION
                End If
            
                WndProc_Bluemetal = hit
                Exit Function
            End If
        Case WM_ACTIVATEAPP
            Select Case LoWord(wParam)
                Case WA_INACTIVE
                    Bluemetal.IsWindowActive = 0
                Case WA_ACTIVE
                    Bluemetal.IsWindowActive = 1
            End Select
            Bluemetal.SetSkinTextures
            Bluemetal.SetSizableSkinTextures
    End Select

    If mPrevProc_Bluemetal <> 0& Then
        WndProc_Bluemetal = CallWindowProc(mPrevProc_Bluemetal, hWnd, uMsg, wParam, lParam)
    Else
        WndProc_Bluemetal = DefWindowProc(hWnd, uMsg, wParam, lParam)
    End If
End Function

Public Function LoWord(dw As Long) As Integer
    If dw And &H8000& Then
        LoWord = &H8000& Or (dw And &H7FFF&)
    Else
        LoWord = dw And &HFFFF&
    End If
End Function

