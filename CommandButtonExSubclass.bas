Attribute VB_Name = "CommandButtonExSubclass"
Option Explicit

Private Declare Function SetWindowSubclass Lib "comctl32" Alias "#410" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long, ByVal dwRefData As Long) As Long
Private Declare Function RemoveWindowSubclass Lib "comctl32" Alias "#412" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long) As Long
Private Declare Function DefSubclassProc Lib "comctl32" Alias "#413" (ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal pDestination As Long, ByVal pSource As Long, ByVal Length As Long)

Private Const WM_NOTIFY As Long = &H4E&
Private Const BCN_FIRST As Long = -1250&
Private Const BCN_DROPDOWN As Long = BCN_FIRST + &H2&
Private Const NM_GETCUSTOMSPLITRECT As Long = BCN_FIRST + &H3&

Private Type NMHDR
    hWndFrom As Long
    idFrom As Long
    code As Long
End Type

Dim Buttons As Collection

Sub HookCommandButtonEx(ByRef Button As CommandButton, ByRef Container As PictureBox, ByRef ctrl)
    On Error Resume Next
    If Buttons Is Nothing Then Set Buttons = New Collection
    If Exists(Buttons, CStr(Button.hWnd)) Then Buttons.Remove CStr(Button.hWnd)
    Buttons.Add ctrl, CStr(Button.hWnd)
    SetWindowSubclass Container.hWnd, AddressOf WndProc, ObjPtr(Container), 0&
End Sub

Sub UnhookCommandButtonEx(ByRef Button As CommandButton, ByRef Container As PictureBox)
    On Error Resume Next
    RemoveWindowSubclass Container.hWnd, AddressOf WndProc, ObjPtr(Container)
    Buttons.Remove CStr(Button.hWnd)
End Sub

Private Function WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal uIdSubclass As Object, ByVal dwRefData As Long) As Long
    Dim NMHDR As NMHDR
   
    If uMsg = WM_NOTIFY Then
        CopyMemory VarPtr(NMHDR), lParam, Len(NMHDR)
        If NMHDR.code = BCN_DROPDOWN Then
            If Exists(Buttons, CStr(NMHDR.hWndFrom)) Then Buttons(CStr(NMHDR.hWndFrom)).ClickDropdown
            WndProc = 1&
        ElseIf NMHDR.code = NM_GETCUSTOMSPLITRECT Then
            WndProc = 0&
        Else
            WndProc = DefSubclassProc(hWnd, uMsg, wParam, lParam)
        End If
    Else
        WndProc = DefSubclassProc(hWnd, uMsg, wParam, lParam)
    End If
End Function


