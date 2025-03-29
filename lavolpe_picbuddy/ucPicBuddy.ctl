VERSION 5.00
Begin VB.UserControl ucPicBuddy 
   BackStyle       =   0  '투명
   ClientHeight    =   825
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   825
   ClipBehavior    =   0  '없음
   HasDC           =   0   'False
   ScaleHeight     =   825
   ScaleWidth      =   825
   Windowless      =   -1  'True
End
Attribute VB_Name = "ucPicBuddy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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


'========================================================================================================
Private Declare Function GetProp Lib "user32.dll" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32.dll" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function RemoveProp Lib "user32.dll" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Declare Function AlphaBlend Lib "msimg32.dll" (ByVal hdcDest As Long, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal nHeightDest As Long, ByVal hdcSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal lBlendFunction As Long) As Long
Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Const GWL_WNDPROC As Long = -4
'-------------------------------------------------------------------------------------------------

Private m_Children As Collection
Private m_Buddy As VB.PictureBox
Private m_BorderOffset As POINTAPI
Private m_NoRedraw As Boolean
Private m_UserMode As Boolean
Private m_Opacity As Long

Public Property Get MyBuddy() As VB.PictureBox

    ' read only property that identifies which picturebox, if any, was assigned
    Set MyBuddy = m_Buddy

End Property

Public Function AttachBuddy(newBuddy As VB.PictureBox, _
                            Optional ByVal CheckBoxesTransparent As Boolean = True, _
                            Optional ByVal OptionButtonsTransparent As Boolean = True) As Boolean
    
    ' call this at runtime, in form_load
    ' i.e., Call AttachBuddy(Picture1)
    
    Me.DetachBuddy
    Set m_Buddy = newBuddy
    If Not m_Buddy Is Nothing Then
        modPicBuddy.ValidateThemeEmployed
        pvMoveMe
        pvSubClassBuddy True, CheckBoxesTransparent, OptionButtonsTransparent
        AttachBuddy = True
    End If

End Function

Public Sub DetachBuddy()

    ' call this to stop subclassing as needed
    
    pvSubClassBuddy False, False, False
    Set m_Buddy = Nothing
    m_NoRedraw = True

End Sub

Public Function AttachChildControl(theControl As Control) As Boolean
    
    ' if adding controls dynamically, at runtime, to the picturebox
    ' call this routine to start subclassing it
    
    Dim lProc As Long
    On Error Resume Next
    
    If m_Buddy Is Nothing Then Exit Function
    If theControl Is Nothing Then Exit Function
    Select Case TypeName(theControl)
    Case "OptionButton", "CheckBox"
        If theControl.Container <> m_Buddy Then Exit Function
        If theControl.Style <> 0 Then Exit Function ' not applicable for Graphical style controls
        If Not m_Children Is Nothing Then
            m_Children.Add theControl, "k" & theControl.hWnd
            If Err Then Exit Function ' already added
        Else
            m_Children.Add theControl, "k" & theControl.hWnd
        End If
        lProc = GetWindowLong(m_Buddy.hWnd, GWL_WNDPROC)
        If lProc = 0& Then
            SetProp m_Buddy.hWnd, "WndProc", lProc
            SetWindowLong m_Buddy.hWnd, GWL_WNDPROC, AddressOf picBuddyWindowProc
        End If
        lProc = GetWindowLong(theControl.hWnd, GWL_WNDPROC)
        SetProp theControl.hWnd, "WndProc", lProc
        SetProp theControl.hWnd, "BtnChild", 1
        SetWindowLong theControl.hWnd, GWL_WNDPROC, AddressOf picBuddyWindowProc
        theControl.Refresh
        AttachChildControl = True
    Case Else
    End Select

End Function

Public Function DetachChildControl(theControl As Control) As Boolean
    
    ' if removing controls dynamically, at runtime, from the picturebox
    ' call this routine to stop subclassing it
    
    Dim lProc As Long
    On Error Resume Next
    
    If m_Buddy Is Nothing Then Exit Function
    If theControl Is Nothing Then Exit Function
    Select Case TypeName(theControl)
    Case "OptionButton", "CheckBox"
        If Not m_Children Is Nothing Then
            m_Children.Remove "k" & theControl.hWnd
            If Err Then Exit Function
            lProc = GetProp(theControl.hWnd, "WndProc")
            If lProc Then
                SetWindowLong theControl.hWnd, GWL_WNDPROC, lProc
                RemoveProp theControl.hWnd, "WndProc"
                RemoveProp theControl.hWnd, "BtnChild"
                RemoveProp theControl.hWnd, "EraseBkg"
                theControl.Refresh
                DetachChildControl = True
            End If
            If m_Children.Count = 0& Then
                Set m_Children = Nothing
                lProc = GetProp(m_Buddy.hWnd, "WndProc")
                If lProc Then
                    SetWindowLong m_Buddy.hWnd, GWL_WNDPROC, lProc
                    RemoveProp m_Buddy.hWnd, "WndProc"
                End If
            End If
        End If
    Case Else
    End Select

End Function

Public Sub Refresh()
    UserControl.Refresh
End Sub

Public Property Get OpacityPercent() As Long
    OpacityPercent = m_Opacity
End Property
Public Property Let OpacityPercent(newValue As Long)
    If newValue >= 0& And newValue <= 100& Then
        If newValue <> m_Opacity Then
            m_Opacity = newValue
            Me.Refresh
            PropertyChanged "OpacityPercent"
        End If
    End If
End Property

Private Sub UserControl_HitTest(X As Single, Y As Single, HitResult As Integer)
    If m_UserMode = False Then HitResult = vbHitResultHit
End Sub

Private Sub UserControl_Initialize()
    UserControl.ScaleMode = vbPixels
    UserControl.FillColor = vbWhite                 ' design-time settings
    UserControl.FillStyle = vbDefault
    UserControl.ForeColor = vbRed
    m_NoRedraw = True
End Sub

Private Sub UserControl_Paint()
    If m_UserMode = False Then
        UserControl.Line (0, 0)-(UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1), UserControl.ForeColor, B
        UserControl.CurrentX = 5: UserControl.CurrentY = 5
        UserControl.Print "Picture"
        UserControl.CurrentX = 5: UserControl.Print "Buddy"
    ElseIf Not m_NoRedraw Then
        If pvReposition = False Then
            Dim lBlend As Long
            m_Buddy.AutoRedraw = True
            If m_Opacity = 0& Then
                BitBlt m_Buddy.hDC, 0, 0, UserControl.ScaleWidth - m_BorderOffset.X * 2, UserControl.ScaleHeight - m_BorderOffset.Y * 2, UserControl.hDC, m_BorderOffset.X, m_BorderOffset.Y, vbSrcCopy
            ElseIf m_Opacity < 100& Then
                lBlend = (((255& * (100& - m_Opacity)) \ 100&) * &H10000)
                AlphaBlend m_Buddy.hDC, 0, 0, UserControl.ScaleWidth - m_BorderOffset.X * 2, UserControl.ScaleHeight - m_BorderOffset.Y * 2, UserControl.hDC, m_BorderOffset.X, m_BorderOffset.Y, UserControl.ScaleWidth - m_BorderOffset.X * 2, UserControl.ScaleHeight - m_BorderOffset.Y * 2, lBlend
            Else
                Set m_Buddy.Picture = Nothing
                m_Buddy.Cls
            End If
            If m_Opacity < 100& Then Set m_Buddy.Picture = m_Buddy.Image
            m_Buddy.AutoRedraw = False
            pvRefreshContainedControls
        End If
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Opacity = PropBag.ReadProperty("Opacity", 0&)
End Sub

Private Sub UserControl_Show()
    On Error Resume Next
    m_UserMode = Ambient.UserMode
    If Err Then m_UserMode = True
    ' possible to get error if control used in another IDE (like Access or when dynamically loading during runtime)
End Sub

Private Sub UserControl_Terminate()
    pvSubClassBuddy False, False, False
End Sub

Private Sub pvSubClassBuddy(Init As Boolean, AddCheckBoxes As Boolean, AddOptionBtns As Boolean)

    Dim lProc As Long, X As Long, ctrlObj As Object
    If Init Then
        If (AddCheckBoxes Or AddOptionBtns) Then
            Set m_Children = New Collection
            On Error Resume Next
            For X = 1 To ParentControls.Count - 1
                Set ctrlObj = ParentControls(X)
                If ctrlObj.Container Is m_Buddy Then
                    If Err Then
                        Err.Clear
                    Else
                        Select Case TypeName(ctrlObj)
                        Case "OptionButton"
                            lProc = Abs(AddOptionBtns)
                        Case "CheckBox"
                            lProc = Abs(AddCheckBoxes)
                        Case Else
                            lProc = 0&
                        End Select
                        If lProc Then
                            If ctrlObj.Style = 0 Then   ' not applicable for Graphical style controls
                                m_Children.Add ctrlObj, "k" & ctrlObj.hWnd
                                lProc = GetWindowLong(ctrlObj.hWnd, GWL_WNDPROC)
                                SetProp ctrlObj.hWnd, "WndProc", lProc
                                SetProp ctrlObj.hWnd, "ChildBtn", 1
                                SetWindowLong ctrlObj.hWnd, GWL_WNDPROC, AddressOf picBuddyWindowProc
                                ctrlObj.Refresh
                            End If
                        End If
                    End If
                End If
            Next
            lProc = GetWindowLong(m_Buddy.hWnd, GWL_WNDPROC)
            SetProp m_Buddy.hWnd, "WndProc", lProc
            SetWindowLong m_Buddy.hWnd, GWL_WNDPROC, AddressOf picBuddyWindowProc
        End If
    
    ElseIf Not m_Buddy Is Nothing Then
        On Error Resume Next
        lProc = GetProp(m_Buddy.hWnd, "WndProc")
        If lProc Then
            SetWindowLong m_Buddy.hWnd, GWL_WNDPROC, lProc
            RemoveProp m_Buddy.hWnd, "WndProc"
            If Not m_Children Is Nothing Then
                For X = 1 To m_Children.Count
                    Set ctrlObj = m_Children.Item(X)
                    If Err Then
                        Err.Clear
                    Else
                        lProc = GetProp(ctrlObj.hWnd, "WndProc")
                        If lProc Then
                            SetWindowLong ctrlObj.hWnd, GWL_WNDPROC, lProc
                            RemoveProp ctrlObj.hWnd, "WndProc"
                            RemoveProp ctrlObj.hWnd, "ChildBtn"
                            RemoveProp ctrlObj.hWnd, "EraseBkg"
                            ctrlObj.Refresh
                        End If
                    End If
                Next
            End If
        End If
        Set m_Children = Nothing
        Set m_Buddy = Nothing
    End If
End Sub

Private Sub pvRefreshContainedControls()

    Dim X As Long, ctrlObj As Object
    If Not m_Children Is Nothing Then
        On Error Resume Next
        For X = 1 To m_Children.Count
            Set ctrlObj = m_Children.Item(X)
            If ctrlObj.Visible Then ctrlObj.Refresh
        Next
    End If

End Sub

Private Sub pvMoveMe()

    Dim cx As Long, cy As Long
    Dim sm As ScaleModeConstants, pm As ScaleModeConstants
    Dim meControl As Control
    
    m_NoRedraw = True
    Set meControl = pvTranslateToControl()
    
    On Error Resume Next
    Set meControl.Container = m_Buddy.Container     ' set our container same as buddy
    If Err Then
        ' error, what error? One can move a windowed control to another form at runtime
        ' but one cannot move a windowless control to another form at runtime
        ' another possibility: can't place non-alignable controls on MDI parent directly
        Set m_Buddy = Nothing
        Exit Sub
    End If
    
    sm = m_Buddy.ScaleMode
    pm = m_Buddy.Container.ScaleMode                ' frame's don't have scalemodes
    If Err Then                                     ' if err, assume scalemode is Twips
        Err.Clear
        pm = vbTwips
    End If
    cx = ScaleX(m_Buddy.ScaleWidth, sm, pm)         ' get scalewidth in relation to buddy's container's scalemode
    cy = ScaleY(m_Buddy.ScaleHeight, sm, pm)        ' get scaleheight in relation to buddy's container's scalemode
    m_BorderOffset.X = ScaleX((m_Buddy.Width - cx) \ 2, pm, vbPixels)   ' calc border width
    m_BorderOffset.Y = ScaleY((m_Buddy.Height - cy) \ 2, pm, vbPixels)   ' calc border height
    meControl.ZOrder                                 ' ensure ZOrder above all other windowless controls
    With m_Buddy
        meControl.Move .Left, .Top, .Width, .Height  ' now move into position
    End With
    Set meControl = Nothing
    
    m_NoRedraw = False                              ' remove flag & refresh
    UserControl.Refresh
End Sub

Private Function pvReposition() As Boolean

    Dim meControl As Control, bMove As Boolean
    Set meControl = pvTranslateToControl()
    If meControl.Left <> m_Buddy.Left Then
        bMove = True
    ElseIf meControl.Top <> m_Buddy.Top Then
        bMove = True
    ElseIf meControl.Width <> m_Buddy.Width Then
        bMove = True
    Else
        bMove = (meControl.Height <> m_Buddy.Height)
    End If
    If bMove Then
        pvMoveMe
    End If
    pvReposition = bMove

End Function

Private Function pvTranslateToControl() As Control

    Dim sControlName As String
    Dim meControl As Control, myForm As Object
    Dim X As Long, Index As Integer
    
    Set myForm = ParentControls(0)                  ' set instance of form/mdi
    sControlName = Ambient.DisplayName              ' get our control's name
    X = InStr(sControlName, "(")                    ' indexed?
    If X Then                                       ' if so, get the index
        Index = val(Mid$(sControlName, X + 1))      ' adjust control name & assign
        sControlName = Left$(sControlName, X - 1)
        Set meControl = myForm.Controls(sControlName)(Index)
    Else                                            ' assign
        Set meControl = myForm.Controls(sControlName)
    End If
    Set myForm = Nothing                            ' done with this
    Set pvTranslateToControl = meControl
    
End Function

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Opacity", m_Opacity, 0&
End Sub
