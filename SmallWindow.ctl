VERSION 5.00
Begin VB.UserControl SmallWindow 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "SmallWindow.ctx":0000
End
Attribute VB_Name = "SmallWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const m_def_Enabled As Boolean = True
Dim m_Enabled As Boolean

Const m_def_Caption = ""
Dim m_Caption As String

Const m_def_BackColor As Long = &H8000000F
Dim m_BackColor As OLE_COLOR

Const m_def_MaximizeBox As Boolean = True
Dim m_MaximizeBox As Boolean

Const m_def_MinimizeBox As Boolean = True
Dim m_MinimizeBox As Boolean

Const m_def_ThickFrame As Boolean = True
Dim m_ThickFrame As Boolean

Const m_def_ControlBox As Boolean = True
Dim m_ControlBox As Boolean

Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property

Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    UserControl.Enabled = m_Enabled
    PropertyChanged "Enabled"
End Property

Property Get Caption() As String
    Caption = m_Caption
End Property

Property Let Caption(ByVal New_Caption As String)
    m_Caption = New_Caption
    PropertyChanged "Caption"
    SetCaption
End Property

Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property

Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    UserControl.BackColor = m_BackColor
    PropertyChanged "BackColor"
End Property

Property Get MaximizeBox() As Boolean
    MaximizeBox = m_MaximizeBox
End Property

Property Let MaximizeBox(ByVal New_MaximizeBox As Boolean)
    m_MaximizeBox = New_MaximizeBox
    PropertyChanged "MaximizeBox"
    SetMaximizeBox
End Property

Property Get MinimizeBox() As Boolean
    MinimizeBox = m_MinimizeBox
End Property

Property Let MinimizeBox(ByVal New_MinimizeBox As Boolean)
    m_MinimizeBox = New_MinimizeBox
    PropertyChanged "MinimizeBox"
    SetMinimizeBox
End Property

Property Get ThickFrame() As Boolean
    ThickFrame = m_ThickFrame
End Property

Property Let ThickFrame(ByVal New_ThickFrame As Boolean)
    m_ThickFrame = New_ThickFrame
    PropertyChanged "ThickFrame"
    SetThickFrame
End Property

Property Get ControlBox() As Boolean
    ControlBox = m_ControlBox
End Property

Property Let ControlBox(ByVal New_ControlBox As Boolean)
    m_ControlBox = New_ControlBox
    PropertyChanged "ControlBox"
    SetControlBox
End Property

Property Get hDC() As Long
    hDC = UserControl.hDC
End Property

Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Private Sub SetCaption()
    SetWindowText UserControl.hWnd, m_Caption
End Sub

Private Sub SetMaximizeBox()
    If m_MaximizeBox Then
        SetWindowLong UserControl.hWnd, GWL_STYLE, GetWindowLong(UserControl.hWnd, GWL_STYLE) Or WS_MAXIMIZEBOX
    Else
        SetWindowLong UserControl.hWnd, GWL_STYLE, GetWindowLong(UserControl.hWnd, GWL_STYLE) And (Not WS_MAXIMIZEBOX)
    End If
    UserControl.Refresh
End Sub

Private Sub SetMinimizeBox()
    If m_MinimizeBox Then
        SetWindowLong UserControl.hWnd, GWL_STYLE, GetWindowLong(UserControl.hWnd, GWL_STYLE) Or WS_MINIMIZEBOX
    Else
        SetWindowLong UserControl.hWnd, GWL_STYLE, GetWindowLong(UserControl.hWnd, GWL_STYLE) And (Not WS_MINIMIZEBOX)
    End If
    UserControl.Refresh
End Sub

Private Sub SetThickFrame()
    If m_ThickFrame Then
        SetWindowLong UserControl.hWnd, GWL_STYLE, GetWindowLong(UserControl.hWnd, GWL_STYLE) Or WS_THICKFRAME
    Else
        SetWindowLong UserControl.hWnd, GWL_STYLE, GetWindowLong(UserControl.hWnd, GWL_STYLE) And (Not WS_THICKFRAME)
    End If
    UserControl.Refresh
End Sub

Private Sub SetControlBox()
    If m_ControlBox Then
        SetWindowLong UserControl.hWnd, GWL_STYLE, GetWindowLong(UserControl.hWnd, GWL_STYLE) Or WS_SYSMENU
    Else
        SetWindowLong UserControl.hWnd, GWL_STYLE, GetWindowLong(UserControl.hWnd, GWL_STYLE) And (Not WS_SYSMENU)
    End If
    UserControl.Refresh
End Sub

Private Sub UserControl_Initialize()
    SetWindowLong UserControl.hWnd, GWL_STYLE, GetWindowLong(UserControl.hWnd, GWL_STYLE) Or WS_BORDER Or WS_OVERLAPPED Or WS_CAPTION
    SetMaximizeBox
    SetMinimizeBox
    SetThickFrame
    SetControlBox
End Sub

Private Sub UserControl_InitProperties()
    m_Enabled = m_def_Enabled
    UserControl.Enabled = m_Enabled
    m_BackColor = m_def_BackColor
    UserControl.BackColor = m_BackColor
    m_Caption = Ambient.DisplayName
    SetCaption
    m_MaximizeBox = m_def_MaximizeBox
    SetMaximizeBox
    m_MinimizeBox = m_def_MinimizeBox
    SetMinimizeBox
    m_ThickFrame = m_def_ThickFrame
    SetThickFrame
    m_ControlBox = m_def_ControlBox
    SetControlBox
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    UserControl.Enabled = m_Enabled
    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    UserControl.BackColor = m_BackColor
    m_Caption = PropBag.ReadProperty("Caption", m_def_Caption)
    SetCaption
    m_MaximizeBox = PropBag.ReadProperty("MaximizeBox", m_def_MaximizeBox)
    SetMaximizeBox
    m_MinimizeBox = PropBag.ReadProperty("MinimizeBox", m_def_MinimizeBox)
    SetMinimizeBox
    m_ThickFrame = PropBag.ReadProperty("ThickFrame", m_def_ThickFrame)
    SetThickFrame
    m_ControlBox = PropBag.ReadProperty("ControlBox", m_def_ControlBox)
    SetControlBox
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Enabled", m_Enabled, m_def_Enabled
    PropBag.WriteProperty "BackColor", m_BackColor, m_def_BackColor
    PropBag.WriteProperty "Caption", m_Caption, m_def_Caption
    PropBag.WriteProperty "MaximizeBox", m_MaximizeBox, m_def_MaximizeBox
    PropBag.WriteProperty "MinimizeBox", m_MinimizeBox, m_def_MinimizeBox
    PropBag.WriteProperty "ThickFrame", m_ThickFrame, m_def_ThickFrame
    PropBag.WriteProperty "ControlBox", m_ControlBox, m_def_ControlBox
End Sub

Sub Refresh()
    UserControl.Refresh
End Sub
