VERSION 5.00
Begin VB.UserControl CheckBoxEx 
   ClientHeight    =   630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1725
   ScaleHeight     =   630
   ScaleWidth      =   1725
   Begin VB.CheckBox chkCheckBox 
      Caption         =   "Check1"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1575
   End
End
Attribute VB_Name = "CheckBoxEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const m_def_Transparent As Boolean = False
Dim m_Transparent As Boolean

Const m_def_VisualStyles As Boolean = True
Dim m_VisualStyles As Boolean

Const m_def_BackColor As Long = &H8000000F
Dim m_BackColor As OLE_COLOR

Const m_def_ForeColor As Long = &H80000012

Private CheckBoxTransparentBrush As Long
Private DesignMode As Boolean

Event Click()

Implements IBSSubclass

Private Function IBSSubclass_MsgResponse(ByVal hWnd As Long, ByVal uMsg As Long) As EMsgResponse
    IBSSubclass_MsgResponse = emrConsume
End Function

Private Sub IBSSubclass_UnsubclassIt()
    If Not DesignMode Then
        DetachMessage Me, UserControl.hWnd, WM_CTLCOLORSTATIC
        DetachMessage Me, UserControl.hWnd, WM_CTLCOLORBTN
    End If
End Sub

Private Function IBSSubclass_WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, wParam As Long, lParam As Long, bConsume As Boolean) As Long
    Select Case uMsg
        Case WM_CTLCOLORSTATIC, WM_CTLCOLORBTN
            If m_Transparent Then
                SetBkMode wParam, 1&
                Dim hDCBmp As Long
                Dim hBmp As Long, hBmpOld As Long
                With chkCheckBox
                    If CheckBoxTransparentBrush = 0& Then
                        hDCBmp = CreateCompatibleDC(wParam)
                        If hDCBmp <> 0& Then
                            hBmp = CreateCompatibleBitmap(wParam, .Width / Screen.TwipsPerPixelX, .Height / Screen.TwipsPerPixelY)
                            If hBmp <> 0& Then
                                Dim hWndParent As Long
                                hWndParent = GetParent(UserControl.hWnd)
                                hBmpOld = SelectObject(hDCBmp, hBmp)
                                Dim WndRect As RECT, P As POINTAPI
                                GetWindowRect .hWnd, WndRect
                                MapWindowPoints hWnd_DESKTOP, hWndParent, WndRect, 2&
                                P.X = WndRect.Left
                                P.Y = WndRect.Top
                                SetViewportOrgEx hDCBmp, -P.X, -P.Y, P
                                SendMessage hWndParent, WM_PAINT, hDCBmp, ByVal 0&
                                SetViewportOrgEx hDCBmp, P.X, P.Y, P
                                CheckBoxTransparentBrush = CreatePatternBrush(hBmp)
                                SelectObject hDCBmp, hBmpOld
                                DeleteObject hBmp
                            End If
                            DeleteDC hDCBmp
                        End If
                    End If
                End With
                If CheckBoxTransparentBrush <> 0& Then
                    IBSSubclass_WindowProc = CheckBoxTransparentBrush
                    Exit Function
                End If
            Else
                IBSSubclass_WindowProc = CreateSolidBrush(WinColor(chkCheckBox.BackColor))
                Exit Function
            End If
    End Select
    IBSSubclass_WindowProc = CallOldWindowProc(hWnd, uMsg, wParam, lParam)
End Function

Property Get Value() As CheckBoxConstants
    Value = chkCheckBox.Value
End Property

Property Let Value(ByVal New_Value As CheckBoxConstants)
    chkCheckBox.Value = New_Value
    PropertyChanged "Value"
End Property

Property Get Caption() As String
    Caption = chkCheckBox.Caption
End Property

Property Let Caption(ByVal New_Caption As String)
    chkCheckBox.Caption = New_Caption
    PropertyChanged "Caption"
End Property

Property Get BackColor() As OLE_COLOR
    BackColor = chkCheckBox.BackColor
End Property

Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    chkCheckBox.BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property

Property Get ForeColor() As OLE_COLOR
    ForeColor = chkCheckBox.ForeColor
End Property

Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    chkCheckBox.ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
End Property

Property Get Enabled() As Boolean
    Enabled = chkCheckBox.Enabled
End Property

Property Let Enabled(ByVal New_Enabled As Boolean)
    chkCheckBox.Enabled = New_Enabled
    PropertyChanged "Enabled"
End Property

Property Get Font() As StdFont
    Set Font = chkCheckBox.Font
End Property

Property Let Font(New_Font As StdFont)
    Set chkCheckBox.Font = New_Font
End Property

Property Get VisualStyles() As Boolean
    VisualStyles = m_VisualStyles
End Property

Property Let VisualStyles(ByVal New_VisualStyles As Boolean)
    m_VisualStyles = New_VisualStyles
    PropertyChanged "VisualStyles"
    SetVisualStyles
End Property

Private Sub SetVisualStyles()
    If m_VisualStyles Then ActivateVisualStyles chkCheckBox.hWnd Else RemoveVisualStyles chkCheckBox.hWnd
End Sub

Property Get Transparent() As Boolean
    Transparent = m_Transparent
End Property

Property Let Transparent(ByVal New_Transparent As Boolean)
    m_Transparent = New_Transparent
    PropertyChanged "Transparent"
    Me.Refresh
End Property

Property Get hWnd() As Long
    hWnd = chkCheckBox.hWnd
End Property

Private Sub chkCheckBox_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_InitProperties()
    DesignMode = Not Ambient.UserMode
    chkCheckBox.Caption = Ambient.DisplayName
    m_VisualStyles = m_def_VisualStyles
    m_Transparent = m_def_Transparent
    Set chkCheckBox.Font = Ambient.Font
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    chkCheckBox.Value = PropBag.ReadProperty("Value", vbUnchecked)
    chkCheckBox.BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    chkCheckBox.ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    chkCheckBox.Enabled = PropBag.ReadProperty("Enabled", True)
    chkCheckBox.Caption = PropBag.ReadProperty("Caption", Ambient.DisplayName)
    m_VisualStyles = PropBag.ReadProperty("VisualStyles", m_def_VisualStyles)
    m_Transparent = PropBag.ReadProperty("Transparent", m_def_Transparent)
    Set chkCheckBox.Font = PropBag.ReadProperty("Font", Ambient.Font)
    SetVisualStyles
    If Not DesignMode Then
        AttachMessage Me, UserControl.hWnd, WM_CTLCOLORSTATIC
        AttachMessage Me, UserControl.hWnd, WM_CTLCOLORBTN
    End If
End Sub

Private Sub UserControl_Resize()
    chkCheckBox.Width = UserControl.Width
    chkCheckBox.Height = UserControl.Height
End Sub

Private Sub UserControl_Terminate()
    If CheckBoxTransparentBrush <> 0& Then DeleteObject CheckBoxTransparentBrush
    IBSSubclass_UnsubclassIt
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Caption", chkCheckBox.Caption, Ambient.DisplayName
    PropBag.WriteProperty "Value", chkCheckBox.Value, vbUnchecked
    PropBag.WriteProperty "BackColor", chkCheckBox.BackColor, m_def_BackColor
    PropBag.WriteProperty "ForeColor", chkCheckBox.ForeColor, m_def_ForeColor
    PropBag.WriteProperty "Enabled", chkCheckBox.Enabled, True
    PropBag.WriteProperty "VisualStyles", m_VisualStyles, m_def_VisualStyles
    PropBag.WriteProperty "Transparent", m_Transparent, m_def_Transparent
    PropBag.WriteProperty "Font", chkCheckBox.Font, Ambient.Font
End Sub

Sub Refresh()
    If CheckBoxTransparentBrush <> 0& Then
        DeleteObject CheckBoxTransparentBrush
        CheckBoxTransparentBrush = 0&
    End If
    UserControl.Refresh
    RedrawWindow UserControl.hWnd, 0&, 0&, RDW_UPDATENOW Or RDW_INVALIDATE Or RDW_ERASE Or RDW_ALLCHILDREN
End Sub
