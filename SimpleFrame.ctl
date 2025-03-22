VERSION 5.00
Begin VB.UserControl SimpleFrame 
   ClientHeight    =   225
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   225
   ScaleWidth      =   4800
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   1215
      X2              =   4800
      Y1              =   90
      Y2              =   90
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   1200
      X2              =   4785
      Y1              =   75
      Y2              =   75
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  '≈ı∏Ì
      Caption         =   "SimpleFrame"
      Height          =   180
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1065
   End
End
Attribute VB_Name = "SimpleFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_Caption As String

Dim m_Font As StdFont

Const m_def_ForeColor = &H80000012
Dim m_ForeColor As OLE_COLOR

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Const WM_PAINT As Long = &HF
Private Declare Function MapWindowPoints Lib "user32" (ByVal hWndFrom As Long, ByVal hWndTo As Long, ByRef lppt As Any, ByVal cPoints As Long) As Long
Private Declare Function SetViewportOrgEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByRef lpPoint As POINTAPI) As Long

Property Get Font() As StdFont
    Set Font = lblCaption.Font
End Property

Property Set Font(New_Font As StdFont)
    Set m_Font = New_Font
    SetFont
    PropertyChanged "Font"
End Property

Private Sub SetFont()
    If Not m_Font Is Nothing Then Set lblCaption.Font = m_Font
End Sub

Property Get ForeColor() As OLE_COLOR
    ForeColor = lblCaption.ForeColor
End Property

Property Let ForeColor(New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    SetForeColor
    PropertyChanged "ForeColor"
End Property

Private Sub SetForeColor()
    lblCaption.ForeColor = m_ForeColor
End Sub

Property Get Caption() As String
    Caption = lblCaption.Caption
End Property

Property Let Caption(New_Caption As String)
    m_Caption = New_Caption
    SetCaption
    PropertyChanged "Caption"
End Property

Private Sub SetCaption()
    lblCaption.Caption = m_Caption
    Line1.X1 = lblCaption.Width + 120
    Line2.X1 = lblCaption.Width + 135
End Sub

Private Sub UserControl_InitProperties()
    m_Caption = Ambient.DisplayName
    m_ForeColor = m_def_ForeColor
    Set m_Font = Ambient.Font
    lblCaption.Caption = m_Caption
    Set lblCaption.Font = m_Font
    Set UserControl.Font = m_Font
    UserControl_Resize
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Caption = PropBag.ReadProperty("Caption", Ambient.DisplayName)
    Set m_Font = PropBag.ReadProperty("Font", Nothing)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    
    SetCaption
    SetFont
    SetForeColor
End Sub

Private Sub UserControl_Resize()
    Line1.X2 = UserControl.Width - Screen.TwipsPerPixelX
    Line2.X2 = UserControl.Width
    UserControl.Height = 255
    SetCaption
    DrawFrame
End Sub

Private Sub DrawFrame()
    Dim WndRect As RECT, P As POINTAPI
    GetWindowRect UserControl.hWnd, WndRect
    MapWindowPoints 0&, GetParent(UserControl.hWnd), WndRect, 2
    P.X = WndRect.Left
    P.Y = WndRect.Top
    SetViewportOrgEx UserControl.hDC, -P.X, -P.Y, P
    SendMessage GetParent(UserControl.hWnd), WM_PAINT, UserControl.hDC, ByVal 0&
    SetViewportOrgEx UserControl.hDC, P.X, P.Y, P
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Caption", m_Caption, Ambient.DisplayName
    PropBag.WriteProperty "Font", m_Font, Nothing
    PropBag.WriteProperty "ForeColor", m_ForeColor, m_def_ForeColor
End Sub
