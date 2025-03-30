VERSION 5.00
Begin VB.UserControl TygemButton 
   ClientHeight    =   1305
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1800
   BeginProperty Font 
      Name            =   "±¼¸²"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   ScaleHeight     =   1305
   ScaleWidth      =   1800
   ToolboxBitmap   =   "TygemButton.ctx":0000
   Begin VB.Timer tmrMouse 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1200
      Top             =   0
   End
   Begin VB.Line Line11 
      BorderColor     =   &H0004D1FD&
      X1              =   30
      X2              =   45
      Y1              =   30
      Y2              =   45
   End
   Begin VB.Image imgOverlay 
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   495
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackStyle       =   0  'Åõ¸í
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1065
   End
   Begin VB.Shape pgFocusRect 
      BorderColor     =   &H00404040&
      BorderStyle     =   3  'Á¡
      Height          =   255
      Left            =   1200
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgIcon 
      Height          =   240
      Left            =   1200
      Stretch         =   -1  'True
      Top             =   600
      Width           =   240
   End
   Begin VB.Line Line10 
      X1              =   975
      X2              =   1005
      Y1              =   0
      Y2              =   30
   End
   Begin VB.Line Line9 
      X1              =   975
      X2              =   1025
      Y1              =   315
      Y2              =   270
   End
   Begin VB.Line Line8 
      X1              =   0
      X2              =   45
      Y1              =   285
      Y2              =   330
   End
   Begin VB.Line Line7 
      X1              =   1005
      X2              =   1005
      Y1              =   30
      Y2              =   285
   End
   Begin VB.Line Line6 
      X1              =   30
      X2              =   990
      Y1              =   315
      Y2              =   315
   End
   Begin VB.Line Line5 
      X1              =   0
      X2              =   30
      Y1              =   30
      Y2              =   0
   End
   Begin VB.Line Line4 
      X1              =   0
      X2              =   0
      Y1              =   30
      Y2              =   285
   End
   Begin VB.Line Line3 
      X1              =   30
      X2              =   990
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0004D1FD&
      X1              =   30
      X2              =   1005
      Y1              =   15
      Y2              =   15
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0004D1FD&
      X1              =   15
      X2              =   15
      Y1              =   30
      Y2              =   315
   End
   Begin VB.Image imgCenter 
      Height          =   285
      Index           =   2
      Left            =   0
      Picture         =   "TygemButton.ctx":0312
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image imgCenter 
      Height          =   285
      Index           =   1
      Left            =   0
      Picture         =   "TygemButton.ctx":0A65
      Stretch         =   -1  'True
      Top             =   480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image imgCenter 
      Height          =   285
      Index           =   0
      Left            =   30
      Picture         =   "TygemButton.ctx":118B
      Stretch         =   -1  'True
      Top             =   30
      Width           =   975
   End
End
Attribute VB_Name = "TygemButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Const m_def_Enabled = True
Dim m_Enabled As Boolean

Const m_def_Caption = ""
Dim m_Caption As String

Const m_def_BackColor = &H8000000F
Dim m_BackColor As OLE_COLOR

Const m_def_FontName = "±¼¸²"
Dim m_FontName As String

Const m_def_FontSize = 9
Dim m_FontSize As Integer

Const m_def_SplitLeft = False
Dim m_SplitLeft As Boolean

Const m_def_SplitRight = False
Dim m_SplitRight As Boolean

Dim m_Icon As IPictureDisp

Event Click()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private bHovering As Boolean
Private bMouseDown As Boolean

Private Sub MouseOut()
    If bMouseDown Then Exit Sub
    bHovering = False
    imgCenter(1).Visible = 0
    If m_Enabled Then
        Line1.BorderColor = IIf(m_SplitRight, RGB(252, 251, 164), RGB(253, 209, 4))
        Line2.BorderColor = RGB(253, 209, 4)
        Line11.BorderColor = RGB(253, 209, 4)
    Else
        Line1.BorderColor = IIf(m_SplitRight, RGB(215, 214, 215), RGB(198, 198, 198))
        Line2.BorderColor = RGB(198, 198, 198)
        Line11.BorderColor = RGB(198, 198, 198)
    End If
    If m_Enabled Then
        lblCaption.ForeColor = &H0&
    Else
        lblCaption.ForeColor = RGB(128, 128, 128)
    End If
    tmrMouse.Enabled = 0
End Sub

Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property

Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
    If m_Enabled Then
        lblCaption.ForeColor = &H0&
        imgCenter(2).Visible = 0
        Line1.BorderColor = IIf(m_SplitRight, RGB(252, 251, 164), RGB(253, 209, 4))
        Line2.BorderColor = RGB(253, 209, 4)
        Line11.BorderColor = RGB(253, 209, 4)
    Else
        lblCaption.ForeColor = RGB(128, 128, 128)
        imgCenter(2).Visible = -1
        Line1.BorderColor = IIf(m_SplitRight, RGB(215, 214, 215), RGB(198, 198, 198))
        Line2.BorderColor = RGB(198, 198, 198)
        Line11.BorderColor = RGB(198, 198, 198)
    End If
End Property

Property Get SplitLeft() As Boolean
    SplitLeft = m_SplitLeft
End Property

Property Let SplitLeft(ByVal New_SplitLeft As Boolean)
    m_SplitLeft = New_SplitLeft
    PropertyChanged "SplitLeft"
    SetSplitLeft
End Property

Property Get SplitRight() As Boolean
    SplitRight = m_SplitRight
End Property

Property Let SplitRight(ByVal New_SplitRight As Boolean)
    m_SplitRight = New_SplitRight
    PropertyChanged "SplitRight"
    SetSplitRight
End Property

Private Sub SetSplitLeft()
    Line9.Visible = Not m_SplitLeft
    Line10.Visible = Not m_SplitLeft
    If m_SplitLeft Then
        Line3.X2 = UserControl.Width
        Line6.X2 = UserControl.Width
        Line7.Y1 = 0
        Line7.Y2 = UserControl.Height
        Line2.X2 = UserControl.Width
    Else
        Line3.X2 = UserControl.Width - 30
        Line6.X2 = UserControl.Width - 30
        Line7.Y1 = 30
        Line7.Y2 = UserControl.Height - 45
        Line2.X2 = UserControl.Width - 30
    End If
End Sub

Private Sub SetSplitRight()
    Line4.Visible = Not m_SplitRight
    Line5.Visible = Not m_SplitRight
    Line8.Visible = Not m_SplitRight
    Line11.Visible = Not m_SplitRight
    Dim i%
    If m_SplitRight Then
        Line3.X1 = 0
        Line6.X1 = 0
        Line2.X1 = 0
        Line1.Y2 = UserControl.Height - 15
        For i = imgCenter.LBound To imgCenter.UBound
            imgCenter(i).Width = (UserControl.Width - 3 * Screen.TwipsPerPixelX) * (imgCenter(0).Picture.Width / 15)
            imgCenter(i).Left = 30 - (UserControl.Width - 3 * Screen.TwipsPerPixelX) * ((imgCenter(0).Picture.Width / 15) - 1)
        Next i
    Else
        Line3.X1 = 30
        Line6.X1 = 30
        Line2.X1 = 30
        Line1.Y2 = UserControl.Height - 30
        For i = imgCenter.LBound To imgCenter.UBound
            imgCenter(i).Width = UserControl.Width - 3 * Screen.TwipsPerPixelX
            imgCenter(i).Left = 30
        Next i
    End If
End Sub

Property Get Caption() As String
    Caption = m_Caption
End Property

Property Let Caption(ByVal New_Caption As String)
    m_Caption = New_Caption
    PropertyChanged "Caption"
    lblCaption.Caption = Trim$(m_Caption)
End Property

Property Get FontName() As String
    FontName = m_FontName
End Property

Property Let FontName(ByVal New_FontName As String)
    m_FontName = New_FontName
    PropertyChanged "FontName"
    SetCaptionFont
End Property

Private Sub SetCaptionFont()
    lblCaption.Font.Name = m_FontName
    lblCaption.Font.Bold = True
    lblCaption.Font.Italic = False
End Sub

Property Get FontSize() As String
    FontSize = m_FontSize
End Property

Property Let FontSize(ByVal New_FontSize As String)
    m_FontSize = New_FontSize
    PropertyChanged "FontSize"
    lblCaption.Font.Size = m_FontSize
End Property

Property Get ButtonIcon() As IPictureDisp
    Set ButtonIcon = m_Icon
End Property

Property Set ButtonIcon(ByVal New_Icon As IPictureDisp)
    Set m_Icon = New_Icon
    PropertyChanged "ButtonIcon"
    If Not New_Icon Is Nothing Then
        If New_Icon.Height < 240 Or UserControl.Width = 255 Then
            imgIcon.Stretch = False
            imgIcon.Top = UserControl.Height / 2 - New_Icon.Height / 2 + 30
        End If
    End If
    Set imgIcon.Picture = New_Icon
    If Not New_Icon Is Nothing Then _
        lblCaption.Caption = "  " & Trim$(lblCaption.Caption)
End Property

Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property

Private Sub imgOverlay_Click()
    If Not m_Enabled Then Exit Sub
    RaiseEvent Click
End Sub

Private Sub tmrMouse_Timer()
    Dim lpPos As POINTAPI
    Dim lhWnd As Long
    GetCursorPos lpPos
    lhWnd = WindowFromPoint(lpPos.X, lpPos.Y)
    If lhWnd <> UserControl.hWnd And bHovering = True Then MouseOut
End Sub

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    RaiseEvent Click
End Sub

Private Sub UserControl_GotFocus()
    pgFocusRect.Visible = True
End Sub

Private Sub UserControl_Initialize()
    Dim i%
    For i = 1 To imgCenter.UBound
        imgCenter(i).Top = imgCenter(0).Top
        imgCenter(i).Left = imgCenter(0).Left
    Next i
    bMouseDown = False
End Sub

Sub ShowAsPressed()
    If Not m_Enabled Then Exit Sub
    lblCaption.Left = 15
    lblCaption.Top = (UserControl.Height / 2 - lblCaption.Height / 2) + 20 + 15
    lblCaption.Tag = "mousedown"
    lblCaption.ForeColor = &H0&
    imgIcon.Left = IIf(UserControl.Width <= 495 And UserControl.Width > 255, UserControl.Width / 2 - imgIcon.Width / 2 + 10, 45)
    imgIcon.Top = UserControl.Height / 2 - imgIcon.Height / 2 + 20
End Sub

Sub ShowAsUnpressed()
    If Not m_Enabled Then Exit Sub
    lblCaption.Left = 0
    lblCaption.Top = UserControl.Height / 2 - lblCaption.Height / 2 + 15
    lblCaption.Tag = ""
    imgIcon.Left = IIf(UserControl.Width <= 495 And UserControl.Width > 255, UserControl.Width / 2 - imgIcon.Width / 2 - 10, 30)
    imgIcon.Top = UserControl.Height / 2 - imgIcon.Height / 2
End Sub

Private Sub imgOverlay_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
    If Not m_Enabled Then Exit Sub
    bMouseDown = True
    ShowAsPressed
End Sub
 
Private Sub imgOverlay_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
    If Not m_Enabled Then Exit Sub
    tmrMouse.Enabled = -1
    If bHovering = False Then
        bHovering = True
        imgCenter(1).Visible = -1
        Line1.BorderColor = IIf(m_SplitRight, RGB(207, 252, 162), RGB(179, 252, 53))
        Line2.BorderColor = RGB(179, 252, 53)
        Line11.BorderColor = RGB(179, 252, 53)
    End If
    If lblCaption.Tag <> "mousedown" Then lblCaption.ForeColor = 255
End Sub
 
Private Sub imgOverlay_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
    If Not m_Enabled Then Exit Sub
    bMouseDown = False
    ShowAsUnpressed
End Sub

Private Sub UserControl_InitProperties()
    m_Caption = Ambient.DisplayName
    m_BackColor = &H8000000F
    m_Enabled = True
    m_SplitLeft = False
    m_SplitRight = False
    UserControl.BackColor = &H8000000F
    lblCaption.Caption = m_Caption
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 32 And (Not bMouseDown) Then ShowAsPressed
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 32 And (Not bMouseDown) Then
        ShowAsUnpressed
        RaiseEvent Click
    End If
End Sub

Private Sub UserControl_LostFocus()
    pgFocusRect.Visible = False
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    Dim i%
    For i = imgCenter.LBound To imgCenter.UBound
        imgCenter(i).Width = UserControl.Width - 3 * Screen.TwipsPerPixelX
        imgCenter(i).Height = UserControl.Height - 3 * Screen.TwipsPerPixelY
    Next i
    imgOverlay.Width = UserControl.Width
    imgOverlay.Height = UserControl.Height
    Line1.Y2 = UserControl.Height - 30
    Line2.X2 = UserControl.Width - 30
    Line3.X2 = UserControl.Width - 30
    Line4.Y2 = UserControl.Height - 30
    Line6.Y1 = UserControl.Height - 15
    Line6.Y2 = UserControl.Height - 15
    Line6.X2 = UserControl.Width - 45
    Line7.X1 = UserControl.Width - 15
    Line7.X2 = UserControl.Width - 15
    Line7.Y2 = UserControl.Height - 45
    Line8.Y1 = UserControl.Height - 45
    Line8.Y2 = UserControl.Height
    Line9.Y1 = UserControl.Height
    Line9.Y2 = UserControl.Height - 60
    Line9.X1 = UserControl.Width - 60
    Line9.X2 = UserControl.Width
    Line10.X1 = UserControl.Width - 45
    Line10.X2 = UserControl.Width - 15
    SetSplitLeft
    SetSplitRight
    lblCaption.Top = UserControl.Height / 2 - lblCaption.Height / 2 + 15
    lblCaption.Width = UserControl.Width
    imgIcon.Top = UserControl.Height / 2 - imgIcon.Height / 2
    imgIcon.Left = IIf(UserControl.Width <= 495 And UserControl.Width > 255, UserControl.Width / 2 - imgIcon.Width / 2 - 10, 30)
    pgFocusRect.Top = 30
    pgFocusRect.Left = 30
    pgFocusRect.Width = UserControl.Width - 60
    pgFocusRect.Height = UserControl.Height - 60
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    m_Caption = PropBag.ReadProperty("Caption", m_def_Caption)
    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_FontName = PropBag.ReadProperty("FontName", m_def_FontName)
    m_FontSize = PropBag.ReadProperty("FontSize", m_def_FontSize)
    Set m_Icon = PropBag.ReadProperty("ButtonIcon", Nothing)
    m_SplitLeft = PropBag.ReadProperty("SplitLeft", m_def_SplitLeft)
    m_SplitRight = PropBag.ReadProperty("SplitRight", m_def_SplitRight)
    If m_Enabled Then
        lblCaption.ForeColor = &H0&
        imgCenter(2).Visible = 0
        Line1.BorderColor = IIf(m_SplitRight, RGB(252, 251, 164), RGB(253, 209, 4))
        Line2.BorderColor = RGB(253, 209, 4)
        Line11.BorderColor = RGB(253, 209, 4)
    Else
        lblCaption.ForeColor = RGB(128, 128, 128)
        imgCenter(2).Visible = -1
        Line1.BorderColor = IIf(m_SplitRight, RGB(215, 214, 215), RGB(198, 198, 198))
        Line2.BorderColor = RGB(198, 198, 198)
        Line11.BorderColor = RGB(198, 198, 198)
    End If
    lblCaption.Caption = Trim$(m_Caption)
    UserControl.BackColor = m_BackColor
    If Not m_Icon Is Nothing Then
        If (m_Icon.Width < 16 And m_Icon.Height < 16) Or UserControl.Width = 255 Then
            imgIcon.Stretch = False
            imgIcon.Top = UserControl.Height / 2 - m_Icon.Height / 2 + 30
        End If
        Set imgIcon.Picture = m_Icon
        lblCaption.Caption = "  " & lblCaption.Caption
    Else
        Set imgIcon.Picture = Nothing
    End If
    SetSplitLeft
    SetSplitRight
    SetCaptionFont
End Sub

Private Sub UserControl_Terminate()
    tmrMouse.Enabled = 0
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("Caption", m_Caption, m_def_Caption)
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("FontSize", m_FontSize, m_def_FontSize)
    On Error Resume Next
    Call PropBag.WriteProperty("ButtonIcon", m_Icon, Nothing)
    Call PropBag.WriteProperty("SplitLeft", m_SplitLeft, m_def_SplitLeft)
    Call PropBag.WriteProperty("SplitRight", m_SplitRight, m_def_SplitRight)
End Sub

