VERSION 5.00
Begin VB.UserControl TygemButton 
   ClientHeight    =   1590
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1665
   BeginProperty Font 
      Name            =   "±¼¸²"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   1590
   ScaleWidth      =   1665
   ToolboxBitmap   =   "TygemButton.ctx":0000
   Begin VB.Timer tmrMouse 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1200
      Top             =   0
   End
   Begin VB.Image imgOverlay 
      Height          =   390
      Left            =   0
      Top             =   0
      Width           =   1080
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
      Width           =   1080
   End
   Begin VB.Image imgCenter 
      Height          =   285
      Index           =   2
      Left            =   60
      Picture         =   "TygemButton.ctx":0312
      Stretch         =   -1  'True
      Top             =   1140
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image imgRight 
      Height          =   285
      Index           =   2
      Left            =   1035
      Picture         =   "TygemButton.ctx":0A65
      Stretch         =   -1  'True
      Top             =   1140
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Image imgBottomRight 
      Height          =   45
      Index           =   2
      Left            =   1035
      Picture         =   "TygemButton.ctx":0B24
      Top             =   1425
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Image imgBottom 
      Height          =   45
      Index           =   2
      Left            =   60
      Picture         =   "TygemButton.ctx":0B64
      Stretch         =   -1  'True
      Top             =   1425
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image imgBottomLeft 
      Height          =   45
      Index           =   2
      Left            =   0
      Picture         =   "TygemButton.ctx":0C05
      Top             =   1425
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.Image imgLeft 
      Height          =   285
      Index           =   2
      Left            =   0
      Picture         =   "TygemButton.ctx":0C47
      Stretch         =   -1  'True
      Top             =   1140
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.Image imgTopRight 
      Height          =   60
      Index           =   2
      Left            =   1035
      Picture         =   "TygemButton.ctx":0D72
      Top             =   1080
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Image imgTop 
      Height          =   60
      Index           =   2
      Left            =   60
      Picture         =   "TygemButton.ctx":0DC1
      Stretch         =   -1  'True
      Top             =   1080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image imgTopLeft 
      Height          =   60
      Index           =   2
      Left            =   0
      Picture         =   "TygemButton.ctx":101B
      Top             =   1080
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.Image imgTopLeft 
      Height          =   60
      Index           =   1
      Left            =   0
      Picture         =   "TygemButton.ctx":106B
      Top             =   540
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.Image imgTop 
      Height          =   60
      Index           =   1
      Left            =   60
      Picture         =   "TygemButton.ctx":10BB
      Stretch         =   -1  'True
      Top             =   540
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image imgTopRight 
      Height          =   60
      Index           =   1
      Left            =   1035
      Picture         =   "TygemButton.ctx":1324
      Top             =   540
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Image imgLeft 
      Height          =   285
      Index           =   1
      Left            =   0
      Picture         =   "TygemButton.ctx":1365
      Stretch         =   -1  'True
      Top             =   600
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.Image imgBottomLeft 
      Height          =   45
      Index           =   1
      Left            =   0
      Picture         =   "TygemButton.ctx":1420
      Top             =   885
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.Image imgBottom 
      Height          =   45
      Index           =   1
      Left            =   60
      Picture         =   "TygemButton.ctx":1461
      Stretch         =   -1  'True
      Top             =   885
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image imgBottomRight 
      Height          =   45
      Index           =   1
      Left            =   1035
      Picture         =   "TygemButton.ctx":15BD
      Top             =   885
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Image imgRight 
      Height          =   285
      Index           =   1
      Left            =   1035
      Picture         =   "TygemButton.ctx":15FD
      Stretch         =   -1  'True
      Top             =   600
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Image imgCenter 
      Height          =   285
      Index           =   1
      Left            =   60
      Picture         =   "TygemButton.ctx":1658
      Stretch         =   -1  'True
      Top             =   600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image imgCenter 
      Height          =   285
      Index           =   0
      Left            =   60
      Picture         =   "TygemButton.ctx":1D7E
      Stretch         =   -1  'True
      Top             =   60
      Width           =   975
   End
   Begin VB.Image imgRight 
      Height          =   285
      Index           =   0
      Left            =   1035
      Picture         =   "TygemButton.ctx":2468
      Stretch         =   -1  'True
      Top             =   60
      Width           =   45
   End
   Begin VB.Image imgBottomRight 
      Height          =   45
      Index           =   0
      Left            =   1035
      Picture         =   "TygemButton.ctx":24C1
      Top             =   345
      Width           =   45
   End
   Begin VB.Image imgBottom 
      Height          =   45
      Index           =   0
      Left            =   60
      Picture         =   "TygemButton.ctx":2501
      Stretch         =   -1  'True
      Top             =   345
      Width           =   975
   End
   Begin VB.Image imgBottomLeft 
      Height          =   45
      Index           =   0
      Left            =   0
      Picture         =   "TygemButton.ctx":2656
      Top             =   345
      Width           =   60
   End
   Begin VB.Image imgLeft 
      Height          =   285
      Index           =   0
      Left            =   0
      Picture         =   "TygemButton.ctx":2697
      Stretch         =   -1  'True
      Top             =   60
      Width           =   60
   End
   Begin VB.Image imgTopRight 
      Height          =   60
      Index           =   0
      Left            =   1035
      Picture         =   "TygemButton.ctx":2751
      Top             =   0
      Width           =   45
   End
   Begin VB.Image imgTop 
      Height          =   60
      Index           =   0
      Left            =   60
      Picture         =   "TygemButton.ctx":2792
      Stretch         =   -1  'True
      Top             =   0
      Width           =   975
   End
   Begin VB.Image imgTopLeft 
      Height          =   60
      Index           =   0
      Left            =   0
      Picture         =   "TygemButton.ctx":29EA
      Top             =   0
      Width           =   60
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
    imgTop(1).Visible = 0
    imgLeft(1).Visible = 0
    imgRight(1).Visible = 0
    imgBottom(1).Visible = 0
    imgTopLeft(1).Visible = 0
    imgTopRight(1).Visible = 0
    imgBottomLeft(1).Visible = 0
    imgBottomRight(1).Visible = 0
    imgCenter(1).Visible = 0
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
        imgTop(2).Visible = 0
        imgLeft(2).Visible = 0
        imgRight(2).Visible = 0
        imgBottom(2).Visible = 0
        imgTopLeft(2).Visible = 0
        imgTopRight(2).Visible = 0
        imgBottomLeft(2).Visible = 0
        imgBottomRight(2).Visible = 0
        imgCenter(2).Visible = 0
    Else
        lblCaption.ForeColor = RGB(128, 128, 128)
        imgTop(2).Visible = -1
        imgLeft(2).Visible = -1
        imgRight(2).Visible = -1
        imgBottom(2).Visible = -1
        imgTopLeft(2).Visible = -1
        imgTopRight(2).Visible = -1
        imgBottomLeft(2).Visible = -1
        imgBottomRight(2).Visible = -1
        imgCenter(2).Visible = -1
    End If
End Property

Property Get Caption() As String
    Caption = m_Caption
End Property

Property Let Caption(ByVal New_Caption As String)
    m_Caption = New_Caption
    PropertyChanged "Caption"
    lblCaption.Caption = m_Caption
End Property

Property Get FontName() As String
    FontName = m_FontName
End Property

Property Let FontName(ByVal New_FontName As String)
    m_FontName = New_FontName
    PropertyChanged "FontName"
    lblCaption.Font.Name = m_FontName
End Property

Property Get FontSize() As String
    FontSize = m_FontSize
End Property

Property Let FontSize(ByVal New_FontSize As String)
    m_FontSize = New_FontSize
    PropertyChanged "FontSize"
    lblCaption.Font.Size = m_FontSize
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
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

Private Sub UserControl_Initialize()
    Dim i%
    For i = 1 To imgCenter.UBound
        imgTop(i).Top = imgTop(0).Top
        imgTop(i).Left = imgTop(0).Left
        
        imgLeft(i).Top = imgLeft(0).Top
        imgLeft(i).Left = imgLeft(0).Left
        
        imgRight(i).Top = imgRight(0).Top
        imgRight(i).Left = imgRight(0).Left
        
        imgBottom(i).Top = imgBottom(0).Top
        imgBottom(i).Left = imgBottom(0).Left
        
        imgTopLeft(i).Top = imgTopLeft(0).Top
        imgTopLeft(i).Left = imgTopLeft(0).Left
        
        imgTopRight(i).Top = imgTopRight(0).Top
        imgTopRight(i).Left = imgTopRight(0).Left
        
        imgBottomLeft(i).Top = imgBottomLeft(0).Top
        imgBottomLeft(i).Left = imgBottomLeft(0).Left
        
        imgBottomRight(i).Top = imgBottomRight(0).Top
        imgBottomRight(i).Left = imgBottomRight(0).Left
        
        imgCenter(i).Top = imgCenter(0).Top
        imgCenter(i).Left = imgCenter(0).Left
    Next i
    bMouseDown = False
End Sub

Private Sub imgOverlay_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
    If Not m_Enabled Then Exit Sub
    lblCaption.Left = 15
    lblCaption.Top = (UserControl.Height / 2 - lblCaption.Height / 2) + 20 + 15
    lblCaption.Tag = "mousedown"
    lblCaption.ForeColor = &H0&
    bMouseDown = True
End Sub
 
Private Sub imgOverlay_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
    If Not m_Enabled Then Exit Sub
    tmrMouse.Enabled = -1
    If bHovering = False Then
        bHovering = True
        imgTop(1).Visible = -1
        imgLeft(1).Visible = -1
        imgRight(1).Visible = -1
        imgBottom(1).Visible = -1
        imgTopLeft(1).Visible = -1
        imgTopRight(1).Visible = -1
        imgBottomLeft(1).Visible = -1
        imgBottomRight(1).Visible = -1
        imgCenter(1).Visible = -1
    End If
    If lblCaption.Tag <> "mousedown" Then lblCaption.ForeColor = 255
End Sub
 
Private Sub imgOverlay_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
    If Not m_Enabled Then Exit Sub
    lblCaption.Left = 0
    lblCaption.Top = UserControl.Height / 2 - lblCaption.Height / 2 + 15
    lblCaption.Tag = ""
    bMouseDown = False
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    Dim i%
    For i = imgCenter.LBound To imgCenter.UBound
        imgTop(i).Width = UserControl.Width - (imgTopLeft(i).Width + imgTopRight(i).Width)
        imgTopRight(i).Left = imgTopLeft(i).Width + imgTop(i).Width
        imgRight(i).Left = imgTopRight(i).Left
        imgLeft(i).Height = UserControl.Height - imgTopLeft(i).Height - imgBottomLeft(i).Height
        imgRight(i).Height = imgLeft(i).Height
        imgBottom(i).Width = imgTop(i).Width
        imgBottom(i).Top = UserControl.Height - imgBottom(i).Height
        imgBottomLeft(i).Top = imgBottom(i).Top
        imgBottomRight(i).Top = imgBottomLeft(i).Top
        imgBottomRight(i).Left = UserControl.Width - imgBottomRight(i).Width
        imgCenter(i).Width = UserControl.Width - imgLeft(i).Width - imgRight(i).Width
        imgCenter(i).Height = UserControl.Height - imgTop(i).Height - imgBottom(i).Height
    Next i
    imgOverlay.Width = UserControl.Width
    imgOverlay.Height = UserControl.Height
    lblCaption.Top = UserControl.Height / 2 - lblCaption.Height / 2 + 15
    lblCaption.Width = UserControl.Width
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    m_Caption = PropBag.ReadProperty("Caption", m_def_Caption)
    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_FontName = PropBag.ReadProperty("FontName", m_def_FontName)
    m_FontSize = PropBag.ReadProperty("FontSize", m_def_FontSize)
    If m_Enabled Then
        lblCaption.ForeColor = &H0&
        imgTop(2).Visible = 0
        imgLeft(2).Visible = 0
        imgRight(2).Visible = 0
        imgBottom(2).Visible = 0
        imgTopLeft(2).Visible = 0
        imgTopRight(2).Visible = 0
        imgBottomLeft(2).Visible = 0
        imgBottomRight(2).Visible = 0
        imgCenter(2).Visible = 0
    Else
        lblCaption.ForeColor = RGB(128, 128, 128)
        imgTop(2).Visible = -1
        imgLeft(2).Visible = -1
        imgRight(2).Visible = -1
        imgBottom(2).Visible = -1
        imgTopLeft(2).Visible = -1
        imgTopRight(2).Visible = -1
        imgBottomLeft(2).Visible = -1
        imgBottomRight(2).Visible = -1
        imgCenter(2).Visible = -1
    End If
    lblCaption.Caption = m_Caption
    UserControl.BackColor = m_BackColor
End Sub

Private Sub UserControl_Terminate()
    tmrMouse.Enabled = 0
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("Caption", m_Caption, m_def_Caption)
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("FontSize", m_FontSize, m_def_FontSize)
End Sub

