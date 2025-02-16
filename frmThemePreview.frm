VERSION 5.00
Begin VB.Form frmThemePreview 
   Caption         =   "Form1"
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7545
   LinkTopic       =   "Form1"
   ScaleHeight     =   5415
   ScaleWidth      =   7545
   StartUpPosition =   3  'Windows ±âº»°ª
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   4200
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      Height          =   2415
      Left            =   1080
      ScaleHeight     =   2355
      ScaleWidth      =   4635
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   4695
   End
End
Attribute VB_Name = "frmThemePreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const WS_VISIBLE As Long = &H10000000
Private Const WS_VSCROLL As Long = &H200000
Private Const WS_TABSTOP As Long = &H10000
Private Const WS_THICKFRAME As Long = &H40000
Private Const WS_MAXIMIZE As Long = &H1000000
Private Const WS_MAXIMIZEBOX As Long = &H10000
Private Const WS_MINIMIZE As Long = &H20000000
Private Const WS_MINIMIZEBOX As Long = &H20000
Private Const WS_SYSMENU As Long = &H80000
Private Const WS_BORDER As Long = &H800000
Private Const WS_CAPTION As Long = &HC00000                  '  WS_BORDER Or WS_DLGFRAME
Private Const WS_CHILD As Long = &H40000000
Private Const WS_CHILDWINDOW As Long = (WS_CHILD)
Private Const WS_CLIPCHILDREN As Long = &H2000000
Private Const WS_CLIPSIBLINGS As Long = &H4000000
Private Const WS_DISABLED As Long = &H8000000
Private Const WS_DLGFRAME As Long = &H400000
Private Const WS_EX_ACCEPTFILES As Long = &H10&
Private Const WS_EX_DLGMODALFRAME As Long = &H1&
Private Const WS_EX_NOPARENTNOTIFY As Long = &H4&
Private Const WS_EX_TOPMOST As Long = &H8&
Private Const WS_EX_TRANSPARENT As Long = &H20&
Private Const WS_EX_WINDOWEDGE As Long = &H100&
Private Const WS_EX_CLIENTEDGE As Long = &H200&
Private Const WS_EX_STATICEDGE As Long = &H20000
Private Const WS_GROUP As Long = &H20000
Private Const WS_HSCROLL As Long = &H100000
Private Const WS_ICONIC As Long = WS_MINIMIZE
Private Const WS_OVERLAPPED As Long = &H0&
Private Const WS_OVERLAPPEDWINDOW As Long = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)

Const WM_LBUTTONDOWN = &H201
Const WM_LBUTTONUP = &H202
Const SW_HIDE = 0
Const SW_NORMAL = 1
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Dim PrevhWnd As Long

Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetActiveWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long

Private Sub Command1_Click()
    Picture1.Visible = True
End Sub

Private Sub Form_Load()
    PrevhWnd = CreateWindowEx(WS_EX_CLIENTEDGE Or WS_EX_TOPMOST, "STATIC", "", WS_CHILD Or WS_VISIBLE Or WS_BORDER Or WS_OVERLAPPED Or WS_CAPTION Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX Or WS_SYSMENU Or WS_THICKFRAME, 10, 10, 100, 100, Picture1.hWnd, 0&, App.hInstance, 0&)
    'SetBkColor GetDC(PrevhWnd), 255&
    'Picture1.Refresh
    'SetWindowRgn hWnd, CreateRectRgn(0, 0, Screen.Width / Screen.TwipsPerPixelX + 300, Screen.Height / Screen.TwipsPerPixelY + 300), True
End Sub
