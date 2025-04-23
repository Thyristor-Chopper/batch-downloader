VERSION 5.00
Begin VB.Form frmTransparentTest 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   Picture         =   "frmTransparentTest.frx":0000
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows ±âº»°ª
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   255
      Left            =   2760
      TabIndex        =   3
      Top             =   600
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   1935
      Left            =   480
      Picture         =   "frmTransparentTest.frx":3542
      ScaleHeight     =   1875
      ScaleWidth      =   3795
      TabIndex        =   0
      Top             =   960
      Width           =   3855
      Begin prjDownloadBooster.CheckBoxEx CheckBoxEx1 
         Height          =   255
         Left            =   600
         TabIndex        =   8
         Top             =   600
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   450
         Transparent     =   -1  'True
      End
      Begin prjDownloadBooster.FrameW FrameW1 
         Height          =   975
         Left            =   960
         TabIndex        =   5
         Top             =   720
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   1720
         Caption         =   "FrameW1"
         Transparent     =   -1  'True
         Begin VB.CheckBox Check3 
            Caption         =   "Check3"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   600
            Width           =   1095
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Option2"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Check2"
         Height          =   255
         Left            =   1320
         TabIndex        =   4
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   1560
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "frmTransparentTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long

Implements IBSSubclass

Dim CheckBoxTransparentBrush As Long

Private Sub Command1_Click()
    If CheckBoxTransparentBrush Then
        DeleteObject CheckBoxTransparentBrush
        CheckBoxTransparentBrush = 0&
    End If
    'RedrawWindow Check1.hWnd, 0&, 0&, RDW_UPDATENOW Or RDW_INVALIDATE Or RDW_ERASE
    'Dim ClientRect As RECT
    'GetClientRect FrameW1.hWnd, ClientRect
    'BitBlt GetDC(Check3.hWnd), 0, 0, ClientRect.Right - ClientRect.Left, ClientRect.Bottom - ClientRect.Top, GetDC(FrameW1.hWnd), 0, 0, vbSrcCopy
End Sub

Private Sub Form_Load()
    'AttachMessage Me, Me.hWnd, WM_CTLCOLORSTATIC
    'AttachMessage Me, Me.hWnd, WM_CTLCOLORBTN
    AttachMessage Me, Picture1.hWnd, WM_CTLCOLORSTATIC
    AttachMessage Me, Picture1.hWnd, WM_CTLCOLORBTN
End Sub

Private Sub Form_Unload(Cancel As Integer)
    IBSSubclass_UnsubclassIt
    DeleteObject CheckBoxTransparentBrush
End Sub

Private Function IBSSubclass_MsgResponse(ByVal hWnd As Long, ByVal uMsg As Long) As EMsgResponse
    IBSSubclass_MsgResponse = emrConsume
End Function

Private Sub IBSSubclass_UnsubclassIt()
    'DetachMessage Me, Me.hWnd, WM_CTLCOLORSTATIC
    'DetachMessage Me, Me.hWnd, WM_CTLCOLORBTN
    DetachMessage Me, Picture1.hWnd, WM_CTLCOLORSTATIC
    DetachMessage Me, Picture1.hWnd, WM_CTLCOLORBTN
End Sub

Private Function IBSSubclass_WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, wParam As Long, lParam As Long, bConsume As Boolean) As Long
    Select Case uMsg
        Case WM_CTLCOLORSTATIC, WM_CTLCOLORBTN
            'WM_CTLCOLORSTATIC = 312
            SetBkMode wParam, 1&
            Dim hDCBmp As Long
            Dim hBmp As Long, hBmpOld As Long
            With Check2
                If CheckBoxTransparentBrush = 0& Then
                    hDCBmp = CreateCompatibleDC(wParam)
                    If hDCBmp <> 0& Then
                        hBmp = CreateCompatibleBitmap(wParam, .Width / Screen.TwipsPerPixelX, .Height / Screen.TwipsPerPixelY)
                        If hBmp <> 0& Then
                            Dim hWndParent As Long
                            hWndParent = GetParent(.hWnd)
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
    End Select
    IBSSubclass_WindowProc = CallOldWindowProc(hWnd, uMsg, wParam, lParam)
End Function
