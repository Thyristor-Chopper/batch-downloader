VERSION 5.00
Begin VB.UserControl WLCheckBox 
   BackStyle       =   0  '≈ı∏Ì
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   KeyPreview      =   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Windowless      =   -1  'True
   Begin VB.Image Image1 
      Height          =   375
      Left            =   120
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "WLCheckBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function DrawThemeBackground Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal lHDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, pRect As RECT, pClipRect As RECT) As Long
Private Declare Function GetThemeBitmap Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal iPartId As Long, ByVal iStateId As Long, ByVal iPropId As Long, ByVal dwFlags As Long, phBitmap As Long) As Long
Private Declare Function OpenThemeData Lib "uxtheme.dll" (ByVal hWnd As Long, ByVal pszClassList As String) As Long
Private Declare Function CloseThemeData Lib "uxtheme.dll" (ByVal hTheme As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function OleCreatePictureIndirect Lib "oleaut32" (ByRef pPictDesc As PICTDESC, ByRef riid As Any, ByVal fPictureOwnsHandle As Long, ByRef pIPicture As IPictureDisp) As Long

Const BP_CHECKBOX As Long = 3&
Const CBS_CHECKEDNORMAL As Long = 5&
Const GBF_DIRECT As Long = 1&

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Type PICTDESC
    cbSizeOfStruct As Long
    PicType As Long
    hImage As Long
    Data1 As Long
    Data2 As Long
End Type

Public Function PictureFromHandle(ByVal Handle As Long, ByVal PicType As VBRUN.PictureTypeConstants) As IPictureDisp
    If Handle = 0& Then Exit Function
    Dim PICD As PICTDESC, IID As CLSID, NewPicture As IPictureDisp
    With PICD
        .cbSizeOfStruct = LenB(PICD)
        .PicType = PicType
        .hImage = Handle
    End With
    With IID
        .Data1 = &H7BF80980
        .Data2 = &HBF32
        .Data3 = &H101A
        .Data4(0) = &H8B
        .Data4(1) = &HBB
        .Data4(3) = &HAA
        .Data4(5) = &H30
        .Data4(6) = &HC
        .Data4(7) = &HAB
    End With
    If OleCreatePictureIndirect(PICD, IID, 1, NewPicture) = 0 Then Set PictureFromHandle = NewPicture
End Function

Private Sub UserControl_Initialize()
    Dim ClientRect As RECT
    GetClientRect UserControl.hWnd, ClientRect
    Dim hTheme As Long
    Dim hBitmap As Long
    hTheme = OpenThemeData(UserControl.hWnd, "Button")
    GetThemeBitmap hTheme, BP_CHECKBOX, CBS_CHECKEDNORMAL, 0&, GBF_DIRECT, hBitmap
    CloseThemeData hTheme
    MsgBox hBitmap
End Sub
