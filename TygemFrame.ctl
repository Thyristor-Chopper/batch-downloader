VERSION 5.00
Begin VB.UserControl TygemFrame 
   BackStyle       =   0  '≈ı∏Ì
   ClientHeight    =   5685
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8520
   ScaleHeight     =   5685
   ScaleWidth      =   8520
   Begin VB.Image imgCenter 
      Height          =   2310
      Left            =   1725
      Stretch         =   -1  'True
      Top             =   435
      Width           =   3570
   End
   Begin VB.Image imgRight 
      Height          =   2310
      Left            =   5295
      Picture         =   "TygemFrame.ctx":0000
      Stretch         =   -1  'True
      Top             =   435
      Width           =   165
   End
   Begin VB.Image imgBottomRight 
      Height          =   180
      Left            =   5310
      Picture         =   "TygemFrame.ctx":15EA
      Top             =   2745
      Width           =   150
   End
   Begin VB.Image imgBottom 
      Height          =   180
      Left            =   1725
      Picture         =   "TygemFrame.ctx":17AC
      Stretch         =   -1  'True
      Top             =   2745
      Width           =   3585
   End
   Begin VB.Image imgBottomLeft 
      Height          =   180
      Left            =   0
      Picture         =   "TygemFrame.ctx":39AE
      Top             =   2745
      Width           =   1725
   End
   Begin VB.Image imgLeft 
      Height          =   2310
      Left            =   0
      Picture         =   "TygemFrame.ctx":4A40
      Stretch         =   -1  'True
      Top             =   435
      Width           =   1725
   End
   Begin VB.Image imgTopRight 
      Height          =   435
      Left            =   5310
      Picture         =   "TygemFrame.ctx":11BDA
      Top             =   0
      Width           =   150
   End
   Begin VB.Image imgTop 
      Height          =   435
      Left            =   1725
      Picture         =   "TygemFrame.ctx":11FBC
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3585
   End
   Begin VB.Image imgTopLeft 
      Height          =   435
      Left            =   0
      Picture         =   "TygemFrame.ctx":1718E
      Top             =   0
      Width           =   1725
   End
End
Attribute VB_Name = "TygemFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Sub UserControl_Resize()
    imgTop.Width = UserControl.Width - imgTopLeft.Width - imgTopRight.Width - 30
    imgBottom.Width = imgTop.Width
    imgLeft.Height = UserControl.Height - imgBottomLeft.Height - imgTopLeft.Height - 30
    imgRight.Height = imgLeft.Height
    
    imgTopRight.Left = UserControl.Width - imgTopRight.Width - 30
    imgBottomRight.Left = imgTopRight.Left
    imgBottomLeft.Top = UserControl.Height - imgBottomLeft.Height - 30
    imgBottomRight.Top = imgBottomLeft.Top
    imgBottomRight.Left = imgTopRight.Left
    
    imgRight.Left = imgBottomRight.Left - 15
    imgBottom.Top = imgBottomLeft.Top
End Sub

Private Sub UserControl_Initialize()
    imgTopLeft.Picture = LoadPngIntoPictureWithAlpha(CachePath & "topleft.png")
    imgTopRight.Picture = LoadPngIntoPictureWithAlpha(CachePath & "topright.png")
    imgTop.Picture = LoadPngIntoPictureWithAlpha(CachePath & "top.png")
    imgLeft.Picture = LoadPngIntoPictureWithAlpha(CachePath & "left.png")
    imgRight.Picture = LoadPngIntoPictureWithAlpha(CachePath & "right.png")
    imgBottom.Picture = LoadPngIntoPictureWithAlpha(CachePath & "bottom.png")
    imgBottomLeft.Picture = LoadPngIntoPictureWithAlpha(CachePath & "bottomleft.png")
    imgBottomRight.Picture = LoadPngIntoPictureWithAlpha(CachePath & "bottomright.png")
    imgCenter.Picture = LoadPngIntoPictureWithAlpha(CachePath & "center.png")
    
    UserControl.BackStyle = 0
End Sub
