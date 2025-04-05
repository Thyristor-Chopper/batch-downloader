VERSION 5.00
Begin VB.UserControl ThreeDLine 
   BackStyle       =   0  '≈ı∏Ì
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Windowless      =   -1  'True
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   0
      X2              =   3585
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   15
      X2              =   3600
      Y1              =   15
      Y2              =   15
   End
End
Attribute VB_Name = "ThreeDLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Sub UserControl_Resize()
    Line1.X2 = UserControl.Width - Screen.TwipsPerPixelX
    Line2.X2 = UserControl.Width
    UserControl.Height = 2 * Screen.TwipsPerPixelY
End Sub
