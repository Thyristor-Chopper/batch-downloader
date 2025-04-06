Attribute VB_Name = "mMain"
Option Explicit

Private Declare Sub InitCommonControls Lib "Comctl32.dll" ()

Public Sub Main()
   
   InitCommonControls
   
   Dim f As New frmTestListView
   f.Show
   
End Sub
