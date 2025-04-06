VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "pcItemGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Stores all the items that the Groups in the
' List View can't.
' Instances of this class are instantiated
' and stored un-referenced in the idGroup
' of the ListView group items

Public Key As String
Public Tag As String
Public ItemData As Long
Public ID As Long

