VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "pcListItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Stores all the items that the ListView
' can't.
' Instances of this class are instantiated
' and stored un-referenced in the lParam
' of the ListView items (fItemData)

Public Key As String
Public Tag As String
Public ToolTipText As String
Public BackColor As OLE_COLOR
Public ForeColor As OLE_COLOR
Public Font As IFont
Public Icon As Long
Public SmallIcon As Long
Public StateIcon As Long
Public ItemData As Long
Private m_bShowInTileView() As Boolean
Private m_lShowInTileViewCount As Long
Public ID As Long

Public Property Get ShowInTileView(ByVal lIndex As Long) As Boolean
   If (m_lShowInTileViewCount < lIndex) Then
      ShowInTileView = False
   Else
      ShowInTileView = m_bShowInTileView(lIndex)
   End If
End Property
Public Property Let ShowInTileView(ByVal lIndex As Long, ByVal bState As Boolean)
   If (m_lShowInTileViewCount < lIndex) And (bState) Then
      m_lShowInTileViewCount = lIndex
      ReDim Preserve m_bShowInTileView(1 To lIndex) As Boolean
      m_bShowInTileView(lIndex) = bState
   Else
      m_bShowInTileView(lIndex) = bState
   End If
End Property

