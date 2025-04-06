VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cSubItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_hWnd As Long
Private m_lPtr As Long
Private m_lId As Long
Private m_lIndex As Long
Private m_lSubItemIndex As Long

Friend Function fInit(ByVal hwnd As Long, ByVal lId As Long, ByVal lPtr As Long, ByVal lSubItemIndex As Long)
   m_hWnd = hwnd
   m_lId = lId
   m_lPtr = lPtr
   m_lSubItemIndex = lSubItemIndex
End Function

Private Function pbVerify(ByRef ctlThis As vbalListViewCtl) As Boolean
Dim lPtr As Long
Dim ctl As vbalListViewCtl
   If IsWindow(m_hWnd) Then
      lPtr = GetProp(m_hWnd, gcObjectProp)
      If Not (lPtr = 0) Then
         Set ctlThis = ObjectFromPtr(lPtr)
         m_lIndex = ctlThis.fVerifylParam(m_lPtr)
         If m_lIndex > 0 Then
            ' Check this is still the right item:
            Dim cI As pcListItem
            Set cI = ObjectFromPtr(m_lPtr)
            If m_lId = cI.ID Then
               If (m_lSubItemIndex > 0) And (m_lSubItemIndex < ctlThis.fColumnCount) Then
                  pbVerify = True
               Else
                  gErr 10, "cSubItem"
               End If
            Else
               gErr 3, "cSubItem"
            End If
         Else
            gErr 3, "cSubItem"
         End If
      Else
         gErr 1, "cSubItem"
      End If
   Else
      gErr 1, "cSubItem"
   End If
End Function

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Gets/sets the text of this subitem."
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      Caption = ctl.fSubItemCaption(m_lIndex, m_lSubItemIndex)
   End If
End Property
Public Property Let Caption(ByVal sText As String)
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      ctl.fSubItemCaption(m_lIndex, m_lSubItemIndex) = sText
   End If
End Property
Public Property Get IconIndex() As Long
Attribute IconIndex.VB_Description = "Gets/sets the icon to display for this sub item  if the control has the SubItemImages property set."
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      IconIndex = ctl.fSubItemIconIndex(m_lIndex, m_lSubItemIndex)
   End If
End Property
Public Property Let IconIndex(ByVal lIconIndex As Long)
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      ctl.fSubItemIconIndex(m_lIndex, m_lSubItemIndex) = lIconIndex
   End If
End Property
Public Property Get ShowInTile() As Boolean
Attribute ShowInTile.VB_Description = "Gets/sets whether to show this sub item in Tile View or not."
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      ShowInTile = ctl.fSubItemShowInTile(m_lIndex, m_lSubItemIndex)
   End If
End Property
Public Property Let ShowInTile(ByVal bState As Boolean)
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      ctl.fSubItemShowInTile(m_lIndex, m_lSubItemIndex) = bState
   End If
End Property
Public Property Get Index() As Long
Attribute Index.VB_Description = "Gets the index of this subitem."
   Index = m_lSubItemIndex
End Property
