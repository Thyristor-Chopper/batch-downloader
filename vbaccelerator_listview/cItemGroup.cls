VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cItemGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_hWnd As Long
Private m_lId As Long
Private m_lPtr As Long
Private m_lIndex As Long

Friend Function fInit(ByVal hwnd As Long, ByVal lId As Long, ByVal lPtr As Long) As Boolean
   m_hWnd = hwnd
   m_lId = lId
   m_lPtr = lPtr
End Function

Private Function pbVerify(ByRef ctlThis As vbalListViewCtl) As Boolean
Dim lPtr As Long
Dim ctl As vbalListViewCtl
   If IsWindow(m_hWnd) Then
      lPtr = GetProp(m_hWnd, gcObjectProp)
      If Not (lPtr = 0) Then
         Set ctlThis = ObjectFromPtr(lPtr)
         If ctlThis.fVerifyGroupID(m_lPtr) Then
            ' Check this is still the right item:
            Dim cI As pcItemGroup
            Set cI = ObjectFromPtr(m_lPtr)
            If m_lId = cI.ID Then
               pbVerify = True
            Else
               gErr 3, "cListItem"
            End If
         Else
            gErr 3, "cListItems"
         End If
      Else
         gErr 1, "cListItems"
      End If
   Else
      gErr 1, "cListItems"
   End If
End Function
Public Property Get Header() As String
Attribute Header.VB_Description = "Gets the text to display in the group."
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      Header = ctl.fGroupHeader(m_lPtr)
   End If
End Property
Public Property Let Header(ByVal sHeader As String)
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      ctl.fGroupHeader(m_lPtr) = sHeader
   End If
End Property

Public Property Get Align() As EItemGroupHeaderAlignConstants
Attribute Align.VB_Description = "Gets/sets the alignment of ths group."
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      Align = ctl.fGroupAlign(m_lPtr)
   End If
End Property
Public Property Let Align(ByVal eAlign As EItemGroupHeaderAlignConstants)
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      ctl.fGroupAlign(m_lPtr) = eAlign
   End If
End Property

Public Property Get State() As EItemGroupStateConstants
Attribute State.VB_Description = "Gets/sets the state of this group."
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      State = ctl.fGroupState(m_lPtr)
   End If
End Property
Public Property Let State(ByVal eState As EItemGroupStateConstants)
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      ctl.fGroupState(m_lPtr) = eState
   End If
End Property

Public Property Get GroupType() As EItemGroupTypeConstants
Attribute GroupType.VB_Description = "Gets/sets the whether this group is a header or a footer."
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      GroupType = ctl.fGroupType(m_lPtr)
   End If
End Property
Public Property Let GroupType(ByVal eType As EItemGroupTypeConstants)
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      ctl.fGroupType(m_lPtr) = eType
   End If
End Property

Public Property Get Key() As String
Attribute Key.VB_Description = "Gets/sets the key of this group."
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      Dim pc As pcItemGroup
      Set pc = ObjectFromPtr(m_lPtr)
      Key = pc.Key
   End If
End Property
Public Property Let Key(ByVal sKey As String)
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      If ctl.fIsDuplicateGroupKey(sKey) Then
         gErr 5, "cItemGroup"
      Else
         Dim pc As pcListItem
         Set pc = ObjectFromPtr(m_lPtr)
         ctl.fChangeGroupKey pc.Key, sKey
         pc.Key = sKey
      End If
   End If
End Property
