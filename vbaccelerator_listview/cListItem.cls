VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cListItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_hWnd As Long
Private m_lPtr As Long
Private m_lId As Long
Private m_lIndex As Long

Friend Function fInit(ByVal hwnd As Long, ByVal lId As Long, ByVal lPtr As Long)
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
         m_lIndex = ctlThis.fVerifylParam(m_lPtr)
         If m_lIndex > 0 Then
            ' Check this is still the right item:
            Dim cI As pcListItem
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
Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "Gets/sets the ToolTip text to display for this item."
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      Dim pc As pcListItem
      Set pc = ObjectFromPtr(m_lPtr)
      ToolTipText = pc.ToolTipText
   End If
End Property
Public Property Let ToolTipText(ByVal sText As String)
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      Dim pc As pcListItem
      Set pc = ObjectFromPtr(m_lPtr)
      pc.ToolTipText = sText
   End If
End Property
Public Property Get Text() As String
Attribute Text.VB_Description = "Gets/sets the text of this item."
Attribute Text.VB_UserMemId = 0
Attribute Text.VB_MemberFlags = "200"
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      Text = ctl.fItemCaption(m_lIndex)
   End If
End Property
Public Property Let Text(ByVal sText As String)
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      ctl.fItemCaption(m_lIndex) = sText
   End If
End Property
Public Property Get Key() As String
Attribute Key.VB_Description = "Gets/sets the key of the item in the control."
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      Dim pc As pcListItem
      Set pc = ObjectFromPtr(m_lPtr)
      Key = pc.Key
   End If
End Property
Public Property Let Key(ByVal sKey As String)
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      If ctl.fIsDuplicateItemKey(sKey) Then
         gErr 5, "cListItem"
      Else
         Dim pc As pcListItem
         Set pc = ObjectFromPtr(m_lPtr)
         ctl.fChangeKey pc.Key, sKey
         pc.Key = sKey
      End If
   End If
End Property
Public Property Get Index() As Long
Attribute Index.VB_Description = "Gets/sets the index of the item within the control."
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      Dim pc As pcListItem
      Set pc = ObjectFromPtr(m_lPtr)
      Index = ctl.fItemIndexForKey(pc.Key)
   End If
End Property
Public Property Get ItemData() As Long
Attribute ItemData.VB_Description = "Gets/sets a custom long value associated with the item."
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      Dim pc As pcListItem
      Set pc = ObjectFromPtr(m_lPtr)
      ItemData = pc.ItemData
   End If
End Property
Public Property Let ItemData(ByVal lItemData As Long)
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      Dim pc As pcListItem
      Set pc = ObjectFromPtr(m_lPtr)
      pc.ItemData = lItemData
   End If
End Property
Public Property Get Tag() As String
Attribute Tag.VB_Description = "Gets/sets a custom string value associated with this item."
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      Dim pc As pcListItem
      Set pc = ObjectFromPtr(m_lPtr)
      Tag = pc.Tag
   End If
End Property
Public Property Let Tag(ByVal sTag As String)
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      Dim pc As pcListItem
      Set pc = ObjectFromPtr(m_lPtr)
      pc.Tag = sTag
   End If
End Property
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Gets/sets the background colour of the item when the control's CustomDraw mode is True.  Set to -1 to make the item transparent."
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      Dim pc As pcListItem
      Set pc = ObjectFromPtr(m_lPtr)
      BackColor = pc.BackColor
   End If
End Property
Public Property Let BackColor(ByVal oColor As OLE_COLOR)
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      Dim pc As pcListItem
      Set pc = ObjectFromPtr(m_lPtr)
      pc.BackColor = oColor
   End If
End Property
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Gets/sets the foreground colour of the item when the control's CustomDraw mode is True.  Set to -1 to use the default."
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      Dim pc As pcListItem
      Set pc = ObjectFromPtr(m_lPtr)
      ForeColor = pc.ForeColor
   End If
End Property
Public Property Let ForeColor(ByVal oColor As OLE_COLOR)
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      Dim pc As pcListItem
      Set pc = ObjectFromPtr(m_lPtr)
      pc.ForeColor = oColor
   End If
End Property
Public Property Get Font() As IFont
Attribute Font.VB_Description = "Gets/sets the item font when the control's CustomDraw mode is True.  Set to nothing to use the default."
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      Dim pc As pcListItem
      Set pc = ObjectFromPtr(m_lPtr)
      Set Font = pc.Font
   End If
End Property
Public Property Let Font(ByVal iFnt As IFont)
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      Dim pc As pcListItem
      Set pc = ObjectFromPtr(m_lPtr)
      Set pc.Font = iFnt
   End If
End Property
Public Property Set Font(ByVal iFnt As IFont)
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      Dim pc As pcListItem
      Set pc = ObjectFromPtr(m_lPtr)
      Set pc.Font = iFnt
   End If
End Property
Public Property Get IconIndex() As Long
Attribute IconIndex.VB_Description = "Gets/sets the 0-based index of the icon to use in the ImageList for this item."
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      Dim pc As pcListItem
      Set pc = ObjectFromPtr(m_lPtr)
      IconIndex = pc.Icon
   End If
End Property
Public Property Let IconIndex(ByVal lIconIndex As Long)
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      Dim pc As pcListItem
      Set pc = ObjectFromPtr(m_lPtr)
      ctl.fItemIconIndex(m_lIndex) = lIconIndex
   End If
End Property
Public Property Get StateIconIndex() As Long
Attribute StateIconIndex.VB_Description = "Gets/sets the state icon index.  Not implemented at this time."
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      Dim pc As pcListItem
      Set pc = ObjectFromPtr(m_lPtr)
      StateIconIndex = pc.StateIcon
   End If
End Property
Public Property Let StateIconIndex(ByVal lIconIndex As Long)
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      Dim pc As pcListItem
      Set pc = ObjectFromPtr(m_lPtr)
      ctl.fItemStateIconIndex(m_lIndex) = lIconIndex
   End If
End Property
Public Property Get Checked() As Boolean
Attribute Checked.VB_Description = "Gets/sets whether the item is checked or not."
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      Checked = ctl.fItemChecked(m_lIndex)
   End If
End Property
Public Property Let Checked(ByVal bChecked As Boolean)
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      ctl.fItemChecked(m_lIndex) = bChecked
   End If
End Property
Public Property Get Hot() As Boolean
Attribute Hot.VB_Description = "Gets/sets whether this item is Hot (shown as if the mouse is tracking over it) or not."
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      Hot = ctl.fItemHot(m_lIndex)
   End If
End Property
Public Property Let Hot(ByVal bHot As Boolean)
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      ctl.fItemHot(m_lIndex) = bHot
   End If
End Property
Public Property Get Cut() As Boolean
Attribute Cut.VB_Description = "Gets/sets whether the item displays faded, as in Explorer when you Edit->Cut an item."
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      Cut = ctl.fItemCut(m_lIndex)
   End If
End Property
Public Property Let Cut(ByVal bCut As Boolean)
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      ctl.fItemCut(m_lIndex) = bCut
   End If
End Property

Public Property Get Indent() As Long
Attribute Indent.VB_Description = "Gets/sets the identation of the item in Report Mode."
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      Indent = ctl.fItemIndent(m_lIndex)
   End If
End Property
Public Property Let Indent(ByVal lIndent As Long)
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      ctl.fItemIndent(m_lIndex) = lIndent
   End If
End Property

Public Sub StartEdit()
Attribute StartEdit.VB_Description = "Starts a label edit action."
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      ctl.fStartEdit m_lIndex
   End If
End Sub
Public Sub EnsureVisible()
Attribute EnsureVisible.VB_Description = "Ensures this item is visible in the control."
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      ctl.fEnsureVisible m_lIndex
   End If
End Sub
Public Property Get Selected() As Boolean
Attribute Selected.VB_Description = "Gets/sets whether this item is selected or not."
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      Selected = ctl.fItemSelected(m_lIndex)
   End If
End Property
Public Property Let Selected(ByVal bState As Boolean)
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      ctl.fItemSelected(m_lIndex) = bState
   End If
End Property
Public Property Get top() As Single
Attribute top.VB_Description = "Gets/sets the Top position of this item within the control."
Dim ctl As vbalListViewCtl
Dim lTop As Long
Dim X As Single, Y As Single
   If pbVerify(ctl) Then
      lTop = ctl.fItemPositionY(m_lIndex)
      ctl.fScale 0, lTop, X, Y
      top = Y
   End If
End Property
Public Property Let top(ByVal lTop As Single)
Dim ctl As vbalListViewCtl
Dim xPixels As Long, yPixels As Long
   If pbVerify(ctl) Then
      ctl.fUnScale 0, lTop, xPixels, yPixels
      ctl.fSetItemPosition m_lIndex, ctl.fItemPositionX(m_lIndex), yPixels
   End If
End Property
Public Property Get left() As Single
Attribute left.VB_Description = "Gets/sets the left position of the item in the control."
Dim ctl As vbalListViewCtl
Dim lLeft As Long
Dim X As Single, Y As Single
   If pbVerify(ctl) Then
      lLeft = ctl.fItemPositionX(m_lIndex)
      ctl.fScale lLeft, 0, X, Y
      left = X
   End If
End Property
Public Property Let left(ByVal lLeft As Single)
Dim ctl As vbalListViewCtl
Dim xPixels As Long, yPixels As Long
   If pbVerify(ctl) Then
      ctl.fUnScale lLeft, 0, xPixels, yPixels
      ctl.fSetItemPosition m_lIndex, xPixels, ctl.fItemPositionY(m_lIndex)
   End If
End Property
Public Property Get SubItems() As cSubItems
Attribute SubItems.VB_Description = "Gets the collection of SubItems associated with this item."
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      Dim cS As New cSubItems
      cS.fInit m_hWnd, m_lId, m_lPtr
      Set SubItems = cS
   End If
End Property
Public Property Get Group() As cItemGroup
Attribute Group.VB_Description = "Gets/sets the group that this item appears in.  Requires ComCtl32.DLL v6.0."
Dim ctl As vbalListViewCtl
Dim lGroupId As Long
Dim lGroupPtr As Long

   If pbVerify(ctl) Then
      ctl.fGetItemGroup m_lIndex, lGroupId, lGroupPtr
      If Not (lGroupPtr = 0) Then
         Dim cG As New cItemGroup
         cG.fInit m_hWnd, lGroupId, lGroupPtr
         Set Group = cG
      End If
   End If
   
End Property
Public Property Let Group(cG As cItemGroup)
   pSetGroup cG
End Property
Public Property Set Group(cG As cItemGroup)
   pSetGroup cG
End Property

Private Sub pSetGroup(cG As cItemGroup)
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      ctl.fSetItemGroup m_lIndex, cG.Key
   End If
End Sub

