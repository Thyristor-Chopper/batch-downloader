VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Private m_hWnd As Long
Private m_lId As Long
Private m_lIndex As Long

Friend Function fInit(ByVal hwnd As Long, ByVal lId As Long)
   m_hWnd = hwnd
   m_lId = lId
End Function
Private Function pbVerify(ByRef ctlThis As vbalListViewCtl) As Boolean
Dim lPtr As Long
Dim lIdx As Long
   If IsWindow(m_hWnd) Then
      lPtr = GetProp(m_hWnd, gcObjectProp)
      If Not (lPtr = 0) Then
         Set ctlThis = ObjectFromPtr(lPtr)
         lIdx = ctlThis.fVerifyColumnID(m_lId)
         If lIdx > 0 Then
            m_lIndex = lIdx
            pbVerify = True
         Else
            gErr 8, "cListItem"
         End If
      Else
         gErr 1, "cListItems"
      End If
   Else
      gErr 1, "cListItems"
   End If
End Function

Public Property Get Width() As Single
Attribute Width.VB_Description = "Gets/sets the width of this column."
Dim ctl As vbalListViewCtl
Dim lWidth As Long
Dim X As Single, Y As Single
   If pbVerify(ctl) Then
      lWidth = ctl.fColumnWidth(m_lIndex)
      ctl.fScale lWidth, 0, X, Y
      Width = X
   End If
End Property
Public Property Let Width(ByVal lWidth As Single)
Dim ctl As vbalListViewCtl
Dim X As Long, Y As Long
   If pbVerify(ctl) Then
      ctl.fUnScale lWidth, 0, X, Y
      ctl.fColumnWidth(m_lIndex) = X
   End If
End Property
Public Property Get ItemData() As Long
Attribute ItemData.VB_Description = "Gets/sets a custom long value associated with this column."
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      ItemData = ctl.fColumnItemData(m_lIndex)
   End If
End Property
Public Property Let ItemData(ByVal lItemData As Long)
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      ctl.fColumnItemData(m_lIndex) = lItemData
   End If
End Property
Public Property Get Tag() As String
Attribute Tag.VB_Description = "Gets/sets a custom string value associated with this column."
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      Tag = ctl.fColumnTag(m_lIndex)
   End If
End Property
Public Property Let Tag(ByVal sTag As String)
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      ctl.fColumnTag(m_lIndex) = sTag
   End If
End Property
Public Property Get Text() As String
Attribute Text.VB_Description = "Gets/sets the text displayed in the column header."
Attribute Text.VB_UserMemId = 0
Attribute Text.VB_MemberFlags = "200"
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      Text = ctl.fColumnText(m_lIndex)
   End If
End Property
Public Property Let Text(ByVal sText As String)
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      ctl.fColumnText(m_lIndex) = sText
   End If
End Property
Public Property Get Key() As String
Attribute Key.VB_Description = "Gets/sets the key of this column."
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      Key = ctl.fColumnKey(m_lIndex)
   End If
End Property
Public Property Let Key(ByVal sKey As String)
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      ctl.fColumnKey(m_lIndex) = sKey
   End If
End Property
Public Property Get IconIndex() As Long
Attribute IconIndex.VB_Description = "Gets/sets the 0-based index of the icon within the ColumnHeader ImageList."
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      IconIndex = ctl.fColumnImage(m_lIndex)
   End If
End Property
Public Property Let IconIndex(ByVal lIconIndex As Long)
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      ctl.fColumnImage(m_lIndex) = lIconIndex
   End If
End Property
Public Property Get ImageOnRight() As Boolean
Attribute ImageOnRight.VB_Description = "Gets/sets whether the image should display on the right of the column."
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      ImageOnRight = ctl.fColumnImageOnRight(m_lIndex)
   End If
End Property
Public Property Let ImageOnRight(ByVal bState As Boolean)
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      ctl.fColumnImageOnRight(m_lIndex) = bState
   End If
End Property
Public Property Get Alignment() As EColumnHeaderAlignConstants
Attribute Alignment.VB_Description = "Gets/sets the alignment of the text and icon."
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      Alignment = ctl.fColumnTextAlign(m_lIndex)
   End If
End Property
Public Property Let Alignment(ByVal eAlign As EColumnHeaderAlignConstants)
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      ctl.fColumnTextAlign(m_lIndex) = eAlign
   End If
End Property
Public Property Get SortOrder() As ESortOrderConstants
Attribute SortOrder.VB_Description = "Gets/sets a value which you can use to maintain the current sort order of this column."
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      SortOrder = ctl.fColumnSortOrder(m_lIndex)
   End If
End Property
Public Property Let SortOrder(ByVal eOrder As ESortOrderConstants)
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      ctl.fColumnSortOrder(m_lIndex) = eOrder
   End If
End Property
Public Property Get SortType() As ESortTypeConstants
Attribute SortType.VB_Description = "Gets/sets a value which you can use to maintain the current sort type of this column."
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      SortType = ctl.fColumnSortType(m_lIndex)
   End If
End Property
Public Property Let SortType(ByVal ESortType As ESortTypeConstants)
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      ctl.fColumnSortType(m_lIndex) = ESortType
   End If
End Property
Public Property Get Position() As Long
Attribute Position.VB_Description = "Gets/sets the position (order) of this header in the control."
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      Position = ctl.fColumnIndex(m_lIndex) + 1
   End If
End Property
Public Property Let Position(ByVal lPosition As Long)
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      ctl.fColumnIndex(m_lIndex) = lPosition - 1
   End If
End Property


