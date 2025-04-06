VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cListItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_hWnd As Long

Friend Function fInit(ByVal hwnd As Long)
   m_hWnd = hwnd
End Function
Private Function pbVerify(ByRef ctlThis As vbalListViewCtl) As Boolean
Dim lPtr As Long
   If IsWindow(m_hWnd) Then
      lPtr = GetProp(m_hWnd, gcObjectProp)
      If Not (lPtr = 0) Then
         Set ctlThis = ObjectFromPtr(lPtr)
         pbVerify = True
      Else
         gErr 1, "cListItems"
      End If
   Else
      gErr 1, "cListItems"
   End If
End Function

Public Function Add( _
      Optional Index As Variant, _
      Optional Key As Variant, _
      Optional Text As String, _
      Optional Icon As Variant, _
      Optional SmallIcon As Variant _
   ) As cListItem
Dim ctl As vbalListViewCtl
Dim cI As cListItem
Dim pc As pcListItem
Dim iUnk As IShellFolderEx_TLB.IUnknown
Dim lIndex As Long
Dim iIcon As Long
Dim iSmallIcon As Long
Dim sKey As String
Dim lPtr As Long

   If pbVerify(ctl) Then
      
      ' Generate an object to store the extra
      ' data for the item:
      Set pc = New pcListItem
      pc.ID = NextItemID
      SetVariant pc.Icon, Icon
      SetVariant pc.SmallIcon, SmallIcon
      If IsMissing(Key) Then
         ' Generate a machine key for the item:
         sKey = "C" & pc.ID
      Else
         ' Verify the key:
         If ctl.fIsDuplicateItemKey(Key) Then
            gErr 5, "cListItems"
            Exit Function
         Else
            If IsNumeric(sKey) Then
               gErr 4, "cListItems"
               Exit Function
            Else
               sKey = Key
            End If
         End If
      End If
      pc.Key = sKey
      
      ' Check for before parameter to add:
      If (IsMissing(Index)) Then
         lIndex = ctl.fCount + 1
      Else
         If (IsNumeric(Index)) Then
            On Error Resume Next
            lIndex = Index
            If (Err.Number = 0) Then
               
            Else
               gErr 4, "cListItems"
            End If
            On Error GoTo 0
         Else
            gErr 4, "cListItems"
         End If
      End If
      
      ' Icons:
      If IsMissing(Icon) Then
         iIcon = -1
      Else
         iIcon = Icon
      End If
            
      ' Lock it:
      Set iUnk = pc
      iUnk.AddRef
      Set iUnk = Nothing
      
      lPtr = ObjPtr(pc)
      If ctl.fAddItem(Text, lIndex, iIcon, 0, lPtr) Then
         ' succeeded in adding.  Add the key to
         ' the control's collection:
         ctl.fAddItemKey sKey, lPtr
         
         ' Set the correct defaults:
         pc.BackColor = -1
         pc.ForeColor = -1
         
         ' Return a reference to the item:
         Set cI = New cListItem
         cI.fInit m_hWnd, pc.ID, lPtr
         Set Add = cI
      Else
         ' failed to add, release the pc object:
         Set iUnk = pc
         iUnk.Release
         Set iUnk = Nothing
         
         gErr 7, "cListItems"
      End If
   End If
   
End Function

Public Sub Clear()
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      ctl.fClearItems
   End If
End Sub

Public Property Get Count() As Long
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      Count = ctl.fCount
   End If
End Property

Public Property Get Exists(Index As Variant) As Boolean
Dim ctl As vbalListViewCtl
Dim lIndex As Long
   If pbVerify(ctl) Then
      On Error Resume Next
      lIndex = plFindIndex(ctl, Index)
      If lIndex > 0 Then
         Exists = (Err.Number = 0)
      End If
   End If
End Property

Public Property Get Item(Index As Variant) As cListItem
Attribute Item.VB_UserMemId = 0
Attribute Item.VB_MemberFlags = "200"
Dim ctl As vbalListViewCtl
Dim lIndex As Long
Dim lPtr As Long
Dim lId As Long
   If pbVerify(ctl) Then
      lIndex = plFindIndex(ctl, Index)
      If lIndex > 0 Then
         lPtr = ctl.fItemData(lIndex)
         Dim pc As pcListItem
         Set pc = ObjectFromPtr(lPtr)
         lId = pc.ID
         Dim cI As New cListItem
         cI.fInit m_hWnd, lId, lPtr
         Set Item = cI
      End If
   End If
End Property

Public Sub Remove(Index As Variant)
Dim ctl As vbalListViewCtl
Dim lIndex As Long
   If pbVerify(ctl) Then
      lIndex = plFindIndex(ctl, Index)
      If lIndex > 0 Then
         ctl.fRemoveItem lIndex
      End If
   End If
End Sub

Public Property Get CountPerPage() As Long
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      CountPerPage = ctl.fItemCountPerPage
   End If
End Property

Private Function plFindIndex(ByRef ctl As vbalListViewCtl, Index As Variant) As Long
   If IsNumeric(Index) Then
      If Index > 0 And Index <= ctl.fCount Then
         plFindIndex = Index
      Else
         gErr 6, "cListItems"
      End If
   Else
      plFindIndex = ctl.fItemIndexForKey(Index)
   End If
End Function

Public Sub SortItems()
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      ctl.fSortItems
   End If
End Sub
