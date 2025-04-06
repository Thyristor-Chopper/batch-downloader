VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cItemGroups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_hWnd As Long

Friend Function fInit(ByVal hWnd As Long)
   m_hWnd = hWnd
End Function
Private Function pbVerify(ByRef ctlThis As vbalListViewCtl) As Boolean
Dim lPtr As Long
   If IsWindow(m_hWnd) Then
      lPtr = GetProp(m_hWnd, gcObjectProp)
      If Not (lPtr = 0) Then
         Set ctlThis = ObjectFromPtr(lPtr)
         pbVerify = True
      Else
         gErr 1, "cItemGroups"
      End If
   Else
      gErr 1, "cItemGroups"
   End If
End Function

Private Function plFindIndex(ByRef ctl As vbalListViewCtl, Index As Variant) As Long
   If IsNumeric(Index) Then
      If Index > 0 And Index <= ctl.fGroupCount Then
         plFindIndex = Index
      Else
         gErr 6, "cItemGroups"
      End If
   Else
      plFindIndex = ctl.fGroupIndexForKey(Index)
   End If
End Function


Public Function Add( _
      Optional Index As Variant, _
      Optional Key As Variant, _
      Optional Text As String _
   ) As cItemGroup
Dim ctl As vbalListViewCtl
Dim cI As cItemGroup
Dim pc As pcItemGroup
Dim iUnk As IShellFolderEx_TLB.IUnknown
Dim lIndex As Long
Dim sKey As String
Dim lPtr As Long

   If pbVerify(ctl) Then
   
      ' Generate an object to store the extra
      ' data for the item:
      Set pc = New pcItemGroup
      pc.ID = NextItemID
      If IsMissing(Key) Then
         ' generate machine key:
         sKey = "C" & pc.ID
      Else
         ' verify the key:
         If ctl.fIsDuplicateGroupKey(Key) Then
            gErr 5, "cItemGroups"
            Exit Function
         Else
            If IsNumeric(sKey) Then
               gErr 4, "cItemGroups"
               Exit Function
            Else
               sKey = Key
            End If
         End If
      End If
      pc.Key = sKey
      
      ' Check for before parameter:
      If IsMissing(Index) Then
         ' Group is placed at the start:
         lIndex = 0
      Else
         ' Get the index of the specified item:
         If IsNumeric(Index) Then
            If (Index > 0) And (Index <= ctl.fCount) Then
               lIndex = Index - 1
            Else
               gErr 3, "cItemGroups"
            End If
         Else
            lIndex = ctl.fItemIndexForKey(Index)
         End If
      End If
      
      ' Lock object:
      Set iUnk = pc
      iUnk.AddRef
      Set iUnk = Nothing
      
      lPtr = ObjPtr(pc)
      If ctl.fAddGroup(Text, lIndex, lPtr) Then
         ' success:
         ctl.fAddGroupKey sKey, lPtr
         
         ' return a reference:
         Set cI = New cItemGroup
         cI.fInit m_hWnd, pc.ID, lPtr
         Set Add = cI
      Else
         ' failed, release object:
         Set iUnk = pc
         iUnk.Release
         Set iUnk = Nothing
         
         gErr 7, "cItemGroups"
      End If
   End If
   
End Function
   
Public Property Get Count() As Long
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      Count = ctl.fGroupCount
   End If
End Property
Public Sub Clear()
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      ctl.fGroupClear
   End If
End Sub
Public Sub Remove(Index As Variant)
Dim ctl As vbalListViewCtl
Dim lIndex As Long
   If pbVerify(ctl) Then
      ctl.fRemoveGroup Index
   End If
End Sub
Public Property Get Enabled() As Boolean
Dim ctl As vbalListViewCtl
   If (pbVerify(ctl)) Then
      Enabled = ctl.fGroupsEnabled()
   End If
End Property
Public Property Let Enabled(ByVal bState As Boolean)
Dim ctl As vbalListViewCtl
   If (pbVerify(ctl)) Then
      ctl.fGroupsEnabled = bState
   End If
End Property

