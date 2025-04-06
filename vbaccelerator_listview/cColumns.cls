VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cColumns"
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
         gErr 1, "cColumns"
      End If
   Else
      gErr 1, "cColumns"
   End If
End Function

Public Function Add( _
      Optional index As Variant, _
      Optional Key As Variant, _
      Optional Text As String, _
      Optional SmallIcon As Variant, _
      Optional Width As Single = 1440 _
   ) As cColumn
Dim ctl As vbalListViewCtl
Dim sKey As String
Dim iIcon As Long
Dim lIndex As Long
Dim lId As Long

   If pbVerify(ctl) Then
      '
      lId = NextColumnID()
      If IsMissing(Key) Then
         ' Generate a machine key for the item:
         sKey = "H" & lId
      Else
         ' Verify the key:
         If ctl.fIsDuplicateColumnKey(Key) Then
            gErr 5, "cColumns"
            Exit Function
         Else
            ' Check for numeric:
            If IsNumeric(sKey) Then
               gErr 4, "cColumns"
               Exit Function
            Else
               sKey = Key
            End If
         End If
      End If
   
      If Not IsMissing(index) Then
         ' Check if the item we're inserting before
         ' exists:
         lIndex = ctl.fIndexForColumnKey(index)
         If lIndex > 0 Then
            ' ok.  We want to add at the zero
            ' based idx:
            lIndex = lIndex - 1
         End If
      Else
         ' Return the current colCount as the index
         lIndex = ctl.fColumnCount
      End If
      
      ' Icons:
      If IsMissing(SmallIcon) Then
         iIcon = -1
      Else
         iIcon = SmallIcon
      End If
   
      lIndex = ctl.fAddColumn(lId, sKey, lIndex, Text, SmallIcon, iIcon, Width)
      If lIndex > 0 Then
         ' We succeeded...
         Dim cH As cColumn
         Set cH = New cColumn
         cH.fInit m_hWnd, lId
         Set Add = cH
      End If
      
   End If
End Function

Public Sub Clear()
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      ctl.fClearColumns
   End If
End Sub

Public Property Get Count() As Long
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      Count = ctl.fColumnCount
   End If
End Property

Public Property Get Exists(index As Variant) As Boolean
Dim col As cColumn
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      On Error Resume Next
      Set col = Item(index)
      Exists = ((Err.Number = 0) And Not (col Is Nothing))
   End If
End Property

Public Property Get Item(index As Variant) As cColumn
Attribute Item.VB_UserMemId = 0
Attribute Item.VB_MemberFlags = "200"
Dim ctl As vbalListViewCtl
Dim lIdx As Long
Dim lId As Long

   If pbVerify(ctl) Then
      If IsNumeric(index) Then
         If index > 0 And index <= ctl.fColumnCount Then
            lIdx = index
         Else
            gErr 8, "vbalListViewCtl"
         End If
      Else
         lIdx = ctl.fIndexForColumnKey(index)
      End If
      
      If lIdx > 0 Then
         Dim cH As New cColumn
         cH.fInit m_hWnd, ctl.fIDForColumnIndex(lIdx)
         Set Item = cH
      End If
      
   End If
End Property
Private Property Get plColIndex(index As Variant) As Long


End Property

Public Sub Remove(index As Variant)
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      Dim lIdx As Long
      ctl.fRemoveColumn lIdx
   End If
End Sub


