Attribute VB_Name = "mListViewSort"
Option Explicit

Private m_eSortType As ESortTypeConstants
Private m_eSortOrder As ESortOrderConstants
Private m_lColumn As Long
Private m_tLV As LVITEM


Public Sub SortInit( _
      ByVal eSortOrder As ESortOrderConstants, _
      ByVal ESortType As ESortTypeConstants, _
      ByVal lColumn As Long _
   )
   m_eSortType = ESortType
   m_eSortOrder = eSortOrder
   m_lColumn = lColumn
End Sub

Public Function LVWSortCompare( _
        ByVal lParam1 As Long, _
        ByVal lParam2 As Long, _
        ByVal hWnd As Long _
    ) As Long
Dim s1 As String, s2 As String
Dim v1 As Variant, v2 As Variant, vt As Variant
Dim cI As pcListItem
Dim lIndex As Long, lR As Long

   'Compare the items
   'Return -ve if lI1<lI2,
   '       0 if lI1 = lI2
   '       +ve if lI1 > lI2
   '
   On Error Resume Next
   Select Case m_eSortType
   Case eLVSortItemData
      ' The lParam directly points to our
      ' structure holding the data:
      v1 = 0: v2 = 0
      If Not (lParam1 = 0) Then
         Set cI = ObjectFromPtr(lParam1)
         v1 = cI.ItemData
      End If
      If Not (lParam2 = 0) Then
         Set cI = ObjectFromPtr(lParam2)
         v2 = cI.ItemData
      End If
      
   Case eLVSortTag
      ' The lParam directly points to our
      ' structure holding the data:
      v1 = "": v2 = ""
      If Not (lParam1 = 0) Then
         Set cI = ObjectFromPtr(lParam1)
         v1 = cI.Tag
      End If
      If Not (lParam2 = 0) Then
         Set cI = ObjectFromPtr(lParam2)
         v2 = cI.Tag
      End If
   
   Case eLVSortNumeric
      ' Get the number equivalent of the text
      ' in the relevant column:
      v1 = 0: v2 = 0
      v1 = CDbl(GetLVTextFromlParam(hWnd, lParam1))
      v2 = CDbl(GetLVTextFromlParam(hWnd, lParam2))
      'Debug.Print v1, v2
   
   Case eLVSortDate
      ' Get the date equivalent of the text
      ' in the relevant column:
      s1 = GetLVTextFromlParam(hWnd, lParam1)
      If IsDate(s1) Then
         v1 = CDate(s1)
      Else
         v1 = DateSerial(100, 1, 1)
      End If
      s2 = CDate(GetLVTextFromlParam(hWnd, lParam2))
      If IsDate(s2) Then
         v2 = CDate(s2)
      Else
         v2 = DateSerial(100, 1, 1)
      End If
   
   Case eLVSortString
       v1 = GetLVTextFromlParam(hWnd, lParam1)
       v2 = GetLVTextFromlParam(hWnd, lParam2)
   
   Case eLVSortStringNoCase
       v1 = UCase$(GetLVTextFromlParam(hWnd, lParam1))
       v2 = UCase$(GetLVTextFromlParam(hWnd, lParam2))
   
   Case eLVSortSelected
      v1 = False: v2 = False
      lIndex = IndexForlParam(hWnd, lParam1)
      If lIndex > -1 Then
         v1 = pIsState(hWnd, lIndex, LVIS_SELECTED)
      End If
      lIndex = IndexForlParam(hWnd, lParam2)
      If lIndex > -1 Then
         v2 = pIsState(hWnd, lIndex, LVIS_SELECTED)
      End If
   
   Case eLVSortChecked
      v1 = False: v2 = False
      lIndex = IndexForlParam(hWnd, lParam1)
      If lIndex > -1 Then
         lR = SendMessage(hWnd, LVM_GETITEMSTATE, lIndex, LVIS_STATEIMAGEMASK)
         v1 = ((lR And &H2000&) = &H2000&)
      End If
      lIndex = IndexForlParam(hWnd, lParam2)
      If lIndex > -1 Then
         lR = SendMessage(hWnd, LVM_GETITEMSTATE, lIndex, LVIS_STATEIMAGEMASK)
         v2 = pIsState(hWnd, lIndex, LVIS_SELECTED)
      End If
      
   Case eLVSortIndent
      v1 = 0: v2 = 0
      lIndex = IndexForlParam(hWnd, lParam1)
      If lIndex > -1 Then
         pGetStyle hWnd, lIndex, LVIF_PARAM
         v1 = m_tLV.iIndent
      End If
      lIndex = IndexForlParam(hWnd, lParam2)
      If lIndex > -1 Then
         pGetStyle hWnd, lIndex, LVIF_PARAM
         v2 = m_tLV.iIndent
      End If
   
   Case eLVSortIcon
      v1 = -1: v2 = -1
      lIndex = IndexForlParam(hWnd, lParam1)
      If lIndex > -1 Then
         pGetStyle hWnd, lIndex, LVIF_IMAGE
         v1 = m_tLV.iImage
      End If
      lIndex = IndexForlParam(hWnd, lParam2)
      If lIndex > -1 Then
         pGetStyle hWnd, lIndex, LVIF_IMAGE
         v2 = m_tLV.iImage
      End If
   
   End Select
        
    If (m_eSortOrder = eSortOrderDescending) Then
        vt = v2
        v2 = v1
        v1 = vt
    End If
    
    If (v1 < v2) Then
        LVWSortCompare = -1
    ElseIf (v1 = v2) Then
        LVWSortCompare = 0
    Else
        LVWSortCompare = 1
    End If
    
End Function

Private Function IndexForlParam( _
      ByVal hWnd As Long, _
      ByVal lParam As Long _
   ) As Long
Dim tVFI As LVFINDINFO

   tVFI.flags = LVFI_PARAM
   tVFI.lParam = lParam
   IndexForlParam = SendMessage(hWnd, LVM_FINDITEM, -1, tVFI)
   
End Function

Private Function GetLVTextFromlParam( _
      ByVal hWnd As Long, _
      ByVal lParam As Long _
   ) As String
Dim lIndex As Long
   lIndex = IndexForlParam(hWnd, lParam)
   If lIndex >= 0 Then
      If m_lColumn = 0 Then
         pGetStyle hWnd, lIndex, LVIF_TEXT
         GetLVTextFromlParam = m_tLV.pszText
      Else
         pGetStyle hWnd, lIndex, LVIF_TEXT, m_lColumn
         GetLVTextFromlParam = m_tLV.pszText
      End If
   End If

End Function
' Retrieves the item info into ItemStyle module variable.
Private Sub pGetStyle(ByVal hWnd As Long, ByVal lIndex As Long, ByVal lMask As Long, Optional ByVal lSubItem As Long = 0)
Dim sBuf As String
Dim lPos As Long
    
   m_tLV.mask = lMask
   sBuf = String(261, 0)
   m_tLV.pszText = sBuf
   m_tLV.cchTextMax = 260
   m_tLV.iItem = lIndex
   m_tLV.iSubItem = lSubItem
   SendMessage hWnd, LVM_GETITEM, 0, m_tLV
   lPos = InStr(m_tLV.pszText, Chr$(0))
   If lPos > 0 Then
      m_tLV.pszText = Left$(m_tLV.pszText, lPos - 1)
   End If
   m_tLV.cchTextMax = Len(m_tLV.pszText)
    
End Sub

Private Function pIsState(ByVal hWnd As Long, ByVal lIndex As Long, ByVal lValue As Long, Optional bUseAsMask As Boolean = False) As Long
   If bUseAsMask Then
      m_tLV.stateMask = lValue
   End If
   m_tLV.iItem = lIndex
   pGetStyle hWnd, lIndex, LVIF_STATE
   pIsState = CBool(m_tLV.state And lValue)
End Function

