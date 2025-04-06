VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cWorkAreas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_hWnd As Long
Private m_lIndex As Long

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
         gErr 1, "cWorkAreas"
      End If
   Else
      gErr 1, "cWorkAreas"
   End If
End Function
Public Property Get Count() As Long
Dim ctl As vbalListViewCtl
   If (pbVerify(ctl)) Then
      Count = ctl.fWorkAreaCount()
   End If
End Property
Public Function Add( _
      Optional Index As Variant, _
      Optional Key As Variant, _
      Optional left As Variant, _
      Optional top As Variant, _
      Optional Width As Variant, _
      Optional Height As Variant _
   ) As cWorkArea

Dim ctl As vbalListViewCtl
Dim i As Long
Dim lIndex As Long
Dim lId As Long

   If (pbVerify(ctl)) Then
      
      ' Check key ok
      If Not IsMissing(Key) Then
         If IsNumeric(Key) Then
            gErr 4, "cWorkAreas"
            Exit Function
         Else
            For i = 1 To ctl.fWorkAreaCount
               If (ctl.fWorkAreaKey(i) = Key) Then
                  gErr 5, "cWorkAreas"
                  Exit Function
               End If
            Next i
         End If
      Else
         Key = ""
      End If
      
      ' Key is ok, is the Index in range?
      lIndex = 0
      If Not IsMissing(Index) Then
         If (Index < 0) Or (Index > ctl.fWorkAreaCount) Then
            gErr 6, "cWorkAreas"
            Exit Function
         Else
            lIndex = Index
         End If
      End If
      
      ' ok
      lId = ctl.fAddWorkArea(lIndex, Key, left, top, Width, Height)
      If (lId > 0) Then
         Dim cWA As New cWorkArea
         cWA.fInit m_hWnd, lId
         Set Add = cWA
      End If
      
   End If
   
End Function
Public Sub Clear()
Dim ctl As vbalListViewCtl
   If (pbVerify(ctl)) Then
      ctl.fWorkAreasClear
   End If
End Sub
Public Property Get Exists(Index As Variant) As Boolean
Dim ctl As vbalListViewCtl
   If (pbVerify(ctl)) Then
      If IsNumeric(Index) Then
         If (Index > 0) And (Index <= ctl.fWorkAreaCount()) Then
            m_lIndex = Index
            Exists = True
         End If
      Else
         Dim i As Long
         For i = 1 To ctl.fWorkAreaCount
            If (ctl.fWorkAreaKey(i) = Index) Then
               m_lIndex = i
               Exists = True
            End If
         Next i
      End If
   End If
End Property
Public Property Get Item(Index As Variant) As cWorkArea
Attribute Item.VB_UserMemId = 0
Dim ctl As vbalListViewCtl
   If (pbVerify(ctl)) Then
      If Exists(Index) Then
         Dim cWA As New cWorkArea
         cWA.fInit m_hWnd, ctl.fWorkAreaIdForIndex(m_lIndex)
         Set Item = cWA
      End If
   End If
End Property
Public Sub Remove(Index As Variant)
Dim ctl As vbalListViewCtl
   If (pbVerify(ctl)) Then
      ctl.fRemoveWorkArea m_lIndex
   End If
End Sub
