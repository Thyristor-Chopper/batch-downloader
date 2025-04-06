VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cSubItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_hWnd As Long
Private m_lPtr As Long
Private m_lId As Long
Private m_lIndex As Long

Friend Function fInit(ByVal hWnd As Long, ByVal lId As Long, ByVal lPtr As Long)
   m_hWnd = hWnd
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
               gErr 3, "cSubItems"
            End If
         Else
            gErr 3, "cSubItems"
         End If
      Else
         gErr 1, "cSubItems"
      End If
   Else
      gErr 1, "cSubItems"
   End If
End Function

Public Property Get Item(ByVal Index As Long) As cSubItem
Attribute Item.VB_UserMemId = 0
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      If (Index > 0) And (Index < ctl.fColumnCount) Then
         Dim cSI As New cSubItem
         cSI.fInit m_hWnd, m_lId, m_lPtr, Index
         Set Item = cSI
      Else
         gErr 6, "cSubItems"
      End If
   End If
End Property

Public Property Get Count() As Long
Dim ctl As vbalListViewCtl
   If pbVerify(ctl) Then
      Count = ctl.fColumnCount - 1
   End If
End Property
