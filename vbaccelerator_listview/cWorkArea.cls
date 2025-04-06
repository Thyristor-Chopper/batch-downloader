VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cWorkArea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_hWnd As Long
Private m_lIndex As Long
Private m_lId As Long

Friend Sub fInit(ByVal hWnd As Long, ByVal lId As Long)
   m_hWnd = hWnd
   m_lId = lId
End Sub
Private Function pbVerify(ByRef ctlThis As vbalListViewCtl) As Boolean
Dim lPtr As Long
   If IsWindow(m_hWnd) Then
      lPtr = GetProp(m_hWnd, gcObjectProp)
      If Not (lPtr = 0) Then
         Set ctlThis = ObjectFromPtr(lPtr)
         m_lIndex = ctlThis.fWorkAreaIndexForId(m_lId)
         If (m_lIndex > 0) Then
            pbVerify = True
         Else
            gErr 9, "cWorkArea"
         End If
      Else
         gErr 1, "cWorkArea"
      End If
   Else
      gErr 1, "cWorkArea"
   End If
End Function
Public Property Get Key() As String
Dim ctl As vbalListViewCtl
   If (pbVerify(ctl)) Then
      Key = ctl.fWorkAreaKey(m_lIndex)
   End If
End Property
Public Property Let Key(ByVal sKey As String)
Dim ctl As vbalListViewCtl
   If (pbVerify(ctl)) Then
      ctl.fWorkAreaKey(m_lIndex) = sKey
   End If
End Property
Public Property Get ItemData() As Long
Dim ctl As vbalListViewCtl
   If (pbVerify(ctl)) Then
      ItemData = ctl.fWorkAreaItemData(m_lIndex)
   End If
End Property
Public Property Let ItemData(ByVal lItemData As Long)
Dim ctl As vbalListViewCtl
   If (pbVerify(ctl)) Then
      ctl.fWorkAreaItemData(m_lIndex) = lItemData
   End If
End Property
Public Property Get Tag() As String
Dim ctl As vbalListViewCtl
   If (pbVerify(ctl)) Then
      Tag = ctl.fWorkAreaTag(m_lIndex)
   End If
End Property
Public Property Let Tag(ByVal sTag As String)
Dim ctl As vbalListViewCtl
   If (pbVerify(ctl)) Then
      ctl.fWorkAreaTag(m_lIndex) = sTag
   End If
End Property
Public Property Get left() As Single
Dim ctl As vbalListViewCtl
   If (pbVerify(ctl)) Then
      Dim X As Single
      ctl.fScale ctl.fWorkAreaLeft(m_lIndex), 0, X, 0
      left = X
   End If
End Property
Public Property Let left(ByVal fLeft As Single)
Dim ctl As vbalListViewCtl
   If (pbVerify(ctl)) Then
      Dim X As Long
      ctl.fUnScale fLeft, 0, X, 0
      ctl.fWorkAreaLeft(m_lIndex) = X
   End If
End Property
Public Property Get top() As Single
Dim ctl As vbalListViewCtl
   If (pbVerify(ctl)) Then
      Dim Y As Single
      ctl.fScale 0, ctl.fWorkAreaTop(m_lIndex), 0, Y
      top = Y
   End If
End Property
Public Property Let top(ByVal fTop As Single)
Dim ctl As vbalListViewCtl
   If (pbVerify(ctl)) Then
      Dim Y As Long
      ctl.fUnScale 0, top, 0, Y
      ctl.fWorkAreaTop(m_lIndex) = Y
   End If
End Property
Public Property Get Width() As Single
Dim ctl As vbalListViewCtl
   If (pbVerify(ctl)) Then
      Dim X As Single
      ctl.fScale ctl.fWorkAreaRight(m_lIndex) - ctl.fWorkAreaLeft(m_lIndex), 0, X, 0
      Width = X
   End If
End Property
Public Property Let Width(ByVal fWidth As Single)
Dim ctl As vbalListViewCtl
   If (pbVerify(ctl)) Then
      Dim X As Long
      ctl.fUnScale fWidth, 0, X, 0
      ctl.fWorkAreaRight(m_lIndex) = X + ctl.fWorkAreaLeft(m_lIndex)
   End If
End Property
Public Property Get Height() As Single
Dim ctl As vbalListViewCtl
   If (pbVerify(ctl)) Then
      Dim Y As Single
      ctl.fScale 0, ctl.fWorkAreaBottom(m_lIndex) - ctl.fWorkAreaTop(m_lIndex), 0, Y
      Height = Y
   End If
End Property
Public Property Let Height(ByVal fHeight As Single)
Dim ctl As vbalListViewCtl
   If (pbVerify(ctl)) Then
      Dim Y As Long
      ctl.fUnScale 0, fHeight, 0, Y
      ctl.fWorkAreaBottom(m_lIndex) = Y + ctl.fWorkAreaTop(m_lIndex)
   End If
End Property
Public Property Get Index() As Long
Dim ctl As vbalListViewCtl
   If (pbVerify(ctl)) Then
      Index = m_lIndex
   End If
End Property
