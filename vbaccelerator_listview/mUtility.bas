Attribute VB_Name = "mUtility"
Option Explicit

Private m_lId As Long
Private m_lColID As Long

Public Const gcObjectProp = "vbalListViewCtl:ObjectPtr"
Public gsInfoTipBuffer As String

#Const DEBUGMODE = 1

Public Property Get NextItemID() As Long
   
   ' Get the ID:
   m_lId = m_lId + 1
   NextItemID = m_lId
   
   ' Wrap around every 4 billion items that
   ' get created :)
   If m_lId > 2147483646 Then
      m_lId = -2147483647
   End If
   
End Property

Public Property Get NextColumnID() As Long
   
   ' Get the ID:
   m_lColID = m_lColID + 1
   NextColumnID = m_lColID
   
   ' Wrap around every 4 billion items that
   ' get created :)
   If m_lColID > 2147483646 Then
      m_lColID = -2147483647
   End If
   
End Property


Public Property Get ObjectFromPtr(ByVal lPtr As Long) As Object
Dim objT As Object
   If Not (lPtr = 0) Then
      ' Turn the pointer into an illegal, uncounted interface
      CopyMemory objT, lPtr, 4
      ' Do NOT hit the End button here! You will crash!
      ' Assign to legal reference
      Set ObjectFromPtr = objT
      ' Still do NOT hit the End button here! You will still crash!
      ' Destroy the illegal reference
      CopyMemory objT, 0&, 4
   End If
End Property

Public Function TranslateColor(ByVal oClr As OLE_COLOR, _
                        Optional hPal As Long = 0) As Long
    ' Convert Automation color to Windows color
    If OleTranslateColor(oClr, hPal, TranslateColor) Then
        TranslateColor = CLR_INVALID
    End If
End Function

Public Sub gErr(ByVal lErrNum As Long, ByVal sSource As String)
Dim sDesc As String
Debug.Assert False
   
   On Error GoTo 0
   
   Select Case lErrNum
   Case 1
      ' Cannot find owner object
      lErrNum = 364
      sDesc = "Object has been unloaded."
   
   Case 2
      ' Bar does not exist
      lErrNum = vbObjectError + 25001
      sDesc = "ListBar does not exist."
      
   Case 3
      ' Item does not exist
      lErrNum = vbObjectError + 25002
      sDesc = "ListItem does not exist."
      
   Case 4
      ' Invalid key: numeric
      lErrNum = 13
      sDesc = "Type Mismatch."
      
   Case 5
      ' Invalid Key: duplicate
      lErrNum = 457
      sDesc = "This key is already associated with an element of this collection."
   
   Case 6
      ' Subscript out of range
      lErrNum = 9
      sDesc = "Subscript out of range."
      
   Case 7
      ' Failed to add a resource/out of memory
      lErrNum = 7
      sDesc = "Out of Memory."
      
   Case 8
      ' Header does not exist
      lErrNum = vbObjectError + 25003
      sDesc = "Header does not exist."
      
   Case 9
      ' can't set grouping
      lErrNum = vbObjectError + 25004
      sDesc = "Failed to set group enable state."
   
   Case 10
      lErrNum = vbObjectError + 25005
      sDesc = "SubItem does not exist."
      
   Case Else
      Debug.Assert "Unexpected Error" = ""
   
   End Select
   
   
   Err.Raise lErrNum, App.EXEName & "." & sSource, sDesc
End Sub

Public Sub SetVariant(vToSet As Variant, vSetWith As Variant)
   If IsMissing(vSetWith) Then
      Set vToSet = Nothing
   ElseIf IsObject(vSetWith) Then
      Set vToSet = vSetWith
   Else
      vToSet = vSetWith
   End If
End Sub
