Attribute VB_Name = "Functions"
Public fso

'https://www.vbforums.com/showthread.php?894947-How-to-test-if-a-font-is-available
Function FontExists(Name As String) As Boolean
    With New StdFont
        .Name = Name
        FontExists = (StrComp(.Name, Name, vbTextCompare) = 0)
    End With
End Function

Function FolderExists(sFullPath As String) As Boolean
    Dim myFSO As Object
    Set myFSO = CreateObject("Scripting.FileSystemObject")
    FolderExists = myFSO.FolderExists(sFullPath)
End Function

Function Floor(ByVal floatval As Double, Optional decimalPlaces As Long = 0) As Long
    Dim intval As Long
    intval = Round(floatval)
    If intval > floatval Then
         intval = intval - 1
    End If

    If decimalPlaces > 0 Then
        floatval = Float / (10 ^ decimalPlaces)
    End If

    Floor = intval
End Function

