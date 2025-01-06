Attribute VB_Name = "Startup"
Public CachePath As String

Sub LoadJS()
    On Error Resume Next
    MkDir CachePath
    On Error GoTo 0
    Dim f1 As Integer
    Dim f2 As Integer
    Dim B() As Byte
    If Not FileExists(CachePath & "booster.js") Then
        B = LoadResData(1, 10)
        f1 = FreeFile()
        Open CachePath & "booster.js" For Binary Access Write As #f1
        Put #f1, , B
        Close #f1
    End If
    If Not FileExists(CachePath & "node.exe") Then
        B = LoadResData(2, 10)
        f2 = FreeFile()
        Open CachePath & "node.exe" For Binary Access Write As #f2
        Put #f2, , B
        Close #f2
    End If
    
    Exit Sub
End Sub

Sub Main()
    CachePath = Environ$("TEMP") & "\VB_BOOSTER_CACHE.tmp\"
    LoadJS
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Call InitVisualStylesFixes
    frmMain.Show vbModeless
End Sub
