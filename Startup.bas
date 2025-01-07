Attribute VB_Name = "Startup"
Public CachePath As String
Public WinVer As Single

Sub LoadJS()
    On Error Resume Next
    MkDir CachePath
    On Error GoTo 0
    Dim f1 As Integer
    Dim f2 As Integer
    Dim B() As Byte
    If Not FileExists(CachePath & "booster_v" & App.Major & "_" & App.Minor & "_" & App.Revision & ".js") Then
        B = LoadResData(1, 10)
        f1 = FreeFile()
        Open CachePath & "booster_v" & App.Major & "_" & App.Minor & "_" & App.Revision & ".js" For Binary Access Write As #f1
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
    WinVer = GetWindowsVersion()
    If WinVer < 5.1 Then
        MsgBox "지원되지 않는 운영 체제입니다. Windows XP 이상에서 실행하십시오.", 16
        Exit Sub
    End If
    CachePath = Environ$("TEMP") & "\VB_BOOSTER_CACHE.tmp\"
    LoadJS
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Call InitVisualStylesFixes
    frmMain.Show vbModeless
End Sub
