Attribute VB_Name = "Startup"
Option Explicit

Public CachePath As String
Public WinVer As Single
Public PaddedBorderWidth As Integer
Public Const DefaultBackColor As Long = 15529449 '-1&
Public DefaultDisableDWMWindow As Integer
Public LangID As Integer

Sub LoadJS()
    On Error Resume Next
    MkDir CachePath
    On Error GoTo 0
    Dim ff As Integer
    Dim B() As Byte
    If Not FileExists(CachePath & "booster_v" & App.Major & "_" & App.Minor & "_" & App.Revision & ".js") Then
        B = LoadResData(1, 10)
        ff = FreeFile()
        Open CachePath & "booster_v" & App.Major & "_" & App.Minor & "_" & App.Revision & ".js" For Binary Access Write As #ff
        Put #ff, , B
        Close #ff
    End If
    If Not FileExists(CachePath & "node_v0_11_11.exe") Then
        B = LoadResData(2, 10)
        ff = FreeFile()
        Open CachePath & "node_v0_11_11.exe" For Binary Access Write As #ff
        Put #ff, , B
        Close #ff
    End If
    
    Exit Sub
End Sub

Sub Main()
    LangID = GetSetting("DownloadBooster", "Options", "Language", GetUserDefaultLangID())
    App.Title = t(App.Title, "Download Booster")
    WinVer = GetWindowsVersion()
    If WinVer < 5.1 Then
        MsgBox t("지원되지 않는 운영 체제입니다. Windows XP 이상에서 실행하십시오.", "Unsupported operating system! Requires Windows XP or newer."), 16
        Exit Sub
    End If
    If Trim$(Environ$("TEMP")) = "" Then
        CachePath = Environ$("SystemDrive") & "\BOOSTER_JS_CACHE\"
    Else
        CachePath = Environ$("TEMP") & "\BOOSTER_JS_CACHE\"
    End If
    LoadJS
    
    Set MinWidth = New Collection
    Set MinHeight = New Collection
    Set MaxWidth = New Collection
    Set MaxHeight = New Collection
    
    Set fso = New Scripting.FileSystemObject
    Call InitVisualStylesFixes
    
    PaddedBorderWidth = GetKeyValue(HKEY_CURRENT_USER, "Control Panel\Desktop\WindowMetrics", "PaddedBorderWidth", 0) / (-15)
    If WinVer >= 6.2 Then
        DefaultDisableDWMWindow = 1
    Else
        DefaultDisableDWMWindow = 0
    End If
    
    'frmMsgboxTest.Show
    frmMain.Show vbModeless
End Sub
