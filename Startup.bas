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
    
    '���̺�ٵ� ������Ų
    If Not FileExists(CachePath & "bottom.png") Then
        B = LoadResData(101, 10)
        ff = FreeFile()
        Open CachePath & "bottom.png" For Binary Access Write As #ff
        Put #ff, , B
        Close #ff
    End If
    If Not FileExists(CachePath & "bottomleft.png") Then
        B = LoadResData(102, 10)
        ff = FreeFile()
        Open CachePath & "bottomleft.png" For Binary Access Write As #ff
        Put #ff, , B
        Close #ff
    End If
    If Not FileExists(CachePath & "bottomright.png") Then
        B = LoadResData(103, 10)
        ff = FreeFile()
        Open CachePath & "bottomright.png" For Binary Access Write As #ff
        Put #ff, , B
        Close #ff
    End If
    If Not FileExists(CachePath & "left.png") Then
        B = LoadResData(104, 10)
        ff = FreeFile()
        Open CachePath & "left.png" For Binary Access Write As #ff
        Put #ff, , B
        Close #ff
    End If
    If Not FileExists(CachePath & "right.png") Then
        B = LoadResData(105, 10)
        ff = FreeFile()
        Open CachePath & "right.png" For Binary Access Write As #ff
        Put #ff, , B
        Close #ff
    End If
    If Not FileExists(CachePath & "top.png") Then
        B = LoadResData(106, 10)
        ff = FreeFile()
        Open CachePath & "top.png" For Binary Access Write As #ff
        Put #ff, , B
        Close #ff
    End If
    If Not FileExists(CachePath & "topleft.png") Then
        B = LoadResData(107, 10)
        ff = FreeFile()
        Open CachePath & "topleft.png" For Binary Access Write As #ff
        Put #ff, , B
        Close #ff
    End If
    If Not FileExists(CachePath & "topright.png") Then
        B = LoadResData(108, 10)
        ff = FreeFile()
        Open CachePath & "topright.png" For Binary Access Write As #ff
        Put #ff, , B
        Close #ff
    End If
    If Not FileExists(CachePath & "center.png") Then
        B = LoadResData(109, 10)
        ff = FreeFile()
        Open CachePath & "center.png" For Binary Access Write As #ff
        Put #ff, , B
        Close #ff
    End If
    
    Exit Sub
End Sub

Sub Main()
    LangID = GetSetting("DownloadBooster", "Options", "Language", GetUserDefaultLangID())
    If LangID = 0 Then LangID = GetUserDefaultLangID()
    App.Title = t(App.Title, "Download Booster")
    WinVer = GetWindowsVersion()
    If WinVer < 5.1 Then
        If (Not (Environ$("BOOSTER_NO_VERSION_CHECK") = "1" Or GetSetting("DownloadBooster", "Options", "DisableVersionCheck", "0") = "1")) Then
            MsgBox t("�������� �ʴ� � ü���Դϴ�. Windows XP �̻󿡼� �����Ͻʽÿ�.", "Unsupported operating system! Requires Windows XP or newer."), 16
            Exit Sub
        End If
    End If
    If Trim$(Environ$("TEMP")) = "" Then
        If Environ$("SystemDrive") = "" Then
            CachePath = "C:\BOOSTER_JS_CACHE\"
        Else
            CachePath = Environ$("SystemDrive") & "\BOOSTER_JS_CACHE\"
        End If
    Else
        CachePath = Environ$("TEMP") & "\BOOSTER_JS_CACHE\"
    End If
    LoadJS
    
    Set MinWidth = New Collection
    Set MinHeight = New Collection
    Set MaxWidth = New Collection
    Set MaxHeight = New Collection
    
    Set SessionHeaders = New Dictionary
    SessionHeaderCache = ""
    
    Set fso = New Scripting.FileSystemObject
    Call InitVisualStylesFixes
    
    PaddedBorderWidth = GetKeyValue(HKEY_CURRENT_USER, "Control Panel\Desktop\WindowMetrics", "PaddedBorderWidth", 0) / (-15)
    If WinVer >= 6.2 Then
        DefaultDisableDWMWindow = 1
    Else
        DefaultDisableDWMWindow = 0
    End If
    
    If GetSetting("DownloadBooster", "UserData", "HeaderSettingsInitialized", "0") = "0" Then
        SaveSetting "DownloadBooster", "UserData", "HeaderSettingsInitialized", 1
        SaveSetting "DownloadBooster", "Options\Headers", "User-Agent", "Mozilla/5.0 (Windows NT 5.1; rv:102.0) Gecko/20100101 Firefox/102.0 PaleMoon/33.2"
    End If
    BuildHeaderCache
    
    'frmMsgboxTest.Show
    Functions.AppExiting = False
    frmMain.Show vbModeless
End Sub
