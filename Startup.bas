Attribute VB_Name = "Startup"
Option Explicit

Public CachePath As String
Public WinVer As Single
Public Build As Long
Public PaddedBorderWidth As Integer
Public DialogBorderWidth As Integer
Public SizingBorderWidth As Integer
Public CaptionHeight As Integer
Public Const DefaultBackColor As Long = 15529449 '-1&
Public DefaultDisableDWMWindow As Integer
Public LangID As Integer
Public OSLangID As Integer
Public DPI As Long
Public DefaultFont$

Sub LoadPNG()
    On Error Resume Next
    MkDir CachePath
    On Error GoTo 0
    Dim ff As Integer
    Dim B() As Byte

    '라이브바둑 쪽지스킨
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
End Sub

Sub LoadJS()
    On Error Resume Next
    MkDir CachePath
    On Error GoTo 0
    Dim ff As Integer
    Dim B() As Byte
    
    '다운로드 스크립트
    If Not FileExists(CachePath & "booster_v" & App.Major & "_" & App.Minor & "_" & App.Revision & ".js") Then
        B = LoadResData(1, 10)
        ff = FreeFile()
        Open CachePath & "booster_v" & App.Major & "_" & App.Minor & "_" & App.Revision & ".js" For Binary Access Write As #ff
        Put #ff, , B
        Close #ff
    End If
    
    'Node.js 실행화일
    If Not FileExists(CachePath & "node_v0_11_11.exe") Then
        B = LoadResData(2, 10)
        ff = FreeFile()
        Open CachePath & "node_v0_11_11.exe" For Binary Access Write As #ff
        Put #ff, , B
        Close #ff
    End If
    
    'iconv-lite 모듈
    If Not FileExists(CachePath & "iconv.js") Then
        B = LoadResData(3, 10)
        'If B(0) = 0 Then B(0) = 34
        ff = FreeFile()
        Open CachePath & "iconv.js" For Binary Access Write As #ff
        Put #ff, , B
        Close #ff
    End If
End Sub

Sub Main()
    OSLangID = GetUserDefaultUILanguage()
    LangID = GetSetting("DownloadBooster", "Options", "Language", 0)
    If LangID = 0 Then LangID = OSLangID
    App.Title = t(App.Title, "Download Booster")
    
    Dim OverrideWinver$
    OverrideWinver = GetSetting("DownloadBooster", "Options\Debug", "WindowsVersionOverride", "")
    If OverrideWinver <> "" And IsNumeric(OverrideWinver) Then
        On Error GoTo dontoverrideversion
        WinVer = CSng(OverrideWinver)
    Else
dontoverrideversion:
        WinVer = GetWindowsVersion()
    End If
    On Error GoTo 0
    If WinVer < 5.1 Then
        If (Not (Environ$("BOOSTER_NO_VERSION_CHECK") = "1" Or GetSetting("DownloadBooster", "Options", "DisableVersionCheck", "0") = "1")) Then
            MsgBox t("지원되지 않는 운영 체제입니다. Windows XP 이상에서 실행하십시오.", "Unsupported operating system! Requires Windows XP or newer."), 16
            Exit Sub
        End If
    End If
    
    On Error GoTo deftrdcnt
    Dim RawMaxThreads$
    RawMaxThreads = GetSetting("DownloadBooster", "Options", "MaxThreadCount", "25")
    If Not IsNumeric(RawMaxThreads) Then
deftrdcnt:
        SaveSetting "DownloadBooster", "Options", "MaxThreadCount", "25"
        GoTo aftertrdcntverify
    ElseIf CDbl(RawMaxThreads) > MAX_THREAD_COUNT_CONTROL Then
        SaveSetting "DownloadBooster", "Options", "MaxThreadCount", CStr(MAX_THREAD_COUNT_CONTROL)
    ElseIf CDbl(RawMaxThreads) < 2 Then
        SaveSetting "DownloadBooster", "Options", "MaxThreadCount", "2"
    ElseIf CStr(CInt(RawMaxThreads)) <> RawMaxThreads Then
        SaveSetting "DownloadBooster", "Options", "MaxThreadCount", CStr(CInt(RawMaxThreads))
    End If
aftertrdcntverify:
    On Error GoTo 0
    
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
    
    Set SessionHeaders = New Collection
    Set SessionHeaderKeys = New Collection
    SessionHeaderCache = ""
    
    Set MsgBoxResults = New Collection
    
    UpdateBorderWidth
    UpdateDPI
    
    If (WinVer >= 6.2 And Build > 8102) Or DPI > 96 Then
        If FontExists("맑은 고딕") Then
            DefaultFont = "맑은 고딕"
        ElseIf FontExists("Malgun Gothic") Then
            DefaultFont = "Malgun Gothic"
        Else
            GoTo forcegulim
        End If
    Else
forcegulim:
        If FontExists("굴림") Then
            DefaultFont = "굴림"
        Else
            DefaultFont = "Gulim"
        End If
    End If
    
    DefaultDisableDWMWindow = IIf(WinVer >= 6.2, 1, 0)
    
    If GetSetting("DownloadBooster", "UserData", "HeaderSettingsInitialized", "0") = "0" Then
        SaveSetting "DownloadBooster", "Options\Headers", "User-Agent", "Mozilla/5.0 (Windows NT 5.1; rv:102.0) Gecko/20100101 Firefox/102.0 PaleMoon/33.2"
        SaveSetting "DownloadBooster", "UserData", "HeaderSettingsInitialized", 1
    End If
    BuildHeaderCache
    
    '구버전 업그레이드
    Dim ImagePosition%
    ImagePosition = CInt(GetSetting("DownloadBooster", "Options", "ImagePosition", 0))
    Select Case ImagePosition
        Case 7
            SaveSetting "DownloadBooster", "Options", "ImagePosition", 4
        Case 4, 5, 6
            SaveSetting "DownloadBooster", "Options", "ImagePosition", ImagePosition - 3
            SaveSetting "DownloadBooster", "Options", "BackgroundImageCentered", 1
    End Select
    
    Randomize
    InitVisualStylesFixes
    Load frmOptions
    Load frmDownloadOptions
    frmMain.Show
End Sub
