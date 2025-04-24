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
Public MainFormOnTop As Boolean

Public ScriptFileName As String
Public NodeFileName As String

Sub Main()
    OSLangID = GetUserDefaultUILanguage()
    LangID = GetSetting("DownloadBooster", "Options", "Language", 0)
    If LangID = 0 Then LangID = OSLangID

    App.Title = t(App.Title, "Download Booster")

'    Dim OverrideWinver$
'    OverrideWinver = GetSetting("DownloadBooster", "Options\Debug", "WindowsVersionOverride", "")
'    If OverrideWinver <> "" And IsNumeric(OverrideWinver) Then
'        On Error GoTo dontoverrideversion
'        WinVer = CSng(OverrideWinver)
'    Else
'dontoverrideversion:
        WinVer = GetWindowsVersion()
'    End If
'    On Error GoTo 0

    If WinVer < 5.1 Then
        If (Not (Environ$("BOOSTER_NO_VERSION_CHECK") = "1" Or GetSetting("DownloadBooster", "Options", "DisableVersionCheck", "0") = "1")) Then
            MsgBox t("지원되지 않는 운영 체제입니다. Windows XP 이상에서 실행하십시오.", "Unsupported operating system! Requires Windows XP or newer."), 16
            Exit Sub
        End If
    End If

'    On Error GoTo deftrdcnt
'    Dim RawMaxThreads$
'    RawMaxThreads = GetSetting("DownloadBooster", "Options", "MaxThreadCount", "25")
'    If Not IsNumeric(RawMaxThreads) Then
'deftrdcnt:
'        SaveSetting "DownloadBooster", "Options", "MaxThreadCount", "25"
'        GoTo aftertrdcntverify
'    ElseIf CDbl(RawMaxThreads) > MAX_THREAD_COUNT_CONTROL Then
'        SaveSetting "DownloadBooster", "Options", "MaxThreadCount", CStr(MAX_THREAD_COUNT_CONTROL)
'    ElseIf CDbl(RawMaxThreads) < 2 Then
'        SaveSetting "DownloadBooster", "Options", "MaxThreadCount", "2"
'    ElseIf CStr(CInt(RawMaxThreads)) <> RawMaxThreads Then
'        SaveSetting "DownloadBooster", "Options", "MaxThreadCount", CStr(CInt(RawMaxThreads))
'    End If
'aftertrdcntverify:
'    On Error GoTo 0

    Dim CachePathSuffix$
    CachePathSuffix = "\BOOSTER_JS_CACHE\"
    If LenB(Trim$(Environ$("TEMP"))) = 0 Then
        If LenB(Environ$("SystemDrive")) = 0 Then
            CachePath = "C:" & CachePathSuffix
        Else
            CachePath = Environ$("SystemDrive") & CachePathSuffix
        End If
    Else
        CachePath = Environ$("TEMP") & CachePathSuffix
    End If
    ScriptFileName = "booster_v" & App.Major & "_" & App.Minor & "_" & App.Revision & ".js"
    NodeFileName = "node_v0_11_11.exe"
    
    ExtractResource 1, RCData, ScriptFileName
    ExtractResource 2, RCData, NodeFileName
    ExtractResource 3, RCData, "iconv.js"

    Set SessionHeaders = New Collection
    Set SessionHeaderKeys = New Collection
    SessionHeaderCache = ""

    Set MsgBoxResults = New Collection

    UpdateBorderWidth
    UpdateDPI

    If (WinVer >= 6.2 And Build > 8102) Or DPI <> 96 Then
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
        ElseIf FontExists("Gulim") Then
            DefaultFont = "Gulim"
        Else
            DefaultFont = "Tahoma"
        End If
    End If

    If WinVer >= 6.2 Then DefaultDisableDWMWindow = 1 Else DefaultDisableDWMWindow = 0

    If GetSetting("DownloadBooster", "UserData", "HeaderSettingsInitialized", "0") = "0" Then
        SaveSetting "DownloadBooster", "Options\Headers", "User-Agent", "Mozilla/5.0 (Windows NT 5.1; rv:102.0) Gecko/20100101 Firefox/102.0 PaleMoon/33.2"
        SaveSetting "DownloadBooster", "UserData", "HeaderSettingsInitialized", 1
    End If
    BuildHeaderCache

    Randomize
    InitVisualStylesFixes
    Load frmDummyForm
    Load frmMain
End Sub
