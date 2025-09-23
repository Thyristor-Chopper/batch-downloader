Attribute VB_Name = "Startup"
Option Explicit

#If BETA Then
Public Const BetaVer As Byte = 0
#End If

Public CachePath As String
Public WinVer As Single
Public Build As Long
Public PaddedBorderWidth As Byte
Public DialogBorderWidth As Byte
Public SizingBorderWidth As Byte
Public ScrollBarWidth As Byte
Public CaptionHeight As Byte
Public Const DefaultBackColor As Long = 15529449 '-1&
Public DefaultDisableDWMWindow As Byte
Public LangID As Integer
Public OSLangID As Integer
Public DPI As Long
Public DefaultFont As String
Public MainFormOnTop As Boolean

Public NodeFileName As String

Public DarkTransparent As IPicture
Public LightTransparent As IPicture
Public Train(1 To 3) As IPicture
Public TygemButtonTexture(0 To 2) As IPicture

Public NodeJS() As Byte
Public DownloadScript As String

Public IPictureIID As IID

Public Is64 As Boolean
Public LaunchFromMemory As Boolean

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
        If (Not (Environ$("BOOSTER_NO_VERSION_CHECK") = "1" Or GetSetting("DownloadBooster", "Options", "DisableVersionCheck", "0") <> "0")) Then
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

    With IPictureIID
        .Data1 = &H7BF80980
        .Data2 = &HBF32
        .Data3 = &H101A
        .Data4(0) = &H8B
        .Data4(1) = &HBB
        .Data4(2) = &H0
        .Data4(3) = &HAA
        .Data4(4) = &H0
        .Data4(5) = &H30
        .Data4(6) = &HC
        .Data4(7) = &HAB
    End With
    
    Is64 = IsWOW64()
    LaunchFromMemory = (GetSetting("DownloadBooster", "Options", "RunDownloaderInMemory", "1") <> "0")

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
    
    DownloadScript = MinifyScript(ConvertUTF8(LoadResData(1, RCData)))
    If LaunchFromMemory Then
        NodeJS = LoadResData(2, RCData)
    Else
        NodeFileName = "node_v5_12_0.exe"
        ExtractResource 2, RCData, NodeFileName
    End If

    Set MsgBoxResults = New Collection
    Set SessionHeaders = New Collection
    Set SessionHeaderKeys = New Collection
    SessionHeaderCache = ""

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

    DefaultDisableDWMWindow = -(WinVer >= 6.2)

    If GetSetting("DownloadBooster", "UserData", "HeaderSettingsInitialized", "0") = "0" Then
        SaveSetting "DownloadBooster", "Options\Headers", "User-Agent", "Mozilla/5.0 (Windows NT 6.1; rv:115.0) Gecko/20100101 Firefox/115.0 PaleMoon/33.7.0"
        SaveSetting "DownloadBooster", "UserData", "HeaderSettingsInitialized", 1
    End If
    BuildHeaderCache
    
    Dim i As Byte
    For i = 1 To 3
        Set TygemButtonTexture(i - 1) = LoadResPicture(i, vbResBitmap)
        Set Train(i) = LoadResPicture(i + 1, vbResIcon)
    Next i

    Randomize
    InitVisualStylesFixes
    Load frmDummyForm
    frmMain.Show vbModeless
    frmMain.SetFrameTexture
End Sub
