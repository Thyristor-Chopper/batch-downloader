Attribute VB_Name = "Functions"
'참고자료
'- https://www.vbforums.com/showthread.php?457171-RESOLVED-How-to-get-Desktop-Path-in-VB
'- https://www.vbforums.com/showthread.php?445574-Reading-shortcut-information
'- https://www.vbforums.com/showthread.php?430704-RESOLVED-Get-drive-size-space
'- https://www.codeguru.com/visual-basic/displaying-the-file-properties-dialog/
'- http://vbcity.com/forums/t/105530.aspx
'- https://www.vbforums.com/showthread.php?644597-How-to-get-String-from-Pointer-in-VB
'- https://www.vbforums.com/showthread.php?903019-UxTheme-dll-Visual-Styles-help-request-button-face-color-when-mouse-hovers-over

Option Explicit

Public MsgBoxResults As Collection
Declare Function MessageBeep Lib "user32" (ByVal wType As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function RtlGetVersion Lib "ntdll" (lpVersionInformation As OSVERSIONINFO) As Long
Declare Function DwmSetWindowAttribute Lib "dwmapi.dll" (ByVal hWnd As Long, ByVal dwAttribute As Long, ByRef pvAttribute As Long, ByVal cbAttribute As Long) As Long
Private Declare Function DwmIsCompositionEnabled Lib "dwmapi.dll" (ByRef pfEnabled As Long) As Long
Declare Function DwmEnableComposition Lib "dwmapi.dll" (ByVal uCompositionAction As Long) As Long
Declare Function DwmExtendFrameIntoClientArea Lib "dwmapi.dll" (ByVal hWnd As Long, Margin As MARGINS) As Long
Declare Sub DwmGetColorizationParameters Lib "dwmapi.dll" Alias "#127" (ByRef Parameters As DWM_COLORIZATION_PARAMS)
Declare Sub DwmSetColorizationParameters Lib "dwmapi.dll" Alias "#131" (ByRef Parameters As DWM_COLORIZATION_PARAMS, Optional ByVal DoNotSaveToRegistry As Boolean = True)
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Declare Function GetUserDefaultUILanguage Lib "kernel32" () As Integer
Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Any) As Long
Declare Function InsertMenuItem Lib "user32" Alias "InsertMenuItemA" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpMII As MENUITEMINFO) As Long
Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpMII As MENUITEMINFO) As Long
Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpMII As MENUITEMINFO) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As Long) As Long
Declare Function CheckMenuRadioItem Lib "user32" (ByVal hMenu As Long, ByVal un1 As Long, ByVal un2 As Long, ByVal un3 As Long, ByVal un4 As Long) As Long
Private Declare Function CryptBinaryToString Lib "crypt32" Alias "CryptBinaryToStringW" (ByVal pbBinary As Long, ByVal cbBinary As Long, ByVal dwFlags As Long, ByVal pszString As Long, ByRef pcchString As Long) As Long
Private Const CRYPT_STRING_BASE64 As Long = 1
Private Declare Function CryptStringToBinary Lib "crypt32" Alias "CryptStringToBinaryW" (ByVal pszString As Long, ByVal cchString As Long, ByVal dwFlags As Long, ByVal pbBinary As Long, ByRef pcbBinary As Long, ByRef pdwSkip As Long, ByRef pdwFlags As Long) As Long
Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ExpandEnvironmentStringsA Lib "kernel32" (ByVal lpSrc As String, ByVal lpDst As String, ByVal nSize As Long) As Long
Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Declare Function PrintWindow Lib "user32" (ByVal hWnd As Long, ByVal hdcBlt As Long, ByVal nFlags As Long) As Long
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyW" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long
Private Declare Function SysAllocStringByteLen Lib "oleaut32.dll" (Optional ByVal pszStrPtr As Long, Optional ByVal Length As Long) As String
Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal IDHook As Long, ByVal lpfn As Long, ByVal hMod As Long, ByVal dwThreadID As Long) As Long
Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function X_GetThemeColor Lib "uxtheme.dll" Alias "GetThemeColor" (ByVal hTheme As Long, ByVal iPartId As Long, ByVal iStateId As Long, ByVal iPropId As Long, pColor As Long) As Long
Private Declare Function IsAppThemed Lib "uxtheme.dll" () As Long
Private Declare Function IsThemeActive Lib "uxtheme.dll" () As Long
Private Declare Function OpenThemeData Lib "uxtheme.dll" (ByVal hWnd As Long, ByVal pszClassList As Long) As Long
Private Declare Function CloseThemeData Lib "uxtheme.dll" (ByVal hTheme As Long) As Long
Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, Optional ByVal dwDuration As Long = 250) As Long
Declare Function FlashWindow Lib "user32" (ByVal hWnd As Long, Optional ByVal bInvert As Long = 1&) As Long
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hWndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal lpRootPathName As String) As Long
Private Declare Function GetDiskFreeSpaceEx Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, lpFreeBytesAvailableToCaller As LARGE_INTEGER, lpTotalNumberOfBytes As LARGE_INTEGER, lpTotalNumberOfFreeBytes As LARGE_INTEGER) As Long
Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (lpChooseColor As ChooseColorStruct) As Long
Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long
Private Declare Function ShellExecuteEx Lib "shell32" (ByRef s As SHELLEXECUTEINFO) As Long
Private Declare Function GetKeyState Lib "user32" (ByVal vKey As Long) As Integer

Public Const DWM_EC_DISABLECOMPOSITION As Long = 0
Public Const DWM_EC_ENABLECOMPOSITION As Long = 1

Public Const SND_APPLICATION = &H80
Public Const SND_ALIAS = &H10000
Public Const SND_ALIAS_ID = &H110000
Public Const SND_ASYNC = &H1
Public Const SND_FILENAME = &H20000
Public Const SND_LOOP = &H8
Public Const SND_MEMORY = &H4
Public Const SND_NODEFAULT = &H2
Public Const SND_NOSTOP = &H10
Public Const SND_NOWAIT = &H2000
Public Const SND_PURGE = &H40
Public Const SND_RESOURCE = &H40004
Public Const SND_SYNC = &H0

Public Const TMT_TEXTCOLOR As Long = 3803

Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2
Public Const HTBOTTOM = 15
Public Const HTLEFT = 10
Public Const HTRIGHT = 11
Public Const HTBOTTOMLEFT = 16
Public Const HTBOTTOMRIGHT = 17

Type DWM_COLORIZATION_PARAMS
    ColorBlue  As Byte
    ColorGreen As Byte
    ColorRed   As Byte
    ColorAlpha As Byte
    
    AfterGlowBlue  As Byte
    AfterGlowGreen As Byte
    AfterGlowRed   As Byte
    AfterGlowAlpha As Byte
    
    ColorBalance     As Long
    AfterGlowBalance As Long
    BlurBalance      As Long
    
    StripesIntensity As Long
    
    Opaque As Boolean
End Type

Type MARGINS
    cxLeftWidth    As Long
    cxRightWidth   As Long
    cyTopHeight    As Long
    cyBottomHeight As Long
End Type

Enum AudioFormat
    Auto = 0
    MP3 = 1
    WAV = 2
    FLAC = 3
End Enum

Enum AudioBitrateType
    VBR = 0
    CBR = 1
End Enum

Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Type CWPSTRUCT
    lParam As Long
    wParam As Long
    Message As Long
    hWnd As Long
End Type

Public Const SWP_FRAMECHANGED As Long = &H20&
Public Const GWL_EXSTYLE As Long = -20&

Public Const RGN_DIFF = 4
Public Const RGN_OR = 2

Public Const WS_VISIBLE As Long = &H10000000
Public Const WS_VSCROLL As Long = &H200000
Public Const WS_TABSTOP As Long = &H10000
Public Const WS_THICKFRAME As Long = &H40000
Public Const WS_MAXIMIZE As Long = &H1000000
Public Const WS_MAXIMIZEBOX As Long = &H10000
Public Const WS_MINIMIZE As Long = &H20000000
Public Const WS_MINIMIZEBOX As Long = &H20000
Public Const WS_SYSMENU As Long = &H80000
Public Const WS_BORDER As Long = &H800000
Public Const WS_CAPTION As Long = &HC00000
Public Const WS_CHILD As Long = &H40000000
Public Const WS_CHILDWINDOW As Long = (WS_CHILD)
Public Const WS_CLIPCHILDREN As Long = &H2000000
Public Const WS_CLIPSIBLINGS As Long = &H4000000
Public Const WS_DISABLED As Long = &H8000000
Public Const WS_DLGFRAME As Long = &H400000
Public Const WS_EX_ACCEPTFILES As Long = &H10&
Public Const WS_EX_DLGMODALFRAME As Long = &H1&
Public Const WS_EX_NOPARENTNOTIFY As Long = &H4&
Public Const WS_EX_TOPMOST As Long = &H8&
Public Const WS_EX_TRANSPARENT As Long = &H20&
Public Const WS_EX_WINDOWEDGE As Long = &H100&
Public Const WS_EX_CLIENTEDGE As Long = &H200&
Public Const WS_EX_STATICEDGE As Long = &H20000
Public Const WS_GROUP As Long = &H20000
Public Const WS_HSCROLL As Long = &H100000
Public Const WS_ICONIC As Long = WS_MINIMIZE
Public Const WS_OVERLAPPED As Long = &H0&

Private Type ItemID
    cb As Long
    abID As Byte
End Type

Private Type ITEMIDLIST
    mkid As ItemID
End Type

Enum DriveTypes
    DRIVE_UNKNOWN = 0
    DRIVE_NO_ROOT_DIR = 1
    DRIVE_REMOVABLE = 2
    DRIVE_FIXED = 3
    DRIVE_REMOTE = 4
    DRIVE_CDROM = 5    'can be a CD or a DVD
    DRIVE_RAMDISK = 6
End Enum


Private Type LARGE_INTEGER
    lowpart As Long
    highpart As Long
End Type


Private Const SW_SHOW = 5
Private Const SEE_MASK_INVOKEIDLIST = &HC
Private Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hWnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    ' optional fields
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type

Public Const CSIDL_DESKTOP = &H0
Public Const CSIDL_INTERNET = &H1
Public Const CSIDL_PROGRAMS = &H2
Public Const CSIDL_CONTROLS = &H3
Public Const CSIDL_PRINTERS = &H4
Public Const CSIDL_PERSONAL = &H5
Public Const CSIDL_FAVORITES = &H6
Public Const CSIDL_STARTUP = &H7
Public Const CSIDL_RECENT = &H8
Public Const CSIDL_SENDTO = &H9
Public Const CSIDL_BITBUCKET = &HA
Public Const CSIDL_STARTMENU = &HB
Public Const CSIDL_DESKTOPDIRECTORY = &H10
Public Const CSIDL_DRIVES = &H11
Public Const CSIDL_NETWORK = &H12
Public Const CSIDL_NETHOOD = &H13
Public Const CSIDL_FONTS = &H14
Public Const CSIDL_TEMPLATES = &H15
Public Const CSIDL_COMMON_STARTMENU = &H16
Public Const CSIDL_COMMON_PROGRAMS = &H17
Public Const CSIDL_COMMON_STARTUP = &H18
Public Const CSIDL_COMMON_DESKTOPDIRECTORY = &H19
Public Const CSIDL_APPDATA = &H1A
Public Const CSIDL_PRINTHOOD = &H1B
Public Const CSIDL_ALTSTARTUP = &H1D
Public Const CSIDL_COMMON_ALTSTARTUP = &H1E
Public Const CSIDL_COMMON_FAVORITES = &H1F
Public Const CSIDL_INTERNET_CACHE = &H20
Public Const CSIDL_COOKIES = &H21
Public Const CSIDL_HISTORY = &H22

Public Const SC_MOVE = &HF010&
Public Const SC_RESTORE = &HF120&
Public Const SC_SIZE = &HF000&
Public Const SC_CLOSE = &HF060&
Global Const MF_BYPOSITION = &H400
Global Const MF_BYCOMMAND = &H0&
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Const ERROR_SUCCESS = 0
Const REG_SZ = 1
Const REG_DWORD = 4

Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL

Public SessionHeaders As Collection
Public SessionHeaderKeys As Collection
Public HeaderCache As String
Public SessionHeaderCache As String

Private Type OSVERSIONINFO
    OSVSize         As Long
    dwVerMajor      As Long
    dwVerMinor      As Long
    dwBuildNumber   As Long
    PlatformID      As Long
    szCSDVersion    As String * 128
End Type

Public Const MIIM_STATE = &H1
Public Const MIIM_ID = &H2
Public Const MIIM_TYPE = &H10
Public Const MFT_SEPARATOR = &H800
Public Const MFT_STRING = &H0
Public Const MFS_ENABLED = &H0
Public Const MFS_GRAYED = &H3
Public Const MFS_DISABLED = MFS_GRAYED
Public Const MFS_CHECKED = &H8

Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hBmpChecked As Long
    hBmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type

Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2

Type ChooseColorStruct
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    RGBResult As Long
    lpCustColors As Long
    Flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Const CC_RGBINIT = &H1&
Const CC_FULLOPEN = &H2&
Const CC_PREVENTFULLOPEN = &H4&
Const CC_SHOWHELP = &H8&
Const CC_ENABLEHOOK = &H10&
Const CC_ENABLETEMPLATE = &H20&
Const CC_ENABLETEMPLATEHANDLE = &H40&
Const CC_SOLIDCOLOR = &H80&
Const CC_ANYCOLOR = &H100&
Const CLR_INVALID = &HFFFF

Enum MsgBoxExIcon
    Critical = 16
    Question = 32
    Exclamation = 48
    Information = 64
    Doraemon = 128
End Enum

Private Const VK_SHIFT As Long = &H10
Private Const VK_CONTROL As Long = &H11
Private Const VK_MENU As Long = &H12
Private Const VK_CAPITAL = &H14
Private Const VK_NUMLOCK = &H90
Private Const VK_SCROLL = &H91

Enum GetKeyStateKeyboardCodes
    gksKeyboardShift = VK_SHIFT
    gksKeyboardctrl = VK_CONTROL
    gksKeyboardalt = VK_MENU
    gksKeyboardCapsLock = VK_CAPITAL
    gksKeyboardNumLock = VK_NUMLOCK
    gksKeyboardScrollLock = VK_SCROLL
End Enum

'https://www.mrexcel.com/board/threads/test-if-shift-key-was-held-when-commandbutton-gets-clicked.194874/
Function IsKeyPressed(ByVal lKey As GetKeyStateKeyboardCodes) As Boolean
    Dim iResult As Integer
    iResult = GetKeyState(lKey)
    
    Select Case lKey
        Case gksKeyboardCapsLock, gksKeyboardNumLock, gksKeyboardScrollLock
            iResult = iResult And 1
        Case Else
            iResult = iResult And &H8000
    End Select
    
    IsKeyPressed = (iResult <> 0)
End Function

Sub DisableDWMWindow(hWnd As Long)
    If WinVer < 6# Then Exit Sub
    On Error Resume Next
    DwmSetWindowAttribute hWnd, 2&, 1&, 4&
End Sub

Sub EnableDWMWindow(hWnd As Long)
    If WinVer < 6# Then Exit Sub
    On Error Resume Next
    DwmSetWindowAttribute hWnd, 2&, 0&, 4&
End Sub

Function IsDWMEnabled() As Boolean
    If WinVer < 6# Then GoTo nodwm
    On Error GoTo nodwm
    Dim DwmEnabled&, ret&
    DwmEnabled = 0
    ret = DwmIsCompositionEnabled(DwmEnabled)
    If ret <> 0 Then GoTo nodwm
    If DwmEnabled <> 0 Then
        IsDWMEnabled = True
    Else
        IsDWMEnabled = False
    End If
    Exit Function
nodwm:
    IsDWMEnabled = False
End Function

Sub ExtendDWMFrame(ByRef frmForm As Form, Top As Long, Right As Long, Bottom As Long, Left As Long)
    Dim Margin As MARGINS
    Margin.cxLeftWidth = Left
    Margin.cxRightWidth = Right
    Margin.cyTopHeight = Top
    Margin.cyBottomHeight = Bottom
    DwmExtendFrameIntoClientArea frmForm.hWnd, Margin
End Sub

Sub SetFormBackgroundColor(frmForm As Form, Optional DisableClassicTheme As Boolean = False)
    Dim clrBackColor As Long
    Dim clrForeColor As Long
    Dim DisableVisualStyle As Boolean
    Dim EnableLBSkin As Boolean
    EnableLBSkin = CBool(CInt(GetSetting("DownloadBooster", "Options", "EnableLiveBadukMemoSkin", 0)))
    DisableVisualStyle = CBool(CInt(GetSetting("DownloadBooster", "Options", "DisableVisualStyle", 0)))
    clrBackColor = GetSetting("DownloadBooster", "Options", "BackColor", DefaultBackColor)
    If clrBackColor < 0 Or clrBackColor > 16777215 Then
        If frmForm.BackColor <> &H8000000F Then frmForm.BackColor = &H8000000F
        clrBackColor = &H8000000F
    Else
        frmForm.BackColor = clrBackColor
    End If
    clrForeColor = GetSetting("DownloadBooster", "Options", "ForeColor", -1)
    If clrForeColor < 0 Or clrForeColor > 16777215 Then
        If frmForm.ForeColor <> &H80000012 Then frmForm.ForeColor = &H80000012
        clrForeColor = &H80000012
    Else
        frmForm.ForeColor = clrForeColor
    End If
    
    On Error Resume Next
    Dim ctrl As Control
    For Each ctrl In frmForm.Controls
        If TypeName(ctrl) = "ImageCombo" Or TypeName(ctrl) = "ToolBar" Or TypeName(ctrl) = "LinkLabel" Or TypeName(ctrl) = "Frame" Or TypeName(ctrl) = "PictureBox" Or TypeName(ctrl) = "Label" Or TypeName(ctrl) = "TabStrip" Or TypeName(ctrl) = "Slider" Or TypeName(ctrl) = "CheckBox" Or TypeName(ctrl) = "OptionButton" Or TypeName(ctrl) = "ProgressBar" Or TypeName(ctrl) = "FrameW" Or TypeName(ctrl) = "CommandButton" Or TypeName(ctrl) = "CommandButtonW" Or TypeName(ctrl) = "OptionButtonW" Or TypeName(ctrl) = "CheckBoxW" Or TypeName(ctrl) = "TextBoxW" Or TypeName(ctrl) = "ComboBoxW" Or TypeName(ctrl) = "StatusBar" Or TypeName(ctrl) = "ListView" Or TypeName(ctrl) = "ListBoxW" Then
            If TypeName(ctrl) = "CommandButtonW" And ctrl.Tag <> "notygchange" Then
                If EnableLBSkin Then
                    ctrl.IsTygemButton = True
                Else
                    ctrl.IsTygemButton = False
                    ctrl.Refresh
                End If
            End If
            If ctrl.Tag <> "novisualstylechange" And ctrl.Tag <> "nobackcolorchange novisualstylechange" Then
                If (Not DisableVisualStyle) And ctrl.VisualStyles = False Then
                    ctrl.VisualStyles = True
                    'If TypeName(ctrl) = "CommandButton" Or TypeName(ctrl) = "CommandButtonW" Then ctrl.Style = 0
                End If
                If DisableVisualStyle And ctrl.VisualStyles = True Then
                    ctrl.VisualStyles = False
                    'If TypeName(ctrl) = "CommandButton" Or TypeName(ctrl) = "CommandButtonW" Then ctrl.Style = 1
                End If
            End If
            If TypeName(ctrl) = "ListView" Or TypeName(ctrl) = "TextBoxW" Or TypeName(ctrl) = "ComboBoxW" Or TypeName(ctrl) = "ListBoxW" Then GoTo nextfor
            If ctrl.Tag <> "nocolorchange" And ctrl.Tag <> "nocolorsizechange" And ctrl.ForeColor <> clrForeColor And ctrl.Name <> "lblOverlay" And frmForm.Name <> "frmOptions" Then ctrl.ForeColor = clrForeColor
            If TypeName(ctrl) = "PictureBox" Then
                If ctrl.AutoRedraw = True Then GoTo nextfor
            End If
            If ctrl.Tag <> "nobackcolorchange" And ctrl.Tag <> "nobackcolorchange novisualstylechange" And ctrl.BackColor <> clrBackColor Then
                ctrl.BackColor = clrBackColor
                If TypeName(ctrl) = "CheckBoxW" Or TypeName(ctrl) = "OptionButtonW" Or TypeName(ctrl) = "FrameW" Then ctrl.Refresh
            End If
        End If
nextfor:
    Next ctrl
    
    SetClassicTheme frmForm, DisableClassicTheme
End Sub

Sub SetClassicTheme(frmForm As Form, Optional DisableClassicTheme As Boolean = False)
    If GetSetting("DownloadBooster", "Options", "UseClassicThemeFrame", 0) <> 0 Then
        SetWindowRgn frmForm.hWnd, CreateRectRgn(0, 0, Screen.Width / Screen.TwipsPerPixelX + 300, Screen.Height / Screen.TwipsPerPixelY + 300), True
    ElseIf DisableClassicTheme Then
        SetWindowRgn frmForm.hWnd, CreateRectRgn(0, 0, Screen.Width / Screen.TwipsPerPixelX + 300, Screen.Height / Screen.TwipsPerPixelY + 300), True
        SetWindowRgn frmForm.hWnd, 0&, True
    End If
End Sub

Function ShowColorDialog(Optional ByVal hParent As Long, Optional ByVal bFullOpen As Boolean, Optional ByVal InitColor As OLE_COLOR, Optional ByVal SolidOnly As Boolean = False) As Long
    Dim CC As ChooseColorStruct
    Static aColorRef(15) As Long
    Dim lInitColor As Long
  
    If InitColor <> 0 Then
        If OleTranslateColor(InitColor, 0, lInitColor) Then
            lInitColor = CLR_INVALID
        End If
    End If
    
    aColorRef(0) = RGB(233, 245, 236)
    aColorRef(1) = RGB(233, 237, 243)
    aColorRef(2) = RGB(185, 209, 234)
    aColorRef(3) = RGB(235, 233, 245)
    aColorRef(4) = RGB(252, 251, 224)
    aColorRef(5) = RGB(244, 232, 232)
    aColorRef(6) = RGB(248, 228, 244)
    aColorRef(7) = RGB(223, 233, 244)
    
    aColorRef(8) = RGB(249, 242, 230)
    aColorRef(9) = RGB(222, 235, 248)
    aColorRef(10) = RGB(227, 244, 232)
    aColorRef(11) = RGB(236, 230, 211)
    aColorRef(12) = RGB(212, 208, 200)
    aColorRef(13) = RGB(192, 192, 192)
    aColorRef(14) = 16777215
    aColorRef(15) = 0&
    
    Dim SolidColor As Long
    If SolidOnly Then
        SolidColor = CC_SOLIDCOLOR
    Else
        SolidColor = 0&
    End If
    
    With CC
        .lStructSize = Len(CC)
        .hWndOwner = hParent
        .lpCustColors = VarPtr(aColorRef(0))
        .RGBResult = lInitColor
        .Flags = SolidColor Or CC_ANYCOLOR Or CC_RGBINIT Or IIf(bFullOpen, CC_FULLOPEN, 0)
    End With
    
    If ChooseColor(CC) Then
        ShowColorDialog = CC.RGBResult
    Else
        ShowColorDialog = -1
    End If
End Function

Function GetKeyValue(ByVal KeyRoot As Long, ByVal KeyName As String, ByVal SubKeyRef As String, Optional ByVal Default As Variant = "") As Variant
    Dim i As Long                                           ' 루프 카운터
    Dim RC As Long                                          ' 반환 코드
    Dim hKey As Long                                        ' 열려 있는 레지스트리 키 처리
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' 레지스트리 키의 데이터 형식
    Dim tmpVal As String                                    ' 레지스트리 키 값을 임시로 저장
    Dim KeyValSize As Long                                  ' 레지스트리 키 변수의 크기
    Dim KeyVal
    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    RC = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' 레지스트리 키를 엽니다.
    
    If (RC <> ERROR_SUCCESS) Then GoTo GetKeyError          ' 오류를 처리합니다...
    
    tmpVal = String$(1024, 0)                             ' 변수의 크기를 할당합니다.
    KeyValSize = 1024                                       ' 변수 크기를 표시합니다.
    
    '------------------------------------------------------------
    ' 레지스트리 키 값을 읽어옵니다...
    '------------------------------------------------------------
    RC = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' 키 값을 가져오고 작성합니다.
                        
    If (RC <> ERROR_SUCCESS) Then GoTo GetKeyError          ' 오류를 처리합니다.
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95는 Null 종료 문자열을 추가합니다...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null을 찾았습니다. 문자열에서 추출합니다.
    Else                                                    ' WinNT는 Null 종료 문자열 추가하지 않습니다...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Null을 찾지 못했습니다. 문자열에서만 추출합니다.
    End If
    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' 데이터 형식을 검색합니다.
    Case REG_DWORD                                          ' 이진 단어 레지스트리 키 데이터 형식
        For i = Len(tmpVal) To 1 Step -1                    ' 각각 비트를 변환합니다.
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' 값 문자를 문자별로 작성합니다.
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' 이진 단어를 문자열로 변환합니다.
    Case Else                                               ' 문자열 레지스트리 키 데이터 형식
        KeyVal = tmpVal                                     ' 문자열 값을 복사합니다.
    End Select
    
    GetKeyValue = KeyVal
    RC = RegCloseKey(hKey)                                  ' 레지스트리 키를 닫습니다.
    Exit Function                                           ' 종료합니다.
    
GetKeyError:      ' 오류가 발생하면 지웁니다...
    GetKeyValue = Default
    RC = RegCloseKey(hKey)                                  ' 레지스트리 키를 닫습니다.
End Function

'https://stackoverflow.com/questions/40651/check-if-a-record-exists-in-a-vb6-collection
Function Exists(ByVal oCol As Collection, ByVal vKey As String) As Boolean
    On Error Resume Next
    oCol.Item CStr(vKey)
    Exists = (Err.Number = 0)
    Err.Clear
End Function

Function TextWidth(ByVal s As String, Optional ByVal FontName As String = "", Optional ByVal FontSize As Integer = -1) As Single
    If FontSize = 0 Then
        TextWidth = 0
        Exit Function
    End If
    On Error Resume Next
    If LangID <> 1042 Then
        If FontName = "" Then
            FontName = "Tahoma"
            If FontSize = -1 Then FontSize = 8
        ElseIf FontSize = -1 Then
            FontSize = 9
        End If
    Else
        If FontName = "" Then FontName = "굴림"
        If FontSize = -1 Then FontSize = 9
    End If
    frmDummyForm.Font.Name = FontName
    frmDummyForm.Font.Size = FontSize
    TextWidth = frmDummyForm.TextWidth(s)
End Function

Function TextHeight(ByVal s As String, Optional ByVal FontName As String = "", Optional ByVal FontSize As Integer = -1) As Single
    If FontSize = 0 Then
        TextHeight = 0
        Exit Function
    End If
    On Error Resume Next
    If LangID <> 1042 Then
        If FontName = "" Then
            FontName = "Tahoma"
            If FontSize = -1 Then FontSize = 8
        ElseIf FontSize = -1 Then
            FontSize = 9
        End If
    Else
        If FontName = "" Then FontName = "굴림"
        If FontSize = -1 Then FontSize = 9
    End If
    frmDummyForm.Font.Name = FontName
    frmDummyForm.Font.Size = FontSize
    TextHeight = frmDummyForm.TextHeight(s)
End Function

Function StrLen(ByVal s As String) As Integer
    StrLen = LenB(StrConv(s, vbFromUnicode))
End Function

Private Function CutLines(ByVal Text As String, ByVal Width As Single) As String()
    Dim Paragraphs() As String
    Dim ParagraphX As Long
    Dim Words() As String
    Dim WordX As Long
    Dim CutLine As String
    Dim NewCutLine As String
    Dim SingleWord As Boolean
    Dim ForceX As Long
    Dim Lines() As String
    Dim LineX As Long
    
    Paragraphs = Split(Text, vbNewLine)
    For ParagraphX = 0 To UBound(Paragraphs)
        Words = Split(Paragraphs(ParagraphX), " ")
        WordX = 0
        Do While WordX <= UBound(Words)
            Do
                If Len(CutLine) = 0 Then
                    NewCutLine = Words(WordX)
                    SingleWord = True
                Else
                    NewCutLine = NewCutLine & " " & Words(WordX)
                End If
                If TextWidth(NewCutLine) > Width Then Exit Do
                CutLine = NewCutLine
                WordX = WordX + 1
                SingleWord = False
            Loop While WordX <= UBound(Words)
            If SingleWord Then
                For ForceX = Len(Words(WordX)) - 1 To 1 Step -1
                    CutLine = Left$(Words(WordX), ForceX)
                    If TextWidth(CutLine) <= Width Then
                        Words(WordX) = Mid$(Words(WordX), ForceX + 1)
                        Exit For
                    End If
                Next
            End If
            ReDim Preserve Lines(LineX)
            Lines(LineX) = CutLine
            LineX = LineX + 1
            CutLine = vbNullString
        Loop
    Next
    CutLines = Lines
End Function

Sub Alert(ByVal Content As String, Optional ByVal Title As String, Optional Icon As MsgBoxExIcon = 64, Optional IsModal As Boolean = True, Optional AlertTimeout As Integer = -1)
    If Title = "" Then Title = App.Title
    If GetSetting("DownloadBooster", "Options", "ForceNativeMessageBox", 0) <> 0 Then
        VBA.MsgBox Content, Icon, Title
        Exit Sub
    End If
    
    Dim MessageBox As frmMessageBox
    Set MessageBox = New frmMessageBox
    MessageBox.MsgBoxMode = 1
    
    On Error Resume Next
    Randomize
    Dim NoIcon As Boolean
    NoIcon = False
    Select Case Icon
        Case 48
            MessageBox.imgMBIconWarning(Int(Rnd * 2)).Visible = True
        Case 16
            MessageBox.imgMBIconError(Int(Rnd * 2)).Visible = True
        Case 64
            MessageBox.imgMBIconInfo(Int(Rnd * 2)).Visible = True
        Case 32
            MessageBox.imgMBIconQuestion(Int(Rnd * 2)).Visible = True
        Case Else
            NoIcon = True
    End Select
    
    Content = Replace(Content, "&", "&&")
    Content = Replace(Content, vbCrLf & vbCrLf, vbCrLf & " " & vbCrLf)
    
    Dim i As Integer
    Dim LineCount As Integer
    Dim LContent As Integer
    Dim MAX_WIDTH As Long
    MAX_WIDTH = Screen.Width / 2
    Content = Join(CutLines(Content, MAX_WIDTH), vbCrLf)
    LContent = 0
    LineCount = UBound(Split(Content, vbLf)) + 1
    Dim s%
    Dim ln$
    Dim CI%, c$
    Dim LineContent$
    For s = 0 To UBound(Split(Content, vbCrLf))
        LineContent = Split(Content, vbCrLf)(s)
        If TextWidth(LineContent) > LContent Then LContent = TextWidth(LineContent)
    Next s
    
    If LContent = 0 Then LContent = frmAbout.TextWidth(Content)
    If LineCount > 1 Then MessageBox.lblContent.Top = 280
    MessageBox.lblContent.Height = 185 * LineCount + 60
    MessageBox.Height = 1615 + LineCount * 180 - 300 + 190 - 60
    MessageBox.Caption = Title
    MessageBox.lblContent.Caption = Content
    MessageBox.Width = Max(2040 + LContent - 640 - 225, 1920)
    If NoIcon Then
        MessageBox.Width = Max(MessageBox.Width - 720, 1920)
        MessageBox.lblContent.Top = MessageBox.lblContent.Top - 180
        MessageBox.lblContent.Left = MessageBox.lblContent.Left - 720
        MessageBox.Height = MessageBox.Height - 240
    End If
    MessageBox.cmdOK.Left = MessageBox.Width / 2 - 810 + 30
    MessageBox.cmdOK.Top = 840 + (LineCount * 185) - 350
    If LineCount < 2 Then
        MessageBox.Height = MessageBox.Height + 180
        MessageBox.cmdOK.Top = MessageBox.cmdOK.Top + 180
    End If
    If NoIcon Then
        MessageBox.cmdOK.Top = MessageBox.cmdOK.Top - 210
    End If
    
    Dim MessageSoundPath$
    Select Case Icon
        Case 48
            If GetSetting("DownloadBooster", "Options", "EnableExclamationSound", 1) = 0 Then
                MessageSoundPath = "-"
            Else
                MessageSoundPath = GetSetting("DownloadBooster", "Options", "ExclamationSound", "")
            End If
        Case 16
            If GetSetting("DownloadBooster", "Options", "EnableErrorSound", 1) = 0 Then
                MessageSoundPath = "-"
            Else
                MessageSoundPath = GetSetting("DownloadBooster", "Options", "ErrorSound", "")
            End If
        Case 64
            If GetSetting("DownloadBooster", "Options", "EnableAsteriskSound", 1) = 0 Then
                MessageSoundPath = "-"
            Else
                MessageSoundPath = GetSetting("DownloadBooster", "Options", "AsteriskSound", "")
            End If
        Case 32
            If GetSetting("DownloadBooster", "Options", "EnableQuestionSound", 1) = 0 Then
                MessageSoundPath = "-"
            Else
                MessageSoundPath = GetSetting("DownloadBooster", "Options", "QuestionSound", "")
            End If
        Case Else
            MessageSoundPath = ""
    End Select
    If MessageSoundPath <> "-" Then PlayWave MessageSoundPath, FallbackSound:=Icon
    
    If AlertTimeout >= 0 Then
        MessageBox.timeout.Interval = AlertTimeout
        MessageBox.timeout.Enabled = -1
    End If
    MessageBox.cmdOK.Caption = t("확인", "OK")
    
    MessageBox.cmdOK.Visible = -1
    MessageBox.cmdCancel.Visible = 0
    MessageBox.cmdYes.Visible = 0
    MessageBox.cmdNo.Visible = 0
    MessageBox.optYes.Visible = 0
    MessageBox.optNo.Visible = 0
    
    MessageBox.cmdAbort.Visible = 0
    MessageBox.cmdRetry.Visible = 0
    MessageBox.cmdIgnore.Visible = 0
    MessageBox.cmdFail.Visible = 0
    MessageBox.cmdHelp.Visible = 0
    
    Dim EnableLBSkin As Boolean
    EnableLBSkin = CBool(CInt(GetSetting("DownloadBooster", "Options", "EnableLiveBadukMemoSkin", 0)))
    
    MessageBox.cmdCancel.Cancel = 0
    MessageBox.cmdCancel.Default = 0
    MessageBox.cmdYes.Cancel = 0
    MessageBox.cmdYes.Default = 0
    MessageBox.cmdNo.Cancel = 0
    MessageBox.cmdNo.Default = 0
    MessageBox.cmdOK.Cancel = -1
    MessageBox.cmdOK.Default = 0
    
    If IsModal Then
        MessageBox.Show vbModal
        Unload MessageBox
        Set MessageBox = Nothing
    Else
        MessageBox.Show
    End If
End Sub

Function Confirm(ByVal Content As String, Optional ByVal Title As String, Optional Icon As MsgBoxExIcon = 32) As VbMsgBoxResult
    If Title = "" Then Title = App.Title
    If GetSetting("DownloadBooster", "Options", "ForceNativeMessageBox", 0) <> 0 Then
        Confirm = VBA.MsgBox(Content, Icon + vbYesNo, Title)
        Exit Function
    End If
    
    Dim MessageBox As frmMessageBox
    Set MessageBox = New frmMessageBox
    MessageBox.MsgBoxMode = 2
    MessageBox.ResultID = CStr(Rnd * 1E+15)
    
    On Error Resume Next
    Randomize
    Dim NoIcon As Boolean
    NoIcon = False
    Select Case Icon
        Case 48
            MessageBox.imgMBIconWarning(Int(Rnd * 2)).Visible = True
        Case 16
            MessageBox.imgMBIconError(Int(Rnd * 2)).Visible = True
        Case 64
            MessageBox.imgMBIconInfo(Int(Rnd * 2)).Visible = True
        Case 32
            MessageBox.imgMBIconQuestion(Int(Rnd * 2)).Visible = True
        Case Else
            NoIcon = True
    End Select
    
    Content = Replace(Content, "&", "&&")
    Content = Replace(Content, vbCrLf & vbCrLf, vbCrLf & " " & vbCrLf)
    
    Dim i As Integer
    Dim LineCount As Integer
    Dim LContent As Integer
    Dim MAX_WIDTH As Long
    MAX_WIDTH = Screen.Width / 2
    Content = Join(CutLines(Content, MAX_WIDTH), vbCrLf)
    LContent = 0
    LineCount = UBound(Split(Content, vbLf)) + 1
    Dim s%
    Dim ln$
    Dim CI%, c$
    Dim LineContent$
    For s = 0 To UBound(Split(Content, vbCrLf))
        LineContent = Split(Content, vbCrLf)(s)
        If TextWidth(LineContent) > LContent Then LContent = TextWidth(LineContent)
    Next s
    
    If LContent = 0 Then LContent = StrLen(Content)
    If LineCount > 1 Then MessageBox.lblContent.Top = 280
    MessageBox.lblContent.Height = 185 * LineCount + 60
    MessageBox.Height = 1615 + LineCount * 180 - 300 + 190 - 60
    MessageBox.Caption = Title
    MessageBox.lblContent.Caption = Content
    MessageBox.Width = Max(2040 + LContent - 640 - 225, 3480)
    If NoIcon Then
        MessageBox.Width = Max(MessageBox.Width - 720, 3480)
        MessageBox.lblContent.Top = MessageBox.lblContent.Top - 180
        MessageBox.lblContent.Left = MessageBox.lblContent.Left - 720
        MessageBox.Height = MessageBox.Height - 240
    End If
    MessageBox.cmdYes.Left = MessageBox.Width / 2 - 810 - MessageBox.cmdYes.Width / 2
    MessageBox.cmdYes.Top = 840 + (LineCount * 185) - 350
    MessageBox.cmdNo.Left = MessageBox.Width / 2 - 810 - MessageBox.cmdYes.Width / 2 - 120 + MessageBox.cmdYes.Width + 240 - 30
    MessageBox.cmdNo.Top = 840 + (LineCount * 185) - 350
    If LineCount < 2 Then
        MessageBox.Height = MessageBox.Height + 180
        MessageBox.cmdYes.Top = MessageBox.cmdYes.Top + 180
        MessageBox.cmdNo.Top = MessageBox.cmdNo.Top + 180
    End If
    If NoIcon Then
        MessageBox.cmdYes.Top = MessageBox.cmdYes.Top - 210
        MessageBox.cmdNo.Top = MessageBox.cmdNo.Top - 210
    End If
    
    Dim MessageSoundPath$
    Select Case Icon
        Case 48
            If GetSetting("DownloadBooster", "Options", "EnableExclamationSound", 1) = 0 Then
                MessageSoundPath = "-"
            Else
                MessageSoundPath = GetSetting("DownloadBooster", "Options", "ExclamationSound", "")
            End If
        Case 16
            If GetSetting("DownloadBooster", "Options", "EnableErrorSound", 1) = 0 Then
                MessageSoundPath = "-"
            Else
                MessageSoundPath = GetSetting("DownloadBooster", "Options", "ErrorSound", "")
            End If
        Case 64
            If GetSetting("DownloadBooster", "Options", "EnableAsteriskSound", 1) = 0 Then
                MessageSoundPath = "-"
            Else
                MessageSoundPath = GetSetting("DownloadBooster", "Options", "AsteriskSound", "")
            End If
        Case 32
            If GetSetting("DownloadBooster", "Options", "EnableQuestionSound", 1) = 0 Then
                MessageSoundPath = "-"
            Else
                MessageSoundPath = GetSetting("DownloadBooster", "Options", "QuestionSound", "")
            End If
        Case Else
            MessageSoundPath = ""
    End Select
    If MessageSoundPath <> "-" Then PlayWave MessageSoundPath, FallbackSound:=Icon
    
    MessageBox.cmdYes.Caption = t("예(&Y)", "&Yes")
    MessageBox.cmdNo.Caption = t("아니요(&N)", "&No")
    
    MessageBox.cmdOK.Visible = 0
    MessageBox.cmdCancel.Visible = 0
    MessageBox.cmdYes.Visible = -1
    MessageBox.cmdNo.Visible = -1
    MessageBox.optYes.Visible = 0
    MessageBox.optNo.Visible = 0
    
    MessageBox.cmdAbort.Visible = 0
    MessageBox.cmdRetry.Visible = 0
    MessageBox.cmdIgnore.Visible = 0
    MessageBox.cmdFail.Visible = 0
    MessageBox.cmdHelp.Visible = 0
    
    MessageBox.cmdCancel.Cancel = 0
    MessageBox.cmdCancel.Default = 0
    MessageBox.cmdYes.Cancel = 0
    MessageBox.cmdYes.Default = 0
    MessageBox.cmdNo.Cancel = 0
    MessageBox.cmdNo.Default = 0
    MessageBox.cmdOK.Cancel = 0
    MessageBox.cmdOK.Default = 0
    
    MessageBox.Init
    MessageBox.Show vbModal
    Confirm = MsgBoxResults(MessageBox.ResultID)
    Unload MessageBox
    Set MessageBox = Nothing
End Function

Function ConfirmEx(ByVal Content As String, Optional ByVal Title As String, Optional ByVal Icon As MsgBoxExIcon = 32, Optional ByVal DefaultOption As VbMsgBoxResult = vbNo) As VbMsgBoxResult
    Dim MessageBox As frmMessageBox
    Set MessageBox = New frmMessageBox
    MessageBox.MsgBoxMode = 3
    MessageBox.ResultID = CStr(Rnd * 1E+15)
    
    If Title = "" Then Title = App.Title
    
    On Error Resume Next
    Randomize
    Dim NoIcon As Boolean
    NoIcon = False
    Select Case Icon
        Case 48
            MessageBox.imgMBIconWarning(Int(Rnd * 2)).Visible = True
        Case 16
            MessageBox.imgMBIconError(Int(Rnd * 2)).Visible = True
        Case 64
            MessageBox.imgMBIconInfo(Int(Rnd * 2)).Visible = True
        Case 32
            MessageBox.imgMBIconQuestion(Int(Rnd * 2)).Visible = True
        Case Else
            NoIcon = True
    End Select
    
    Content = Replace(Content, "&", "&&")
    Content = Replace(Content, vbCrLf & vbCrLf, vbCrLf & " " & vbCrLf)
    
    Dim i As Integer
    Dim LineCount As Integer
    Dim LContent As Integer
    Dim MAX_WIDTH As Long
    MAX_WIDTH = Screen.Width / 2
    Content = Join(CutLines(Content, MAX_WIDTH), vbCrLf)
    LContent = 0
    LineCount = UBound(Split(Content, vbLf)) + 1
    Dim s%
    Dim ln$
    Dim CI%, c$
    Dim LineContent$
    For s = 0 To UBound(Split(Content, vbCrLf))
        LineContent = Split(Content, vbCrLf)(s)
        If TextWidth(LineContent) > LContent Then LContent = TextWidth(LineContent)
    Next s
    
    If LContent = 0 Then LContent = StrLen(Content)
    If LineCount > 1 Then MessageBox.lblContent.Top = 280
    MessageBox.lblContent.Height = 185 * LineCount + 60
    MessageBox.Height = 1615 + LineCount * 180 - 300 + 190 + 705 - 30
    MessageBox.Caption = Title
    MessageBox.lblContent.Caption = Content
    MessageBox.Width = Max(2040 + LContent - 640, 3480)
    If NoIcon Then
        MessageBox.Width = Max(MessageBox.Width - 720, 3480)
        MessageBox.lblContent.Top = MessageBox.lblContent.Top - 180
        MessageBox.lblContent.Left = MessageBox.lblContent.Left - 720
        MessageBox.Height = MessageBox.Height - 240
    End If
    MessageBox.cmdOK.Left = MessageBox.Width / 2 - 810 - MessageBox.cmdOK.Width / 2
    MessageBox.cmdOK.Top = 840 + (LineCount * 185) - 350 + 705
    MessageBox.cmdCancel.Left = MessageBox.Width / 2 - 810 - MessageBox.cmdOK.Width / 2 - 120 + MessageBox.cmdOK.Width + 240 - 30
    MessageBox.cmdCancel.Top = 840 + (LineCount * 185) - 350 + 705
    MessageBox.optYes.Top = MessageBox.cmdOK.Top - 620
    MessageBox.optNo.Top = MessageBox.cmdOK.Top - 320
    If LineCount > 1 Then
        MessageBox.optYes.Top = MessageBox.optYes.Top - 80
        MessageBox.optNo.Top = MessageBox.optNo.Top - 80
    End If
    If IsEmpty(DefaultOption) Then
        MessageBox.optYes.Value = False
        MessageBox.optNo.Value = False
        MessageBox.cmdOK.Enabled = False
    ElseIf DefaultOption = vbYes Then
        MessageBox.optYes.Value = True
        MessageBox.cmdOK.Enabled = True
    Else
        MessageBox.optNo.Value = True
        MessageBox.cmdOK.Enabled = True
    End If
    If LineCount < 2 Then
        MessageBox.Height = MessageBox.Height + 180
        MessageBox.cmdOK.Top = MessageBox.cmdOK.Top + 180
        MessageBox.cmdCancel.Top = MessageBox.cmdCancel.Top + 180
    End If
    If NoIcon Then
        MessageBox.cmdOK.Top = MessageBox.cmdOK.Top - 210
        MessageBox.cmdCancel.Top = MessageBox.cmdCancel.Top - 210
        MessageBox.optYes.Top = MessageBox.optYes.Top - 210
        MessageBox.optNo.Top = MessageBox.optNo.Top - 210
    End If
    MessageBox.optYes.Caption = t("예(&Y)", "&Yes")
    MessageBox.optNo.Caption = t("아니요(&N)", "&No")
    MessageBox.cmdOK.Caption = t("확인", "OK")
    MessageBox.cmdCancel.Caption = t("취소", "Cancel")
    
    Dim MessageSoundPath$
    Select Case Icon
        Case 48
            If GetSetting("DownloadBooster", "Options", "EnableExclamationSound", 1) = 0 Then
                MessageSoundPath = "-"
            Else
                MessageSoundPath = GetSetting("DownloadBooster", "Options", "ExclamationSound", "")
            End If
        Case 16
            If GetSetting("DownloadBooster", "Options", "EnableErrorSound", 1) = 0 Then
                MessageSoundPath = "-"
            Else
                MessageSoundPath = GetSetting("DownloadBooster", "Options", "ErrorSound", "")
            End If
        Case 64
            If GetSetting("DownloadBooster", "Options", "EnableAsteriskSound", 1) = 0 Then
                MessageSoundPath = "-"
            Else
                MessageSoundPath = GetSetting("DownloadBooster", "Options", "AsteriskSound", "")
            End If
        Case 32
            If GetSetting("DownloadBooster", "Options", "EnableQuestionSound", 1) = 0 Then
                MessageSoundPath = "-"
            Else
                MessageSoundPath = GetSetting("DownloadBooster", "Options", "QuestionSound", "")
            End If
        Case Else
            MessageSoundPath = ""
    End Select
    If MessageSoundPath <> "-" Then PlayWave MessageSoundPath, FallbackSound:=Icon
    
    MessageBox.cmdOK.Visible = -1
    MessageBox.cmdCancel.Visible = -1
    MessageBox.cmdYes.Visible = 0
    MessageBox.cmdNo.Visible = 0
    MessageBox.optYes.Visible = -1
    MessageBox.optNo.Visible = -1
    
    MessageBox.cmdAbort.Visible = 0
    MessageBox.cmdRetry.Visible = 0
    MessageBox.cmdIgnore.Visible = 0
    MessageBox.cmdFail.Visible = 0
    MessageBox.cmdHelp.Visible = 0
    
    MessageBox.cmdCancel.Cancel = -1
    MessageBox.cmdCancel.Default = 0
    MessageBox.cmdYes.Cancel = 0
    MessageBox.cmdYes.Default = 0
    MessageBox.cmdNo.Cancel = 0
    MessageBox.cmdNo.Default = 0
    MessageBox.cmdOK.Cancel = 0
    MessageBox.cmdOK.Default = -1
    
    MessageBox.Show vbModal
    ConfirmEx = MsgBoxResults(MessageBox.ResultID)
    Unload MessageBox
    Set MessageBox = Nothing
End Function

Function ConfirmCancel(ByVal Content As String, Optional ByVal Title As String, Optional Icon As MsgBoxExIcon = 32) As VbMsgBoxResult
    If Title = "" Then Title = App.Title
    If GetSetting("DownloadBooster", "Options", "ForceNativeMessageBox", 0) <> 0 Then
        ConfirmCancel = VBA.MsgBox(Content, Icon + vbYesNoCancel, Title)
        Exit Function
    End If
    
    Dim MessageBox As frmMessageBox
    Set MessageBox = New frmMessageBox
    MessageBox.MsgBoxMode = 4
    MessageBox.ResultID = CStr(Rnd * 1E+15)
    
    On Error Resume Next
    Randomize
    Dim NoIcon As Boolean
    NoIcon = False
    Select Case Icon
        Case 48
            MessageBox.imgMBIconWarning(Int(Rnd * 2)).Visible = True
        Case 16
            MessageBox.imgMBIconError(Int(Rnd * 2)).Visible = True
        Case 64
            MessageBox.imgMBIconInfo(Int(Rnd * 2)).Visible = True
        Case 32
            MessageBox.imgMBIconQuestion(Int(Rnd * 2)).Visible = True
        Case Else
            NoIcon = True
    End Select
    
    Content = Replace(Content, "&", "&&")
    Content = Replace(Content, vbCrLf & vbCrLf, vbCrLf & " " & vbCrLf)
    
    Dim i As Integer
    Dim LineCount As Integer
    Dim LContent As Integer
    Dim MAX_WIDTH As Long
    MAX_WIDTH = Screen.Width / 2
    Content = Join(CutLines(Content, MAX_WIDTH), vbCrLf)
    LContent = 0
    LineCount = UBound(Split(Content, vbLf)) + 1
    Dim s%
    Dim ln$
    Dim CI%, c$
    Dim LineContent$
    For s = 0 To UBound(Split(Content, vbCrLf))
        LineContent = Split(Content, vbCrLf)(s)
        If TextWidth(LineContent) > LContent Then LContent = TextWidth(LineContent)
    Next s
    
    If LContent = 0 Then LContent = StrLen(Content)
    If LineCount > 1 Then MessageBox.lblContent.Top = 280
    MessageBox.lblContent.Height = 185 * LineCount
    MessageBox.Height = 1615 + LineCount * 180 - 300 + 190 - 30
    MessageBox.Caption = Title
    MessageBox.lblContent.Caption = Content
    MessageBox.Width = Max(2040 + LContent - 640, 4920)
    If NoIcon Then
        MessageBox.Width = Max(MessageBox.Width - 720, 4920)
        MessageBox.lblContent.Top = MessageBox.lblContent.Top - 180
        MessageBox.lblContent.Left = MessageBox.lblContent.Left - 720
        MessageBox.Height = MessageBox.Height - 240
    End If
    MessageBox.cmdYes.Left = MessageBox.Width / 2 - 900 - MessageBox.cmdYes.Width
    MessageBox.cmdYes.Top = 840 + (LineCount * 185) - 350
    MessageBox.cmdNo.Left = MessageBox.Width / 2 - 810 + 15
    MessageBox.cmdNo.Top = 840 + (LineCount * 185) - 350
    MessageBox.cmdCancel.Left = MessageBox.Width / 2 - 900 + MessageBox.cmdYes.Width + 190 + 30
    MessageBox.cmdCancel.Top = 840 + (LineCount * 185) - 350
    If LineCount < 2 Then
        MessageBox.Height = MessageBox.Height + 180
        MessageBox.cmdYes.Top = MessageBox.cmdYes.Top + 180
        MessageBox.cmdNo.Top = MessageBox.cmdNo.Top + 180
        MessageBox.cmdCancel.Top = MessageBox.cmdCancel.Top + 180
    End If
    If NoIcon Then
        MessageBox.cmdCancel.Top = MessageBox.cmdCancel.Top - 210
        MessageBox.cmdYes.Top = MessageBox.cmdYes.Top - 210
        MessageBox.cmdNo.Top = MessageBox.cmdNo.Top - 210
    End If
    
    Dim MessageSoundPath$
    Select Case Icon
        Case 48
            If GetSetting("DownloadBooster", "Options", "EnableExclamationSound", 1) = 0 Then
                MessageSoundPath = "-"
            Else
                MessageSoundPath = GetSetting("DownloadBooster", "Options", "ExclamationSound", "")
            End If
        Case 16
            If GetSetting("DownloadBooster", "Options", "EnableErrorSound", 1) = 0 Then
                MessageSoundPath = "-"
            Else
                MessageSoundPath = GetSetting("DownloadBooster", "Options", "ErrorSound", "")
            End If
        Case 64
            If GetSetting("DownloadBooster", "Options", "EnableAsteriskSound", 1) = 0 Then
                MessageSoundPath = "-"
            Else
                MessageSoundPath = GetSetting("DownloadBooster", "Options", "AsteriskSound", "")
            End If
        Case 32
            If GetSetting("DownloadBooster", "Options", "EnableQuestionSound", 1) = 0 Then
                MessageSoundPath = "-"
            Else
                MessageSoundPath = GetSetting("DownloadBooster", "Options", "QuestionSound", "")
            End If
        Case Else
            MessageSoundPath = ""
    End Select
    If MessageSoundPath <> "-" Then PlayWave MessageSoundPath, FallbackSound:=Icon
    
    MessageBox.cmdOK.Visible = 0
    MessageBox.cmdCancel.Visible = -1
    MessageBox.cmdYes.Visible = -1
    MessageBox.cmdNo.Visible = -1
    MessageBox.optYes.Visible = 0
    MessageBox.optNo.Visible = 0
    
    MessageBox.cmdYes.Caption = t("예(&Y)", "&Yes")
    MessageBox.cmdNo.Caption = t("아니요(&N)", "&No")
    MessageBox.cmdCancel.Caption = t("취소", "Cancel")
    
    MessageBox.cmdAbort.Visible = 0
    MessageBox.cmdRetry.Visible = 0
    MessageBox.cmdIgnore.Visible = 0
    MessageBox.cmdFail.Visible = 0
    MessageBox.cmdHelp.Visible = 0
    
    MessageBox.cmdCancel.Cancel = -1
    MessageBox.cmdCancel.Default = 0
    MessageBox.cmdYes.Cancel = 0
    MessageBox.cmdYes.Default = 0
    MessageBox.cmdNo.Cancel = 0
    MessageBox.cmdNo.Default = 0
    MessageBox.cmdOK.Cancel = 0
    MessageBox.cmdOK.Default = 0
    
    MessageBox.Show vbModal
    ConfirmCancel = MsgBoxResults(MessageBox.ResultID)
    Unload MessageBox
    Set MessageBox = Nothing
End Function

Function MsgBoxAbortRetryIgnore(ByVal Content As String, Optional ByVal Title As String, Optional Icon As MsgBoxExIcon = 0) As VbMsgBoxResult
    If Title = "" Then Title = App.Title
    If GetSetting("DownloadBooster", "Options", "ForceNativeMessageBox", 0) <> 0 Then
        MsgBoxAbortRetryIgnore = VBA.MsgBox(Content, Icon + vbAbortRetryIgnore, Title)
        Exit Function
    End If
    
    Dim MessageBox As frmMessageBox
    Set MessageBox = New frmMessageBox
    MessageBox.MsgBoxMode = 5
    MessageBox.ResultID = CStr(Rnd * 1E+15)
    
    On Error Resume Next
    Randomize
    Dim NoIcon As Boolean
    NoIcon = False
    Select Case Icon
        Case 48
            MessageBox.imgMBIconWarning(Int(Rnd * 2)).Visible = True
        Case 16
            MessageBox.imgMBIconError(Int(Rnd * 2)).Visible = True
        Case 64
            MessageBox.imgMBIconInfo(Int(Rnd * 2)).Visible = True
        Case 32
            MessageBox.imgMBIconQuestion(Int(Rnd * 2)).Visible = True
        Case Else
            NoIcon = True
    End Select
    
    Content = Replace(Content, "&", "&&")
    Content = Replace(Content, vbCrLf & vbCrLf, vbCrLf & " " & vbCrLf)
    
    Dim i As Integer
    Dim LineCount As Integer
    Dim LContent As Integer
    Dim MAX_WIDTH As Long
    MAX_WIDTH = Screen.Width / 2
    Content = Join(CutLines(Content, MAX_WIDTH), vbCrLf)
    LContent = 0
    LineCount = UBound(Split(Content, vbLf)) + 1
    Dim s%
    Dim ln$
    Dim CI%, c$
    Dim LineContent$
    For s = 0 To UBound(Split(Content, vbCrLf))
        LineContent = Split(Content, vbCrLf)(s)
        If TextWidth(LineContent) > LContent Then LContent = TextWidth(LineContent)
    Next s
    
    If LContent = 0 Then LContent = StrLen(Content)
    If LineCount > 1 Then MessageBox.lblContent.Top = 280
    MessageBox.lblContent.Height = 185 * LineCount
    MessageBox.Height = 1615 + LineCount * 180 - 300 + 190 - 30
    MessageBox.Caption = Title
    MessageBox.lblContent.Caption = Content
    MessageBox.Width = Max(2040 + LContent - 640, 4920)
    If NoIcon Then
        MessageBox.Width = Max(MessageBox.Width - 720, 4920)
        MessageBox.lblContent.Top = MessageBox.lblContent.Top - 180
        MessageBox.lblContent.Left = MessageBox.lblContent.Left - 720
        MessageBox.Height = MessageBox.Height - 240
    End If
    MessageBox.cmdAbort.Left = MessageBox.Width / 2 - 900 - MessageBox.cmdAbort.Width
    MessageBox.cmdAbort.Top = 840 + (LineCount * 185) - 350
    MessageBox.cmdRetry.Left = MessageBox.Width / 2 - 810 + 15
    MessageBox.cmdRetry.Top = 840 + (LineCount * 185) - 350
    MessageBox.cmdIgnore.Left = MessageBox.Width / 2 - 900 + MessageBox.cmdAbort.Width + 190 + 30
    MessageBox.cmdIgnore.Top = 840 + (LineCount * 185) - 350
    If LineCount < 2 Then
        MessageBox.Height = MessageBox.Height + 180
        MessageBox.cmdAbort.Top = MessageBox.cmdAbort.Top + 180
        MessageBox.cmdRetry.Top = MessageBox.cmdRetry.Top + 180
        MessageBox.cmdIgnore.Top = MessageBox.cmdIgnore.Top + 180
    End If
    If NoIcon Then
        MessageBox.cmdIgnore.Top = MessageBox.cmdIgnore.Top - 210
        MessageBox.cmdAbort.Top = MessageBox.cmdAbort.Top - 210
        MessageBox.cmdRetry.Top = MessageBox.cmdRetry.Top - 210
    End If
    
    Dim MessageSoundPath$
    Select Case Icon
        Case 48
            If GetSetting("DownloadBooster", "Options", "EnableExclamationSound", 1) = 0 Then
                MessageSoundPath = "-"
            Else
                MessageSoundPath = GetSetting("DownloadBooster", "Options", "ExclamationSound", "")
            End If
        Case 16
            If GetSetting("DownloadBooster", "Options", "EnableErrorSound", 1) = 0 Then
                MessageSoundPath = "-"
            Else
                MessageSoundPath = GetSetting("DownloadBooster", "Options", "ErrorSound", "")
            End If
        Case 64
            If GetSetting("DownloadBooster", "Options", "EnableAsteriskSound", 1) = 0 Then
                MessageSoundPath = "-"
            Else
                MessageSoundPath = GetSetting("DownloadBooster", "Options", "AsteriskSound", "")
            End If
        Case 32
            If GetSetting("DownloadBooster", "Options", "EnableQuestionSound", 1) = 0 Then
                MessageSoundPath = "-"
            Else
                MessageSoundPath = GetSetting("DownloadBooster", "Options", "QuestionSound", "")
            End If
        Case Else
            MessageSoundPath = ""
    End Select
    If MessageSoundPath <> "-" Then PlayWave MessageSoundPath, FallbackSound:=Icon
    
    MessageBox.cmdOK.Visible = 0
    MessageBox.cmdCancel.Visible = 0
    MessageBox.cmdYes.Visible = 0
    MessageBox.cmdNo.Visible = 0
    MessageBox.optYes.Visible = 0
    MessageBox.optNo.Visible = 0
    
    MessageBox.cmdAbort.Caption = t("중단(&A)", "&Abort")
    MessageBox.cmdRetry.Caption = t("다시 시도(&R)", "&Retry")
    MessageBox.cmdIgnore.Caption = t("무시(&I)", "&Ignore")
    
    MessageBox.cmdAbort.Visible = -1
    MessageBox.cmdRetry.Visible = -1
    MessageBox.cmdIgnore.Visible = -1
    MessageBox.cmdFail.Visible = 0
    MessageBox.cmdHelp.Visible = 0
    
    MessageBox.cmdCancel.Cancel = -1
    MessageBox.cmdCancel.Default = 0
    MessageBox.cmdYes.Cancel = 0
    MessageBox.cmdYes.Default = 0
    MessageBox.cmdNo.Cancel = 0
    MessageBox.cmdNo.Default = 0
    MessageBox.cmdOK.Cancel = 0
    MessageBox.cmdOK.Default = 0
    
    MessageBox.Init
    MessageBox.Show vbModal
    MsgBoxAbortRetryIgnore = MsgBoxResults(MessageBox.ResultID)
    Unload MessageBox
    Set MessageBox = Nothing
End Function

Function MsgBoxRetryCancel(ByVal Content As String, Optional ByVal Title As String, Optional Icon As MsgBoxExIcon = 0) As VbMsgBoxResult
    If Title = "" Then Title = App.Title
    If GetSetting("DownloadBooster", "Options", "ForceNativeMessageBox", 0) <> 0 Then
        MsgBoxRetryCancel = VBA.MsgBox(Content, Icon + vbRetryCancel, Title)
        Exit Function
    End If
    
    Dim MessageBox As frmMessageBox
    Set MessageBox = New frmMessageBox
    MessageBox.MsgBoxMode = 6
    MessageBox.ResultID = CStr(Rnd * 1E+15)
    
    On Error Resume Next
    Randomize
    Dim NoIcon As Boolean
    NoIcon = False
    Select Case Icon
        Case 48
            MessageBox.imgMBIconWarning(Int(Rnd * 2)).Visible = True
        Case 16
            MessageBox.imgMBIconError(Int(Rnd * 2)).Visible = True
        Case 64
            MessageBox.imgMBIconInfo(Int(Rnd * 2)).Visible = True
        Case 32
            MessageBox.imgMBIconQuestion(Int(Rnd * 2)).Visible = True
        Case Else
            NoIcon = True
    End Select
    
    Content = Replace(Content, "&", "&&")
    Content = Replace(Content, vbCrLf & vbCrLf, vbCrLf & " " & vbCrLf)
    
    Dim i As Integer
    Dim LineCount As Integer
    Dim LContent As Integer
    Dim MAX_WIDTH As Long
    MAX_WIDTH = Screen.Width / 2
    Content = Join(CutLines(Content, MAX_WIDTH), vbCrLf)
    LContent = 0
    LineCount = UBound(Split(Content, vbLf)) + 1
    Dim s%
    Dim ln$
    Dim CI%, c$
    Dim LineContent$
    For s = 0 To UBound(Split(Content, vbCrLf))
        LineContent = Split(Content, vbCrLf)(s)
        If TextWidth(LineContent) > LContent Then LContent = TextWidth(LineContent)
    Next s
    
    If LContent = 0 Then LContent = StrLen(Content)
    If LineCount > 1 Then MessageBox.lblContent.Top = 280
    MessageBox.lblContent.Height = 185 * LineCount + 60
    MessageBox.Height = 1615 + LineCount * 180 - 300 + 190 - 60
    MessageBox.Caption = Title
    MessageBox.lblContent.Caption = Content
    MessageBox.Width = Max(2040 + LContent - 640 - 225, 3480)
    If NoIcon Then
        MessageBox.Width = Max(MessageBox.Width - 720, 3480)
        MessageBox.lblContent.Top = MessageBox.lblContent.Top - 180
        MessageBox.lblContent.Left = MessageBox.lblContent.Left - 720
        MessageBox.Height = MessageBox.Height - 240
    End If
    MessageBox.cmdRetry.Left = MessageBox.Width / 2 - 810 - MessageBox.cmdRetry.Width / 2
    MessageBox.cmdRetry.Top = 840 + (LineCount * 185) - 350
    MessageBox.cmdCancel.Left = MessageBox.Width / 2 - 810 - MessageBox.cmdCancel.Width / 2 - 120 + MessageBox.cmdRetry.Width + 240 - 30
    MessageBox.cmdCancel.Top = 840 + (LineCount * 185) - 350
    If LineCount < 2 Then
        MessageBox.Height = MessageBox.Height + 180
        MessageBox.cmdRetry.Top = MessageBox.cmdRetry.Top + 180
        MessageBox.cmdCancel.Top = MessageBox.cmdCancel.Top + 180
    End If
    If NoIcon Then
        MessageBox.cmdRetry.Top = MessageBox.cmdRetry.Top - 210
        MessageBox.cmdCancel.Top = MessageBox.cmdCancel.Top - 210
    End If
    
    Dim MessageSoundPath$
    Select Case Icon
        Case 48
            If GetSetting("DownloadBooster", "Options", "EnableExclamationSound", 1) = 0 Then
                MessageSoundPath = "-"
            Else
                MessageSoundPath = GetSetting("DownloadBooster", "Options", "ExclamationSound", "")
            End If
        Case 16
            If GetSetting("DownloadBooster", "Options", "EnableErrorSound", 1) = 0 Then
                MessageSoundPath = "-"
            Else
                MessageSoundPath = GetSetting("DownloadBooster", "Options", "ErrorSound", "")
            End If
        Case 64
            If GetSetting("DownloadBooster", "Options", "EnableAsteriskSound", 1) = 0 Then
                MessageSoundPath = "-"
            Else
                MessageSoundPath = GetSetting("DownloadBooster", "Options", "AsteriskSound", "")
            End If
        Case 32
            If GetSetting("DownloadBooster", "Options", "EnableQuestionSound", 1) = 0 Then
                MessageSoundPath = "-"
            Else
                MessageSoundPath = GetSetting("DownloadBooster", "Options", "QuestionSound", "")
            End If
        Case Else
            MessageSoundPath = ""
    End Select
    If MessageSoundPath <> "-" Then PlayWave MessageSoundPath, FallbackSound:=Icon
    
    MessageBox.cmdRetry.Caption = t("다시 시도(&R)", "&Retry")
    MessageBox.cmdCancel.Caption = t("취소", "Cancel")
    
    MessageBox.cmdOK.Visible = 0
    MessageBox.cmdCancel.Visible = -1
    MessageBox.cmdYes.Visible = 0
    MessageBox.cmdNo.Visible = 0
    MessageBox.optYes.Visible = 0
    MessageBox.optNo.Visible = 0
    
    MessageBox.cmdAbort.Visible = 0
    MessageBox.cmdRetry.Visible = -1
    MessageBox.cmdIgnore.Visible = 0
    MessageBox.cmdFail.Visible = 0
    MessageBox.cmdHelp.Visible = 0
    
    MessageBox.cmdCancel.Cancel = -1
    MessageBox.cmdCancel.Default = 0
    MessageBox.cmdYes.Cancel = 0
    MessageBox.cmdYes.Default = 0
    MessageBox.cmdNo.Cancel = 0
    MessageBox.cmdNo.Default = 0
    MessageBox.cmdOK.Cancel = 0
    MessageBox.cmdOK.Default = 0
    
    MessageBox.Show vbModal
    MsgBoxRetryCancel = MsgBoxResults(MessageBox.ResultID)
    Unload MessageBox
    Set MessageBox = Nothing
End Function

Function MsgBoxOKCancel(ByVal Content As String, Optional ByVal Title As String, Optional Icon As MsgBoxExIcon = 0) As VbMsgBoxResult
    If Title = "" Then Title = App.Title
    If GetSetting("DownloadBooster", "Options", "ForceNativeMessageBox", 0) <> 0 Then
        MsgBoxOKCancel = VBA.MsgBox(Content, Icon + vbOKCancel, Title)
        Exit Function
    End If
    
    Dim MessageBox As frmMessageBox
    Set MessageBox = New frmMessageBox
    MessageBox.MsgBoxMode = 7
    MessageBox.ResultID = CStr(Rnd * 1E+15)
    
    On Error Resume Next
    Randomize
    Dim NoIcon As Boolean
    NoIcon = False
    Select Case Icon
        Case 48
            MessageBox.imgMBIconWarning(Int(Rnd * 2)).Visible = True
        Case 16
            MessageBox.imgMBIconError(Int(Rnd * 2)).Visible = True
        Case 64
            MessageBox.imgMBIconInfo(Int(Rnd * 2)).Visible = True
        Case 32
            MessageBox.imgMBIconQuestion(Int(Rnd * 2)).Visible = True
        Case Else
            NoIcon = True
    End Select
    
    Content = Replace(Content, "&", "&&")
    Content = Replace(Content, vbCrLf & vbCrLf, vbCrLf & " " & vbCrLf)
    
    Dim i As Integer
    Dim LineCount As Integer
    Dim LContent As Integer
    Dim MAX_WIDTH As Long
    MAX_WIDTH = Screen.Width / 2
    Content = Join(CutLines(Content, MAX_WIDTH), vbCrLf)
    LContent = 0
    LineCount = UBound(Split(Content, vbLf)) + 1
    Dim s%
    Dim ln$
    Dim CI%, c$
    Dim LineContent$
    For s = 0 To UBound(Split(Content, vbCrLf))
        LineContent = Split(Content, vbCrLf)(s)
        If TextWidth(LineContent) > LContent Then LContent = TextWidth(LineContent)
    Next s
    
    If LContent = 0 Then LContent = StrLen(Content)
    If LineCount > 1 Then MessageBox.lblContent.Top = 280
    MessageBox.lblContent.Height = 185 * LineCount + 60
    MessageBox.Height = 1615 + LineCount * 180 - 300 + 190 - 60
    MessageBox.Caption = Title
    MessageBox.lblContent.Caption = Content
    MessageBox.Width = Max(2040 + LContent - 640 - 225, 3480)
    If NoIcon Then
        MessageBox.Width = Max(MessageBox.Width - 720, 3480)
        MessageBox.lblContent.Top = MessageBox.lblContent.Top - 180
        MessageBox.lblContent.Left = MessageBox.lblContent.Left - 720
        MessageBox.Height = MessageBox.Height - 240
    End If
    MessageBox.cmdOK.Left = MessageBox.Width / 2 - 810 - MessageBox.cmdOK.Width / 2
    MessageBox.cmdOK.Top = 840 + (LineCount * 185) - 350
    MessageBox.cmdCancel.Left = MessageBox.Width / 2 - 810 - MessageBox.cmdCancel.Width / 2 - 120 + MessageBox.cmdOK.Width + 240 - 30
    MessageBox.cmdCancel.Top = 840 + (LineCount * 185) - 350
    If LineCount < 2 Then
        MessageBox.Height = MessageBox.Height + 180
        MessageBox.cmdOK.Top = MessageBox.cmdOK.Top + 180
        MessageBox.cmdCancel.Top = MessageBox.cmdCancel.Top + 180
    End If
    If NoIcon Then
        MessageBox.cmdOK.Top = MessageBox.cmdOK.Top - 210
        MessageBox.cmdCancel.Top = MessageBox.cmdCancel.Top - 210
    End If
    
    Dim MessageSoundPath$
    Select Case Icon
        Case 48
            If GetSetting("DownloadBooster", "Options", "EnableExclamationSound", 1) = 0 Then
                MessageSoundPath = "-"
            Else
                MessageSoundPath = GetSetting("DownloadBooster", "Options", "ExclamationSound", "")
            End If
        Case 16
            If GetSetting("DownloadBooster", "Options", "EnableErrorSound", 1) = 0 Then
                MessageSoundPath = "-"
            Else
                MessageSoundPath = GetSetting("DownloadBooster", "Options", "ErrorSound", "")
            End If
        Case 64
            If GetSetting("DownloadBooster", "Options", "EnableAsteriskSound", 1) = 0 Then
                MessageSoundPath = "-"
            Else
                MessageSoundPath = GetSetting("DownloadBooster", "Options", "AsteriskSound", "")
            End If
        Case 32
            If GetSetting("DownloadBooster", "Options", "EnableQuestionSound", 1) = 0 Then
                MessageSoundPath = "-"
            Else
                MessageSoundPath = GetSetting("DownloadBooster", "Options", "QuestionSound", "")
            End If
        Case Else
            MessageSoundPath = ""
    End Select
    If MessageSoundPath <> "-" Then PlayWave MessageSoundPath, FallbackSound:=Icon
    
    MessageBox.cmdOK.Caption = t("확인", "OK")
    MessageBox.cmdCancel.Caption = t("취소", "Cancel")
    
    MessageBox.cmdOK.Visible = -1
    MessageBox.cmdCancel.Visible = -1
    MessageBox.cmdYes.Visible = 0
    MessageBox.cmdNo.Visible = 0
    MessageBox.optYes.Visible = 0
    MessageBox.optNo.Visible = 0
    
    MessageBox.cmdAbort.Visible = 0
    MessageBox.cmdRetry.Visible = 0
    MessageBox.cmdIgnore.Visible = 0
    MessageBox.cmdFail.Visible = 0
    MessageBox.cmdHelp.Visible = 0
    
    MessageBox.cmdCancel.Cancel = -1
    MessageBox.cmdCancel.Default = 0
    MessageBox.cmdYes.Cancel = 0
    MessageBox.cmdYes.Default = 0
    MessageBox.cmdNo.Cancel = 0
    MessageBox.cmdNo.Default = 0
    MessageBox.cmdOK.Cancel = 0
    MessageBox.cmdOK.Default = 0
    
    MessageBox.Show vbModal
    MsgBoxOKCancel = MsgBoxResults(MessageBox.ResultID)
    Unload MessageBox
    Set MessageBox = Nothing
End Function

'https://www.vbforums.com/showthread.php?894947-How-to-test-if-a-font-is-available
Function FontExists(ByVal Name As String) As Boolean
    With New StdFont
        .Name = Name
        FontExists = (StrComp(.Name, Name, vbTextCompare) = 0)
    End With
End Function

Function FolderExists(ByVal sFullPath As String) As Boolean
    On Error GoTo nonexist
    FolderExists = ((GetAttr(sFullPath) And (vbDirectory Or vbVolume)) <> 0)
    Exit Function
nonexist:
    FolderExists = False
End Function

Function Floor(ByVal floatval As Double, Optional ByVal decimalPlaces As Long = 0) As Long
    Dim intval As Long
    intval = Round(floatval)
    If intval > floatval Then
         intval = intval - 1
    End If

    If decimalPlaces > 0 Then
        floatval = floatval / (10 ^ decimalPlaces)
    End If

    Floor = intval
End Function

Function ParseSize(ByVal Size As Double, Optional ByVal ShowBytes As Boolean = False, Optional ByVal Suffix As String = "") As String
    If Size < 0 Then
        ParseSize = "-"
        Exit Function
    End If

    On Error GoTo ErrLn4
    Dim ret@
    If Size >= (1024@ * 1024@ * 1024@ * 1024@) Then
        ret = Fix(Size / 1024@ / 1024@ / 1024@ / 1024@ * 100) / 100
        'If ret >= 10@ Then ret = Fix(ret * 10) / 10
        'ElseIf ret >= 100@ Then ret = Fix(ret)
        ParseSize = ret & "TB" & Suffix
    ElseIf Size >= (1024@ * 1024@ * 1024@) Then
        ret = Fix(Size / 1024@ / 1024@ / 1024@ * 100) / 100
        'If ret >= 10@ Then ret = Fix(ret * 10) / 10
        'ElseIf ret >= 100@ Then ret = Fix(ret)
        ParseSize = ret & "GB" & Suffix
    ElseIf Size >= (1024@ * 1024@) Then
        ret = Fix(Size / 1024@ / 1024@ * 100) / 100
        'If ret >= 10@ Then ret = Fix(ret * 10) / 10
        'ElseIf ret >= 100@ Then ret = Fix(ret)
        ParseSize = ret & "MB" & Suffix
    ElseIf Size >= (1024@) Then
        ret = Fix(Size / 1024@ * 100) / 100
        'If ret >= 10@ Then ret = Fix(ret * 10) / 10
        'ElseIf ret >= 100@ Then ret = Fix(ret)
        ParseSize = ret & "KB" & Suffix
    Else
        ParseSize = CStr(Size) & " " & t("바이트", "Bytes")
    End If
    
    If Size >= (1024@) And ShowBytes Then
        ParseSize = ParseSize & " (" & Size & " " & t("바이트", "Bytes") & Suffix & ")"
    End If
    Exit Function
ErrLn4:
    ParseSize = "0 " & t("바이트", "Bytes")
End Function

Function FilterFilename(ByVal FileName As String, Optional ByVal PreserveBackslash As Boolean = False) As String
    Dim Str As String
    Dim ret As String
    ret = ""
    Str = StrConv(FileName, vbProperCase)
    Dim i%
    For i = 1 To Len(Str)
        If Mid(Str, i, 1) = "?" Then
            ret = ret & "_"
        Else
            ret = ret & Mid(FileName, i, 1)
        End If
    Next i
    ret = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(ret, "?", "_"), "*", "_"), "|", "_"), """", "_"), ":", "_"), "<", "_"), ">", "_"), "/", "_")
    If Not PreserveBackslash Then
        ret = Replace(ret, "\", "_")
    ElseIf Mid$(ret, 2, 1) = "_" Then
        ret = Left$(ret, 1) & ":" & Right$(ret, Len(ret) - 2)
    End If
    FilterFilename = ret
End Function

'https://gist.github.com/jvarn/5e11b1fd741b5f79d8a516c9c2368f17
Function URLDecode(ByVal strIn As String) As String
    On Error GoTo ErrorHandler
    
    Dim sl As Long, tl As Long
    Dim Key As String, kl As Long
    Dim hh As String, Hi As String, hl As String
    Dim a As Long
    
    Key = "%"
    kl = Len(Key)
    sl = 1: tl = 1
    sl = InStr(sl, strIn, Key, vbTextCompare)
    Do While sl > 0
        If (tl = 1 And sl <> 1) Or tl < sl Then
            URLDecode = URLDecode & Mid(strIn, tl, sl - tl)
        End If
        
        Select Case UCase(Mid(strIn, sl + kl, 1))
            Case "U"
                a = Val("&H" & Mid(strIn, sl + kl + 1, 4))
                URLDecode = URLDecode & ChrW(a)
                sl = sl + 6
            Case "E"
                hh = Mid(strIn, sl + kl, 2)
                a = Val("&H" & hh)
                If a < 128 Then
                    sl = sl + 3
                    URLDecode = URLDecode & Chr(a)
                Else
                    Hi = Mid(strIn, sl + 3 + kl, 2)
                    hl = Mid(strIn, sl + 6 + kl, 2)
                    a = ((Val("&H" & hh) And &HF) * 2 ^ 12) Or ((Val("&H" & Hi) And &H3F) * 2 ^ 6) Or (Val("&H" & hl) And &H3F)
                    URLDecode = URLDecode & ChrW(a)
                    sl = sl + 9
                End If
            Case Else
                hh = Mid(strIn, sl + kl, 2)
                a = Val("&H" & hh)
                If a < 128 Then
                    sl = sl + 3
                Else
                    Hi = Mid(strIn, sl + 3 + kl, 2)
                    a = ((Val("&H" & hh) - 194) * 64) + Val("&H" & Hi)
                    sl = sl + 6
                End If
                URLDecode = URLDecode & ChrW(a)
        End Select
        
        tl = sl
        sl = InStr(sl, strIn, Key, vbTextCompare)
    Loop
    
    URLDecode = URLDecode & Mid(strIn, tl)
    Exit Function
    
ErrorHandler:
    URLDecode = strIn
End Function

Function GetWindowsVersion() As Single
    Dim osv As OSVERSIONINFO
    Dim ver As Single
    osv.OSVSize = Len(osv)

    If GetVersionEx(osv) = 1 Then
        Select Case osv.PlatformID
            Case VER_PLATFORM_WIN32s
                GetWindowsVersion = 3.1
            Case VER_PLATFORM_WIN32_NT
                GetWindowsVersion = 3.1
                ver = osv.dwVerMajor + (CSng(osv.dwVerMinor) * 0.1)
                If ver >= 6.2 Then ver = fWinVer()
                GetWindowsVersion = ver
        
            Case VER_PLATFORM_WIN32_WINDOWS:
                Select Case osv.dwVerMinor
                    Case 0
                        GetWindowsVersion = 4#
                    Case 90
                        GetWindowsVersion = 4.9
                    Case Else
                        GetWindowsVersion = 4.1
                End Select
        End Select
    Else
        GetWindowsVersion = 5.2
    End If
End Function

Function fWinVer() As Single
    Dim osv As OSVERSIONINFO
    osv.OSVSize = Len(osv)
    If GetVersionEx(osv) <> 1 Then
        fWinVer = 5.1
        Exit Function
    End If

    If osv.PlatformID = VER_PLATFORM_WIN32_NT Then
        If RtlGetVersion(osv) <> 0 Then
            fWinVer = 5.1
            Exit Function
        End If
    End If

    fWinVer = osv.dwVerMajor + (CSng(osv.dwVerMinor) * 0.1)
End Function

Function t(ByVal k, ByVal e) As Variant
    If LangID = 1042 Then
        t = k
    Else
        t = e
    End If
End Function

Sub SetFont(frm As Form)
    On Error Resume Next
    If LangID = 1042 Then Exit Sub
    frm.Font.Name = "Tahoma"
    frm.Font.Size = 8
    Dim ctrl As Control
    For Each ctrl In frm.Controls
        If ctrl.Name <> "lvDummyScroll" Then
            ctrl.Font.Name = "Tahoma"
            If ctrl.Tag <> "nocolorsizechange" And ctrl.Tag <> "nosizechange" Then ctrl.Font.Size = 8
            ctrl.FontName = "Tahoma"
            If ctrl.Tag <> "nocolorsizechange" And ctrl.Tag <> "nosizechange" Then ctrl.FontSize = 8
        End If
    Next ctrl
End Sub

Function FormatTime(Sec) As String
    Dim Hour As Integer, Minutes As Integer, Seconds As Integer
    Dim ret As String
    If Sec >= 3600 Then
        ret = CStr(Floor(Sec / 3600)) & t("시간 ", " hours, ")
    Else
        ret = ""
    End If
    
    If Sec >= 60 Then
        ret = ret & Floor((Sec Mod 3600) / 60) & t("분 ", " minutes and ")
    End If
    ret = ret & (Sec Mod 60) & t("초", " seconds")
    FormatTime = ret
End Function

Function btoa(Str As String) As String
    On Error Resume Next
    Dim Data() As Byte
    Data = StrConv(Str, vbFromUnicode)
    Dim ss As String, s As Long
    ss = String$(2 * UBound(Data) + 6, 0)
    s = Len(ss) + 1
    CryptBinaryToString VarPtr(Data(0)), UBound(Data) + 1, CRYPT_STRING_BASE64, StrPtr(ss), s
    btoa = Left$(ss, s)
End Function

Sub BuildHeaderCache()
    Dim Headers() As String
    Dim RawHeaders As String
    RawHeaders = ""
    On Error GoTo dontbuild
    Headers = GetAllSettings("DownloadBooster", "Options\Headers")
    On Error GoTo 0
    Dim i%
    For i = LBound(Headers) To UBound(Headers)
        RawHeaders = RawHeaders & LCase(Headers(i, 0)) & ": " & Headers(i, 1) & vbLf
    Next i
    If Right$(RawHeaders, 1) = vbLf Then RawHeaders = Left$(RawHeaders, Len(RawHeaders) - 1)
    HeaderCache = btoa(RawHeaders)
dontbuild:
End Sub

Function DecodeHeaderCache(ByVal HeaderCache As String) As Collection
    Set DecodeHeaderCache = New Collection
    Dim Headers As Collection
    Dim HeaderKeys As Collection
    Set Headers = New Collection
    Set HeaderKeys = New Collection
    Dim RawHeaders$
    RawHeaders = StrConv(atob(HeaderCache), vbUnicode)
    Dim HeaderSplit() As String
    HeaderSplit = Split(RawHeaders, vbLf)
    Dim HeaderLine$
    Dim ColonPos%
    Dim i%
    For i = LBound(HeaderSplit) To UBound(HeaderSplit)
        HeaderLine = HeaderSplit(i)
        Debug.Print HeaderLine
        ColonPos = InStr(HeaderLine, ": ")
        If ColonPos < 1 Then GoTo continue
        Headers.Add Mid$(HeaderLine, ColonPos + 2), Left$(HeaderLine, ColonPos - 1)
        HeaderKeys.Add Left$(HeaderLine, ColonPos - 1)
continue:
    Next i
    
    DecodeHeaderCache.Add HeaderKeys, "keys"
    DecodeHeaderCache.Add Headers, "values"
End Function

Function GetSpecialfolder(CSIDL As Long) As String
    Dim lngRetVal As Long
    Dim IDL As ITEMIDLIST
    Dim strPath As String
    lngRetVal = SHGetSpecialFolderLocation(100, CSIDL, IDL)
    If lngRetVal = 0 Then
        strPath$ = Space$(512)
        lngRetVal = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal strPath$)
        GetSpecialfolder = Left$(strPath, InStr(strPath, Chr$(0)) - 1)
        Exit Function
    End If
    GetSpecialfolder = ""
End Function

Sub GetDiskSpace(sDrive As String, ByRef dblTotal As Double, ByRef dblFree As Double)
    Dim lresult As Long
    Dim liAvailable As LARGE_INTEGER
    Dim liTotal As LARGE_INTEGER
    Dim liFree As LARGE_INTEGER
    If Right(sDrive, 1) <> "" Then sDrive = sDrive & ""
    lresult = GetDiskFreeSpaceEx(sDrive, liAvailable, liTotal, liFree)
    
    dblTotal = CLargeInt(liTotal.lowpart, liTotal.highpart)
    dblFree = CLargeInt(liFree.lowpart, liFree.highpart)
End Sub
 
Private Function CLargeInt(Lo As Long, Hi As Long) As Double
    Dim dblLo As Double, dblHi As Double
    
    If Lo < 0 Then
        dblLo = 2 ^ 32 + Lo
    Else
        dblLo = Lo
    End If
    
    If Hi < 0 Then
        dblHi = 2 ^ 32 + Hi
    Else
        dblHi = Hi
    End If
    CLargeInt = dblLo + dblHi * 2 ^ 32
End Function

Sub DisplayFileProperties(ByVal sFullFileAndPathName As String)
    Dim shInfo As SHELLEXECUTEINFO

    With shInfo
        .cbSize = LenB(shInfo)
        .lpFile = sFullFileAndPathName
        .nShow = SW_SHOW
        .fMask = SEE_MASK_INVOKEIDLIST
        .lpVerb = "properties"
    End With

    ShellExecuteEx shInfo
End Sub

Function GetShortcutTarget(sPath As String) As String
    Dim shl As Shell, file As FolderItem, fld As shell32.Folder
    Dim lnk As ShellLinkObject, i As Long, folderPath As String
    Dim Shortcutname As String
    
    On Error GoTo ErrRtn
    folderPath = GetParentFolderName(sPath)
    Set shl = New Shell
    Set fld = shl.NameSpace(folderPath)
    Set file = fld.Items.Item(GetFilename(sPath))
    If Err <> 0 Then
        GetShortcutTarget = " Not Accesible"
        Err.Clear
        GoTo exit_sub
   Else
        If file.IsLink Then
            Set lnk = file.GetLink
            GetShortcutTarget = lnk.Path
        Else
            GetShortcutTarget = " Not decoded"
        End If
    End If
exit_sub:
    Set lnk = Nothing
    Set file = Nothing
    Set fld = Nothing
    Set shl = Nothing
    Exit Function
ErrRtn:
    Err.Clear
    Resume exit_sub
End Function

Function atob(sText As String) As Byte()
    Dim lSize           As Long
    Dim dwDummy         As Long
    Dim baOutput()      As Byte
    
    lSize = Len(sText) + 1
    ReDim baOutput(0 To lSize - 1) As Byte
    CryptStringToBinary StrPtr(sText), Len(sText), CRYPT_STRING_BASE64, VarPtr(baOutput(0)), lSize, 0, dwDummy
    If lSize > 0 Then
        ReDim Preserve baOutput(0 To lSize - 1) As Byte
        atob = baOutput
    Else
        atob = vbNullString
    End If
End Function

Function FormatModified(DateTime) As String
    If t(1, 2) = 1 Then
        FormatModified = Replace(Replace(Format(DateTime, "yyyy-mm-dd AM/PM h:mm"), "AM", "오전"), "PM", "오후")
    Else
        FormatModified = Replace(Format(DateTime, "m-d-yyyy h:mm AM/PM"), "-", "/")
    End If
End Function

Function GetParentFolderName(ByVal Path As String) As String
    On Error GoTo errfso
    Do While Right$(Path, 1) = "\"
        Path = Left$(Path, Len(Path) - 1)
    Loop
    If InStrRev(Path, "\") = 0 Then GoTo errfso
    GetParentFolderName = Left$(Path, InStrRev(Path, "\") - 1)
    Do While Right$(GetParentFolderName, 1) = "\"
        GetParentFolderName = Left$(GetParentFolderName, Len(GetParentFolderName) - 1)
    Loop
    If Len(GetParentFolderName) = 2 And Right$(GetParentFolderName, 1) = ":" Then GetParentFolderName = GetParentFolderName & "\"
    Exit Function
errfso:
    GetParentFolderName = ""
End Function

Function GetFilename(ByVal Path As String) As String
    On Error GoTo errfso
    Do While Right$(Path, 1) = "\"
        Path = Left$(Path, Len(Path) - 1)
    Loop
    GetFilename = Mid$(Path, InStrRev(Path, "\") + 1)
    Exit Function
errfso:
    GetFilename = ""
End Function

Function GetExtensionName(ByVal Path As String) As String
    On Error GoTo errfso
    Path = GetFilename(Path)
    If InStrRev(Path, ".") = 0 Then GoTo errfso
    GetExtensionName = Mid$(Path, InStrRev(Path, ".") + 1)
    Exit Function
errfso:
    GetExtensionName = ""
End Function

Function Includes(Target, toFind) As Boolean
    Dim TargetType As VbVarType
    TargetType = VarType(Target)
    If TargetType = vbString Then
stringproc:
        Includes = (InStr(CStr(Target), CStr(toFind)) <> 0)
        Exit Function
    ElseIf TargetType < vbArray Then
        If (TargetType >= vbInteger And TargetType <= vbDouble) Or TargetType = vbByte Or TargetType = vbDecimal Then GoTo stringproc
        Includes = False
        Exit Function
    End If
    
    Dim i%
    For i = LBound(Target) To UBound(Target)
        If Target(i) = toFind Then
            Includes = True
            Exit Function
        End If
    Next i
    Includes = False
End Function

Function GetStrFromPtr(ByVal Ptr As Long) As String
    GetStrFromPtr = SysAllocStringByteLen(Ptr, lstrlen(Ptr))
End Function

Sub UpdateBorderWidth()
    Startup.DialogBorderWidth = GetSystemMetrics(8)
    Startup.SizingBorderWidth = GetSystemMetrics(33)
    Startup.PaddedBorderWidth = SizingBorderWidth - DialogBorderWidth
End Sub

Function ExpandEnvironmentStrings(ByVal strInput As String) As String
    Dim ret As Long
    Dim strOutput As String
    ret = ExpandEnvironmentStringsA(strInput, strOutput, ret)
    If ret = 0& Then
        ExpandEnvironmentStrings = strInput
        Exit Function
    End If
    strOutput = Space$(ret)
    ret = ExpandEnvironmentStringsA(strInput, strOutput, ret)
    ExpandEnvironmentStrings = strOutput
End Function

Function StartsWith(ByVal Str As String, ByVal s As String) As Boolean
    StartsWith = (Left$(Str, Len(s)) = s)
End Function

Function EndsWith(ByVal Str As String, ByVal s As String) As Boolean
    EndsWith = (Right$(Str, Len(s)) = s)
End Function

Function ExcludeParameters(ByVal URL As String) As String
    If Includes(URL, "?") Then
        ExcludeParameters = Left$(URL, InStr(URL, "?") - 1)
    Else
        ExcludeParameters = URL
    End If
End Function

Function Col(Expression, ByRef IfFalse)
    If Expression Then
        Col = Expression
    Else
        Col = IfFalse
    End If
End Function

Function IsYtdlSupported(ByVal URL As String) As Boolean
    If EndsWith(LCase(ExcludeParameters(URL)), ".m3u8") Then
        IsYtdlSupported = True
        Exit Function
    End If
    
    Dim HostName$
    If Includes(URL, "://") Then
        HostName = Mid$(URL, InStr(URL, "://") + 3)
    Else
        HostName = URL
    End If
    HostName = Left$(HostName, Col(InStr(HostName, "/"), Len(HostName) + 1) - 1)
    If Includes(HostName, ":") Then
        HostName = Left$(HostName, InStrRev(HostName, ":") - 1)
    End If
    HostName = LCase(HostName)
    
    IsYtdlSupported = Includes(Array("youtube.com", "soundcloud.com", "ok.ru", "bilibili.tv", "dailymotion.com"), HostName)
End Function

Sub tr(ByRef ctrl As Control, ByVal EnglishCaption As String)
    On Error Resume Next
    ctrl.Caption = t(ctrl.Caption, EnglishCaption)
End Sub

Function GetThemeColor(ByVal hWnd As Long, ByVal ClassList As String, Optional ByVal Part As Long = 0&, Optional ByVal State As Long = 0&, Optional ByVal Prop As Long = TMT_TEXTCOLOR, Optional ByVal DefaultColor As Long = 0&) As Long
    On Error GoTo returndefault
    Dim hTheme As Long
    Dim clr As Long
    
    If IsAppThemed() = 0 Or IsThemeActive() = 0 Then GoTo returndefault
    hTheme = OpenThemeData(hWnd, StrPtr(ClassList))
    If hTheme = 0& Then GoTo returndefault
    If X_GetThemeColor(hTheme, Part, State, Prop, clr) <> 0 Then GoTo returndefault
    CloseThemeData hTheme
    
    GetThemeColor = clr
    Exit Function
    
returndefault:
    If hTheme <> 0& Then CloseThemeData hTheme
    GetThemeColor = DefaultColor
    Exit Function
End Function

Sub PlayWave(ByVal Path As String, Optional ByVal LoopWave As Boolean = False, Optional ByVal StopPreviousWave As Boolean = True, Optional ByVal FallbackSound As MsgBoxExIcon = -1)
    If FileExists(Path) Then
        PlaySound Path, 0&, SND_FILENAME Or SND_ASYNC Or IIf(LoopWave, SND_LOOP, 0&) Or IIf(StopPreviousWave, 0&, SND_NOSTOP)
    ElseIf FallbackSound >= 0 Then
        MessageBeep FallbackSound
    End If
End Sub

Sub CopyText(ByVal Content As String)
    Clipboard.Clear
    Clipboard.SetText Content
End Sub

Sub EnableFrameControls(ByRef fFrame As FrameW, ByRef Except As Control, Optional ByVal Enable As Boolean = True)
    Dim ctrl As Control
    For Each ctrl In fFrame.ContainedControls
        If ctrl.Name <> Except.Name Then ctrl.Enabled = Enable
    Next ctrl
End Sub

Function Max(ByRef L, ByRef R)
    If L > R Then
        Max = L
    Else
        Max = R
    End If
End Function

Function MsgBox(Prompt, Optional Buttons As VbMsgBoxStyle = vbOKOnly, Optional Title = "", Optional HelpFile, Optional Context) As VbMsgBoxResult
    If Title = "" Then Title = App.Title
    Select Case Buttons
        Case vbOKOnly
            Alert Prompt, Title, 0
            MsgBox = vbOK
        Case vbYesNo
            MsgBox = Confirm(Prompt, Title, 0)
        Case vbYesNoCancel
            MsgBox = ConfirmCancel(Prompt, Title, 0)
        Case vbOKCancel
            MsgBox = MsgBoxOKCancel(Prompt, Title, 0)
        Case vbAbortRetryIgnore
            MsgBox = MsgBoxAbortRetryIgnore(Prompt, Title, 0)
        Case vbRetryCancel
            MsgBox = MsgBoxRetryCancel(Prompt, Title, 0)
        
        Case vbCritical, vbQuestion, vbExclamation, vbInformation
            Alert Prompt, Title, Buttons
            MsgBox = vbOK
        Case vbCritical + vbYesNo, vbQuestion + vbYesNo, vbExclamation + vbYesNo, vbInformation + vbYesNo
            MsgBox = Confirm(Prompt, Title, Buttons - vbYesNo)
        Case vbCritical + vbYesNoCancel, vbQuestion + vbYesNoCancel, vbExclamation + vbYesNoCancel, vbInformation + vbYesNoCancel
            MsgBox = ConfirmCancel(Prompt, Title, Buttons - vbYesNoCancel)
        Case vbCritical + vbOKCancel, vbQuestion + vbOKCancel, vbExclamation + vbOKCancel, vbInformation + vbOKCancel
            MsgBox = MsgBoxOKCancel(Prompt, Title, Buttons - vbOKCancel)
        Case vbCritical + vbAbortRetryIgnore, vbQuestion + vbAbortRetryIgnore, vbExclamation + vbAbortRetryIgnore, vbInformation + vbAbortRetryIgnore
            MsgBox = MsgBoxAbortRetryIgnore(Prompt, Title, Buttons - vbAbortRetryIgnore)
        Case vbCritical + vbRetryCancel, vbQuestion + vbRetryCancel, vbExclamation + vbRetryCancel, vbInformation + vbRetryCancel
            MsgBox = MsgBoxRetryCancel(Prompt, Title, Buttons - vbRetryCancel)
        
        Case Else
            MsgBox = VBA.MsgBox(Prompt, Buttons, Title, HelpFile, Context)
    End Select
End Function

Function Right(Str As String, Length As Long) As String
    On Error GoTo errproc
    Right = VBA.Right$(Str, Length)
    Exit Function
errproc:
    Right = ""
End Function
