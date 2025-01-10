Attribute VB_Name = "Functions"
Public fso As Scripting.FileSystemObject
Public ConfirmResult As VbMsgBoxResult
Declare Function MessageBeep Lib "user32" (ByVal wType As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Declare Function DwmSetWindowAttribute Lib "dwmapi.dll" (ByVal hWnd As Long, ByVal dwAttribute As Long, ByRef pvAttribute As Long, ByVal cbAttribute As Long) As Long
Declare Function DwmIsCompositionEnabled Lib "dwmapi.dll" (ByRef pfEnabled As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode null 종료 문자열
Const REG_DWORD = 4                      ' 32비트 숫자

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

Private Type OSVERSIONINFO
  OSVSize         As Long
  dwVerMajor      As Long
  dwVerMinor      As Long
  dwBuildNumber   As Long
  PlatformID      As Long
  szCSDVersion    As String * 128
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

Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" _
    (lpChooseColor As ChooseColorStruct) As Long
Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor _
    As Long, ByVal lHPalette As Long, lColorRef As Long) As Long
    
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

Sub DisableDWMWindow(hWnd As Long)
    If WinVer < 6.1 Then Exit Sub
    DwmSetWindowAttribute hWnd, 2, 1, 4
End Sub

Sub EnableDWMWindow(hWnd As Long)
    If WinVer < 6.1 Then Exit Sub
    DwmSetWindowAttribute hWnd, 2, 0, 4
End Sub

Function IsDWMEnabled() As Boolean
    If WinVer < 6.1 Then
        IsDWMEnabled = False
        Exit Function
    End If
    Dim DwmEnabled As Long
    DwmEnabled = 0
    DwmIsCompositionEnabled DwmEnabled
    If DwmEnabled > 0 Then
        IsDWMEnabled = True
    Else
        IsDWMEnabled = False
    End If
End Function

Sub SetFormBackgroundColor(frmForm As Form)
    Dim clrBackColor As Long
    Dim clrForeColor As Long
    clrBackColor = GetSetting("DownloadBooster", "Options", "BackColor", DefaultBackColor)
    If clrBackColor < 0 Or clrBackColor > 16777215 Then
        frmForm.BackColor = &H8000000F
        clrBackColor = &H8000000F
    Else
        frmForm.BackColor = clrBackColor
    End If
    clrForeColor = GetSetting("DownloadBooster", "Options", "ForeColor", -1)
    If clrForeColor < 0 Or clrForeColor > 16777215 Then
        frmForm.ForeColor = &H80000012
        clrForeColor = &H80000012
    Else
        frmForm.ForeColor = clrForeColor
    End If
    
    On Error Resume Next
    Dim ctrl As Control
    For Each ctrl In frmForm.Controls
        If TypeName(ctrl) = "Frame" Or TypeName(ctrl) = "PictureBox" Or TypeName(ctrl) = "Label" Or TypeName(ctrl) = "TabStrip" Or TypeName(ctrl) = "Slider" Or TypeName(ctrl) = "CheckBox" Or TypeName(ctrl) = "OptionButton" Or TypeName(ctrl) = "ProgressBar" Or TypeName(ctrl) = "FrameW" Or TypeName(ctrl) = "CommandButton" Or TypeName(ctrl) = "CommandButtonW" Then
            ctrl.BackColor = clrBackColor
            ctrl.ForeColor = clrForeColor
        End If
    Next ctrl
End Sub

Function ShowColorDialog(Optional ByVal hParent As Long, Optional ByVal bFullOpen As Boolean, Optional ByVal InitColor As OLE_COLOR) As Long
    Dim CC As ChooseColorStruct
    Dim aColorRef(15) As Long
    Dim lInitColor As Long
  
    If InitColor <> 0 Then
        If OleTranslateColor(InitColor, 0, lInitColor) Then
            lInitColor = CLR_INVALID
        End If
    End If
    
    With CC
        .lStructSize = Len(CC)
        .hWndOwner = hParent
        .lpCustColors = VarPtr(aColorRef(0))
        .RGBResult = lInitColor
        .Flags = CC_SOLIDCOLOR Or CC_ANYCOLOR Or CC_RGBINIT Or IIf(bFullOpen, _
            CC_FULLOPEN, 0)
    End With
    
    If ChooseColor(CC) Then
        ShowColorDialog = CC.RGBResult
    Else
        ShowColorDialog = -1
    End If
End Function

'Function ReadRegistry(ByVal KeyPath As String, ByVal KeyName, Optional ByVal Default) As Variant
'    On Error GoTo RegReadFail
'    Dim WShell As Object
'    Set WShell = CreateObject("WScript.Shell")
'    If Right$(KeyPath, 1) <> "\" Then KeyPath = KeyPath & "\"
'    ReadRegistry = WShell.RegRead(KeyPath & KeyName)
'    Exit Function
'RegReadFail:
'    ReadRegistry = Default
'End Function

Public Function GetKeyValue(ByVal KeyRoot As Long, ByVal KeyName As String, ByVal SubKeyRef As String, Optional ByVal Default As Variant = "") As Variant
    Dim i As Long                                           ' 루프 카운터
    Dim RC As Long                                          ' 반환 코드
    Dim hKey As Long                                        ' 열려 있는 레지스트리 키 처리
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' 레지스트리 키의 데이터 형식
    Dim tmpVal As String                                    ' 레지스트리 키 값을 임시로 저장
    Dim KeyValSize As Long                                  ' 레지스트리 키 변수의 크기
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
    Case REG_SZ                                             ' 문자열 레지스트리 키 데이터 형식
        KeyVal = tmpVal                                     ' 문자열 값을 복사합니다.
    Case REG_DWORD                                          ' 이진 단어 레지스트리 키 데이터 형식
        For i = Len(tmpVal) To 1 Step -1                    ' 각각 비트를 변환합니다.
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' 값 문자를 문자별로 작성합니다.
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' 이진 단어를 문자열로 변환합니다.
    End Select
    
    GetKeyValue = KeyVal
    RC = RegCloseKey(hKey)                                  ' 레지스트리 키를 닫습니다.
    Exit Function                                           ' 종료합니다.
    
GetKeyError:      ' 오류가 발생하면 지웁니다...
    GetKeyValue = Default
    RC = RegCloseKey(hKey)                                  ' 레지스트리 키를 닫습니다.
End Function

'https://stackoverflow.com/questions/40651/check-if-a-record-exists-in-a-vb6-collection
Function Exists(ByVal oCol As Collection, ByVal vKey As Variant) As Boolean
    On Error Resume Next
    oCol.Item CStr(vKey)
    Exists = (Err.Number = 0)
    Err.Clear
End Function

Function TextWidth(ByVal s As String) As Single
    TextWidth = ConfirmMsgBox.TextWidth(s)
End Function

Function TextHeight(ByVal s As String) As Single
    TextHeight = ConfirmMsgBox.TextHeight(s)
End Function

Function strlen(ByVal s As String) As Integer
    strlen = LenB(StrConv(s, vbFromUnicode))
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

Function ConfirmEx(ByVal Content As String, ByVal Title As String, OwnerForm As Form, Optional ByVal Icon As MsgBoxExIcon = 32, Optional ByVal DefaultOption As VbMsgBoxResult = vbNo, Optional ByVal YesCaption As String = "", Optional ByVal NoCaption As String = "") As VbMsgBoxResult
    If Title = "" Then Title = App.Title
    If YesCaption = "" Then YesCaption = "예(&Y)"
    If NoCaption = "" Then NoCaption = "아니요(&N)"
    Select Case Icon
        Case 48
            ConfirmMsgBox.imgMBIconWarning.Visible = True
        Case 16
            ConfirmMsgBox.imgMBIconError.Visible = True
        Case 64
            ConfirmMsgBox.imgMBIconInfo.Visible = True
        Case 32
            ConfirmMsgBox.imgMBIconQuestion.Visible = True
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
    
    If LContent = 0 Then LContent = strlen(Content)
    If LineCount > 1 Then ConfirmMsgBox.lblContent.Top = 280
    ConfirmMsgBox.lblContent.Height = 185 * LineCount
    ConfirmMsgBox.Height = 1615 + LineCount * 180 - 300 + 190 + 705
    ConfirmMsgBox.Caption = Title
    ConfirmMsgBox.lblContent.Caption = Content
    ConfirmMsgBox.Width = 2040 + LContent - 640
    ConfirmMsgBox.cmdOK.Left = ConfirmMsgBox.Width / 2 - 810 - ConfirmMsgBox.cmdOK.Width / 2
    ConfirmMsgBox.cmdOK.Top = 840 + (LineCount * 185) - 350 + 705
    ConfirmMsgBox.cmdCancel.Left = ConfirmMsgBox.Width / 2 - 810 - ConfirmMsgBox.cmdOK.Width / 2 - 120 + ConfirmMsgBox.cmdOK.Width + 240
    ConfirmMsgBox.cmdCancel.Top = 840 + (LineCount * 185) - 350 + 705
    ConfirmMsgBox.optYes.Top = ConfirmMsgBox.cmdOK.Top - 620
    ConfirmMsgBox.optNo.Top = ConfirmMsgBox.cmdOK.Top - 320
    If LineCount > 1 Then
        ConfirmMsgBox.optYes.Top = ConfirmMsgBox.optYes.Top - 80
        ConfirmMsgBox.optNo.Top = ConfirmMsgBox.optNo.Top - 80
    End If
    If IsEmpty(DefaultOption) Then
        ConfirmMsgBox.optYes.Value = False
        ConfirmMsgBox.optNo.Value = False
        ConfirmMsgBox.cmdOK.Enabled = False
    ElseIf DefaultOption = vbYes Then
        ConfirmMsgBox.optYes.Value = True
        ConfirmMsgBox.cmdOK.Enabled = True
    Else
        ConfirmMsgBox.optNo.Value = True
        ConfirmMsgBox.cmdOK.Enabled = True
    End If
    If LineCount < 2 Then
        ConfirmMsgBox.Height = ConfirmMsgBox.Height + 180
        ConfirmMsgBox.cmdOK.Top = ConfirmMsgBox.cmdOK.Top + 180
        ConfirmMsgBox.cmdCancel.Top = ConfirmMsgBox.cmdCancel.Top + 180
    End If
    ConfirmMsgBox.optYes.Caption = YesCaption
    ConfirmMsgBox.optNo.Caption = NoCaption
    ConfirmMsgBox.BeepSnd = Icon
    MessageBeep Icon
    ConfirmMsgBox.Show vbModal, OwnerForm
    
    ConfirmEx = ConfirmResult
End Function

'https://www.vbforums.com/showthread.php?894947-How-to-test-if-a-font-is-available
Function FontExists(ByVal Name As String) As Boolean
    With New StdFont
        .Name = Name
        FontExists = (StrComp(.Name, Name, vbTextCompare) = 0)
    End With
End Function

Function FolderExists(ByVal sFullPath As String) As Boolean
    FolderExists = fso.FolderExists(sFullPath)
End Function

Function Floor(ByVal floatval As Double, Optional ByVal decimalPlaces As Long = 0) As Long
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

Function ParseSize(ByVal Size As Double, Optional ByVal ShowBytes As Boolean = False, Optional ByVal Suffix As String = "") As String
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
        ParseSize = CStr(Size) & " 바이트"
    End If
    
    If Size >= (1024@) And ShowBytes Then
        ParseSize = ParseSize & " (" & Size & " 바이트" & Suffix & ")"
    End If
    Exit Function
ErrLn4:
    ParseSize = "0 바이트"
End Function

Function FilterFilename(ByVal FileName As String) As String
    Dim str As String
    Dim ret As String
    ret = ""
    str = StrConv(FileName, vbProperCase)
    Dim i%
    For i = 1 To Len(str)
        If Mid(str, i, 1) = "?" Then
            ret = ret & "_"
        Else
            ret = ret & Mid(FileName, i, 1)
        End If
    Next i
    FilterFilename = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(ret, "\", "_"), "?", "_"), "*", "_"), "|", "_"), """", "_"), ":", "_"), "<", "_"), ">", "_"), "/", "_")
End Function

'https://gist.github.com/jvarn/5e11b1fd741b5f79d8a516c9c2368f17
Function URLDecode(ByVal strIn As String) As String
    On Error GoTo ErrorHandler
    
    Dim sl As Long, tl As Long
    Dim Key As String, kl As Long
    Dim hh As String, hi As String, hl As String
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
                    hi = Mid(strIn, sl + 3 + kl, 2)
                    hl = Mid(strIn, sl + 6 + kl, 2)
                    a = ((Val("&H" & hh) And &HF) * 2 ^ 12) Or ((Val("&H" & hi) And &H3F) * 2 ^ 6) Or (Val("&H" & hl) And &H3F)
                    URLDecode = URLDecode & ChrW(a)
                    sl = sl + 9
                End If
            Case Else
                hh = Mid(strIn, sl + kl, 2)
                a = Val("&H" & hh)
                If a < 128 Then
                    sl = sl + 3
                Else
                    hi = Mid(strIn, sl + 3 + kl, 2)
                    a = ((Val("&H" & hh) - 194) * 64) + Val("&H" & hi)
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
    osv.OSVSize = Len(osv)

    If GetVersionEx(osv) = 1 Then
        Select Case osv.PlatformID
            Case VER_PLATFORM_WIN32s
                GetWindowsVersion = 3.1
            Case VER_PLATFORM_WIN32_NT
                GetWindowsVersion = 3.1
                GetWindowsVersion = osv.dwVerMajor + (CSng(osv.dwVerMinor) * 0.1)
        
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
