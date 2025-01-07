Attribute VB_Name = "Functions"
Public fso
Public ConfirmResult As VbMsgBoxResult
Declare Function MessageBeep Lib "user32" (ByVal wType As Long) As Long

Enum MsgBoxExIcon
    Critical = 16
    Question = 32
    Exclamation = 48
    Information = 64
    Doraemon = 128
End Enum

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

Function ConfirmEx(Content As String, Title As String, OwnerForm As Form, Optional Icon As MsgBoxExIcon = 32, Optional DefaultOption As VbMsgBoxResult = vbNo, Optional YesCaption As String = "", Optional NoCaption As String = "") As VbMsgBoxResult
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
    Dim ci%, c$
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

Function ParseSize(Size As Double, Optional ShowBytes As Boolean = False, Optional Suffix As String = "") As String
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

Function FilterFilename(FileName As String) As String
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
    FilterFilename = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(ret, "\", "_"), "?", "_"), "*", "_"), "|", "_"), """", "_"), ":", "_"), "<", "_"), ">", "_")
End Function

'https://gist.github.com/jvarn/5e11b1fd741b5f79d8a516c9c2368f17
Function URLDecode(ByVal strIn As String) As String
    On Error GoTo ErrorHandler
    
    Dim sl As Long, tl As Long
    Dim key As String, kl As Long
    Dim hh As String, hi As String, hl As String
    Dim a As Long
    
    key = "%"
    kl = Len(key)
    sl = 1: tl = 1
    sl = InStr(sl, strIn, key, vbTextCompare)
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
        sl = InStr(sl, strIn, key, vbTextCompare)
    Loop
    
    URLDecode = URLDecode & Mid(strIn, tl)
    Exit Function
    
ErrorHandler:
    URLDecode = strIn
End Function
