VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "정보"
   ClientHeight    =   5265
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   7650
   BeginProperty Font 
      Name            =   "굴림"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   7650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.PictureBox pbLicenses 
      BorderStyle     =   0  '없음
      Height          =   3255
      Index           =   6
      Left            =   2640
      ScaleHeight     =   3255
      ScaleWidth      =   4815
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1440
      Width           =   4815
      Begin VB.TextBox txtVbal 
         Height          =   3255
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   19
         Top             =   0
         Width           =   4815
      End
   End
   Begin prjDownloadBooster.CommandButtonW cmdOK 
      Cancel          =   -1  'True
      Default         =   -1  'True
      Height          =   345
      Left            =   6120
      TabIndex        =   5
      Top             =   4800
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   609
      Caption         =   "확인"
   End
   Begin VB.PictureBox pbLicenses 
      BorderStyle     =   0  '없음
      Height          =   3255
      Index           =   3
      Left            =   2640
      ScaleHeight     =   3255
      ScaleWidth      =   4815
      TabIndex        =   17
      Top             =   1440
      Width           =   4815
      Begin VB.TextBox txtShellPipe 
         Height          =   3255
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  '양방향
         TabIndex        =   4
         Top             =   0
         Width           =   4815
      End
   End
   Begin VB.PictureBox pbLicenses 
      BorderStyle     =   0  '없음
      Height          =   3255
      Index           =   2
      Left            =   2640
      ScaleHeight     =   3255
      ScaleWidth      =   4815
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1440
      Width           =   4815
      Begin VB.TextBox txtLicensePlaceholder 
         Enabled         =   0   'False
         Height          =   270
         Left            =   0
         Locked          =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   7
         Top             =   0
         Width           =   1215
      End
      Begin prjDownloadBooster.ProgressBar pbLicenseLoadProgress 
         Height          =   255
         Left            =   0
         Top             =   3000
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   450
         Max             =   812
         Step            =   10
      End
      Begin VB.TextBox txtLicense 
         Enabled         =   0   'False
         Height          =   2970
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  '양방향
         TabIndex        =   8
         Top             =   0
         Width           =   4815
      End
   End
   Begin VB.PictureBox pbLicenses 
      BorderStyle     =   0  '없음
      Height          =   3255
      Index           =   5
      Left            =   2640
      ScaleHeight     =   3255
      ScaleWidth      =   4815
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1440
      Width           =   4815
      Begin VB.TextBox txtPNG 
         Height          =   3255
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   11
         Top             =   0
         Width           =   4815
      End
   End
   Begin VB.PictureBox pbLicenses 
      BorderStyle     =   0  '없음
      Height          =   3255
      Index           =   1
      Left            =   2640
      ScaleHeight     =   3255
      ScaleWidth      =   4815
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1440
      Width           =   4815
      Begin VB.TextBox txtCC 
         Height          =   3255
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   10
         Top             =   0
         Width           =   4815
      End
   End
   Begin VB.PictureBox pbLicenses 
      BorderStyle     =   0  '없음
      Height          =   3255
      Index           =   4
      Left            =   2640
      ScaleHeight     =   3255
      ScaleWidth      =   4815
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1440
      Width           =   4815
      Begin VB.TextBox txtIconv 
         Height          =   3255
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   9
         Top             =   0
         Width           =   4815
      End
   End
   Begin prjDownloadBooster.ImageList imgItems 
      Left            =   360
      Top             =   4320
      _ExtentX        =   1005
      _ExtentY        =   1005
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   16711935
      InitListImages  =   "frmAbout.frx":000C
   End
   Begin prjDownloadBooster.FrameW FrameW1 
      Height          =   3255
      Left            =   1080
      TabIndex        =   6
      Top             =   1440
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   5741
      BorderStyle     =   0
      Caption         =   "라이선스(&L)"
      Begin prjDownloadBooster.ListView lvItems 
         Height          =   3255
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   5741
         Icons           =   "imgItems"
         Arrange         =   2
         LabelEdit       =   2
         HideSelection   =   0   'False
         ShowInfoTips    =   -1  'True
         ShowLabelTips   =   -1  'True
         ShowColumnTips  =   -1  'True
         SnapToGrid      =   -1  'True
      End
   End
   Begin VB.Timer timLicenseLoader 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   5880
      Top             =   120
   End
   Begin VB.PictureBox pbLicenses 
      BorderStyle     =   0  '없음
      Height          =   3255
      Index           =   7
      Left            =   2640
      ScaleHeight     =   3255
      ScaleWidth      =   4815
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1440
      Width           =   4815
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  '투명
      Caption         =   "버전"
      Height          =   225
      Left            =   1050
      TabIndex        =   1
      Top             =   600
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  '투명
      Caption         =   "응용 프로그램 제목"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1050
      TabIndex        =   0
      Top             =   240
      Width           =   3885
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  '투명
      Caption         =   "응용 프로그램 설명"
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   1050
      TabIndex        =   2
      Top             =   960
      Width           =   6405
   End
   Begin VB.Image picIcon 
      Height          =   480
      Left            =   240
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim LineNum As Integer

Const Mit1 As String = "Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the ""Software""), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:" & vbCrLf & vbCrLf & "The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software."
Const Mit2 As String = "THE SOFTWARE IS PROVIDED ""AS IS"", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE."

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    lvItems.SetFocus
    'SavePicture imgItems.ListImages(1).ExtractIcon(), "F:\1호선저항.ico"
    'SavePicture imgItems.ListImages(2).ExtractIcon(), "F:\2호선저항.ico"
End Sub

Private Sub Form_Load()
    InitForm Me
    
    imgItems.ListImages.Add Picture:=frmDummyForm.imgTrain(1).Picture
    imgItems.ListImages.Add Picture:=frmDummyForm.imgTrain(2).Picture
    imgItems.ListImages.Add Picture:=frmDummyForm.imgTrain(4).Picture
    imgItems.ListImages.Add Picture:=frmDummyForm.imgTrain(5).Picture
    
    LineNum = 1
    Me.Caption = t(App.Title & " 정보", "About " & App.Title)
    Set picIcon.Picture = frmMain.Icon
    lblVersion.Caption = t("버전 ", "Version ") & App.Major & "." & App.Minor & IIf(App.Revision > 0, "." & App.Revision, "")
    lblTitle.Caption = App.Title
    lblDescription.Caption = "This product includes software developed by vbAccelerator. " & t("이 프로그램에는 외부 라이브러리가 사용됐으며 라이선스는 아래와 같습니다.", "This program includes external libraries. Check out the license of them below.")
    txtLicensePlaceholder.Width = txtLicense.Width
    txtLicensePlaceholder.Height = txtLicense.Height
    txtLicensePlaceholder.Top = txtLicense.Top
    txtLicensePlaceholder.Left = txtLicense.Left
    pbLicenseLoadProgress.Width = txtLicense.Width
    pbLicenseLoadProgress.Top = txtLicense.Top + txtLicense.Height + 30
    pbLicenseLoadProgress.Left = txtLicense.Left
    cmdOK.Caption = t(cmdOK.Caption, "OK")
    
    timLicenseLoader.Enabled = True
    
    Dim EasterEggMultiplier As Byte
    EasterEggMultiplier = -(RandInt(1, 10000) = 2094) * 2
    
    lvItems.ListItems.Add , , "Krool's Comctl", 1 + EasterEggMultiplier
    lvItems.ListItems.Add , , "Node.js (v0.11.11)", 2 + EasterEggMultiplier
    lvItems.ListItems.Add , , "ShellPipe (v7)", 1 + EasterEggMultiplier
    lvItems.ListItems.Add , , "iconv-lite (v0.6.3)", 2 + EasterEggMultiplier
    lvItems.ListItems.Add , , "PNG with alpha", 1 + EasterEggMultiplier
    lvItems.ListItems.Add , , "vbAccelerator SSubTmr", 2 + EasterEggMultiplier
    'lvItems.ListItems.Add , , t("기타 출처", "Other references"), 1 + EasterEggMultiplier
    lvItems.ListItems(1).Selected = True
    
    txtIconv.Text = txtIconv.Text & "Copyright (c) 2011 Alexander Shtuchkin" & vbCrLf & vbCrLf
    txtIconv.Text = txtIconv.Text & Mit1 & vbCrLf & vbCrLf
    txtIconv.Text = txtIconv.Text & Mit2
    
    txtCC.Text = "https://github.com/Kr00l/VBCCR/tree/master/Standard%20EXE%20Version" & vbCrLf & vbCrLf
    txtCC.Text = txtCC.Text & "Copyright (c) 2012-present Krool" & vbCrLf & vbCrLf
    txtCC.Text = txtCC.Text & Mit1 & vbCrLf & vbCrLf
    txtCC.Text = txtCC.Text & Mit2
    
    txtPNG.Text = "https://www.vbforums.com/showthread.php?896878" & vbCrLf & vbCrLf
    txtPNG.Text = txtPNG.Text & "Elroy, LaVolpe, Dilettante, Wqweto, Schmidt, & The Trick" & vbCrLf & vbCrLf
    txtPNG.Text = txtPNG.Text & "Any software I (Elroy) post in these forums (VBForums) written by me is provided ""AS IS"" without warranty of any kind, expressed or implied, and permission is hereby granted, free of charge and without restriction, to any person obtaining a copy. To all, peace and happiness." & vbCrLf & vbCrLf
    
    txtShellPipe.Text = txtShellPipe.Text & "https://www.vbforums.com/showthread.php?660014 (dilettante)" & vbCrLf
    txtShellPipe.Text = txtShellPipe.Text & "" & vbCrLf
    txtShellPipe.Text = txtShellPipe.Text & "No License"
    
    txtVbal.Text = txtVbal.Text & "vbAccelerator Software License" & vbCrLf
    txtVbal.Text = txtVbal.Text & "" & vbCrLf
    txtVbal.Text = txtVbal.Text & "Version 1.0" & vbCrLf
    txtVbal.Text = txtVbal.Text & "" & vbCrLf
    txtVbal.Text = txtVbal.Text & "Copyright (c) 2002 vbAccelerator.com" & vbCrLf
    txtVbal.Text = txtVbal.Text & "" & vbCrLf
    txtVbal.Text = txtVbal.Text & "Redistribution and use in source and binary forms, with or without modification, are permitted provided that the following conditions are met:" & vbCrLf
    txtVbal.Text = txtVbal.Text & "" & vbCrLf
    txtVbal.Text = txtVbal.Text & "    Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer" & vbCrLf
    txtVbal.Text = txtVbal.Text & "    Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer in the documentation and/or other materials provided with the distribution." & vbCrLf
    txtVbal.Text = txtVbal.Text & "    The end-user documentation included with the redistribution, if any, must include the following acknowledgment:" & vbCrLf
    txtVbal.Text = txtVbal.Text & "" & vbCrLf
    txtVbal.Text = txtVbal.Text & "    ""This product includes software developed by vbAccelerator (/index.html).""" & vbCrLf
    txtVbal.Text = txtVbal.Text & "" & vbCrLf
    txtVbal.Text = txtVbal.Text & "    Alternately, this acknowledgment may appear in the software itself, if and wherever such third-party acknowledgments normally appear." & vbCrLf
    txtVbal.Text = txtVbal.Text & "    The names ""vbAccelerator"" and ""vbAccelerator.com"" must not be used to endorse or promote products derived from this software without prior written permission. For written permission, please contact vbAccelerator through steve@vbaccelerator.com." & vbCrLf
    txtVbal.Text = txtVbal.Text & "    Products derived from this software may not be called ""vbAccelerator"", nor may ""vbAccelerator"" appear in their name, without prior written permission of vbAccelerator." & vbCrLf
    txtVbal.Text = txtVbal.Text & "" & vbCrLf
    txtVbal.Text = txtVbal.Text & Mit2
    
'    lvMisc.ColumnHeaders.Add , , t("주소", "URL"), 3135
'    lvMisc.ColumnHeaders.Add(, , t("작성자", "Author"), 1215).Alignment = LvwColumnHeaderAlignmentCenter
'
'    lvMisc.ListItems.Add(, , "https://www.vbforums.com/showthread.php?457171", , 1).ListSubItems.Add , , "MartinLiss"
'    lvMisc.ListItems.Add(, , "https://www.vbforums.com/showthread.php?430704", , 1).ListSubItems.Add , , "DanCool999"
'    lvMisc.ListItems.Add(, , "https://www.codeguru.com/visual-basic/displaying-the-file-properties-dialog/", , 1).ListSubItems.Add , , "Lothar A. Haensler"
'    lvMisc.ListItems.Add(, , "http://vbcity.com/forums/t/105530.aspx", , 1).ListSubItems.Add , , "IanB"
'    lvMisc.ListItems.Add(, , "https://www.vbforums.com/showthread.php?696217", , 1).ListSubItems.Add , , "dilettante"
'    lvMisc.ListItems.Add(, , "https://www.vbforums.com/showthread.php?644597", , 1).ListSubItems.Add , , "Bonnie West"
'    lvMisc.ListItems.Add(, , "https://www.vbforums.com/showthread.php?903019", , 1).ListSubItems.Add , , "AAraya"
'    lvMisc.ListItems.Add(, , "https://www.mrexcel.com/board/threads/194874/", , 1).ListSubItems.Add , , "JoeWeis"
'    lvMisc.ListItems.Add(, , "https://stackoverflow.com/questions/40651", , 1).ListSubItems.Add , , "Christian Hayter"
'    lvMisc.ListItems.Add(, , "https://www.vbforums.com/showthread.php?894947", , 1).ListSubItems.Add , , "wqweto"
'    lvMisc.ListItems.Add(, , "https://gist.github.com/jvarn/5e11b1fd741b5f79d8a516c9c2368f17", , 1).ListSubItems.Add , , "jvarn"
'    lvMisc.ListItems.Add(, , "https://www.vbforums.com/showthread.php?842795", , 1).ListSubItems.Add , , "Elroy"
'    lvMisc.ListItems.Add(, , "https://stackoverflow.com/questions/1230333", , 1).ListSubItems.Add , , "Robert Harvey"
'    lvMisc.ListItems.Add(, , "https://www.vbforums.com/showthread.php?704979", , 1).ListSubItems.Add , , "Max187Boucher"
    
    FrameW1.Caption = t(FrameW1.Caption, "&License")
End Sub

Private Sub lvItems_ItemSelect(ByVal Item As LvwListItem, ByVal Selected As Boolean)
    On Error Resume Next
    If Selected = False Then Exit Sub
    
    Dim i%
    For i = pbLicenses.LBound To pbLicenses.UBound
        pbLicenses(i).Visible = (i = Item.Index)
    Next i
End Sub

'Private Sub lvMisc_ItemDblClick(ByVal Item As LvwListItem, ByVal Button As Integer)
'    Shell "cmd /c start """" " & Item.Text
'End Sub

Private Sub timLicenseLoader_Timer()
    If LineNum > 812 Then
        timLicenseLoader.Enabled = 0
        pbLicenseLoadProgress.Visible = 0
        txtLicense.Height = txtLicense.Height + pbLicenseLoadProgress.Height + 30
        txtLicense.Enabled = -1
        txtLicensePlaceholder.Visible = 0
        Exit Sub
    End If
    
    'On Error GoTo LicenseFail
    txtLicense.Text = txtLicense.Text & LoadResString(LineNum) & vbCrLf
    pbLicenseLoadProgress.Value = LineNum
    txtLicensePlaceholder.Text = t("라이선스를 불러오는 중... (", "Loading the license text... (") & Floor(LineNum / 812 * 100) & "%)"
    LineNum = LineNum + 1
'    Exit Sub
'
'LicenseFail:
'    txtLicense.Text = t("라이선스를 불러올 수 없습니다. 다음 링크에서 확인할 수 있습니다.", "Unable to load the license. Check this URL:") & vbCrLf & " https://raw.githubusercontent.com/nodejs/node/refs/heads/v0.10/LICENSE"
'    timLicenseLoader.Enabled = 0
'    pbLicenseLoadProgress.Visible = 0
'    txtLicense.Height = txtLicense.Height + pbLicenseLoadProgress.Height + 30
'    txtLicense.Enabled = -1
'    txtLicensePlaceholder.Visible = 0
End Sub

