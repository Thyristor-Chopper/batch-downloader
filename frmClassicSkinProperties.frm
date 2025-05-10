VERSION 5.00
Begin VB.Form frmClassicSkinProperties 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "스킨 설정"
   ClientHeight    =   1260
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "굴림"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmClassicSkinProperties.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1260
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin prjDownloadBooster.CommandButtonW cmdOK 
      Default         =   -1  'True
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   720
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "확인"
   End
   Begin prjDownloadBooster.CommandButtonW cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   720
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "취소"
   End
   Begin prjDownloadBooster.CheckBoxW chkRoundClassicButtons 
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   450
      Caption         =   "둥근 단추 사용(&U)"
   End
End
Attribute VB_Name = "frmClassicSkinProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    SaveSetting "DownloadBooster", "Options", "RoundClassicButtons", chkRoundClassicButtons.Value
    frmOptions.VisualStyleChanged = True
    frmOptions.cmdApply.Enabled = True
    frmOptions.cmdSample.RoundButton = chkRoundClassicButtons.Value
    Unload Me
End Sub

Private Sub Form_Load()
    InitForm Me
    
    chkRoundClassicButtons.Value = GetSetting("DownloadBooster", "Options", "RoundClassicButtons", 0)
    
    tr Me, "Skin Settings"
    tr chkRoundClassicButtons, "&Use rounded buttons"
    tr cmdOK, "OK"
    tr cmdCancel, "Cancel"
End Sub
