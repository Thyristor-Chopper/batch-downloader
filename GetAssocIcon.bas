Attribute VB_Name = "GetAssocIconMod"
'https://www.vbforums.com/showthread.php?696217-How-do-I-load-an-EXE-or-DLL-file-icon

Option Explicit

Private Const S_OK As Long = 0
Private Const MAX_PATH As Long = 260
Private Const SHGFI_ICON As Long = &H100&
Private Const SHGFI_LARGEICON As Long = &H0&  '32x32 pixels.
Private Const SHGFI_SMALLICON As Long = &H1&  '16x16 pixels.
Private Const SHGFI_USEFILEATTRIBUTES As Long = &H10&

Private Type SHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * MAX_PATH
    szTypeName As String * 80
End Type

Private Type PictDesc_Icon
    cbSizeOfStruct As Long
    PicType As Long
    hIcon As Long
End Type

Private Type Guid
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Declare Function SHGetFileInfo Lib "shell32" Alias "SHGetFileInfoW" ( _
    ByVal pszPath As Long, _
    ByVal dwFileAttributes As Long, _
    ByVal psfi As Long, _
    ByVal cbSizeFileInfo As Long, _
    ByVal uFlags As Long) As Long

Private Declare Function GetFullPathName Lib "kernel32" Alias "GetFullPathNameW" ( _
    ByVal lpFileName As Long, _
    ByVal nBufferLength As Long, _
    ByVal lpBuffer As Long, _
    ByVal lpFilePart As Long) As Long

Private Declare Function OleCreatePictureIndirect Lib "olepro32" ( _
    ByVal lpPictDesc As Long, _
    ByVal riid As Long, _
    ByVal fOwn As Long, _
    ByRef lplpvObj As IPicture) As Long

Private IPictureIID As Guid

Public Function GetAssocIcon( _
    ByVal PathToFile As String, _
    Optional ByVal Attributes As Long = 0, _
    Optional ByVal LargeIcon As Boolean = False, _
    Optional ByVal Extension As Boolean = False) As StdPicture
    'Returns a StdPicture object on success.
    '
    'On any error (or no associated icon) Nothing is returned.
    '
    'PathToFile
    '
    '   This should be an absolute or relative file path or extension,
    '   such as:
    '
    '       o An executable file (EXE, DLL, OCX?), or
    '       o An .ico file, or
    '       o A data file of some "type" (as defined by its file extension
    '         Value) that has a file association, or
    '       o A file extension in the form ".ext" or "anything.ext" but
    '         only if Extension = True!
    '
    'LargeIcon
    '
    '   True:   Returns 32x32 icon.
    '   False:  Returns 16x16 icon.
    '
    'Extension
    '
    '   True:   The PathToFile is not actually examined, and can even be
    '           just an empty name and extension, e.g. ".txt" alone, or
    '           pass just "" for the "generic" icon.
    '   False:  PathToFile must exist.
    '
    Dim SFI As SHFILEINFO
    Dim Desc As PictDesc_Icon
    
    If Len(PathToFile) = 0 And Extension Then PathToFile = "x" 'Win7 "generic icon" request fix.
    If SHGetFileInfo(StrPtr(PathToFile), Attributes, VarPtr(SFI), LenB(SFI), SHGFI_ICON Or IIf(LargeIcon, SHGFI_LARGEICON, SHGFI_SMALLICON) Or IIf(Extension, SHGFI_USEFILEATTRIBUTES, 0)) = 0 Then
        If LargeIcon Then
            Set GetAssocIcon = frmExplorer.imgFolder.ListImages(2).Picture
        Else
            Set GetAssocIcon = frmExplorer.imgFolderSmall.ListImages(2).Picture
        End If
        Exit Function
    End If
    If IPictureIID.Data1 = 0 Then
        'Initialize once on first call.
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
    End If
    With Desc
       .cbSizeOfStruct = Len(Desc)
       .PicType = vbPicTypeIcon
       .hIcon = SFI.hIcon
    End With
    If OleCreatePictureIndirect(VarPtr(Desc), VarPtr(IPictureIID), True, GetAssocIcon) <> S_OK Then
        If LargeIcon Then
            Set GetAssocIcon = frmExplorer.imgFolder.ListImages(2).Picture
        Else
            Set GetAssocIcon = frmExplorer.imgFolderSmall.ListImages(2).Picture
        End If
    End If
End Function
