Attribute VB_Name = "AlphaPNG"
'https://www.vbforums.com/showthread.php?896878-PNG-with-alpha-channel-into-standard-VB6-image-control

Option Explicit

Private mlGdipToken             As Long
Private Type GdiplusStartupInput
    GdiplusVersion              As Long
    DebugEventCallback          As Long
    SuppressBackgroundThread    As Long
    SuppressExternalCodecs      As Long
End Type
Private Declare Function GdiplusStartup Lib "gdiplus" (ByRef token As Long, ByRef lpInput As GdiplusStartupInput, Optional ByRef lpOutput As Long) As Long
Private Declare Function GdipLoadImageFromFile Lib "gdiplus" (ByVal sFilename As Long, hImage As Long) As Long
'Private Declare Function GdipCreateBitmapFromFile Lib "gdiplus" (ByVal pFilename As Long, ByRef pbitmap As Long) As Long
Private Type BitmapData
    Width               As Long
    Height              As Long
    Stride              As Long
    PixelFormat         As Long ' ImageColorFormatConstants can be used here.
    Scan0               As Long
    Reserved            As Long
End Type
Private Const ImageLockModeRead As Long = &H1&
Private Const PixelFormat32bppPARGB As Long = &HE200B             ' 32 bits per pixel; 8 bits each are used for the alpha, red, green, and blue components. The red, green, and blue components are premultiplied according to the alpha component.
Private Declare Function GdipBitmapLockBits Lib "gdiplus" (ByVal hBitmap As Long, lpRect As Any, ByVal lFlags As Long, ByVal lPixelFormat As Long, uLockedBitmapData As BitmapData) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Type BITMAPINFOHEADER
    biSize              As Long
    biWidth             As Long
    biHeight            As Long
    biPlanes            As Integer
    biBitCount          As Integer
    biCompression       As Long
    biSizeImage         As Long
    biXPelsPerMeter     As Long
    biYPelsPerMeter     As Long
    biClrUsed           As Long
    biClrImportant      As Long
End Type
Private Const DIB_RGB_COLORS As Long = 0&
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As Long, lpBitsInfo As BITMAPINFOHEADER, ByVal wUsage As Long, lpBitsOut As Long, ByVal hSection As Long, ByVal offset As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GdipBitmapUnlockBits Lib "gdiplus" (ByVal hBitmap As Long, uLockedBitmapData As BitmapData) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Type SIZEL
    CX As Long
    CY As Long
End Type
Private Declare Sub AtlPixelToHiMetric Lib "atl" (lpSizeInPix As SIZEL, lpSizeInHiMetric As SIZEL)
Private Type RECT
    Left    As Long
    Top     As Long
    Right   As Long
    Bottom  As Long
End Type
'Private Declare Function CreateEnhMetaFile Lib "gdi32" Alias "CreateEnhMetaFileA" (ByVal hdcRef As Long, ByVal lpFileName As Long, lpRect As RECT, ByVal lpDescription As Long) As Long
Private Declare Function CreateEnhMetaFileW Lib "gdi32" (ByVal hdcRef As Long, ByVal lpFileName As Long, lpRect As Any, ByVal lpDescription As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Type BLENDFUNCTION
    BlendOp             As Byte
    BlendFlags          As Byte
    SourceConstantAlpha As Byte
    AlphaFormat         As Byte
End Type
Private Const AC_SRC_ALPHA As Byte = 1
Private Declare Function GetMem4 Lib "msvbvm60" (ByRef Source As Any, ByRef Dest As Any) As Long ' Always ignore the returned value, it's useless.
Private Declare Function GdiAlphaBlend Lib "gdi32" (ByVal hdcDest As Long, ByVal xoriginDest As Long, ByVal yoriginDest As Long, ByVal wDest As Long, ByVal hDest As Long, ByVal hdcSrc As Long, ByVal xoriginSrc As Long, ByVal yoriginSrc As Long, ByVal wSrc As Long, ByVal hSrc As Long, ByVal ftn As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CloseEnhMetaFile Lib "gdi32" (ByVal hDC As Long) As Long
Private Type PICTDESC
    cbSize          As Long
    PicType         As Long
    hgdiObj         As Long
    hPalOrXYExt     As Long
    Reserved        As Long
End Type
Private Type RECTL
    Left    As Long
    Top     As Long
    Right   As Long
    Bottom  As Long
End Type
Private Type ENHMETAHEADER
    iType As Long
    nSize As Long
    rclBounds As RECTL
    rclFrame As RECTL
    dSignature As Long
    nVersion As Long
    nBytes As Long
    nRecords As Long
    nHandles As Integer
    sReserved As Integer
    nDescription As Long
    offDescription As Long
    nPalEntries As Long
    szlDevice As SIZEL
    szlMillimeters As SIZEL
    cbPixelFormat As Long
    offPixelFormat As Long
    bOpenGL As Long
    szlMicrometers As SIZEL
End Type
Private Declare Function GetEnhMetaFileHeader Lib "gdi32" (ByVal hEmf As Long, ByVal cbBuffer As Long, ByRef lpemh As ENHMETAHEADER) As Long
Private Declare Function DeleteEnhMetaFile Lib "gdi32" (ByVal hEmf As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "oleaut32" (lpPictDesc As PICTDESC, riid As IID, ByVal fOwn As Boolean, lplpvObj As Object) As Long
Private Type IID
    Data1       As Long
    Data2       As Integer
    Data3       As Integer
    Data4(7&)   As Byte
End Type
Private Declare Function IIDFromString Lib "ole32" (ByVal lpsz As Long, ByRef CLSID As IID) As Long
Private Declare Function GdiplusShutdown Lib "gdiplus" (ByVal token As Long) As Long

Public Function LoadPngIntoPictureWithAlpha(sFilename As String, Optional ByVal bbOverallOpacity As Byte = &HFF, Optional ByVal dScalingFactor As Single = 1!) As IPicture
    ' The sFilename should be a valid PNG file.
    ' We can put the return of this directly into an Image.Picture or a PictureBox.Picture, and it will correctly show any alpha channel.
    '
    ' Get the GDI+ going.
    Dim StartupInput As GdiplusStartupInput
    StartupInput.GdiplusVersion = 1&
    GErr GdiplusStartup(mlGdipToken, StartupInput, 0&)
    '
    ' Open the file into hGdipImage.
    Dim hGdipImage         As Long
    GErr GdipLoadImageFromFile(StrPtr(sFilename), hGdipImage)
    'GErr GdipCreateBitmapFromFile(StrPtr(sFilename), hGdipImage)   ' This also works, as hGdipBitmap implements hGdipImage.
    '
    ' Creates a temporary buffer of the hGdipImage's or hGdipBitmap's bits.
    ' The bits in this temp buffer don't have to be the same format as the original bits format.
    ' The uData can be both an input and output, but only output here because uData flag is zero.
    ' Also PNG files aren't pre-multiplied, but we let this function do that for us, as that's what we need.
    Dim uData           As BitmapData
    GErr GdipBitmapLockBits(hGdipImage, ByVal 0&, ImageLockModeRead, PixelFormat32bppPARGB, uData)
    '
    ' Get screen compatible DC.
    Dim hMemDC          As Long
    hMemDC = CreateCompatibleDC(0&)
    '
    ' Create BITMAPINFOHEADER header for making DIB.
    Dim uHdr As BITMAPINFOHEADER
    uHdr.biSize = Len(uHdr)
    uHdr.biPlanes = 1
    uHdr.biBitCount = 32
    uHdr.biWidth = uData.Width      ' Pixels.
    uHdr.biHeight = -uData.Height   ' Pixels.
    uHdr.biSizeImage = uData.Stride * uData.Height
    '
    ' Creates an EMPTY buffer associated with the DC for image (DIB) bits,
    ' and returns pointer to the buffer (lpBits).
    ' CreateDIBSection does not use the BITMAPINFOHEADER biXPelsPerMeter or biYPelsPerMeter
    ' and will not provide resolution information in the BITMAPINFO structure.
    Dim hDib            As Long
    Dim lpBits          As Long
    hDib = ApiZ(CreateDIBSection(hMemDC, uHdr, DIB_RGB_COLORS, lpBits, 0&, 0&))
    '
    ' Copy the actual image (PARGB bits) from uData (uData.Scan0) into our DIBs bits (lpBits).
    Call CopyMemory(ByVal lpBits, ByVal uData.Scan0, uData.Stride * uData.Height)
    '
    ' We're done with the uData buffer as well as the hGdipImage so clean them up.
    GErr GdipBitmapUnlockBits(hGdipImage, uData)
    GErr GdipDisposeImage(hGdipImage)
    '
    ' Put our DIB into our memory DC.  Second time we've used it.  First time was just to create a compatible DIB.
    Dim hPrevDib        As Long
    hPrevDib = ApiZ(SelectObject(hMemDC, hDib))         ' Save the initial DIB as we're suppose to put it back when done.
    '
    ' Create an EMPTY EMF in a primary monitor DC with no initial size.
    ' The DC returned by CreateEnhMetaFile can be passed to any GDI function.
    Dim hEmfDC          As Long
    hEmfDC = ApiZ(CreateEnhMetaFileW(0&, 0&, ByVal 0&, 0&))  ' It actually returns an EMF DC.
    '
    ' Calculate the EMF scaling factors from its DC, so we can use it in GdiAlphaBlend, as GdiAlphaBlend scales based on the hEmfDC.
    Const HORZSIZE   As Long = 4&:  Const VERTSIZE   As Long = 6&
    Const HORZRES    As Long = 8&:  Const VERTRES    As Long = 10&
    Const LOGPIXELSX As Long = 88&: Const LOGPIXELSY As Long = 90&
    Dim Xscale As Double, Yscale As Double
    Xscale = CDbl(GetDeviceCaps(hEmfDC, HORZRES)) / CDbl(GetDeviceCaps(hEmfDC, HORZSIZE)) * 25.4 / CDbl(GetDeviceCaps(hEmfDC, LOGPIXELSX))
    Yscale = CDbl(GetDeviceCaps(hEmfDC, VERTRES)) / CDbl(GetDeviceCaps(hEmfDC, VERTSIZE)) * 25.4 / CDbl(GetDeviceCaps(hEmfDC, LOGPIXELSY))
    '
    ' Create BLENDFUNCTION structure for an Alpha Blend.
    Const AC_SRC_OVER  As Byte = 0: Const AC_SRC_ALPHA As Byte = 1
    Dim bf As BLENDFUNCTION
    bf.BlendOp = AC_SRC_OVER
    bf.AlphaFormat = AC_SRC_ALPHA
    bf.SourceConstantAlpha = bbOverallOpacity           ' Full opacity for overall image, other than alpha channel applied.
    Dim ftn As Long
    GetMem4 bf, ftn                                     ' Must put into a Long so we can pass it ByVal.
    '
    ' Copy our DIB that's in our memory DC into our EMF+ using its DC, and our scale factors.
    ApiZ GdiAlphaBlend(hEmfDC, 0&, 0&, CLng((CDbl(uData.Width)) * Xscale * dScalingFactor) + 1&, _
                                       CLng((CDbl(uData.Height)) * Yscale * dScalingFactor) + 1&, _
                       hMemDC, 0&, 0&, uData.Width, uData.Height, ftn), "AlphaBlend"
    '
    ' Done with hMemDC and hDib, so clean them up.
    ApiZ SelectObject(hMemDC, hPrevDib)
    ApiZ DeleteDC(hMemDC)
    ApiZ DeleteObject(hDib)
    '
    ' Done with hEmfDC so clean it up and just save the hEmf (which is returned from CloseEnhMetaFile).
    Dim hEmf            As Long
    hEmf = ApiZ(CloseEnhMetaFile(hEmfDC), "CloseEnhMetaFile")
    '
    ' Setup PictDesc with EMF type.
    Dim uDesc As PICTDESC
    uDesc.cbSize = Len(uDesc)
    uDesc.PicType = vbPicTypeEMetafile
    uDesc.hgdiObj = hEmf
    '
    ' Wrap our EMF (from uDesc) into an IPicture object, which is our function's return.
    ApiE OleCreatePictureIndirect(uDesc, IPictureIID, 1&, LoadPngIntoPictureWithAlpha)
    '
    ' Shutdown the GDI+.
    GErr GdiplusShutdown(mlGdipToken)
End Function

Private Function IPictureIID() As IID
    ApiE IIDFromString(StrPtr("{7BF80980-BF32-101A-8BBB-00AA00300CAB}"), IPictureIID)
End Function



Public Function ApiZ(ApiReturn As Long, Optional sApiCall As String) As Long
    ' This one is for API calls that report error by returning ZERO.
    '
    If ApiReturn <> 0& Then
        ApiZ = ApiReturn
        Exit Function
    End If
    '
    Dim sErr As String
    If Len(sApiCall) Then
        sErr = sApiCall & " error"
    Else
        sErr = "API Error"
    End If
    '
    Dim InIDE As Boolean: Debug.Assert MakeTrue(InIDE)
    If False Then
        Debug.Print sErr
        Stop
    Else
        'Err.Raise vbObjectError + 1147221504, , sErr
    End If
End Function

Public Sub ApiE(ApiReturn As Long, Optional sApiCall As String)
    ' Just a general error processing procedure for non-GDI+ API errors.
    ' For API calls where 0& is OK.
    '
    If ApiReturn = 0& Then Exit Sub
    '
    Dim sErr As String
    If Len(sApiCall) Then
        sErr = sApiCall & " error " & CStr(ApiReturn)
    Else
        sErr = "API Error " & CStr(ApiReturn)
    End If
    '
    Dim InIDE As Boolean: Debug.Assert MakeTrue(InIDE)
    If False Then
        Debug.Print sErr
        Stop
    Else
        Err.Raise vbObjectError + 1147221504 - ApiReturn, , sErr
    End If
End Sub

Public Sub GErr(ByVal GdipReturn As Long)
    ' Just to check for any errors during development.
    ' It's public because there are a few cases where GDI+ is used outside of this module.
    '
    If GdipReturn = 0& Then Exit Sub ' All is well.
    '
    Dim sErr As String
    Select Case GdipReturn
    Case 1&:    sErr = "Generic Error"
    Case 2&:    sErr = "Invalid Parameter/Argument"
    Case 3&:    sErr = "Out Of Memory"
    Case 4&:    sErr = "Object Busy, already in use in another thread"
    Case 5&:    sErr = "Insufficient Buffer, buffer specified as an argument in the API call is not large enough"
    Case 6&:    sErr = "Method Not Implemented"
    Case 7&:    sErr = "Win32 Error"
    Case 8&:    sErr = "Wrong State"
    Case 9&:    sErr = "Method Aborted"
    Case 10&:   sErr = "File Not Found"
    Case 11&:   sErr = "Value Overflow, arithmetic operation that produced a numeric overflow"
    Case 12&:   sErr = "Access Denied"
    Case 13&:   sErr = "Unknown Image Format"
    Case 14&:   sErr = "Font Family Not Found"
    Case 15&:   sErr = "Font Style Not Found"
    Case 16&:   sErr = "Not TrueType Font"
    Case 17&:   sErr = "Unsupported Gdiplus Version"
    Case 18&:   sErr = "Gdiplus Not Initialized"
    Case 19&:   sErr = "Property Not Found, does not exist in the image"
    Case 20&:   sErr = "Property Not Supported, not supported by the format of the image"
    Case 21&:   sErr = "Profile Not Found, color profile required to save an image in CMYK format was not found"
    Case Else:  sErr = "Error Not Specified": GdipReturn = 99&
    End Select
    '
    sErr = "GDI+ Error:  " & sErr
    Dim InIDE As Boolean: Debug.Assert MakeTrue(InIDE)
    If False Then
        Debug.Print sErr
        Stop
    Else
        If GdipReturn = 3& Or GdipReturn = 2& Then Exit Sub
        'Err.Raise vbObjectError + 1147221504 - GdipReturn, , sErr
    End If
End Sub

' If you wish to interpret those errors if/when they occur, here's an Enum that will help.
'Public Enum GdipErrors
'    ' This is a custom enum for handling GDI+ errors, if they're trapped.
'    ' It's here (rather than the enum module) because they're specific to errors.
'    ' These correspond to what's raised in the Err.Raise within the GErr procedure.
'    '
'    GdipErrGeneric = -1000000001
'    GdipErrInvalidParamArg = -1000000002
'    GdipErrOutOfMemory = -1000000003
'    GdipErrObjectBusy = -1000000004
'    GdipErrInsufficientBuffer = -1000000005
'    GdipErrMethodNotImplemented = -1000000006
'    GdipErrWin32Error = -1000000007
'    GdipErrWrongState = -1000000008
'    GdipErrMethodAborted = -1000000009
'    GdipErrFileNotFound = -1000000010
'    GdipErrValueOverflow = -1000000011
'    GdipErrAccessDenied = -1000000012
'    GdipErrUnknownImageFormat = -1000000013
'    GdipErrFontFamilyNotFound = -1000000014
'    GdipErrFontStyleNotFound = -1000000015
'    GdipErrNotTrueTypeFont = -1000000016
'    GdipErrUnsupportedGdiplusVersion = -1000000017
'    GdipErrGdiplusNotInitialized = -1000000018
'    GdipErrPropertyNotFound = -1000000019
'    GdipErrPropertyNotSupported = -1000000020
'    GdipErrProfileNotFound = -1000000021
'    GdipErrNotSpecified = -1000000099
'    '
'    #If False Then ' IntelliSense fix.
'        Dim GdipErrGeneric, GdipErrInvalidParamArg, GdipErrOutOfMemory, GdipErrObjectBusy, GdipErrInsufficientBuffer, GdipErrMethodNotImplemented, GdipErrWin32Error, GdipErrWrongState
'        Dim GdipErrMethodAborted, GdipErrFileNotFound, GdipErrValueOverflow, GdipErrAccessDenied, GdipErrUnknownImageFormat, GdipErrFontFamilyNotFound, GdipErrFontStyleNotFound
'        Dim GdipErrNotTrueTypeFont, GdipErrUnsupportedGdiplusVersion, GdipErrGdiplusNotInitialized, GdipErrPropertyNotFound, GdipErrPropertyNotSupported, GdipErrProfileNotFound, GdipErrNotSpecified
'    #End If
'End Enum



Private Function MakeTrue(ByRef B As Boolean) As Boolean
    B = True
    MakeTrue = True
End Function

