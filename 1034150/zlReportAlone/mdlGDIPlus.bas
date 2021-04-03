Attribute VB_Name = "mdlGDIPlus"
Option Explicit
 
Private Type GUID
   Data1    As Long
   Data2    As Integer
   Data3    As Integer
   Data4(7) As Byte
End Type
 
Private Type PICTDESC
   size     As Long
   Type     As Long
   hBmp     As Long
   hPal     As Long
   Reserved As Long
End Type
 
Private Type GdiplusStartupInput
    GdiplusVersion           As Long
    DebugEventCallback       As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs   As Long
End Type
 
Private Type PWMFRect16
    left   As Integer
    top    As Integer
    Right  As Integer
    Bottom As Integer
End Type
 
Private Type wmfPlaceableFileHeader
    Key         As Long
    hMf         As Integer
    BoundingBox As PWMFRect16
    Inch        As Integer
    Reserved    As Long
    CheckSum    As Integer
End Type

Public Type Clsid
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(0 To 7) As Byte
End Type

' Image file format identifiers
Public Enum GpImageFormatIdentifiers
    GpImageFormatUndefined = 0
    GpImageFormatMemoryBMP = 1
    GpImageFormatBMP = 2
    GpImageFormatEMF = 3
    GpImageFormatWMF = 4
    GpImageFormatJPEG = 5
    GpImageFormatPNG = 6
    GpImageFormatGIF = 7
    GpImageFormatTIFF = 8
    GpImageFormatEXIF = 9
    GpImageFormatIcon = 10
End Enum

' NOTE: Enums evaluate to a Long
Public Enum GpStatus   ' aka Status
   Ok = 0
   GenericError = 1
   InvalidParameter = 2
   OutOfMemory = 3
   ObjectBusy = 4
   InsufficientBuffer = 5
   NotImplemented = 6
   Win32Error = 7
   WrongState = 8
   Aborted = 9
   FileNotFound = 10
   ValueOverflow = 11
   AccessDenied = 12
   UnknownImageFormat = 13
   FontFamilyNotFound = 14
   FontStyleNotFound = 15
   NotTrueTypeFont = 16
   UnsupportedGdiplusVersion = 17
   GdiplusNotInitialized = 18
   PropertyNotFound = 19
   PropertyNotSupported = 20
End Enum

' Image Format
Public Const ImageFormatSuffix        As String = "-0728-11D3-9D7B-0000F81EF32E}"
Public Const ImageFormatUndefined     As String = "{B96B3CA9" & ImageFormatSuffix
Public Const ImageFormatMemoryBMP     As String = "{B96B3CAA" & ImageFormatSuffix
Public Const ImageFormatBMP           As String = "{B96B3CAB" & ImageFormatSuffix
Public Const ImageFormatEMF           As String = "{B96B3CAC" & ImageFormatSuffix
Public Const ImageFormatWMF           As String = "{B96B3CAD" & ImageFormatSuffix
Public Const ImageFormatJPEG          As String = "{B96B3CAE" & ImageFormatSuffix
Public Const ImageFormatPNG           As String = "{B96B3CAF" & ImageFormatSuffix
Public Const ImageFormatGIF           As String = "{B96B3CB0" & ImageFormatSuffix
Public Const ImageFormatTIFF          As String = "{B96B3CB1" & ImageFormatSuffix
Public Const ImageFormatEXIF          As String = "{B96B3CB2" & ImageFormatSuffix
Public Const ImageFormatIcon          As String = "{B96B3CB5" & ImageFormatSuffix
 
' GDI Functions
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PICTDESC, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function PatBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
 
' GDI+ functions
Private Declare Function GdiplusStartup Lib "gdiplus.dll" (Token As Long, gdipInput As GdiplusStartupInput, GdiplusStartupOutput As Long) As Long
Private Declare Function GdipCreateFromHDC Lib "gdiplus.dll" (ByVal hDC As Long, GpGraphics As Long) As Long
Private Declare Function GdipSetInterpolationMode Lib "gdiplus.dll" (ByVal Graphics As Long, ByVal InterMode As Long) As Long
Private Declare Function GdipDrawImageRectI Lib "gdiplus.dll" (ByVal Graphics As Long, ByVal img As Long, ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus.dll" (ByVal Graphics As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus.dll" (ByVal Image As Long) As Long
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "gdiplus.dll" (ByVal hBmp As Long, ByVal hPal As Long, GpBitmap As Long) As Long
Private Declare Function GdipGetImageWidth Lib "gdiplus.dll" (ByVal Image As Long, Width As Long) As Long
Private Declare Function GdipGetImageHeight Lib "gdiplus.dll" (ByVal Image As Long, Height As Long) As Long
Private Declare Function GdipCreateMetafileFromWmf Lib "gdiplus.dll" (ByVal hWmf As Long, ByVal deleteWmf As Long, WmfHeader As wmfPlaceableFileHeader, Metafile As Long) As Long
Private Declare Function GdipCreateMetafileFromEmf Lib "gdiplus.dll" (ByVal hEmf As Long, ByVal deleteEmf As Long, Metafile As Long) As Long
Private Declare Function GdipCreateBitmapFromHICON Lib "gdiplus.dll" (ByVal hIcon As Long, GpBitmap As Long) As Long
Private Declare Function GdipDrawImageRectRectI Lib "gdiplus.dll" (ByVal Graphics As Long, ByVal GpImage As Long, ByVal dstx As Long, ByVal dsty As Long, ByVal dstwidth As Long, ByVal dstheight As Long, ByVal srcx As Long, ByVal srcy As Long, ByVal srcwidth As Long, ByVal srcheight As Long, ByVal srcUnit As Long, ByVal imageAttributes As Long, ByVal callback As Long, ByVal callbackData As Long) As Long
Private Declare Sub GdiplusShutdown Lib "gdiplus.dll" (ByVal Token As Long)

Private Declare Function GdipGetImageRawFormat Lib "gdiplus" (ByVal Image As Long, format As Clsid) As GpStatus
Public Declare Function GdipLoadImageFromFile Lib "gdiplus.dll" (ByVal FileName As Long, GpImage As Long) As Long
 
' GDI and GDI+ constants
Private Const PLANES = 14            '  Number of planes
Private Const BITSPIXEL = 12         '  Number of bits per pixel
Private Const PATCOPY = &HF00021     ' (DWORD) dest = pattern
Private Const PICTYPE_BITMAP = 1     ' Bitmap type
Private Const InterpolationModeHighQualityBicubic = 7
Private Const GDIP_WMF_PLACEABLEKEY = &H9AC6CDD7
Private Const UnitPixel = 2
 
' Initialises GDI Plus
Public Function InitGDIPlus() As Long
    Dim Token    As Long
    Dim gdipInit As GdiplusStartupInput
    
    gdipInit.GdiplusVersion = 1
    GdiplusStartup Token, gdipInit, ByVal 0&
    InitGDIPlus = Token
End Function
 
' Frees GDI Plus
Public Sub FreeGDIPlus(Token As Long)
    GdiplusShutdown Token
End Sub
 
' Loads the picture (optionally resized)
Public Function LoadPictureGDIPlus(PicFile As String, Optional Width As Long = -1, Optional Height As Long = -1 _
    , Optional ByVal BackColor As Long = vbWhite, Optional RetainRatio As Boolean = False) As IPicture
    
    Dim hDC     As Long
    Dim hBitmap As Long
    Dim img     As Long
        
    ' Load the image
    If GdipLoadImageFromFile(StrPtr(PicFile), img) <> 0 Then
        Err.Raise 999, "GDI+ Module", "Error loading picture " & PicFile
        Exit Function
    End If
    
    ' Calculate picture's width and height if not specified
    If Width = -1 Or Height = -1 Then
        GdipGetImageWidth img, Width
        GdipGetImageHeight img, Height
    End If
    
    ' Initialise the hDC
    InitDC hDC, hBitmap, BackColor, Width, Height
 
    ' Resize the picture
    gdipResize img, hDC, Width, Height, RetainRatio
    GdipDisposeImage img
    
    ' Get the bitmap back
    GetBitmap hDC, hBitmap
 
    ' Create the picture
    Set LoadPictureGDIPlus = CreatePicture(hBitmap)
End Function
 
' Initialises the hDC to draw
Private Sub InitDC(hDC As Long, hBitmap As Long, BackColor As Long, Width As Long, Height As Long)
    Dim hBrush As Long
        
    ' Create a memory DC and select a bitmap into it, fill it in with the backcolor
    hDC = CreateCompatibleDC(ByVal 0&)
    hBitmap = CreateBitmap(Width, Height, GetDeviceCaps(hDC, PLANES), GetDeviceCaps(hDC, BITSPIXEL), ByVal 0&)
    hBitmap = SelectObject(hDC, hBitmap)
    hBrush = CreateSolidBrush(BackColor)
    hBrush = SelectObject(hDC, hBrush)
    PatBlt hDC, 0, 0, Width, Height, PATCOPY
    DeleteObject SelectObject(hDC, hBrush)
End Sub
 
' Resize the picture using GDI plus
Private Sub gdipResize(img As Long, hDC As Long, Width As Long, Height As Long, Optional RetainRatio As Boolean = False)
    Dim Graphics   As Long      ' Graphics Object Pointer
    Dim OrWidth    As Long      ' Original Image Width
    Dim OrHeight   As Long      ' Original Image Height
    Dim OrRatio    As Double    ' Original Image Ratio
    Dim DesRatio   As Double    ' Destination rect Ratio
    Dim DestX      As Long      ' Destination image X
    Dim DestY      As Long      ' Destination image Y
    Dim DestWidth  As Long      ' Destination image Width
    Dim DestHeight As Long      ' Destination image Height
    
    GdipCreateFromHDC hDC, Graphics
    GdipSetInterpolationMode Graphics, InterpolationModeHighQualityBicubic
    
    If RetainRatio Then
        GdipGetImageWidth img, OrWidth
        GdipGetImageHeight img, OrHeight
        
        OrRatio = OrWidth / OrHeight
        DesRatio = Width / Height
        
        ' Calculate destination coordinates
        DestWidth = IIF(DesRatio < OrRatio, Width, Height * OrRatio)
        DestHeight = IIF(DesRatio < OrRatio, Width / OrRatio, Height)
'        DestX = (Width - DestWidth) / 2
'        DestY = (Height - DestHeight) / 2
 
        DestX = 0
        DestY = 0
 
        GdipDrawImageRectRectI Graphics, img, DestX, DestY, DestWidth, DestHeight, 0, 0, OrWidth, OrHeight, UnitPixel, 0, 0, 0
    Else
        GdipDrawImageRectI Graphics, img, 0, 0, Width, Height
    End If
    GdipDeleteGraphics Graphics
End Sub
 
' Replaces the old bitmap of the hDC, Returns the bitmap and Deletes the hDC
Private Sub GetBitmap(hDC As Long, hBitmap As Long)
    hBitmap = SelectObject(hDC, hBitmap)
    DeleteDC hDC
End Sub
 
' Creates a Picture Object from a handle to a bitmap
Private Function CreatePicture(hBitmap As Long) As IPicture
    Dim IID_IDispatch As GUID
    Dim pic           As PICTDESC
    Dim IPic          As IPicture
    
    ' Fill in OLE IDispatch Interface ID
    IID_IDispatch.Data1 = &H20400
    IID_IDispatch.Data4(0) = &HC0
    IID_IDispatch.Data4(7) = &H46
        
    ' Fill Pic with necessary parts
    pic.size = Len(pic)        ' Length of structure
    pic.Type = PICTYPE_BITMAP  ' Type of Picture (bitmap)
    pic.hBmp = hBitmap         ' Handle to bitmap
 
    ' Create the picture
    OleCreatePictureIndirect pic, IID_IDispatch, True, IPic
    Set CreatePicture = IPic
End Function
 
' Returns a resized version of the picture
Public Function Resize(Handle As Long, PicType As PictureTypeConstants, Width As Long, Height As Long, Optional BackColor As Long = vbWhite, Optional RetainRatio As Boolean = False) As IPicture
    Dim img       As Long
    Dim hDC       As Long
    Dim hBitmap   As Long
    Dim WmfHeader As wmfPlaceableFileHeader
    
    ' Determine pictyre type
    Select Case PicType
    Case vbPicTypeBitmap
         GdipCreateBitmapFromHBITMAP Handle, ByVal 0&, img
    Case vbPicTypeMetafile
         FillInWmfHeader WmfHeader, Width, Height
         GdipCreateMetafileFromWmf Handle, False, WmfHeader, img
    Case vbPicTypeEMetafile
         GdipCreateMetafileFromEmf Handle, False, img
    Case vbPicTypeIcon
         ' Does not return a valid Image object
         GdipCreateBitmapFromHICON Handle, img
    End Select
    
    ' Continue with resizing only if we have a valid image object
    If img Then
        InitDC hDC, hBitmap, BackColor, Width, Height
        gdipResize img, hDC, Width, Height, RetainRatio
        GdipDisposeImage img
        GetBitmap hDC, hBitmap
        Set Resize = CreatePicture(hBitmap)
    End If
End Function
 
' Fills in the wmfPlacable header
Private Sub FillInWmfHeader(WmfHeader As wmfPlaceableFileHeader, Width As Long, Height As Long)
    WmfHeader.BoundingBox.Right = Width
    WmfHeader.BoundingBox.Bottom = Height
    WmfHeader.Inch = 1440
    WmfHeader.Key = GDIP_WMF_PLACEABLEKEY
End Sub

Private Function hexPad(ByVal value As Long, ByVal padSize As Long) As String
    Dim sRet As String
    Dim lMissing As Long
    
    sRet = Hex$(value)
    lMissing = padSize - Len(sRet)
    If (lMissing > 0) Then
       sRet = String$(lMissing, "0") & sRet
    ElseIf (lMissing < 0) Then
       sRet = Mid$(sRet, -lMissing + 1)
    End If
    hexPad = sRet
End Function

' Get the string for the GUID
Public Function GetGuidString(GUID As Clsid) As String
    Dim i As Long
    Dim sGuid As String

    sGuid = "{" & hexPad(GUID.Data1, 8) & "-" & hexPad(GUID.Data2, 4) & "-" & hexPad(GUID.Data3, 4) & "-"
    sGuid = sGuid & hexPad(GUID.Data4(0), 2) & hexPad(GUID.Data4(1), 2) & "-"
    For i = 2 To 7
        sGuid = sGuid & hexPad(GUID.Data4(i), 2)
    Next i
    sGuid = sGuid & "}"
    GetGuidString = sGuid
End Function

'gets identifies the format of this Image object.
Public Function GetRawFormat(ByVal alngImage As Long) As GpImageFormatIdentifiers
    Dim FormatID As Clsid
    Dim enmStatus As GpStatus
    
    enmStatus = GdipGetImageRawFormat(alngImage, FormatID)
    Select Case GetGuidString(FormatID)
        Case ImageFormatUndefined
          GetRawFormat = GpImageFormatUndefined
        Case ImageFormatMemoryBMP
          GetRawFormat = GpImageFormatMemoryBMP
        Case ImageFormatBMP
          GetRawFormat = GpImageFormatBMP
        Case ImageFormatEMF
          GetRawFormat = GpImageFormatEMF
        Case ImageFormatWMF
          GetRawFormat = GpImageFormatWMF
        Case ImageFormatJPEG
          GetRawFormat = GpImageFormatJPEG
        Case ImageFormatPNG
          GetRawFormat = GpImageFormatPNG
        Case ImageFormatGIF
          GetRawFormat = GpImageFormatGIF
        Case ImageFormatTIFF
          GetRawFormat = GpImageFormatTIFF
        Case ImageFormatEXIF
          GetRawFormat = GpImageFormatEXIF
        Case ImageFormatIcon
          GetRawFormat = GpImageFormatIcon
    End Select
End Function

