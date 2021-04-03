VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDIB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'######################################################################################################################
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Var() As Any) As Long

Private Type SAFEARRAY
    cDims As Integer         '��������м�ά��
    fFeatures As Integer     '���������ʲô���ԣ�
    cbElements As Long       '�����ÿ��Ԫ���ж��
    cLocks As Long           '������鱻���������Σ�
    pvData As Long           '�������������ݷ���ʲô�ط���
    'rgsabound() As SFArrayBOUND
End Type

Private Type SAFEARRAYBOUND
    cElements As Long      '��һά�ж��ٸ�Ԫ�أ�
    lLbound As Long        '���������Ӽ���ʼ��
End Type

Private Type SAFEARRAY1
    cDims As Integer         '��������м�ά��
    fFeatures As Integer     '���������ʲô���ԣ�
    cbElements As Long       '�����ÿ��Ԫ���ж��
    cLocks As Long           '������鱻���������Σ�
    pvData As Long           '�������������ݷ���ʲô�ط���
    CE0 As Long              '��һά�ж��ٸ�Ԫ�أ�
    LB0 As Long              '���������Ӽ���ʼ��
End Type

Private Type SAFEARRAY2
    cDims As Integer         '��������м�ά��
    fFeatures As Integer     '���������ʲô���ԣ�
    cbElements As Long       '�����ÿ��Ԫ���ж��
    cLocks As Long           '������鱻���������Σ�
    pvData As Long           '�������������ݷ���ʲô�ط���
    CE0 As Long              '��һά�ж��ٸ�Ԫ�أ�
    LB0 As Long              '���������Ӽ���ʼ��
    CE1 As Long
    LB1 As Long
End Type

Private Const pvDataPtrAdd As Long = 12&


'-- �ؼ����� --------------------------------------
Private InitPtrFlag As Boolean

Private pLongAll(0 To 0) As Long
Private pLongAllPtr(0 To 0) As Long
Private OldpLongAll As Long
Private OldpLongAllPtr As Long

'-- 2��3Byteָ�� ----------------------------------
Private p3Byte0(0 To 2) As Byte
Private p3Byte0Ptr(0 To 0) As Long
Private Oldp3Byte0 As Long
Private Oldp3Byte0Ptr As Long

Private p3Byte1(0 To 2) As Byte
Private p3Byte1Ptr(0 To 0) As Long
Private Oldp3Byte1 As Long
Private Oldp3Byte1Ptr As Long

'-- 2��Longָ�� -----------------------------------
Private pLong0(0 To 0) As Long
Private pLong0Ptr(0 To 0) As Long
Private OldpLong0 As Long
Private OldpLong0Ptr As Long

Private pLong1(0 To 0) As Long
Private pLong1Ptr(0 To 0) As Long
Private OldpLong1 As Long
Private OldpLong1Ptr As Long

'######################################################################################################################
'## �豸�������� ###############################
'== Fun ========================================
'CreateCompatibleDC ����һ�����ض��豸����һ�µ��ڴ��豸����������ֵ��Long�����豸�����������������Ϊ��
'DeleteDC           ɾ��ר���豸��������Ϣ�������ͷ�������ش�����Դ����Ҫ��������GetDC����ȡ�ص��豸����������ֵ��Long��ִ�гɹ���Ϊ���㣬ʧ����Ϊ��
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long

'== Const ======================================
'== Type =======================================


'## ��դ���㺯�� ###############################
'== Fun ========================================
'BitBlt            ��һ��λͼ��һ���豸�������Ƶ���һ����Դ��Ŀ��DC�໥�������ݡ�����ֵ��Long�������ʾ�ɹ������ʾʧ�ܡ�������GetLastError
'SetStretchBltMode ָ��StretchBlt �� StretchDIBits����������ģʽ����������ģʽ������Windows��ζ������������޳���ɨ���߽��п��ơ�����VB����Ϳؼ���������API��ͼ������ʹ���������������ָ�ԭ����StretchBltģʽ
'StretchBlt        ��һ��λͼ��һ���豸�������Ƶ���һ����Դ��Ŀ��DC�໥�������ݡ�������������豸�����ж���һ��Ŀ����Σ�����λͼ�ж���һ��Դͼ��Դ���λ������Ҫ�����������Ա���Ŀ����εĴ�С���������ֵ��Long�������ʾ�ɹ������ʾʧ�ܡ�������GetLastError
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

'== Const ======================================
'== Type =======================================
Private Const STRETCH_ANDSCANS = 1    'Ĭ�����á��޳����߶���ʣ�µ��߶ν���AND���㡣���ģʽͨ��Ӧ���ڲ����˰�ɫ�����ĵ�ɫλͼ
Private Const STRETCH_ORSCANS = 2     '�޳����߶α��򵥵���������ģʽͨ�����ڲ�ɫλͼ
Private Const STRETCH_DELETESCANS = 3 '�޳����߶���ʣ�µ��߶ν���OR���㡣���ģʽͨ��Ӧ���ڲ����˰�ɫ�����ĵ�ɫλͼ
Private Const STRETCH_HALFTONE = 4    'Ŀ��λͼ�ϵ����ؿ鱻��ΪԴλͼ�ϴ��½��ƵĿ顣���ģʽҪ������������ģʽ


'## ��ͼ���� ###################################
'== Fun ========================================
'DeleteObject ���������ɾ��GDI���󣬱��续�ʡ�ˢ�ӡ����塢λͼ�������Լ���ɫ��ȵȡ�����ʹ�õ�����ϵͳ��Դ���ᱻ�ͷš�����ֵ��Long�������ʾ�ɹ������ʾʧ��
'GetObject    ȡ�ö�ָ���������˵����һ���ṹ��windows�ֲὨ����GetObject������������øú�����GetObjectAPI��vb�����ڱ�����GetObject�ؼ��ֻ���������ֵ��Long����lpObject��Ϊ�㣨��ByVal As Long����������´���һ��NULL����������������û������ĳ��ȡ���ִ�гɹ�����������ṹ�ڲ���ʵ���ֽ�������ʧ�ܣ�������ֵ
'SelectObject ÿ���豸������������ѡ�����е�ͼ�ζ������а���λͼ��ˢ�ӡ����塢�����Լ�����ȵȡ�һ��ѡ���豸������ֻ����һ������ѡ���Ķ�������豸�����Ļ�ͼ������ʹ�á����磬��ǰѡ���Ļ��ʾ��������豸�����������߶���ɫ����ʽ������ֵ��Long������ǰѡ���豸��������ͬhObject���͵�һ������ľ�������ʾ��������ѡ���Ķ�����һ������Region��������������г���֮һ��SIMPLEREGION�� COMPLEXREGION �� NULLREGION ��������������� GDI_ERROR ��ʾ����
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long

'== Const ======================================
'== Type =======================================
Private Type BitMap
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type


'## Ӳ����ϵͳ���� #############################
'== Fun ========================================
'CopyMemory  �����ڴ�
'timeGetTime ȡ��ʱ�䡣��λΪ ����
'ZeroMemory  ����ڴ�
'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)


'######################################################################################################################

Private Const CW_USEDEFAULT As Long = &H80000000

'## λͼ���� ###################################
'== Fun ========================================
'CreateDIBSection ����һ��DIBSection������һ��GDI���󣬿���һ�����豸�й�λͼ����ʹ�á����ǣ������ڲ���Ϊһ�����豸�޹�λͼ���档����ֵ��Long��ִ�гɹ�����DIBSectionλͼ�ľ�������ʾʧ�ܡ�������GetLastError
'LoadImage        ����һ��λͼ��ͼ���ָ��
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As Long, pBitmapInfo As Any, ByVal un As Long, lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long

'== Const ======================================
Private Const DIB_RGB_COLORS = 0 'BITMAPINFO������һ����ɫ�������б�����32λ��ɫ��RGBQUAD��
Private Const DIB_PAL_COLORS = 1 'BITMAPINFO������һ��16λ��ɫ������������


Private Const IMAGE_BITMAP      As Long = 0 '
Private Const IMAGE_ICON        As Long = 1 '
Private Const IMAGE_CURSOR      As Long = 2 '
Private Const IMAGE_ENHMETAFILE As Long = 3 '

Private Const LR_DEFAULTCOLOR        As Long = &H0 '�Գ��淽ʽ����ͼ��
Private Const LR_MONOCHROME          As Long = &H1 '��ͼ��ת���ɵ�ɫ
Private Const LR_COLOR               As Long = &H2 '
Private Const LR_COPYRETURNORG       As Long = &H4 'Creates an exact copy of the image, ignoring the cxDesired and cyDesired parameters.
Private Const LR_COPYDELETEORG       As Long = &H8 'Deletes the original image after creating the copy.
Private Const LR_LOADFROMFILE       As Long = &H10 '��hInstΪ�㣬lpsz�ʹ���Ҫ�����ʵ����͵�һ���ļ������֣���������Win95
Private Const LR_LOADTRANSPARENT    As Long = &H20 '��ͼ���е�һ������������������ض���ϵͳ�滻
Private Const LR_DEFAULTSIZE        As Long = &H40 '����ͼ��������Ŵ���������n1��n2������
Private Const LR_VGACOLOR           As Long = &H80 'ʹ�����ɫ��Uses true VGA colors.
Private Const LR_LOADMAP3DCOLORS  As Long = &H1000 '��ͼ���е���ҡ��ҡ��Լ�ǳ�����ض��滻��COLOR_3DSHADOW��COLOR_3DFACE�Լ�COLOR_3DLIGHT�ĵ�ǰ����
Private Const LR_CREATEDIBSECTION As Long = &H2000 '���ָ����IMAGE_BITMAP���ͷ���DIBSection�ľ����������λͼ�ľ��
Private Const LR_COPYFROMRESOURCE As Long = &H4000 'Tries to reload an icon or cursor resource from the original resource file rather than simply copying the current image. This is useful for creating a different-sized copy when the resource file contains multiple sizes of the resource. Without this flag, CopyImage stretches the original image to the new size. If this flag is set, CopyImage uses the size in the resource file closest to the desired size.
Private Const LR_SHARED           As Long = &H8000 '��ͼ����Ϊһ��������Դ���롣��NT 4.0��װ�ع�����ԴʱҪ�õ��������

'## λͼ���� ###################################
'== Fun ========================================
Private Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As Any, ByVal wUsage As Long) As Long
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As Any, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Private Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Private Declare Function GetDIBColorTable Lib "gdi32" (ByVal hDC As Long, ByVal un1 As Long, ByVal un2 As Long, pRGBQuad As RGBQuad) As Long
Private Declare Function SetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Private Declare Function SetDIBColorTable Lib "gdi32" (ByVal hDC As Long, ByVal un1 As Long, ByVal un2 As Long, pcRGBQuad As RGBQuad) As Long

'## Other ######################################
Private Type BITMAPFILEHEADER
    bfType(1 To 2) As Byte
    bfSize As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits As Long
End Type

Private Type BitMapInfoHeader
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type
Private Const BI_RGB       As Long = 0&
Private Const BI_RLE8      As Long = 1&
Private Const BI_RLE4      As Long = 2&
Private Const BI_BitFields As Long = 3&
Private Const BI_JPEG      As Long = 4&
Private Const BI_PNG       As Long = 5&

Private Type RGBQuad
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type

Private Type BitMapInfo
    bmiHeader As BitMapInfoHeader
    bmiColors(0 To &HFF) As RGBQuad
End Type

'##################################################

Private Enum ScaleModeConst
    SMC_Min = 0
    SMC_Nearest = 0    '���ڽ���ֵ
    SMC_StretchBlt = 1 'StretchBlt
    SMC_BiliNear = 2   '˫�����ڲ�ֵ
    SMC_ThreeOrder = 3 '���ξ�����
    SMC_Max = 3
End Enum

'######################################################################################################################
Private MyhDC As Long
Private MyBMI As BitMapInfo
Private MyhDIB As Long
Private MyPtr As Long
Private m_WidthBytes As Long
Private hOldMap As Long

'######################################################################################################################
Private Sub Class_initialize()
    MyhDC = 0
End Sub

Private Sub Class_Terminate()
    Free
End Sub

Public Property Get DIBDC() As Long
    DIBDC = MyhDC
End Property

Public Property Get DataSize() As Long
    DataSize = MyBMI.bmiHeader.biSizeImage
End Property

Public Property Get Width() As Long
    Width = MyBMI.bmiHeader.biWidth
End Property

Public Property Get Height() As Long
    Height = MyBMI.bmiHeader.biHeight
End Property

Public Property Get WidthBytes() As Long
    WidthBytes = m_WidthBytes
End Property

Public Property Get ColorBit() As Long
    ColorBit = MyBMI.bmiHeader.biBitCount
End Property

Public Property Get DataPtr() As Long
    DataPtr = MyPtr
End Property

'����ͼƬ
Public Function Create(ByVal Width As Long, ByVal Height As Long, Optional ByVal CBit As Long = 32, Optional ClsMap As Boolean = True) As Boolean

    If Width <= 0 Or Height <= 0 Then
        Create = False
        Exit Function
    End If
    
    Select Case CBit
    Case 1, 4, 8, 16, 24, 32
    Case Else
        Create = False
        Exit Function
    End Select
    
    Free
    
    With MyBMI.bmiHeader
        .biSize = Len(MyBMI.bmiHeader)
        .biWidth = Width
        .biHeight = Height
        .biPlanes = 1
        .biBitCount = CBit
        .biCompression = BI_RGB
        m_WidthBytes = GetWidthBytes(.biWidth, .biBitCount)
        m_WidthBytes = (m_WidthBytes + 3) And &H7FFFFFFC
        .biSizeImage = m_WidthBytes * .biHeight
        
    End With
    MyhDC = CreateCompatibleDC(0)
    MyhDIB = CreateDIBSection(MyhDC, MyBMI, IIf(CBit > 8, DIB_RGB_COLORS, DIB_PAL_COLORS), MyPtr, 0, 0)
    If MyhDIB Then
        hOldMap = SelectObject(MyhDC, MyhDIB)
        If ClsMap Then Cls ': SetRectAlpha &HFF
        Create = True
    End If
    
End Function

'�ͷ�ͼƬ
Public Sub Free()
    If Me.DataPtr Then
        If hOldMap Then DeleteObject SelectObject(MyhDC, hOldMap)
        hOldMap = 0
        DeleteDC MyhDC
        MyhDC = 0
        MyBMI.bmiHeader.biBitCount = 0
        MyBMI.bmiHeader.biWidth = 0
        MyBMI.bmiHeader.biHeight = 0
        MyBMI.bmiHeader.biSizeImage = 0
        MyhDIB = 0
        MyPtr = 0
        
    End If
    
End Sub

'��ʼ��ͼƬ
Public Sub Cls()
    If Me.DataPtr Then ZeroMemory ByVal MyPtr, MyBMI.bmiHeader.biSizeImage
End Sub


Public Function PutTo(ByVal hDestDC As Long, _
        Optional ByVal X As Long = 0, Optional ByVal Y As Long = 0, _
        Optional ByVal nWidth As Long = CW_USEDEFAULT, _
        Optional ByVal nHeight As Long = CW_USEDEFAULT, _
        Optional ByVal xSrc As Long = 0, Optional ByVal ySrc As Long = 0) As Boolean
    If Me.DataPtr = 0 Then Exit Function
    If nWidth = CW_USEDEFAULT Then nWidth = Me.Width
    If nHeight = CW_USEDEFAULT Then nHeight = Me.Height
    
    PutTo = SetDIBitsToDevice(hDestDC, X, Y, Width, Height, _
            xSrc, ySrc, _
            0, Me.Height, _
            ByVal Me.DataPtr, MyBMI, _
            IIf(Me.ColorBit > 8, DIB_RGB_COLORS, DIB_PAL_COLORS))
    
End Function

Public Function PutToEx(ByVal hDestDC As Long, _
        Optional ByVal X As Long = 0, Optional ByVal Y As Long = 0, _
        Optional ByVal nWidth As Long = CW_USEDEFAULT, _
        Optional ByVal nHeight As Long = CW_USEDEFAULT, _
        Optional ByVal xSrc As Long = 0, Optional ByVal ySrc As Long = 0, _
        Optional ByVal nSrcWidth As Long = CW_USEDEFAULT, _
        Optional ByVal nSrcHeight As Long = CW_USEDEFAULT, _
        Optional ByVal dwRop As RasterOpConstants = vbSrcCopy) As Boolean
    If Me.DataPtr = 0 Then Exit Function
    If nWidth = CW_USEDEFAULT Then nWidth = Me.Width
    If nHeight = CW_USEDEFAULT Then nHeight = Me.Height
    If nSrcWidth = CW_USEDEFAULT Then nSrcWidth = Me.Width
    If nSrcHeight = CW_USEDEFAULT Then nSrcHeight = Me.Height
    
    PutToEx = StretchDIBits(hDestDC, X, Y, nWidth, nHeight, _
            xSrc, ySrc, nSrcWidth, nSrcHeight, _
            ByVal Me.DataPtr, MyBMI, _
            IIf(Me.ColorBit > 8, DIB_RGB_COLORS, DIB_PAL_COLORS), _
            dwRop)
    
End Function


Private Function ChkFileWrite(FileName As String) As Boolean
    Dim FileNum As Integer
    
    FileNum = FreeFile
    
    On Error Resume Next
    
    Open FileName For Output As #FileNum
    
    If Err.Number Then
        '
    Else
        Close #FileNum
        ChkFileWrite = True
    End If
    
    On Error GoTo 0
    
End Function

Public Function SaveBMP(FileName As String) As Boolean
    Dim FileNum As Integer
    Dim TempBMFH As BITMAPFILEHEADER
    Dim TCB() As RGBQuad
    Dim TempBytes() As Byte
    
    If (ChkFileWrite(FileName) = False) Or (MyPtr = 0) Then Exit Function
    
    TempBMFH.bfType(1) = Asc("B")
    TempBMFH.bfType(2) = Asc("M")
    TempBMFH.bfOffBits = Len(TempBMFH) + Len(MyBMI.bmiHeader)
    If MyBMI.bmiHeader.biBitCount <= 8 Then
        TempBMFH.bfOffBits = TempBMFH.bfOffBits + 4 * 2 ^ MyBMI.bmiHeader.biBitCount
        
        ReDim TCB(1 To 2 ^ MyBMI.bmiHeader.biBitCount)
        GetDIBColorTable MyhDC, 0, 2 ^ MyBMI.bmiHeader.biBitCount, TCB(1)
        
    End If
    
    TempBMFH.bfSize = TempBMFH.bfOffBits + MyBMI.bmiHeader.biSizeImage
    
    ReDim TempBytes(1 To MyBMI.bmiHeader.biSizeImage)
    CopyMemory TempBytes(1), ByVal MyPtr, MyBMI.bmiHeader.biSizeImage
    
    FileNum = FreeFile
    
    Open FileName For Binary As #FileNum
    
    Put #FileNum, , TempBMFH
    Put #FileNum, , MyBMI.bmiHeader
    If MyBMI.bmiHeader.biBitCount <= 8 Then Put #FileNum, , TCB
    Put #FileNum, , TempBytes
    
    Close #FileNum
    
    SaveBMP = True
    
End Function


'######################################################################################################################
Public Function GetWidthBytes(ByVal Width As Long, ByVal ColorBits As Long) As Long
    If Width <= 0 Then
        GetWidthBytes = 0
        Exit Function
    End If
    
    Select Case ColorBits
    Case 1, 4, 8, 16, 24, 32
        Dim TempNum As Long
        
        If ColorBits < 8 Then
            TempNum = (Width * ColorBits + 7) \ 8
        Else
            TempNum = Width * (ColorBits \ 8)
        End If
        GetWidthBytes = (TempNum + 3) And &H7FFFFFFC
        
    Case Else
        GetWidthBytes = 0
    End Select
    
End Function

Public Function DIBLoadMap(FileName As String, Optional ByVal AutoSize As Boolean = True, Optional ByVal CBit As Long = 24) As Boolean
    Dim hMap As Long
    Dim hDC As Long
    Dim hOldMap As Long
    Dim TempPic As IPictureDisp
    Dim BM As BitMap
    
'    If dDIB Is Nothing Then Set dDIB = New clsDIB
    If DIBDC = 0 Then AutoSize = True
    
    If AutoSize Then
        If CBit = 16 Or CBit = 24 Or CBit = 32 Then
        Else
            Exit Function
        End If
        
    End If
        
    hMap = LoadImage(0, FileName, IMAGE_BITMAP, 0, 0, LR_DEFAULTSIZE Or LR_CREATEDIBSECTION)
    If hMap Then
        '
    Else
        On Error Resume Next
        Set TempPic = LoadPicture(FileName)
        On Error GoTo 0
        If TempPic Is Nothing Then
        Else
            If TempPic.Type = vbPicTypeBitmap Then
                hMap = TempPic.handle
            Else
                Set TempPic = Nothing
            End If
        End If
        
    End If
    
    If hMap Then
        hDC = CreateCompatibleDC(0)
        If hDC = 0 Then GoTo FreeMe
        
        GetObjectAPI hMap, Len(BM), BM
        If AutoSize Then
            If Create(BM.bmWidth, BM.bmHeight, CBit, False) Then
            Else
                GoTo FreeMe
                
            End If
            
        End If
        
        hOldMap = SelectObject(hDC, hMap)
        SetStretchBltMode DIBDC, STRETCH_HALFTONE
        StretchBlt DIBDC, 0, 0, Width, Height, hDC, 0, 0, BM.bmWidth, BM.bmHeight, vbSrcCopy
        Call SelectObject(hDC, hOldMap)
        
        DIBLoadMap = True
        
    End If
    
    
FreeMe:
    If hDC Then DeleteDC hDC
    If TempPic Is Nothing Then DeleteObject hMap
    Set TempPic = Nothing
    
End Function

Public Function ScaleModeName(ByVal ScaleMode As Byte) As String
    Dim s As String
    
    Select Case ScaleMode
    Case ScaleModeConst.SMC_Nearest
        s = "���ڽ���ֵ"
    Case ScaleModeConst.SMC_StretchBlt
        s = "StretchBlt"
    Case ScaleModeConst.SMC_BiliNear
        s = "˫�����ڲ�ֵ"
    Case ScaleModeConst.SMC_ThreeOrder
        s = "���ξ�����"
    End Select
    
    ScaleModeName = s
    
End Function



Private Function Sinxx(ByVal X As Single) As Single
    X = Abs(X)
    If X < 1 Then
        Sinxx = 1 - 2 * X * X + X * X * X
    ElseIf X < 2 Then
        Sinxx = 4 - 8 * X + 5 * X * X - X * X * X
    Else
        Sinxx = 0
    End If
    
End Function

Public Function DIBScale(ByVal sDIB As clsDIB, ByVal W As Long, ByVal H As Long) As Boolean
                
'    If sDIB Is Nothing Then Exit Function
    If sDIB.DataPtr = 0 Then Exit Function
    If sDIB.ColorBit <> 24 Then Exit Function
    
    If W < 2 Then Exit Function
    If H < 2 Then Exit Function
    
    'If dDIB Is Nothing Then Set dDIB = New clsDIB
    
    If Create(W, H, 24, False) = False Then Exit Function
    
    If W = sDIB.Width And H = sDIB.Height Then
        CopyMemory ByVal DataPtr, ByVal sDIB.DataPtr, sDIB.DataSize
        DIBScale = True
        Exit Function
    End If
    
    Dim pMapData() As Byte
    Dim pMapDataInfo As SAFEARRAY1
    
    pMapDataInfo.cDims = 1
    pMapDataInfo.fFeatures = 0
    pMapDataInfo.cbElements = 1
    pMapDataInfo.cLocks = 0
    pMapDataInfo.pvData = 0
    pMapDataInfo.CE0 = 12
    pMapDataInfo.LB0 = 0
    pLongAllPtr(0) = VarPtrArray(pMapData)
    pLongAll(0) = VarPtr(pMapDataInfo)
    
    
    Dim sMaxX As Long, sMaxY As Long
    Dim dMaxX As Long, dMaxY As Long
    
    sMaxX = sDIB.Width - 1
    sMaxY = sDIB.Height - 1
    dMaxX = Width - 1
    dMaxY = Height - 1
    
    Dim sPtr As Long, dPtr As Long
    Dim sLineBytes As Long, dLineBytes As Long
    
    sPtr = sDIB.DataPtr
    dPtr = DataPtr
    sLineBytes = sDIB.WidthBytes
    dLineBytes = WidthBytes
    
    Dim sPixelAdd As Long, sLineAdd As Long
    Dim dPixelAdd As Long, dLineAdd As Long
    
    sPixelAdd = 3
    sLineAdd = sLineBytes - sDIB.Width * sPixelAdd
    dPixelAdd = 3
    dLineAdd = dLineBytes - Width * dPixelAdd
    
    Dim X As Single, Y As Single
    Dim nX As Long, nY As Long
    Dim u As Single, v As Single
    Dim u_ As Single, v_ As Single
    
    Dim LinenX() As Long
    Dim Lineu() As Single
    
    Dim I As Long, J As Long
    Dim fI As Long, fJ As Long
    
    Dim LinePtr(0 To 3) As Long
    Dim CurPtr As Long, CurIdx As Long
    
    Dim OldX As Long, OldY As Long
    Dim TempLng As Long

    
    SetStretchBltMode DIBDC, STRETCH_HALFTONE
    StretchBlt DIBDC, 0, 0, Width, Height, sDIB.DIBDC, 0, 0, sDIB.Width, sDIB.Height, vbSrcCopy
        
    
    pLongAllPtr(0) = VarPtrArray(pMapData)
    pLongAll(0) = 0
    
    DIBScale = True
    
End Function

'######################################################################################################################
Public Sub MakePoint(ByVal DataArrPtr As Long, ByVal pDataArrPtr As Long, ByRef OldArrPtr As Long, ByRef OldpArrPtr As Long)
    Dim TempLng As Long
    Dim TempPtr As Long
    
    If InitPtrFlag Then
        Dim OldPtr As Long
        
        OldPtr = pLongAllPtr(0)
        pLongAllPtr(0) = DataArrPtr
        TempLng = pLongAll(0) + pvDataPtrAdd
        pLongAllPtr(0) = pDataArrPtr
        TempPtr = pLongAll(0) + pvDataPtrAdd
        pLongAllPtr(0) = TempPtr
        OldpArrPtr = pLongAll(0)
        pLongAll(0) = TempLng
        pLongAllPtr(0) = TempLng
        OldArrPtr = pLongAll(0)
        pLongAllPtr(0) = OldPtr
        
    Else
        CopyMemory TempLng, ByVal DataArrPtr, 4 '�õ�DataArrPtr��SAFEARRAY�ṹ�ĵ�ַ
        TempLng = TempLng + pvDataPtrAdd '���ָ��ƫ��12���ֽں����pvDataָ��
        CopyMemory TempPtr, ByVal pDataArrPtr, 4 '�õ�pDataArrPtr��SAFEARRAY�ṹ�ĵ�ַ
        TempPtr = TempPtr + pvDataPtrAdd '���ָ��ƫ��12���ֽں����pvDataָ��
        CopyMemory OldpArrPtr, ByVal TempPtr, 4  '����ɵ�ַ
        CopyMemory ByVal TempPtr, TempLng, 4  'ʹpDataArrPtrָ��DataArrPtr��SAFEARRAY�ṹ��pvDataָ��
        CopyMemory OldArrPtr, ByVal TempLng, 4 '����ɵ�ַ
        
    End If
    
End Sub

Public Sub FreePoint(ByVal DataArrPtr As Long, ByVal pDataArrPtr As Long, ByVal OldArrPtr As Long, ByVal OldpArrPtr As Long)
    Dim TempPtr As Long
    
    If InitPtrFlag Then
        pLongAllPtr(0) = DataArrPtr
        pLongAllPtr(0) = pLongAll(0) + pvDataPtrAdd
        pLongAll(0) = OldArrPtr
        pLongAllPtr(0) = pDataArrPtr
        pLongAllPtr(0) = pLongAll(0) + pvDataPtrAdd
        pLongAll(0) = OldpArrPtr
        
    Else
        CopyMemory TempPtr, ByVal DataArrPtr, 4 '�õ�DataArrPtr��SAFEARRAY�ṹ�ĵ�ַ
        CopyMemory ByVal (TempPtr + pvDataPtrAdd), OldArrPtr, 4 '�ָ��ɵ�ַ
        CopyMemory TempPtr, ByVal pDataArrPtr, 4 '�õ�pDataArrPtr��SAFEARRAY�ṹ�ĵ�ַ
        CopyMemory ByVal (TempPtr + pvDataPtrAdd), OldpArrPtr, 4 '�ָ��ɵ�ַ
        
    End If
    
End Sub

Public Sub PointInit()
    If InitPtrFlag Then Exit Sub
    
    MakePoint VarPtrArray(pLongAll), VarPtrArray(pLongAllPtr), OldpLongAll, OldpLongAllPtr
    InitPtrFlag = True
    
    MakePoint VarPtrArray(p3Byte0), VarPtrArray(p3Byte0Ptr), Oldp3Byte0, Oldp3Byte0Ptr
    MakePoint VarPtrArray(p3Byte1), VarPtrArray(p3Byte1Ptr), Oldp3Byte1, Oldp3Byte1Ptr
    
    MakePoint VarPtrArray(pLong0), VarPtrArray(pLong0Ptr), OldpLong0, OldpLong0Ptr
    MakePoint VarPtrArray(pLong1), VarPtrArray(pLong1Ptr), OldpLong1, OldpLong1Ptr
    
End Sub

Public Sub PointFree()
    If InitPtrFlag = False Then Exit Sub
    
    FreePoint VarPtrArray(p3Byte0), VarPtrArray(p3Byte0Ptr), Oldp3Byte0, Oldp3Byte0Ptr
    FreePoint VarPtrArray(p3Byte1), VarPtrArray(p3Byte1Ptr), Oldp3Byte1, Oldp3Byte1Ptr
    
    FreePoint VarPtrArray(pLong0), VarPtrArray(pLong0Ptr), OldpLong0, OldpLong0Ptr
    FreePoint VarPtrArray(pLong1), VarPtrArray(pLong1Ptr), OldpLong1, OldpLong1Ptr
    
    InitPtrFlag = False
    FreePoint VarPtrArray(pLongAll), VarPtrArray(pLongAllPtr), OldpLongAll, OldpLongAllPtr
    
End Sub








