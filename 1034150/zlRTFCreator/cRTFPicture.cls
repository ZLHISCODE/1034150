VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRTFPicture"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit
'��������ֵ�ľֲ�����
Private mvarFileName As String          'ʵ��ͼƬ16��������
Private mvarScaleX As Long              'ˮƽ���ű�����Ĭ�ϣ�100
Private mvarScaleY As Long              '��ֱ���ű�����Ĭ�ϣ�100
Private mvarCropLeft As Long            '��˼���ֵ��Ĭ�ϣ�0
Private mvarCropRight As Long           '�Ҷ˼���ֵ��Ĭ�ϣ�0
Private mvarCropTop As Long             '�϶˼���ֵ��Ĭ�ϣ�0
Private mvarCropBottom As Long          '�¶˼���ֵ��Ĭ�ϣ�0
Private mvarWidth As Long               'ͼƬ���ؿ���
Private mvarHeight As Variant           'ͼƬ���ظ߶�
Private mvarWidthOld As Long            'ͼƬԭʼ����
Private mvarHeightOld As Long           'ͼƬԭʼ�߶�
Private mvarPicType As PicTypeEnum      'Ŀǰֻ��Ϊjpegblip��Ĭ�ϣ�JPG
Private mvarPicID As Long               'ͼƬIDֵ������Ϊ��ֵ����\bliptag-10292
Private mvarSingleLine As Boolean       'ͼƬ�Ƿ񵥶�һ��
Private mvarAlignment As AlignTypeEnum
Private mvarIsInTable As Boolean

Public Property Let IsInTable(ByVal vData As Boolean)
Attribute IsInTable.VB_Description = "�Ƿ�����ڱ����С�Ĭ��:False"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.IsInTable = 5
    mvarIsInTable = vData
End Property


Public Property Get IsInTable() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.IsInTable
    IsInTable = mvarIsInTable
End Property


Public Property Let Alignment(ByVal vData As AlignTypeEnum)
Attribute Alignment.VB_Description = "ͼƬ���뷽ʽ��"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Alignment = 5
    mvarAlignment = vData
End Property


Public Property Get Alignment() As AlignTypeEnum
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Alignment
    Alignment = mvarAlignment
End Property


Public Property Let SingleLine(ByVal vData As Boolean)
Attribute SingleLine.VB_Description = "ͼƬ�Ƿ��ǵ���һ�С�Ĭ�ϣ�True��"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.SingleLine = 5
    mvarSingleLine = vData
End Property


Public Property Get SingleLine() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.SingleLine
    SingleLine = mvarSingleLine
End Property




Public Property Let PicID(ByVal vData As Long)
Attribute PicID.VB_Description = "ͼƬIDֵ������Ϊ������Ŀǰû��ʹ�á�"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.PicID = 5
    mvarPicID = -Abs(vData)
End Property


Public Property Get PicID() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.PicID
    PicID = -Abs(mvarPicID)
End Property



Public Property Let PicType(ByVal vData As PicTypeEnum)
Attribute PicType.VB_Description = "ͼƬ���͡�"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.PicType = 5
    mvarPicType = vData
End Property


Public Property Get PicType() As PicTypeEnum
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.PicType
    PicType = mvarPicType
End Property



Public Property Let HeightOld(ByVal vData As Long)
Attribute HeightOld.VB_Description = "ͼƬԭʼ�߶ȡ�"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.HeightOld = 5
    mvarHeightOld = vData
End Property


Public Property Get HeightOld() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.HeightOld
    HeightOld = mvarHeightOld
End Property



Public Property Let WidthOld(ByVal vData As Long)
Attribute WidthOld.VB_Description = "ͼƬԭʼ���ȡ�"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.WidthOld = 5
    mvarWidthOld = vData
End Property


Public Property Get WidthOld() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.WidthOld
    WidthOld = mvarWidthOld
End Property



Public Property Let Height(ByVal vData As Variant)
Attribute Height.VB_Description = "ͼƬ���ظ߶ȡ�"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.mHeight = 5
    mvarHeight = vData
End Property


Public Property Set Height(ByVal vData As Variant)
'������ָ�ɶ���ʱʹ�ã�λ�� Set ������ߡ�
'Syntax: Set x.mHeight = Form1
    Set mvarHeight = vData
End Property


Public Property Get Height() As Variant
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.mHeight
    If IsObject(mvarHeight) Then
        Set Height = mvarHeight
    Else
        Height = mvarHeight
    End If
End Property



Public Property Let Width(ByVal vData As Long)
Attribute Width.VB_Description = "ͼƬ���ؿ��ȡ�"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.mWidth = 5
    mvarWidth = vData
End Property


Public Property Get Width() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.mWidth
    Width = mvarWidth
End Property



Public Property Let CropBottom(ByVal vData As Long)
Attribute CropBottom.VB_Description = "�¶˼���ֵ��Ĭ�ϣ�0"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.CropBottom = 5
    mvarCropBottom = vData
End Property


Public Property Get CropBottom() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.CropBottom
    CropBottom = mvarCropBottom
End Property



Public Property Let CropTop(ByVal vData As Long)
Attribute CropTop.VB_Description = "�϶˼���ֵ��Ĭ�ϣ�0"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.CropTop = 5
    mvarCropTop = vData
End Property


Public Property Get CropTop() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.CropTop
    CropTop = mvarCropTop
End Property



Public Property Let CropRight(ByVal vData As Long)
Attribute CropRight.VB_Description = "�Ҷ˼���ֵ��Ĭ�ϣ�0"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.CropRight = 5
    mvarCropRight = vData
End Property


Public Property Get CropRight() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.CropRight
    CropRight = mvarCropRight
End Property



Public Property Let CropLeft(ByVal vData As Long)
Attribute CropLeft.VB_Description = "��˼���ֵ��Ĭ�ϣ�0"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.CropLeft = 5
    mvarCropLeft = vData
End Property


Public Property Get CropLeft() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.CropLeft
    CropLeft = mvarCropLeft
End Property



Public Property Let ScaleY(ByVal vData As Long)
Attribute ScaleY.VB_Description = "��ֱ���ű�����Ĭ�ϣ�100"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ScaleY = 5
    mvarScaleY = vData
End Property


Public Property Get ScaleY() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ScaleY
    ScaleY = mvarScaleY
End Property



Public Property Let ScaleX(ByVal vData As Long)
Attribute ScaleX.VB_Description = "ˮƽ���ű�����Ĭ�ϣ�100"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ScaleX = 5
    mvarScaleX = vData
End Property


Public Property Get ScaleX() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ScaleX
    ScaleX = mvarScaleX
End Property



Public Property Let FileName(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Data = 5
    mvarFileName = vData
End Property


Public Property Get FileName() As String
Attribute FileName.VB_Description = "ͼƬԴ���ļ�����"
Attribute FileName.VB_UserMemId = 0
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Data
    FileName = mvarFileName
End Property


Public Function GetTextRTF() As String
Attribute GetTextRTF.VB_Description = "��ȡͼƬ��RTF�ַ�����"
'��ȡ����RTF�ı�
    Dim strR As String
    
    Dim strAlignment As String
    Select Case mvarAlignment
    Case alm2_���ж���
        strAlignment = "\qc"
    Case alm3_���Ҷ���
        strAlignment = "\qr"
    Case alm4_���˶���
        strAlignment = "\qj"
    Case alm5_��ɢ����
        strAlignment = "\qd"
    Case Else
        strAlignment = "\ql"
    End Select
    Dim strIsInTable As String
    If mvarIsInTable Then strIsInTable = "\intbl"
    
    strR = "\pard\plain " & strAlignment & " \li0\ri0\nowidctlpar" & strIsInTable & "\aspalpha\aspnum\faauto\adjustright\rin0\lin0\yts15 \fs18\lang1033\langfe2052\kerning2\loch\af0\hich\af0\dbch\af13\cgrid\langnp1033\langfenp2052 "
    strR = strR & "{\*\shppict{\pict"
    If mvarScaleX > 0 Then strR = strR & "\picscalex" & mvarScaleX
    If mvarScaleY > 0 Then strR = strR & "\picscaley" & mvarScaleY
    If mvarCropLeft > 0 Then strR = strR & "\piccrop" & mvarCropLeft
    If mvarCropRight > 0 Then strR = strR & "\piccropr" & mvarCropRight
    If mvarCropTop > 0 Then strR = strR & "\piccropt" & mvarCropTop
    If mvarCropBottom > 0 Then strR = strR & "\piccropb" & mvarCropBottom
    If mvarHeight > 0 Then strR = strR & "\pichgoal" & mvarHeight
    If mvarWidth > 0 Then strR = strR & "\picwgoal" & mvarWidth
    If mvarHeightOld > 0 Then strR = strR & "\pich" & mvarHeightOld
    If mvarWidthOld > 0 Then strR = strR & "\picw" & mvarWidthOld
    If mvarPicType = pic_JPG Then   'Ŀǰֻ����һ��JPEG����
        strR = strR & "\jpegblip "
    Else
        strR = strR & "\jpegblip "   'Ҳ��JPEG����������
    End If
    'If mvarPicID > 0 Then  mvarpicID��ʱû��ʹ�á�
    
    strR = strR & PicToASC(mvarFileName) & "}}"
    If SingleLine And Not mvarIsInTable Then strR = strR & "\par "

    GetTextRTF = strR
End Function
