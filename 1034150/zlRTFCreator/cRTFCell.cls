VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRTFCell"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

'��������ֵ�ľֲ�����
Private mvarAlignmentV As AlignVTypeEnum    '��Ԫ��ֱ���룬1��3��Ĭ�ϣ�0
Private mvarAlignment As AlignTypeEnum      '��Ԫ��ˮƽ���룬1��3��Ĭ�ϣ�0
Private mvarBorderLeft As Boolean       '��Ԫ����߿�
Private mvarBorderRight As Boolean      '��Ԫ���ұ߿�
Private mvarBorderTop As Boolean        '��Ԫ���ϱ߿�
Private mvarBorderBottom As Boolean     '��Ԫ���±߿�
Private mvarBorderLeftDown As Boolean   '������б��\
Private mvarBorderLeftUp As Boolean     '������б��/
Private mvarBorderWidth As Long         '�߿���ȣ�0��75��Ĭ�ϣ�15��
Private mvarBorderStyle As BorderStyleEnum         '�߿����͡�1��28�����͡�
Private mvarBorderColor As ColorEnum                   '��ɫ���͡�
Private mvarCellPatType As PatTypeEnum             '��Ԫ�񱳾�ͼ�����͡�1��12��Ĭ�ϣ�0
Private mvarCellPatBackColor As ColorEnum              '��Ԫ�񱳾�ͼ����������ɫ��
Private mvarCellPatForeColor As ColorEnum              '��Ԫ�񱳾�ͼ���ı�����ɫ��
Private mvarCellPatPercentage As Long                  '��Ԫ����������ٷֱȡ�
Private mvarWidth As Long                              '��Ԫ����ѡ����
Private mvarWidthUnit As CellLengthUnitEnum            '��Ԫ����ѡ���ȵ�λ��0��3����0���գ�1���Զ���2���ٷֱȣ�3��羣�Ĭ�ϣ���
Private mvarRight As Long                              '��Ԫ���ұ߽�λ��
Private mvarMergeVStart As Boolean      '��ֱ�ϲ��ĵ�һ����Ԫ��
Private mvarMergeV As Boolean           '��ֱ��ǰһ��Ԫ��ϲ�
Private mvarMergeStart As Boolean       'ˮƽ�ϲ��ĵ�һ����Ԫ��
Private mvarMerge As Boolean            'ˮƽ��ǰһ��Ԫ��ϲ�
Private mvarRow As Long                 '����
Private mvarCol As Long                 '����
Private mvarFitText As Boolean          '�ı���Ӧ��Ԫ�����
Private mvarNoWrap As Boolean           '��Ԫ�������ı�����
Private mvarNested As Boolean           '�Ƿ�Ƕ�ף�Ĭ��Ϊ��False
Private mvarData As Variant                             '��Ԫ�����ݣ��ı���ͼƬ��Ƕ�ױ���
Private mvarDataType As CellContentTypeEnum             '�������͡�0���ı���1��ͼƬ��2��Ƕ�ױ���Ĭ�ϣ�0
Private mvarText As String
Private mvarBold As Boolean
Private mvarItalic As Boolean
Private mvarForeColor As ColorEnum
Private mvarFontSizeFixup As FontSizeEnum
Private mvarFontSize As Long

Public Property Let FontSize(ByVal vData As Long)
Attribute FontSize.VB_Description = "���ı��������С��Ϊ����ֵ��"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.FontSize = 5
    mvarFontSize = vData
End Property


Public Property Get FontSize() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.FontSize
    FontSize = mvarFontSize
End Property




Public Property Let FontSizeFixup(ByVal vData As FontSizeEnum)
Attribute FontSizeFixup.VB_Description = "���ı��������С����ֵ��"
'������ָ�ɶ���ʱʹ�ã�λ�� Set ������ߡ�
'Syntax: Set x.FontSizeFixup = Form1
    mvarFontSizeFixup = vData
End Property

Public Property Get FontSizeFixup() As FontSizeEnum
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.FontSizeFixup
    FontSizeFixup = mvarFontSizeFixup
End Property

Public Property Let ForeColor(ByVal vData As ColorEnum)
Attribute ForeColor.VB_Description = "���ı���ǰ��ɫ��"
'������ָ�ɶ���ʱʹ�ã�λ�� Set ������ߡ�
'Syntax: Set x.ForeColor = Form1
    mvarForeColor = vData
End Property

Public Property Get ForeColor() As ColorEnum
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ForeColor
    ForeColor = mvarForeColor
End Property



Public Property Let Italic(ByVal vData As Boolean)
Attribute Italic.VB_Description = "���ı���б�塣"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Italic = 5
    mvarItalic = vData
End Property


Public Property Get Italic() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Italic
    Italic = mvarItalic
End Property



Public Property Let Bold(ByVal vData As Boolean)
Attribute Bold.VB_Description = "���ı��Ĵ������"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Bold = 5
    mvarBold = vData
End Property


Public Property Get Bold() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Bold
    Bold = mvarBold
End Property




Public Property Let Text(ByVal vData As String)
Attribute Text.VB_Description = "���ı���"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Text = 5
    mvarText = vData
End Property

Public Property Get Text() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Text
    Text = mvarText
End Property

Public Property Let CellPatPercentage(ByVal vData As Long)
Attribute CellPatPercentage.VB_Description = "��Ԫ����������ٷֱȡ�"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.CellPatPercentage = 5
    mvarCellPatPercentage = vData
End Property

Public Property Get CellPatPercentage() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.CellPatPercentage
    CellPatPercentage = mvarCellPatPercentage
End Property

Public Property Let BorderWidth(ByVal vData As Long)
Attribute BorderWidth.VB_Description = "�߿�Ļ��߿��ȡ�0��75��Ĭ�ϣ�15"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.BorderWidth = 5
    mvarBorderWidth = vData
End Property

Public Property Get BorderWidth() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.BorderWidth
    BorderWidth = mvarBorderWidth
End Property

Public Function GetCellDefineRTF() As String
Attribute GetCellDefineRTF.VB_Description = "��ȡ��Ԫ�����RTF�ַ�����"
'��ȡ��Ԫ����RTF
    Dim strR As String
    
    '�ϲ���Ԫ����
    If mvarMergeVStart Then
        strR = strR & "\clvmgf"
    End If
    If mvarMergeStart Then
        strR = strR & "\clmgf"
    End If
    If mvarMergeV Then
        strR = strR & "\clvmrg"
    End If
    If mvarMerge Then
        strR = strR & "\clmrg"
    End If
    
    '��Ԫ���ڴ�ֱ���뷽ʽ��
    If mvarAlignmentV = alv2_���ж��� Then
        strR = strR & "\clvertalc" & vbCrLf
    ElseIf mvarAlignmentV = alv3_�׶˶��� Then
        strR = strR & "\clvertalb" & vbCrLf
    Else
        strR = strR & "\clvertalt" & vbCrLf
    End If
    Dim strBorderStyle As String, strBorderWidth As String, strBorderColor As String
    Select Case mvarBorderStyle
    Case cbs01_�ޱ߿�
        strBorderStyle = "\brdrnone"
    Case cbs02_������ȱ߿�
        strBorderStyle = "\brdrs"
    Case cbs03_˫����ȱ߿�
        strBorderStyle = "\brdrth"
    Case cbs04_��Ӱ�߿�
        strBorderStyle = "\brdrsh"
    Case cbs05_˫�߿�
        strBorderStyle = "\brdrdb"
    Case cbs06_���߱߿�
        strBorderStyle = "\brdrdot"
    Case cbs07_���߱߿�
        strBorderStyle = "\brdrdash"
    Case cbs08_ϸ�߱߿�
        strBorderStyle = "\brdrhair"
    Case cbs09_����ʽ�߿�
        strBorderStyle = "\brdrinset"
    Case cbs10_����ϸ�߿�
        strBorderStyle = "\brdrdashsm"
    Case cbs11_�㻮�߱߿�
        strBorderStyle = "\brdrdashd"
    Case cbs12_��㻮�߱߿�
        strBorderStyle = "\brdrdashdd"
    Case cbs13_��ʼ�߿�
        strBorderStyle = "\brdroutset"
    Case cbs14_���ر߿�
        strBorderStyle = "\brdrtriple"
    Case cbs15_�񱡵���ϸ�߿�
        strBorderStyle = "\brdrtnthsg"
    Case cbs16_�������ϸ�߿�
        strBorderStyle = "\brdrthtnsg"
    Case cbs17_���񱡵���ϸ�߿�
        strBorderStyle = "\brdrtnthtnsg"
    Case cbs18_�񱡵����б߿�
        strBorderStyle = "\brdrtnthmg"
    Case cbs19_��������б߿�
        strBorderStyle = "\brdrthtnmg"
    Case cbs20_���񱡵����б߿�
        strBorderStyle = "\brdrtnthtnmg"
    Case cbs21_�񱡵��Ӵֱ߿�
        strBorderStyle = "\brdrtnthlg"
    Case cbs22_������Ӵֱ߿�
        strBorderStyle = "\brdrthtnlg"
    Case cbs23_���񱡵��Ӵֱ߿�
        strBorderStyle = "\brdrtnthtnlg"
    Case cbs24_�����߱߿�
        strBorderStyle = "\brdrwavy"
    Case cbs25_˫�����߱߿�
        strBorderStyle = "\brdrwavydb"
    Case cbs26_���Ʊ߿�
        strBorderStyle = "\brdrdashdotstr"
    Case cbs27_����߿�
        strBorderStyle = "\brdremboss"
    Case cbs28_��̱߿�
        strBorderStyle = "\brdrengrave"
    Case Else   'Ĭ�ϵ��߿�
        strBorderStyle = "\brdrs"
    End Select
    
        
    '�߿����
    If mvarBorderWidth <> 0 Then
        strBorderWidth = "\brdrw" & mvarBorderWidth
    Else    'Ĭ�Ͽ���15
        strBorderWidth = "\brdrw15"
    End If
    
    If mvarBorderColor <> clr00_Ĭ�� Then
        strBorderColor = "\brdrcf" & mvarBorderColor
    End If
    
    '���ɵ�Ԫ��߿������
    If mvarBorderTop Then
        strR = strR & "\clbrdrt" & strBorderStyle & strBorderWidth & strBorderColor & vbCrLf
    Else
        strR = strR & "\clbrdrt\brdrnone" & vbCrLf
    End If
    If mvarBorderLeft Then
        strR = strR & "\clbrdrl" & strBorderStyle & strBorderWidth & strBorderColor & vbCrLf
    Else
        strR = strR & "\clbrdrl\brdrnone" & vbCrLf
    End If
    If mvarBorderBottom Then
        strR = strR & "\clbrdrb" & strBorderStyle & strBorderWidth & strBorderColor & vbCrLf
    Else
        strR = strR & "\clbrdrb\brdrnone" & vbCrLf
    End If
    If mvarBorderRight Then
        strR = strR & "\clbrdrr" & strBorderStyle & strBorderWidth & strBorderColor & vbCrLf
    Else
        strR = strR & "\clbrdrr\brdrnone" & vbCrLf
    End If
    If mvarBorderLeftDown Then
        strR = strR & "\cldglu" & strBorderStyle & strBorderWidth & strBorderColor & vbCrLf
    End If
    If mvarBorderLeftUp Then
        strR = strR & "\cldgll" & strBorderStyle & strBorderWidth & strBorderColor & vbCrLf
    End If
    
    Dim strCellPatType As String
    '����ͼ������
    Select Case mvarCellPatType
    Case ptt00_��
        strCellPatType = "\clbghoriz"
    Case ptt02_����
        strCellPatType = "\clbgvert"
    Case ptt03_��б��
        strCellPatType = "\clbgfdiag"
    Case ptt04_��б��
        strCellPatType = "\clbgbdiag"
    Case ptt05_������
        strCellPatType = "\clbgcross"
    Case ptt06_б������
        strCellPatType = "\clbgdcross"
    Case ptt07_�ֺ���
        strCellPatType = "\clbgdkhor"
    Case ptt08_������
        strCellPatType = "\clbgdkvert"
    Case ptt09_����б��
        strCellPatType = "\clbgdkfdiag"
    Case ptt10_����б��
        strCellPatType = "\clbgdkbdiag"
    Case ptt11_��������
        strCellPatType = "\clbgdkcross"
    Case ptt12_��б������
        strCellPatType = "\clbgdkdcross"
    Case Else
        strCellPatType = "\clshdrawnil"
    End Select
    
    '����ǰ��ɫ
    If mvarCellPatForeColor <> clr00_Ĭ�� Then strR = strR & "\clcfpat" & mvarCellPatForeColor
    '���Ʊ���ɫ
    If mvarCellPatBackColor <> clr00_Ĭ�� Then strR = strR & "\clcbpat" & mvarCellPatBackColor
    '��������
    If mvarCellPatType <> ptt00_�� Then
        strR = strR & strCellPatType
    End If
    '������ɫ�ٷֱ�
    If mvarCellPatPercentage <> 0 Then
        strR = strR & "\clshdng" & mvarCellPatPercentage
    End If
    
    '�ı�����Ĭ�ϣ�
    strR = strR & "\cltxlrtb"
    
    '��Ԫ�����
    strR = strR & "\clftsWidth" & mvarWidthUnit
    strR = strR & "\clwWidth" & mvarWidth
    
    '���Ʊ���ɫ����ʽ��
    If mvarCellPatBackColor <> clr00_Ĭ�� Then strR = strR & "\clcbpatraw" & mvarCellPatBackColor
    '����ǰ��ɫ����ʽ��
    If mvarCellPatForeColor <> clr00_Ĭ�� Then strR = strR & "\clcfpatraw" & mvarCellPatForeColor
    '�������ͣ���ʽ�����ظ�һ��
    If mvarCellPatType <> ptt00_�� Then
        strR = strR & strCellPatType
    End If
    '������ɫ�ٷֱȣ���ʽ��
    If mvarCellPatPercentage <> 0 Then
        strR = strR & "\clshdngraw" & mvarCellPatPercentage
    End If
    If mvarRight = 0 Then
        mvarRight = Col * 4000
    End If
    strR = strR & "\cellx" & mvarRight & vbCrLf
    
    GetCellDefineRTF = strR
End Function

Public Function GetCellContentRTF() As String
Attribute GetCellContentRTF.VB_Description = "��ȡ��Ԫ�����ݵ�RTF�ַ�����"
'��ȡ��Ԫ������RTF
    Dim strR As String, strIsInTable As String
    If mvarNested Then strIsInTable = "\intbl"
    
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
    
    strR = "\pard\plain " & strAlignment & "\li0\ri0\nowidctlpar" & strIsInTable & "\aspalpha\aspnum\faauto\adjustright\rin0\lin0\itap0 \fs18\lang1033\langfe2052\kerning2\loch\af0\hich\af0\dbch\af13\cgrid\langnp1033\langfenp2052 "
    Select Case mvarDataType
    Case cct1_����
        strR = strR & mvarData.GetTextRTF()
    Case cct2_ͼƬ
        strR = strR & mvarData.GetTextRTF()
    Case cct3_����
        strR = strR & mvarData.GetTextRTF()
    Case Else   '���ı�
        '���ļ��Ĵ���:��Ԫ���е�ͼƬ���Դ�����ı������ı�ֻ�ṩ����������
        Dim strFontSize As String
        If mvarFontSize <> 0 Then   '�����С����Ȩ��FontSize > FontSizeFixup > FontSizeDefault
            strFontSize = "\fs" & mvarFontSize * 2
        Else
            If mvarFontSizeFixup <> fs00_Ĭ�� Then
                strFontSize = "\fs" & mvarFontSizeFixup * 2
            Else
                strFontSize = "\fs18"   '��Ԫ����Ĭ�ϴ�СΪ 9
            End If
        End If
        strR = strR & strFontSize
        If mvarBold Then
            strR = strR & "\b"
        Else
            strR = strR & "\b0"
        End If
        If mvarItalic Then
            strR = strR & "\i"
        Else
            strR = strR & "\i0"
        End If
        If mvarForeColor <> clr00_Ĭ�� Then
            strR = strR & "\cf" & mvarForeColor
        End If
        strR = strR & StrToASC(mvarText)
    End Select
    
    
    If Nested Then
        strR = strR & "\nestcell"
    Else
        strR = strR & "\cell"
    End If
    
    GetCellContentRTF = strR
End Function

Public Property Let Nested(ByVal vData As Boolean)
Attribute Nested.VB_Description = "�Ƿ�Ƕ�ס�Ĭ��:False"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Nested = 5
    mvarNested = vData
End Property


Public Property Get Nested() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Nested
    Nested = mvarNested
End Property


Public Property Let NoWrap(ByVal vData As Boolean)
Attribute NoWrap.VB_Description = "��Ԫ�������ı����С�"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.NoWrap = 5
    mvarNoWrap = vData
End Property


Public Property Get NoWrap() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.NoWrap
    NoWrap = mvarNoWrap
End Property


Public Property Let FitText(ByVal vData As Boolean)
Attribute FitText.VB_Description = "�ı��Զ���Ӧ��Ԫ����ȡ�"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.FitText = 5
    mvarFitText = vData
End Property


Public Property Get FitText() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.FitText
    FitText = mvarFitText
End Property



Public Property Let DataType(ByVal vData As CellContentTypeEnum)
Attribute DataType.VB_Description = "��Ԫ���������͡�0-���ı���1-�߼��ı���2-ͼƬ��3-Ƕ�ױ���"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.DataType = 5
    mvarDataType = vData
End Property


Public Property Get DataType() As CellContentTypeEnum
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.DataType
    DataType = mvarDataType
End Property



Public Property Let Data(ByVal vData As Variant)
Attribute Data.VB_Description = "��Ԫ���������ݡ���DataType�������DataTypeȡֵ�����ǣ�0-���ı���1-�߼��ı���2-ͼƬ��3-Ƕ�ױ���"
Attribute Data.VB_UserMemId = 0
'������ָ�ɶ���ʱʹ�ã�λ�� Set ������ߡ�
'Syntax: Set x.Data = Form1
    Select Case mvarDataType
    Case cct1_����
        Set mvarData = New cRTFText
        Set mvarData = vData
        mvarData.IsInTable = True
    Case cct2_ͼƬ
        Set mvarData = New cRTFPicture
        Set mvarData = vData
        mvarData.IsInTable = True
    Case cct3_����
        Set mvarData = New cRTFTable
        Set mvarData = vData
        mvarData.Nested = True
    Case Else   '����Ϊ���ı�
        mvarData = vData
        mvarText = vData
    End Select
End Property


Public Property Set Data(ByVal vData As Variant)
'������ָ�ɶ���ʱʹ�ã�λ�� Set ������ߡ�
'Syntax: Set x.Data = Form1
    Select Case mvarDataType
    Case cct1_����
        Set mvarData = New cRTFText
        Set mvarData = vData
        mvarData.IsInTable = True
    Case cct2_ͼƬ
        Set mvarData = New cRTFPicture
        Set mvarData = vData
        mvarData.IsInTable = True
    Case cct3_����
        Set mvarData = New cRTFTable
        Set mvarData = vData
        mvarData.Nested = True
    Case Else   '����Ϊ���ı�
        mvarData = vData
        mvarText = vData
    End Select
End Property


Public Property Get Data() As Variant
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Data
    If IsObject(mvarData) Then
        Set Data = mvarData
    Else
        Data = mvarData
    End If
End Property



Public Property Let Col(ByVal vData As Long)
Attribute Col.VB_Description = "��ǰ����ֵ��"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Col = 5
    mvarCol = vData
End Property


Public Property Get Col() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Col
    Col = mvarCol
End Property



Public Property Let Row(ByVal vData As Long)
Attribute Row.VB_Description = "������"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Row = 5
    mvarRow = vData
End Property


Public Property Get Row() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Row
    Row = mvarRow
End Property



Public Property Let Merge(ByVal vData As Boolean)
Attribute Merge.VB_Description = "ˮƽ��ǰһ��Ԫ��ϲ���"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Merge = 5
    mvarMerge = vData
End Property


Public Property Get Merge() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Merge
    Merge = mvarMerge
End Property



Public Property Let MergeStart(ByVal vData As Boolean)
Attribute MergeStart.VB_Description = "ˮƽ�ϲ��ĵ�һ����Ԫ��"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.MergeStart = 5
    mvarMergeStart = vData
End Property


Public Property Get MergeStart() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.MergeStart
    MergeStart = mvarMergeStart
End Property



Public Property Let MergeV(ByVal vData As Boolean)
Attribute MergeV.VB_Description = "��ֱ��ǰһ��Ԫ��ϲ���"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.MergeV = 5
    mvarMergeV = vData
End Property


Public Property Get MergeV() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.MergeV
    MergeV = mvarMergeV
End Property



Public Property Let MergeVStart(ByVal vData As Boolean)
Attribute MergeVStart.VB_Description = "��ֱ�ϲ��ĵ�һ����Ԫ��"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.MergeVStart = 5
    mvarMergeVStart = vData
End Property


Public Property Get MergeVStart() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.MergeVStart
    MergeVStart = mvarMergeVStart
End Property



Public Property Let Right(ByVal vData As Long)
Attribute Right.VB_Description = "��Ԫ���ұ߽�λ�á�CellX"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Right = 5
    mvarRight = vData
End Property


Public Property Get Right() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Right
    Right = mvarRight
End Property



Public Property Let WidthUnit(ByVal vData As CellLengthUnitEnum)
Attribute WidthUnit.VB_Description = "��Ԫ����ѡ���ȵ�λ��0��3��Ĭ�ϣ�3"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.WidthUnit = 5
    mvarWidthUnit = vData
End Property


Public Property Get WidthUnit() As CellLengthUnitEnum
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.WidthUnit
    WidthUnit = mvarWidthUnit
End Property



Public Property Let Width(ByVal vData As Long)
Attribute Width.VB_Description = "��Ԫ����ѡ���ȡ�"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Width = 5
    mvarWidth = vData
End Property


Public Property Get Width() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Width
    Width = mvarWidth
End Property



Public Property Let CellPatForeColor(ByVal vData As ColorEnum)
Attribute CellPatForeColor.VB_Description = "��Ԫ�����ǰ��ɫ����ֵ��"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.CellPatForeColor = 5
    mvarCellPatForeColor = vData
End Property


Public Property Get CellPatForeColor() As ColorEnum
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.CellPatForeColor
    CellPatForeColor = mvarCellPatForeColor
End Property



Public Property Let CellPatBackColor(ByVal vData As ColorEnum)
Attribute CellPatBackColor.VB_Description = "��Ԫ�����������ɫ����ֵ��"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.CellPatBackColor = 5
    mvarCellPatBackColor = vData
End Property


Public Property Get CellPatBackColor() As ColorEnum
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.CellPatBackColor
    CellPatBackColor = mvarCellPatBackColor
End Property



Public Property Let CellPatType(ByVal vData As PatTypeEnum)
Attribute CellPatType.VB_Description = "��Ԫ��������͡�1��12��Ĭ�ϣ�0"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.CellPatType = 5
    mvarCellPatType = vData
End Property


Public Property Get CellPatType() As PatTypeEnum
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.CellPatType
    CellPatType = mvarCellPatType
End Property



Public Property Let BorderColor(ByVal vData As ColorEnum)
Attribute BorderColor.VB_Description = "�߿���ɫ����ֵ��"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.BorderColor = 5
    mvarBorderColor = vData
End Property


Public Property Get BorderColor() As ColorEnum
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.BorderColor
    BorderColor = mvarBorderColor
End Property



Public Property Let BorderStyle(ByVal vData As BorderStyleEnum)
Attribute BorderStyle.VB_Description = "�߿����͡�1��28�����͡�Ĭ�ϣ�0"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.BorderStyle = 5
    mvarBorderStyle = vData
End Property


Public Property Get BorderStyle() As BorderStyleEnum
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.BorderStyle
    BorderStyle = mvarBorderStyle
End Property




Public Property Let BorderLeftUp(ByVal vData As Boolean)
Attribute BorderLeftUp.VB_Description = "�Ƿ���ʾ��Ԫ��������б��/��Ĭ��Ϊ:False"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.BorderLeftUp = 5
    mvarBorderLeftUp = vData
End Property


Public Property Get BorderLeftUp() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.BorderLeftUp
    BorderLeftUp = mvarBorderLeftUp
End Property



Public Property Let BorderLeftDown(ByVal vData As Boolean)
Attribute BorderLeftDown.VB_Description = "�Ƿ���ʾ��Ԫ��������б��\\��Ĭ��Ϊ:False"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.BorderLeftDown = 5
    mvarBorderLeftDown = vData
End Property


Public Property Get BorderLeftDown() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.BorderLeftDown
    BorderLeftDown = mvarBorderLeftDown
End Property



Public Property Let BorderBottom(ByVal vData As Boolean)
Attribute BorderBottom.VB_Description = "�Ƿ���ʾ��Ԫ���±߿�Ĭ�ϣ�True"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.BorderBottom = 5
    mvarBorderBottom = vData
End Property


Public Property Get BorderBottom() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.BorderBottom
    BorderBottom = mvarBorderBottom
End Property



Public Property Let BorderTop(ByVal vData As Boolean)
Attribute BorderTop.VB_Description = "�Ƿ���ʾ��Ԫ���ϱ߿�Ĭ�ϣ�True"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.BorderTop = 5
    mvarBorderTop = vData
End Property


Public Property Get BorderTop() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.BorderTop
    BorderTop = mvarBorderTop
End Property



Public Property Let BorderRight(ByVal vData As Boolean)
Attribute BorderRight.VB_Description = "�Ƿ���ʾ��Ԫ���ұ߿�Ĭ�ϣ�True"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.BorderRight = 5
    mvarBorderRight = vData
End Property


Public Property Get BorderRight() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.BorderRight
    BorderRight = mvarBorderRight
End Property



Public Property Let BorderLeft(ByVal vData As Boolean)
Attribute BorderLeft.VB_Description = "�Ƿ���ʾ��Ԫ����߿�Ĭ�ϣ�True"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.BorderLeft = 5
    mvarBorderLeft = vData
End Property


Public Property Get BorderLeft() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.BorderLeft
    BorderLeft = mvarBorderLeft
End Property

Public Property Let Alignment(ByVal vData As AlignTypeEnum)
Attribute Alignment.VB_Description = "��Ԫ��ˮƽ���롣"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Alignment = 5
    If mvarDataType = cct2_ͼƬ Then
        mvarData.Alignment = vData
    ElseIf mvarDataType = cct3_���� Then
        '������뷽ʽ��
        'mvarData.Alignment = vData
    Else    'Ĭ��Ϊ�ı���ʽ��
        If IsObject(mvarData) Then mvarData.Alignment = vData
    End If
    mvarAlignment = vData
End Property

Public Property Get Alignment() As AlignTypeEnum
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Alignment
    Alignment = mvarAlignment
End Property


Public Property Let AlignmentV(ByVal vData As AlignVTypeEnum)
Attribute AlignmentV.VB_Description = "��Ԫ��ֱ���롣"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.AlignmentV = 5
    mvarAlignmentV = vData
End Property


Public Property Get AlignmentV() As AlignVTypeEnum
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.AlignmentV
    AlignmentV = mvarAlignmentV
End Property
