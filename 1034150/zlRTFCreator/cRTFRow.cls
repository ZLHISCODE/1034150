VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRTFRow"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Private mvarRow As Long             '����
Private mvarRowBand As Long         '�������ο������С�������Ϊ-1
Private mvarIsLastRow As Boolean    '�Ƿ�����ĩһ��
Private mvarCellSpace As Long       '��Ԫ����
Private mvarWidth As Long                           'Ĭ���п�
Private mvarWidthUnit As CellLengthUnitEnum         'Ĭ���п���λ��1��3��Ĭ�ϣ�0
Private mvarWidthB As Long                          '��ǰ���ɼ���Ԫ�����
Private mvarWidthBUnit As CellLengthUnitEnum        '��ǰ���ɼ���Ԫ����ȵ�λ��1��3��Ĭ�ϣ�0
Private mvarWidthA As Long                          '��ĩ���ɼ���Ԫ�����
Private mvarWidthAUnit As CellLengthUnitEnum        '��ĩ���ɼ���Ԫ����ȵ�λ��1��3��Ĭ�ϣ�0
Private mvarPatType As PatTypeEnum                         '�������ͣ�1��12��Ĭ�ϣ�0
Private mvarPatForeColor As ColorEnum               '����ǰ��ɫ
Private mvarPatBackColor As ColorEnum               '���Ʊ���ɫ
Private mvarPatPercentage As Long   '���������ٷֱ�
Private mvarAutoFit As Boolean      '�������Ƿ����Զ���Ӧ��
Private mvarIsHeadRow As Boolean    '�Ƿ�����У�������ÿҳ������ʾ��trhdr
Private mvarKeep As Boolean         '���ָ�����ǰһ��ͬҳ��Ĭ�ϣ�False
Private mvarKeepFollow As Boolean   '���ָ������һ��ͬҳ��Ĭ�ϣ�False
Private mvarAlignment As RowAlignTypeEnum       '�ж��뷽ʽ���������еĵ�Ԫ��1��3��Ĭ��0
Private mvarHeight As Long          '�и�
Private mvarPadLeft As Long         'Ĭ�����е�Ԫ����߾�
Private mvarPadRight As Long        'Ĭ�����е�Ԫ���ұ߾�
Private mvarPadTop As Long          'Ĭ�����е�Ԫ���ϱ߾�
Private mvarPadBottom As Long       'Ĭ�����е�Ԫ���±߾�
Private mvarBorderLeft As Boolean   '����߿�Ĭ��:True
Private mvarBorderRight As Boolean  '���ұ߿�Ĭ��:True
Private mvarBorderTop As Boolean    '���ϱ߿�Ĭ��:True
Private mvarBorderBottom As Boolean '���±߿�Ĭ��:True
Private mvarBorderH As Boolean      '��ˮƽ�ڲ��߿�Ĭ��:True
Private mvarBorderV As Boolean      '�д�ֱ�ڲ��߿�Ĭ��:True
Private mvarBorderStyle As BorderStyleEnum  '�б߿�����
Private mvarBorderColor As ColorEnum            '�б߿���ɫ
Private mvarNested As Boolean       '�Ƿ�Ƕ�ף�Ĭ��:False
Private mvarCells As cRTFCells      '��Ԫ�񼯺�
Private mvarBorderWidth As Long
Private mvarColCount As Long

Public Function LastCell() As cRTFCell
Attribute LastCell.VB_Description = "��ȡ������ӵĵ�Ԫ�����"
    Set LastCell = mvarCells(mvarCells.Count)
End Function

Public Property Let ColCount(ByVal vData As Long)
Attribute ColCount.VB_Description = "��������"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ColCount = 5
    mvarColCount = vData
End Property

Public Property Get ColCount() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ColCount
    ColCount = mvarCells.Count
End Property

Public Function AddCell(Optional Row As Long, Optional Col As Long, _
Optional DataType As CellContentTypeEnum = cct0_�ı�, _
Optional Right As Long, _
Optional Width As Long = 0, _
Optional WidthUnit As CellLengthUnitEnum = clt0_Ĭ��, _
Optional sKey As String, _
Optional AlignmentV As AlignVTypeEnum = alv0_Ĭ�϶���, _
Optional Alignment As AlignTypeEnum = alm0_Ĭ�϶���, Optional BorderLeft As Boolean = True, _
Optional BorderRight As Boolean = True, Optional BorderTop As Boolean = True, _
Optional BorderBottom As Boolean = True, Optional BorderLeftDown As Boolean = False, _
Optional BorderLeftUp As Boolean = False, Optional BorderStyle As BorderStyleEnum = cbs00_Ĭ�ϱ߿���ʽ, _
Optional BorderColor As ColorEnum = clr00_Ĭ��, Optional CellPatType As PatTypeEnum = ptt00_��, _
Optional CellPatBackColor As ColorEnum = clr00_Ĭ��, Optional CellPatForeColor As ColorEnum = clr00_Ĭ��, _
Optional MergeVStart As Boolean = False, Optional MergeV As Boolean = False, _
Optional MergeStart As Boolean = False, Optional Merge As Boolean = False, _
Optional FitText As Boolean = True, Optional NoWrap As Boolean = False, _
Optional Nested As Boolean = False, Optional BorderWidth As Long = 0, _
Optional CellPatPercentage As Long = 100) As Long
Attribute AddCell.VB_Description = "����һ����Ԫ��"
    
    Dim i As Long

    i = mvarCells.Add(Row, Col, DataType, Right, Width, WidthUnit, sKey, AlignmentV, Alignment, BorderLeft, BorderRight, _
        BorderTop, BorderBottom, BorderLeftDown, BorderLeftUp, BorderStyle, BorderColor, CellPatType, _
        CellPatBackColor, CellPatForeColor, MergeVStart, MergeV, MergeStart, Merge, FitText, _
        NoWrap, Nested, BorderWidth, CellPatPercentage)
    mvarCells(i).Row = mvarRow
    mvarCells(i).Col = mvarCells.Count
    mvarColCount = mvarCells.Count
    mvarRow = Row
    AddCell = i
End Function


Public Property Let BorderWidth(ByVal vData As Long)
Attribute BorderWidth.VB_Description = "Ĭ�ϱ߿���ȡ�"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.BorderWidth = 5
    mvarBorderWidth = vData
End Property

Public Property Get BorderWidth() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.BorderWidth
    BorderWidth = mvarBorderWidth
End Property


Public Property Let BorderColor(ByVal vData As ColorEnum)
Attribute BorderColor.VB_Description = "�߿���ɫ��"
'������ָ�ɶ���ʱʹ�ã�λ�� Set ������ߡ�
'Syntax: Set x.BorderColor = Form1
    mvarBorderColor = vData
End Property


Public Property Get BorderColor() As ColorEnum
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.BorderColor
    BorderColor = mvarBorderColor
End Property


Public Property Let BorderStyle(ByVal vData As BorderStyleEnum)
Attribute BorderStyle.VB_Description = "�߿���ʽ��"
'������ָ�ɶ���ʱʹ�ã�λ�� Set ������ߡ�
'Syntax: Set x.BorderStyle = Form1
    mvarBorderStyle = vData
End Property


Public Property Get BorderStyle() As BorderStyleEnum
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.BorderStyle
    BorderStyle = mvarBorderStyle
End Property




Public Function GetTextRTF() As String
Attribute GetTextRTF.VB_Description = "��ȡ�е�RTF�ַ�����"
    Dim strCellDefine As String, strCellData As String, strR As String, strRowDefine As String
    Dim i As Long, j As Long
    
    '�ж����ַ���
    strRowDefine = "\pard\plain \trowd\irow" & mvarRow
    '�����������
    strRowDefine = strRowDefine & "\irowband" & mvarRow
    
    '�Ƿ����һ��
    If mvarIsLastRow Then strRowDefine = strRowDefine & "\lastrow"
    
    'Ĭ�ϲ�����ʽts15
    strRowDefine = strRowDefine & "\ts15"
    '�ж��뷽ʽ
    Select Case mvarAlignment
    Case ral2_���ж���
        strRowDefine = strRowDefine & "\trqc"
    Case ral3_���Ҷ���
        strRowDefine = strRowDefine & "\trqr"
    Case Else
        strRowDefine = strRowDefine & "\trql"
    End Select
    '��Ԫ����
    If mvarCellSpace <> 0 Then strRowDefine = strRowDefine & "\trgaph" & mvarCellSpace
    '�и�
    If mvarHeight <> 0 Then strRowDefine = strRowDefine & "\trrh" & mvarHeight
    '�������λ�ã�Ĭ��Ϊ-108
    strRowDefine = strRowDefine & "\trleft-108"
    Dim strBorderStyle As String, strBorderWidth As String, strBorderColor As String
    '�߿���ʽ
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
    '�߿���ɫ
    If mvarBorderColor <> clr00_Ĭ�� Then
        strBorderColor = "\brdrcf" & mvarBorderColor
    End If
    '�����б߿������
    If mvarBorderTop Then
        strRowDefine = strRowDefine & "\trbrdrt" & strBorderStyle & strBorderWidth & strBorderColor & vbCrLf
    Else
        strRowDefine = strRowDefine & "\trbrdrt\brdrnone" & vbCrLf
    End If
    If mvarBorderLeft Then
        strRowDefine = strRowDefine & "\trbrdrl" & strBorderStyle & strBorderWidth & strBorderColor & vbCrLf
    Else
        strRowDefine = strRowDefine & "\trbrdrl\brdrnone" & vbCrLf
    End If
    If mvarBorderBottom Then
        strRowDefine = strRowDefine & "\trbrdrb" & strBorderStyle & strBorderWidth & strBorderColor & vbCrLf
    Else
        strRowDefine = strRowDefine & "\trbrdrb\brdrnone" & vbCrLf
    End If
    If mvarBorderRight Then
        strRowDefine = strRowDefine & "\trbrdrr" & strBorderStyle & strBorderWidth & strBorderColor & vbCrLf
    Else
        strRowDefine = strRowDefine & "\trbrdrr\brdrnone" & vbCrLf
    End If
    If mvarBorderH Then
        strRowDefine = strRowDefine & "\trbrdrh" & strBorderStyle & strBorderWidth & strBorderColor & vbCrLf
    Else
        strRowDefine = strRowDefine & "\trbrdrh\brdrnone" & vbCrLf
    End If
    If mvarBorderV Then
        strRowDefine = strRowDefine & "\trbrdrv" & strBorderStyle & strBorderWidth & strBorderColor & vbCrLf
    Else
        strRowDefine = strRowDefine & "\trbrdrv\brdrnone" & vbCrLf
    End If
    
    '��Ԫ�����
    strRowDefine = strRowDefine & "\trftsWidth" & mvarWidthUnit
    If mvarWidth <> 0 Then strRowDefine = strRowDefine & "\trwWidth" & mvarWidth
    strRowDefine = strRowDefine & "\trftsWidthA" & mvarWidthUnit
    If mvarWidthA <> 0 Then strRowDefine = strRowDefine & "\trwWidthA" & mvarWidthA
    strRowDefine = strRowDefine & "\trftsWidthB" & mvarWidthUnit
    If mvarWidthB <> 0 Then strRowDefine = strRowDefine & "\trwWidthB" & mvarWidthB
    '���Զ���Ӧ
    If mvarAutoFit Then
        strRowDefine = strRowDefine & "\trautofit1"
    Else
        strRowDefine = strRowDefine & "\trautofit0"
    End If

    If mvarKeep Then strRowDefine = strRowDefine & "\trkeep"
    If mvarKeepFollow Then strRowDefine = strRowDefine & "\trkeepfollow"
    If mvarPadBottom <> 0 Then strRowDefine = strRowDefine & "\trpaddb" & mvarPadBottom
    If mvarPadLeft <> 0 Then strRowDefine = strRowDefine & "\trpaddl" & mvarPadLeft
    If mvarPadRight <> 0 Then strRowDefine = strRowDefine & "\trpaddr" & mvarPadRight
    If mvarPadTop <> 0 Then strRowDefine = strRowDefine & "\trpaddt" & mvarPadTop
    '�����λĬ�Ͼ�Ϊ��羣��������ص�
    strRowDefine = strRowDefine & "\trpaddfl3\trpaddft3\trpaddfb3\trpaddfr3\tabsnoovrlp1"
    
    '����ͼ�����ͣ�ͨ��Cell�ı�����ʵ��
    For i = 1 To mvarCells.Count
        If mvarPatType <> ptt00_�� Then mvarCells(i).CellPatType = mvarPatType
        If mvarPatForeColor <> ptt00_�� Then mvarCells(i).CellPatForeColor = mvarPatForeColor
        If mvarPatBackColor <> ptt00_�� Then mvarCells(i).CellPatBackColor = mvarPatBackColor
        If mvarPatPercentage <> ptt00_�� Then mvarCells(i).CellPatPercentage = mvarPatPercentage
    Next
    
    strRowDefine = strRowDefine & "\tbllkhdrrows\tbllklastrow\tbllkhdrcols\tbllklastcol"
    
    strCellDefine = ""
    For i = 1 To mvarCells.Count
        'ȡÿ����Ԫ������
        strCellDefine = strCellDefine & mvarCells(i).GetCellDefineRTF
    Next
    strCellData = ""
    For i = 1 To mvarCells.Count
        strCellData = strCellData & mvarCells(i).GetCellContentRTF
    Next
    
    strR = strRowDefine + strCellDefine + strCellData + strRowDefine + strCellDefine '����<define>+<data>+<define>��ʽ
    
    If Nested Then
        strR = strR & "\nestrow"
    Else
        strR = strR & "\row"    'Ƕ�ױ��н���
    End If
    GetTextRTF = strR
End Function

Public Property Set Cells(ByVal vData As cRTFCells)
Attribute Cells.VB_Description = "��ǰ�е����е�Ԫ��ļ��ϡ�"
'������ָ�ɶ���ʱʹ�ã�λ�� Set ������ߡ�
'Syntax: Set x.Cells = Form1
    Set mvarCells = vData
End Property


Public Property Get Cells() As cRTFCells
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Cells
    Set Cells = mvarCells
End Property


Public Property Let BorderV(ByVal vData As Boolean)
Attribute BorderV.VB_Description = "�Ƿ���ʾ���ڴ�ֱ�߿�"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.BorderV = 5
    mvarBorderV = vData
End Property


Public Property Get BorderV() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.BorderV
    BorderV = mvarBorderV
End Property



Public Property Let BorderH(ByVal vData As Boolean)
Attribute BorderH.VB_Description = "�Ƿ���ʾ����ˮƽ�߿�"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.BorderH = 5
    mvarBorderH = vData
End Property


Public Property Get BorderH() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.BorderH
    BorderH = mvarBorderH
End Property



Public Property Let BorderBottom(ByVal vData As Boolean)
Attribute BorderBottom.VB_Description = "�Ƿ���ʾ���±߿�"
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
Attribute BorderTop.VB_Description = "�Ƿ���ʾ���ϱ߿�"
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
Attribute BorderRight.VB_Description = "�Ƿ���ʾ���ұ߿�"
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
Attribute BorderLeft.VB_Description = "�Ƿ���ʾ����߿�"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.BorderLeft = 5
    mvarBorderLeft = vData
End Property


Public Property Get BorderLeft() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.BorderLeft
    BorderLeft = mvarBorderLeft
End Property

Public Property Let Nested(ByVal vData As Boolean)
Attribute Nested.VB_Description = "�Ƿ���Ƕ���У�������Ƕ�ױ���Ĭ��:False"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Nested = 5
    Dim i As Long
    For i = 1 To mvarCells.Count
        mvarCells(i).Nested = vData
    Next
    mvarNested = vData
End Property

Public Property Get Nested() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Nested
    Nested = mvarNested
End Property


Public Property Let PadBottom(ByVal vData As Long)
Attribute PadBottom.VB_Description = "Ĭ�����е�Ԫ���±߾ࡣ"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.PatBottom = 5
    mvarPadBottom = vData
End Property


Public Property Get PadBottom() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.PatBottom
    PadBottom = mvarPadBottom
End Property


Public Property Let PadTop(ByVal vData As Long)
Attribute PadTop.VB_Description = "Ĭ�����е�Ԫ���ϱ߾ࡣ"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.PatTop = 5
    mvarPadTop = vData
End Property


Public Property Get PadTop() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.PatTop
    PadTop = mvarPadTop
End Property


Public Property Let PadRight(ByVal vData As Long)
Attribute PadRight.VB_Description = "Ĭ�����е�Ԫ���ұ߾ࡣ"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.PadRight = 5
    mvarPadRight = vData
End Property


Public Property Get PadRight() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.PadRight
    PadRight = mvarPadRight
End Property



Public Property Let PadLeft(ByVal vData As Long)
Attribute PadLeft.VB_Description = "Ĭ�����е�Ԫ����߾ࡣ"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.PadLeft = 5
    mvarPadLeft = vData
End Property


Public Property Get PadLeft() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.PadLeft
    PadLeft = mvarPadLeft
End Property



Public Property Let Height(ByVal vData As Long)
Attribute Height.VB_Description = "�иߡ�"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Height = 5
    mvarHeight = vData
End Property


Public Property Get Height() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Height
    Height = mvarHeight
End Property



Public Property Let Alignment(ByVal vData As RowAlignTypeEnum)
Attribute Alignment.VB_Description = "�ж��뷽ʽ���������е�Ԫ��1��3��Ĭ��:0"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Alignment = 5
    mvarAlignment = vData
End Property


Public Property Get Alignment() As RowAlignTypeEnum
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Alignment
    Alignment = mvarAlignment
End Property



Public Property Let KeepFollow(ByVal vData As Boolean)
Attribute KeepFollow.VB_Description = "���ָ������һ��ͬҳ��"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.KeepFollow = 5
    mvarKeepFollow = vData
End Property


Public Property Get KeepFollow() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.KeepFollow
    KeepFollow = mvarKeepFollow
End Property



Public Property Let Keep(ByVal vData As Boolean)
Attribute Keep.VB_Description = "���ָ�����ǰһ��ͬҳ��"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Keep = 5
    mvarKeep = vData
End Property


Public Property Get Keep() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Keep
    Keep = mvarKeep
End Property



Public Property Let IsHeadRow(ByVal vData As Boolean)
Attribute IsHeadRow.VB_Description = "�Ƿ��Ǳ����У����н���ÿҳ������ʾ��Ĭ��:False"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.IsHeadRow = 5
    mvarIsHeadRow = vData
End Property


Public Property Get IsHeadRow() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.IsHeadRow
    IsHeadRow = mvarIsHeadRow
End Property



Public Property Let AutoFit(ByVal vData As Boolean)
Attribute AutoFit.VB_Description = "�������Ƿ����Զ���Ӧ��"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.AutoFit = 5
    mvarAutoFit = vData
End Property


Public Property Get AutoFit() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.AutoFit
    AutoFit = mvarAutoFit
End Property



Public Property Let PatPercentage(ByVal vData As Long)
Attribute PatPercentage.VB_Description = "���������ٷֱȡ�"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.PatPercentage = 5
    mvarPatPercentage = vData
End Property


Public Property Get PatPercentage() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.PatPercentage
    PatPercentage = mvarPatPercentage
End Property



Public Property Let PatBackColor(ByVal vData As ColorEnum)
Attribute PatBackColor.VB_Description = "���Ʊ���ɫ����ֵ��"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.PatBackColor = 5
    mvarPatBackColor = vData
End Property


Public Property Get PatBackColor() As ColorEnum
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.PatBackColor
    PatBackColor = mvarPatBackColor
End Property



Public Property Let PatForeColor(ByVal vData As ColorEnum)
Attribute PatForeColor.VB_Description = "����ǰ��ɫ����ֵ��"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.PatForeColor = 5
    mvarPatForeColor = vData
End Property


Public Property Get PatForeColor() As ColorEnum
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.PatForeColor
    PatForeColor = mvarPatForeColor
End Property



Public Property Let PatType(ByVal vData As PatTypeEnum)
Attribute PatType.VB_Description = "�������͡�"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.PatType = 5
    mvarPatType = vData
End Property


Public Property Get PatType() As PatTypeEnum
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.PatType
    PatType = mvarPatType
End Property



Public Property Let WidthAUnit(ByVal vData As CellLengthUnitEnum)
Attribute WidthAUnit.VB_Description = "��ĩ���ɼ���Ԫ����ȵ�λ��1��3��Ĭ�ϣ�0"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.WidthAUnit = 5
    mvarWidthAUnit = vData
End Property


Public Property Get WidthAUnit() As CellLengthUnitEnum
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.WidthAUnit
    WidthAUnit = mvarWidthAUnit
End Property



Public Property Let WidthA(ByVal vData As Long)
Attribute WidthA.VB_Description = "��ĩ���ɼ���Ԫ����ȡ�"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.WidthA = 5
    mvarWidthA = vData
End Property


Public Property Get WidthA() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.WidthA
    WidthA = mvarWidthA
End Property



Public Property Let WidthBUnit(ByVal vData As CellLengthUnitEnum)
Attribute WidthBUnit.VB_Description = "��ǰ���ɼ���Ԫ����ȵ�λ��1��3��Ĭ�ϣ�0"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.WidthBUnit = 5
    mvarWidthBUnit = vData
End Property


Public Property Get WidthBUnit() As CellLengthUnitEnum
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.WidthBUnit
    WidthBUnit = mvarWidthBUnit
End Property



Public Property Let WidthB(ByVal vData As Long)
Attribute WidthB.VB_Description = "��ǰ���ɼ���Ԫ����ȡ�"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.WidthB = 5
    mvarWidthB = vData
End Property


Public Property Get WidthB() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.WidthB
    WidthB = mvarWidthB
End Property



Public Property Let WidthUnit(ByVal vData As CellLengthUnitEnum)
Attribute WidthUnit.VB_Description = "Ĭ���п���λ��1��3��Ĭ�ϣ�0"
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
Attribute Width.VB_Description = "Ĭ���п���"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Width = 5
    mvarWidth = vData
End Property


Public Property Get Width() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Width
    Width = mvarWidth
End Property



Public Property Let CellSpace(ByVal vData As Long)
Attribute CellSpace.VB_Description = "��Ԫ���ࡣ"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.CellSpace = 5
    mvarCellSpace = vData
End Property


Public Property Get CellSpace() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.CellSpace
    CellSpace = mvarCellSpace
End Property



Public Property Let IsLastRow(ByVal vData As Boolean)
Attribute IsLastRow.VB_Description = "�Ƿ�����ĩһ�С�Ĭ�ϣ�False"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.IsLastRow = 5
    mvarIsLastRow = vData
End Property


Public Property Get IsLastRow() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.IsLastRow
    IsLastRow = mvarIsLastRow
End Property



Public Property Let RowBand(ByVal vData As Long)
Attribute RowBand.VB_Description = "�������ο������С�������Ϊ-1"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.RowBand = 5
    mvarRowBand = vData
End Property


Public Property Get RowBand() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.RowBand
    RowBand = mvarRowBand
End Property



Public Property Let Row(ByVal vData As Long)
Attribute Row.VB_Description = "��ǰ��ֵ��"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Row = 5
    mvarRow = vData
End Property


Public Property Get Row() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Row
    Row = mvarRow
End Property

Private Sub Class_Initialize()
    Set mvarCells = New cRTFCells
End Sub

Private Sub Class_Terminate()
    Set mvarCells = Nothing
End Sub