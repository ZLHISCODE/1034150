VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRTFNode"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

'ȫ�ֹ������Ͷ���Ӧ����һ��ȫ�ֹ������ж��壺


'ͼƬ����
Public Enum PicTypeEnum
    pic_JPG = 0
    pic_GIF = 1
    pic_BMP = 2
    pic_PNG = 3
End Enum

'�̶���ɫ����
Public Enum ColorEnum
    clr00_Ĭ�� = 0
    clr01_��ɫ = 1
    clr02_��ɫ = 2
    clr03_���� = 3
    clr04_��ɫ = 4
    clr05_�ۺ� = 5
    clr06_��� = 6
    clr07_��ɫ = 7
    clr08_��ɫ = 8
    clr09_���� = 9
    clr10_���� = 10
    clr11_���� = 11
    clr12_ǳ�� = 12
    clr13_���� = 13
    clr14_���� = 14
    clr15_��� = 15
    clr16_ǳ�� = 16
End Enum

'����ߴ磺
Public Enum FontSizeEnum
    fs00_Ĭ�� = 0
    fs01_���� = 42
    fs02_С�� = 36
    fs03_һ�� = 26
    fs04_Сһ = 24
    fs05_���� = 22
    fs06_С�� = 18
    fs07_���� = 16
    fs08_С�� = 15
    fs09_�ĺ� = 14
    fs10_С�� = 12
    fs11_��� = 10.5
    fs12_С�� = 9
    fs13_���� = 7.5
    fs14_С�� = 6.5
    fs15_�ߺ� = 5.5
    fs16_�˺� = 5
End Enum

'������ʽ��
Public Enum FontStyleEnum
    ft0_���� = 0
    ft1_����1 = 1
    ft2_����2 = 2
    ft3_����3 = 3
End Enum

'�ַ���ʽ��
Public Enum CharStyleEnum
    csĬ�� = 0
End Enum

'��̬�ı���
Public Enum AnimTypeEnum
    ani0_�� = 0
    ani1_�߲��޺� = 1
    ani2_��˸���� = 2
    ani3_������ = 3
    ani4_�������� = 4
    ani5_��ˮ���� = 5
    ani6_������� = 6
End Enum

'������
Public Enum EmboEnum
    emb0_�� = 0
    emb1_����Ч�� = 1
    emb2_���Ч�� = 2
End Enum

'�»�������
Public Enum UnderLineTypeEnum
    udl00_�� = 0
    udl01_������ = 1
    udl02_���� = 2
    udl03_�̻��� = 3
    udl04_�㻮�� = 4
    udl05_˫�㻮�� = 5
    udl06_˫�� = 6
    udl07_���ز����� = 7
    udl08_������ = 8
    udl09_���� = 9
    udl10_�ֵ��� = 10
    udl11_�ֶ̻��� = 11
    udl12_�ֵ㻮�� = 12
    udl13_��˫�㻮�� = 13
    udl14_�ֳ����� = 14
    udl15_˫������ = 15
    udl16_���¼��� = 16
    udl17_������ = 17
End Enum

'����ͼ������
Public Enum PatTypeEnum
    ptt00_�� = 0
    ptt01_���� = 1
    ptt02_���� = 2
    ptt03_��б�� = 3
    ptt04_��б�� = 4
    ptt05_������ = 5
    ptt06_б������ = 6
    ptt07_�ֺ��� = 7
    ptt08_������ = 8
    ptt09_����б�� = 9
    ptt10_����б�� = 10
    ptt11_�������� = 11
    ptt12_��б������ = 12
End Enum

'���뷽ʽ
Public Enum AlignTypeEnum
    alm0_Ĭ�϶��� = 0
    alm1_��˶��� = 1
    alm2_���ж��� = 2
    alm3_���Ҷ��� = 3
    alm4_���˶��� = 4
    alm5_��ɢ���� = 5
End Enum

'��Ŀ��������
Public Enum ListTypeEnum
    ltp00_�� = 0
    ltp01_�ı�1 = 1
    ltp02_�ı�2 = 2
    ltp03_�ı�3 = 3
    ltp04_�ı�4 = 4
    ltp05_�ı�5 = 5
    ltp06_�ı�6 = 6
    ltp07_�ı�7 = 7
    ltp08_�ı�8 = 8
    ltp09_����1 = 9
    ltp10_����2 = 10
    ltp11_����3 = 11
    ltp12_����4 = 12
    ltp13_����5 = 13
    ltp14_����6 = 14
    ltp15_����7 = 15
End Enum


'��ֱ���뷽ʽ
Public Enum AlignVTypeEnum
    alv0_Ĭ�϶��� = 0
    alv1_���˶��� = 1
    alv2_���ж��� = 2
    alv3_�׶˶��� = 3
End Enum

'��Ԫ��߿���ʽ
Public Enum BorderStyleEnum
    cbs00_Ĭ�ϱ߿���ʽ = 0
    cbs01_�ޱ߿� = 1
    cbs02_������ȱ߿� = 2
    cbs03_˫����ȱ߿� = 3
    cbs04_��Ӱ�߿� = 4
    cbs05_˫�߿� = 5
    cbs06_���߱߿� = 6
    cbs07_���߱߿� = 7
    cbs08_ϸ�߱߿� = 8
    cbs09_����ʽ�߿� = 9
    cbs10_����ϸ�߿� = 10
    cbs11_�㻮�߱߿� = 11
    cbs12_��㻮�߱߿� = 12
    cbs13_��ʼ�߿� = 13
    cbs14_���ر߿� = 14
    cbs15_�񱡵���ϸ�߿� = 15
    cbs16_�������ϸ�߿� = 16
    cbs17_���񱡵���ϸ�߿� = 17
    cbs18_�񱡵����б߿� = 18
    cbs19_��������б߿� = 19
    cbs20_���񱡵����б߿� = 20
    cbs21_�񱡵��Ӵֱ߿� = 21
    cbs22_������Ӵֱ߿� = 22
    cbs23_���񱡵��Ӵֱ߿� = 23
    cbs24_�����߱߿� = 24
    cbs25_˫�����߱߿� = 25
    cbs26_���Ʊ߿� = 26
    cbs27_����߿� = 27
    cbs28_��̱߿� = 28
End Enum

'��Ԫ�񳤶ȵ�λ
Public Enum CellLengthUnitEnum
    clt0_Ĭ�� = 0
    clt1_�Զ� = 1
    clt2_�ٷֱ� = 2
    clt3_� = 3
End Enum

'��Ԫ����������
Public Enum CellContentTypeEnum
    cct0_�ı� = 0
    cct1_���� = 1
    cct2_ͼƬ = 2
    cct3_���� = 3
End Enum


'�ж��뷽ʽ
Public Enum RowAlignTypeEnum
    ral0_Ĭ�϶��� = 0
    ral1_��˶��� = 1
    ral2_���ж��� = 2
    ral3_���Ҷ��� = 3
End Enum

'�ĵ���ͼ����
Public Enum DrawingDirectionEnum
    dwd0_ˮƽ��ͼ = 0
    dwd1_��ֱ��ͼ = 1
End Enum

'ҳ��ߴ�
Public Enum PaperSizeEnum
    pps0_Default = 0
    pps1_Letter = 1
    pps2_Tabloid = 2
    pps3_Legal = 3
    pps4_A3 = 4
    pps5_A4 = 5
    pps6_A5 = 6
    pps7_B4 = 7
    pps8_B5 = 8
End Enum

'ҳ����ͼ
Public Enum ViewKindEnum
    vkd0_�� = 0
    vkd1_ҳ��ͼ = 1
    vkd2_�����ͼ = 2
    vkd3_���ĵ���ͼ = 3
    vkd4_������ͼ = 4
    vkd5_������ͼ = 5
End Enum

'��������ֵ�ľֲ�����
Private mvarSubID As Long
Private mvarNodeType As String

Public Property Let NodeType(ByVal vData As String)
Attribute NodeType.VB_Description = "�ַ�������ʾ�ڵ����͡��ı���ͼƬ������ȡ�"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.NodeType = 5
    mvarNodeType = vData
End Property


Public Property Get NodeType() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.NodeType
    NodeType = mvarNodeType
End Property



Public Property Let SubID(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.SubID = 5
    mvarSubID = vData
End Property


Public Property Get SubID() As Long
Attribute SubID.VB_Description = "��ID���ı���ͼƬ���߱����ID��"
Attribute SubID.VB_UserMemId = 0
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.SubID
    SubID = mvarSubID
End Property
