VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RPTItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Public Key As String

Private mvarID As Integer  '�ֲ�����
Private mvar��ʽ�� As Byte  '�ֲ�����
Private mvar���� As String
Private mvar�ϼ�ID As Integer  '�ֲ�����
Private mvar��� As Integer '�ֲ�����
Private mvar���� As Byte '�ֲ�����
Private mvar���� As String '�ֲ�����
Private mvarͼƬ As StdPicture
Private mvar���� As String
Private mvar���� As Byte
Private mvar��ͷ As String '�ֲ�����
Private mvarX As Long '�ֲ�����
Private mvarY As Long '�ֲ�����
Private mvarW As Long '�ֲ�����
Private mvarH As Long '�ֲ�����
Private mvar�и� As Long '�ֲ�����
Private mvar���� As Byte '�ֲ�����
Private mvar�Ե� As Boolean
Private mvar���� As String '�ֲ�����
Private mvar�ֺ� As Single '�ֲ�����
Private mvar���� As Boolean '�ֲ�����
Private mvar�����߼Ӵ� As Boolean '�ֲ�����
Private mvar���� As Boolean '�ֲ�����
Private mvarб�� As Boolean '�ֲ�����
Private mvar���� As Long '�ֲ�����
Private mvarǰ�� As Long '�ֲ�����
Private mvar���� As Long '�ֲ�����
Private mvar�߿� As Boolean '�ֲ�����
Private mvarϵͳ As Boolean '�ֲ�����
Private mvar���� As Byte '�ֲ�����
Private mvar���� As String '�ֲ�����
Private mvar��ʽ As String '�ֲ�����
Private mvar���� As String '�ֲ�����
Private mvarCopyIDs As RelatIDs '�ֲ�����
Private mvarSubIDs As RelatIDs '�ֲ�����
Private mvarÿҳ���� As Integer '�ֲ�����
Private mvar��ID As Integer '�ֲ�����
Private mvar����Դ As String '�ֲ�����
Private mvar���¼�� As Long '�ֲ�����
Private mvar���Ҽ�� As Long '�ֲ�����
Private mvar������� As Long '�ֲ�����
Private mvar������� As Long '�ֲ�����
Private mvarԴ�к� As Long '�ֲ�����
Private mvarRelations As RPTRelations '�ֲ�����
Private mvarColProtertys As RPTColProtertys '�ֲ�����
Private mstrValue As String
Private mbln����Ӧ�и� As Boolean                   '����Ӧ�и�
Private mblnˮƽ��ת As Boolean
Private mbln��ֵ�Ԫ�� As Boolean
Private mstr�������� As String                      '��ǩ�ο��������
Private mbln�Զ���� As Boolean

Public Property Get �Զ����() As Boolean
    �Զ���� = mbln�Զ����
End Property
Public Property Let �Զ����(ByVal vData As Boolean)
    mbln�Զ���� = vData
End Property

'��ǩר������
Public Property Get ��������() As String
    �������� = mstr��������
End Property
Public Property Let ��������(ByVal vData As String)
    mstr�������� = vData
End Property

Public Property Get ˮƽ��ת() As Boolean
    ˮƽ��ת = mblnˮƽ��ת
End Property
Public Property Let ˮƽ��ת(ByVal vData As Boolean)
    mblnˮƽ��ת = vData
End Property

Public Property Get ��ֵ�Ԫ��() As Boolean
    ��ֵ�Ԫ�� = mbln��ֵ�Ԫ��
End Property
Public Property Let ��ֵ�Ԫ��(ByVal vData As Boolean)
    mbln��ֵ�Ԫ�� = vData
End Property

Public Property Get ����Ӧ�и�() As Boolean
    ����Ӧ�и� = mbln����Ӧ�и�
End Property
Public Property Let ����Ӧ�и�(ByVal vData As Boolean)
    mbln����Ӧ�и� = vData
End Property

Private Sub Class_Initialize()
    Set CopyIDs = New RelatIDs
    Set SubIDs = New RelatIDs
    Set Relations = New RPTRelations
    Set ColProtertys = New RPTColProtertys
    Set ͼƬ = Nothing
End Sub

Public Property Let Value(ByVal vData As String)
    mstrValue = vData
End Property

Public Property Get Value() As String
    Value = mstrValue
End Property

Public Property Let ����(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvar���� = vData
End Property

Public Property Get ����() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����
    ���� = mvar����
End Property

Public Property Let ����(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvar���� = vData
End Property

Public Property Get ����() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����
    ���� = mvar����
End Property

Public Property Let ����(ByVal vData As Byte)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvar���� = vData
End Property

Public Property Get ����() As Byte
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����
    ���� = mvar����
End Property

Public Property Let �Ե�(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�Ե� = 5
    mvar�Ե� = vData
End Property

Public Property Get �Ե�() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�Ե�
    �Ե� = mvar�Ե�
End Property

Public Property Set ͼƬ(ByVal vData As StdPicture)
'������ָ�ɶ���ʱʹ�ã�λ�� Set ������ߡ�
'Syntax: Set x.ͼƬ = Form1
    Set mvarͼƬ = vData
End Property

Public Property Get ͼƬ() As StdPicture
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ͼƬ
    Set ͼƬ = mvarͼƬ
End Property

Public Property Set SubIDs(ByVal vData As RelatIDs)
'������ָ�ɶ���ʱʹ�ã�λ�� Set ������ߡ�
'Syntax: Set x.SubIDs = Form1
    Set mvarSubIDs = vData
End Property

Public Property Get SubIDs() As RelatIDs
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.SubIDs
    Set SubIDs = mvarSubIDs
End Property

Public Property Set CopyIDs(ByVal vData As RelatIDs)
'������ָ�ɶ���ʱʹ�ã�λ�� Set ������ߡ�
'Syntax: Set x.CopyIDs = Form1
    Set mvarCopyIDs = vData
End Property

Public Property Get CopyIDs() As RelatIDs
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.CopyIDs
    Set CopyIDs = mvarCopyIDs
End Property

Public Property Let ����(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvar���� = vData
End Property

Public Property Get ����() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����
    ���� = mvar����
End Property

Public Property Let ����(ByVal vData As Byte)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvar���� = vData
End Property

Public Property Get ����() As Byte
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����
    ���� = mvar����
End Property

Public Property Let ����(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvar���� = vData
End Property

Public Property Get ����() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����
    ���� = mvar����
End Property

Public Property Let ǰ��(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ǰ�� = 5
    mvarǰ�� = vData
End Property

Public Property Get ǰ��() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ǰ��
    ǰ�� = mvarǰ��
End Property

Public Property Let ����(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ǰ�� = 5
    mvar���� = vData
End Property

Public Property Get ����() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ǰ��
    ���� = mvar����
End Property

Public Property Let б��(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.б�� = 5
    mvarб�� = vData
End Property

Public Property Get б��() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.б��
    б�� = mvarб��
End Property

Public Property Let ����(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvar���� = vData
End Property

Public Property Get ����() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����
    ���� = mvar����
End Property

Public Property Let ����(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvar���� = vData
End Property

Public Property Get ����() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����
    ���� = mvar����
End Property

Public Property Let �����߼Ӵ�(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvar�����߼Ӵ� = vData
End Property

Public Property Get �����߼Ӵ�() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����
    �����߼Ӵ� = mvar�����߼Ӵ�
End Property

Public Property Let �ֺ�(ByVal vData As Single)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�ֺ� = 5
    mvar�ֺ� = vData
End Property

Public Property Get �ֺ�() As Single
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�ֺ�
    �ֺ� = mvar�ֺ�
End Property

Public Property Let ����(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvar���� = vData
End Property

Public Property Get ����() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����
    ���� = mvar����
End Property

Public Property Let ����(ByVal vData As Byte)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvar���� = vData
End Property

Public Property Get ����() As Byte
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����
    ���� = mvar����
End Property

Public Property Let �и�(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�и� = 5
    mvar�и� = vData
End Property

Public Property Get �и�() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�и�
    �и� = mvar�и�
End Property

Public Property Let H(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.H = 5
    mvarH = vData
End Property

Public Property Get H() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.H
    H = mvarH
End Property

Public Property Let W(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.W = 5
    mvarW = vData
End Property

Public Property Get W() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.W
    W = mvarW
End Property

Public Property Let Y(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Y = 5
    mvarY = vData
End Property

Public Property Get Y() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Y
    Y = mvarY
End Property

Public Property Let X(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.X = 5
    mvarX = vData
End Property

Public Property Get X() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.X
    X = mvarX
End Property

Public Property Let ��ͷ(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��ͷ = 5
    mvar��ͷ = vData
End Property

Public Property Get ��ͷ() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.��ͷ
    ��ͷ = mvar��ͷ
End Property

Public Property Let ����(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvar���� = vData
End Property

Public Property Get ����() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����
    ���� = mvar����
End Property

Public Property Let ����(ByVal vData As Byte)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'˵����
'1-������2-��ǩ/����Դ��ǩ��3-δʹ�ã�4-���ɱ��񣨱�ͷ����6-���ɱ��񣨱��壩
'5-���ܱ���7-���ܱ��������ࣩ��8-���ܱ��񣨺����ࣩ��9-���ܱ���ͳ�����10-����
'11-ͼƬ��12-ͼ����13-���룻14-��Ƭ
    mvar���� = vData
End Property

Public Property Get ����() As Byte
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
    ���� = mvar����
End Property

Public Property Let ���(ByVal vData As Integer)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��� = 5
    mvar��� = vData
End Property

Public Property Get ���() As Integer
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.���
    ��� = mvar���
End Property

Public Property Let �ϼ�ID(ByVal vData As Integer)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�ϼ�ID = 5
    mvar�ϼ�ID = vData
End Property

Public Property Get �ϼ�ID() As Integer
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�ϼ�ID
    �ϼ�ID = mvar�ϼ�ID
End Property

Public Property Let id(ByVal vData As Integer)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ID = 5
    mvarID = vData
End Property

Public Property Get id() As Integer
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ID
    id = mvarID
End Property

Public Property Let �߿�(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvar�߿� = vData
End Property

Public Property Get �߿�() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����
    �߿� = mvar�߿�
End Property

Public Property Let ��ʽ(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvar��ʽ = vData
End Property

Public Property Get ��ʽ() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����
    ��ʽ = mvar��ʽ
End Property

Public Property Let ����(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvar���� = vData
End Property

Public Property Get ����() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����
    ���� = mvar����
End Property

Public Property Let ��ʽ��(ByVal vData As Byte)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��ʽ�� = 5
    mvar��ʽ�� = vData
End Property

Public Property Get ��ʽ��() As Byte
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.��ʽ��
    ��ʽ�� = mvar��ʽ��
End Property

Public Property Let ϵͳ(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ϵͳ = 5
    mvarϵͳ = vData
End Property

Public Property Get ϵͳ() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ϵͳ
    ϵͳ = mvarϵͳ
End Property

Public Property Let ÿҳ����(ByVal vData As Integer)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ID = 5
    mvarÿҳ���� = vData
End Property

Public Property Get ÿҳ����() As Integer
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ID
    ÿҳ���� = mvarÿҳ����
End Property


Public Property Get ��ID() As Integer
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�ϼ�ID
    ��ID = mvar��ID
End Property

Public Property Let ��ID(ByVal vData As Integer)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�ϼ�ID = 5
    mvar��ID = vData
End Property

Public Property Get ����Դ() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�ϼ�ID
    ����Դ = mvar����Դ
End Property

Public Property Let ����Դ(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�ϼ�ID = 5
    mvar����Դ = vData
End Property

Public Property Get ���¼��() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�ϼ�ID
    ���¼�� = mvar���¼��
End Property

Public Property Let ���¼��(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�ϼ�ID = 5
    mvar���¼�� = vData
End Property

Public Property Get ���Ҽ��() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�ϼ�ID
    ���Ҽ�� = mvar���Ҽ��
End Property

Public Property Let ���Ҽ��(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�ϼ�ID = 5
    mvar���Ҽ�� = vData
End Property

Public Property Get �������() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�ϼ�ID
    ������� = mvar�������
End Property

Public Property Let �������(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�ϼ�ID = 5
    mvar������� = vData
End Property

Public Property Get �������() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�ϼ�ID
    ������� = mvar�������
End Property

Public Property Let �������(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�ϼ�ID = 5
    mvar������� = vData
End Property

Public Property Get Դ�к�() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�ϼ�ID
    Դ�к� = mvarԴ�к�
End Property

Public Property Let Դ�к�(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�ϼ�ID = 5
    mvarԴ�к� = vData
End Property

Public Property Set Relations(ByVal vData As RPTRelations)
'������ָ�ɶ���ʱʹ�ã�λ�� Set ������ߡ�
'Syntax: Set x.Items = Form1
    Set mvarRelations = vData
End Property

Public Property Get Relations() As RPTRelations
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Items
    Set Relations = mvarRelations
End Property

Public Property Set ColProtertys(ByVal vData As RPTColProtertys)
'������ָ�ɶ���ʱʹ�ã�λ�� Set ������ߡ�
'Syntax: Set x.Items = Form1
    Set mvarColProtertys = vData
End Property

Public Property Get ColProtertys() As RPTColProtertys
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Items
    Set ColProtertys = mvarColProtertys
End Property

