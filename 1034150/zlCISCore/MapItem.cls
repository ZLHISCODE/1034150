VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MapItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'��������ֵ�ľֲ�����
Private mvar���� As Byte '�ֲ�����
Private mvar���� As String '�ֲ�����
Private mvar�㼯 As String '�ֲ�����
Private mvarX1 As Long '�ֲ�����
Private mvarY1 As Long '�ֲ�����
Private mvarX2 As Long '�ֲ�����
Private mvarY2 As Long '�ֲ�����
Private mvar���ɫ As Long '�ֲ�����
Private mvar��䷽ʽ As Integer '�ֲ�����
Private mvar����ɫ As Long '�ֲ�����
Private mvar���� As Byte  '�ֲ�����
Private mvar�߿� As Byte  '�ֲ�����
Private mvar���� As String   '�ֲ�����

Private mvarKey As String   '�ֲ�����
Private mvarOper As Byte    '�ֲ�����

Public Property Let Key(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���ɫ = 5
    mvarKey = vData
End Property

Public Property Get Key() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.���ɫ
    Key = mvarKey
End Property

Public Property Let Oper(ByVal vData As Byte)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���ɫ = 5
    mvarOper = vData
End Property

Public Property Get Oper() As Byte
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.���ɫ
    Oper = mvarOper
End Property

Public Property Let ����(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���ɫ = 5
    mvar���� = vData
End Property

Public Property Get ����() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.���ɫ
    ���� = mvar����
End Property

Public Property Let �߿�(ByVal vData As Byte)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���ɫ = 5
    mvar�߿� = vData
End Property

Public Property Get �߿�() As Byte
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.���ɫ
    �߿� = mvar�߿�
End Property

Public Property Let ����(ByVal vData As Byte)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���ɫ = 5
    mvar���� = vData
End Property

Public Property Get ����() As Byte
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.���ɫ
    ���� = mvar����
End Property

Public Property Let ����ɫ(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���ɫ = 5
    mvar����ɫ = vData
End Property

Public Property Get ����ɫ() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.���ɫ
    ����ɫ = mvar����ɫ
End Property

Public Property Let ��䷽ʽ(ByVal vData As Integer)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��䷽ʽ = 5
    mvar��䷽ʽ = vData
End Property

Public Property Get ��䷽ʽ() As Integer
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.��䷽ʽ
    ��䷽ʽ = mvar��䷽ʽ
End Property


Public Property Let ���ɫ(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���ɫ = 5
    mvar���ɫ = vData
End Property

Public Property Get ���ɫ() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.���ɫ
    ���ɫ = mvar���ɫ
End Property

Public Property Let Y2(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Y2 = 5
    mvarY2 = vData
End Property

Public Property Get Y2() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Y2
    Y2 = mvarY2
End Property

Public Property Let X2(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.X2 = 5
    mvarX2 = vData
End Property

Public Property Get X2() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.X2
    X2 = mvarX2
End Property

Public Property Let Y1(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Y1 = 5
    mvarY1 = vData
End Property

Public Property Get Y1() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Y1
    Y1 = mvarY1
End Property

Public Property Let X1(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.X1 = 5
    mvarX1 = vData
End Property

Public Property Get X1() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.X1
    X1 = mvarX1
End Property

Public Property Let �㼯(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�㼯 = 5
    mvar�㼯 = vData
End Property

Public Property Get �㼯() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�㼯
    �㼯 = mvar�㼯
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