VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RPTPar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Public Key As String

'��������ֵ�ľֲ�����
Private mvar���� As String '�ֲ�����
Private mvar��� As Byte '�ֲ�����
Private mvar���� As String '�ֲ�����
Private mvar���� As Byte '�ֲ�����
Private mvarȱʡֵ As String '�ֲ�����
Private mvar��ʽ As Byte '�ֲ�����
Private mvarֵ�б� As String '�ֲ�����
Private mvar����SQL As String '�ֲ�����
Private mvar��ϸSQL As String '�ֲ�����
Private mvar�����ֶ� As String '�ֲ�����
Private mvar��ϸ�ֶ� As String '�ֲ�����
Private mvar���� As String '�ֲ�����
Private mvarReserve As String '�ֲ�����
Private mvar�Ƿ����� As Boolean '�ֲ�����

Public Property Let ����(ByVal vData As String)
    mvar���� = vData
End Property

Public Property Get ����() As String
    ���� = mvar����
End Property

Public Property Let ����(ByVal vData As String)
    mvar���� = vData
End Property

Public Property Get ����() As String
    ���� = mvar����
End Property

Public Property Let ��ϸ�ֶ�(ByVal vData As String)
    mvar��ϸ�ֶ� = vData
End Property

Public Property Get ��ϸ�ֶ�() As String
    ��ϸ�ֶ� = mvar��ϸ�ֶ�
End Property

Public Property Let �����ֶ�(ByVal vData As String)
    mvar�����ֶ� = vData
End Property

Public Property Get �����ֶ�() As String
    �����ֶ� = mvar�����ֶ�
End Property

Public Property Let ��ϸSQL(ByVal vData As String)
    mvar��ϸSQL = vData
End Property

Public Property Get ��ϸSQL() As String
    ��ϸSQL = mvar��ϸSQL
End Property

Public Property Let ����SQL(ByVal vData As String)
    mvar����SQL = vData
End Property

Public Property Get ����SQL() As String
    ����SQL = mvar����SQL
End Property

Public Property Let ֵ�б�(ByVal vData As String)
    mvarֵ�б� = vData
End Property

Public Property Get ֵ�б�() As String
    ֵ�б� = mvarֵ�б�
End Property

Public Property Let ȱʡֵ(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ȱʡֵ = 5
    mvarȱʡֵ = vData
End Property

Public Property Get ȱʡֵ() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ȱʡֵ
    ȱʡֵ = mvarȱʡֵ
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

Public Property Let ���(ByVal vData As Byte)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��� = 5
    mvar��� = vData
End Property

Public Property Get ���() As Byte
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.���
    ��� = mvar���
End Property

Public Property Let Reserve(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ȱʡֵ = 5
    mvarReserve = vData
End Property

Public Property Get Reserve() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ȱʡֵ
    Reserve = mvarReserve
End Property

Public Property Let ��ʽ(ByVal vData As Byte)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��ʽ = 5
    mvar��ʽ = vData
End Property

Public Property Get ��ʽ() As Byte
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.��ʽ
    ��ʽ = mvar��ʽ
End Property

Public Property Let �Ƿ�����(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��ʽ = 5
    mvar�Ƿ����� = vData
End Property

Public Property Get �Ƿ�����() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.��ʽ
    �Ƿ����� = mvar�Ƿ�����
End Property

