VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FuncPar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Key As String

'��������ֵ�ľֲ�����
Private mvar���� As String '�ֲ�����
Private mvar��� As Byte '�ֲ�����
Private mvar���� As String '�ֲ�����
Private mvar������ As String '�ֲ�����
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

Public Property Let Reserve(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Reserve = 5
    mvarReserve = vData
End Property

Public Property Get Reserve() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Reserve
    Reserve = mvarReserve
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

Public Property Let ��ϸ�ֶ�(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��ϸ�ֶ� = 5
    mvar��ϸ�ֶ� = vData
End Property

Public Property Get ��ϸ�ֶ�() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.��ϸ�ֶ�
    ��ϸ�ֶ� = mvar��ϸ�ֶ�
End Property

Public Property Let �����ֶ�(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�����ֶ� = 5
    mvar�����ֶ� = vData
End Property

Public Property Get �����ֶ�() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�����ֶ�
    �����ֶ� = mvar�����ֶ�
End Property

Public Property Let ��ϸSQL(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��ϸSQL = 5
    mvar��ϸSQL = vData
End Property

Public Property Get ��ϸSQL() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.��ϸSQL
    ��ϸSQL = mvar��ϸSQL
End Property

Public Property Let ����SQL(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.����SQL = 5
    mvar����SQL = vData
End Property

Public Property Get ����SQL() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����SQL
    ����SQL = mvar����SQL
End Property

Public Property Let ֵ�б�(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ֵ�б� = 5
    mvarֵ�б� = vData
End Property

Public Property Get ֵ�б�() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ֵ�б�
    ֵ�б� = mvarֵ�б�
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

Public Property Let ������(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.������ = 5
    mvar������ = vData
End Property

Public Property Get ������() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.������
    ������ = mvar������
End Property

Public Property Let ����(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvar���� = vData
End Property

Public Property Get ����() As String
Attribute ����.VB_UserMemId = 0
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