VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Bill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "��Һ��"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'��������ֵ�ľֲ�����
Private mvar����ID As Long '�ֲ�����
Private mvarҽ��ID As Long '�ֲ�����
Private mvarҽ������ As String '�ֲ�����

'��������ֵ�ľֲ�����
Private mvar���� As Currency  '�ֲ�����
Private mvar��λ As String '�ֲ�����
Private mvar���� As Long '�ֲ�����
Private mvarҽ������ As String '�ֲ�����
Private mvar��� As Integer '�ֲ�����
Private mvarʱ�� As Long
Private mvar��� As Currency '
Private mvar�Ʒ�״̬ As Integer

'��������ֵ�ľֲ�����

Public Property Let ���(ByVal vData As Integer)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��� = 5
    mvar��� = vData
End Property


Public Property Get ���() As Integer
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.���
    ��� = mvar���
End Property


Public Property Let ҽ������(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ҽ������ = 5
    mvarҽ������ = MidUni(vData, 1, 200)
End Property


Public Property Get ҽ������() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.ҽ������
    ҽ������ = mvarҽ������
End Property

Public Property Let ����(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvar���� = vData
End Property


Public Property Get ����() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.����
    ���� = mvar����
End Property

Public Property Let ��λ(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��λ = 5
    mvar��λ = vData
End Property


Public Property Get ��λ() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.��λ
    ��λ = mvar��λ
End Property

Public Property Let ���(ByVal vData As Currency)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvar��� = vData
End Property


Public Property Get ���() As Currency
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.����
    ��� = mvar���
End Property

Public Property Let ����(ByVal vData As Currency)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvar���� = vData
End Property


Public Property Get ����() As Currency
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.����
    ���� = mvar����
End Property

Public Property Get ʱ��() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.����
    ʱ�� = mvarʱ��
End Property

Public Property Let ʱ��(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvarʱ�� = vData
End Property

Public Property Let ҽ������(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ҽ������ = 5
    mvarҽ������ = vData
End Property


Public Property Get ҽ������() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.ҽ������
    ҽ������ = mvarҽ������
End Property



Public Property Let ҽ��ID(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ҽ��ID = 5
    mvarҽ��ID = vData
End Property


Public Property Get ҽ��ID() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.ҽ��ID
    ҽ��ID = mvarҽ��ID
End Property



Public Property Let ����ID(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.����ID = 5
    mvar����ID = vData
End Property


Public Property Get ����ID() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.����ID
    ����ID = mvar����ID
End Property

Public Property Let ��ϸ�Ʒ�״̬(ByVal vData As Integer)
    mvar�Ʒ�״̬ = vData
End Property

Public Property Get ��ϸ�Ʒ�״̬() As Integer
    ��ϸ�Ʒ�״̬ = mvar�Ʒ�״̬
End Property