VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MediBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "�Ĵ�ҩƷ"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'��������ֵ�ľֲ�����
Private mvar��� As Integer '�ֲ�����
Private mvarҽ��ID As Long '�ֲ�����
Private mvar���ͺ� As Long '
Private mvarҩƷID As Long '�ֲ�����
Private mvarҩƷ���� As String '�ֲ�����
Private mvar��� As String '�ֲ�����
Private mvarִ�з��� As Integer '�ֲ�����
Private mvarʹ��״̬ As Integer '�ֲ�����
Private mvar���ϵ�� As Integer '�ֲ�����
Private mvar���㵥λ As String '�ֲ�����
Private mvar���ﵥλ As String '�ֲ�����
Private mvar�����װ As Double '�ֲ�����
Private mvar������λ As String '�ֲ�����
Private mvar����ϵ�� As Double '�ֲ�����
Private mvar���� As Double '�ֲ�����
Private mvar���� As Double '�ֲ�����
Private mvar�������� As Double
Private mvar���� As Double '�ֲ�����
Private mvar��� As Double '�ֲ�����
Private mvar�Ǽ�ʱ�� As Date
Private mvarʹ��ժҪ As String
Private mvar������ As String

Public Property Let ���(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��� = 5
    mvar��� = vData
End Property

Public Property Get ���() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.���
    ��� = mvar���
End Property

Public Property Let ����(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvar���� = vData
End Property


Public Property Get ����() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.����
    ���� = mvar����
End Property

Public Property Let ����(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvar���� = vData
End Property


Public Property Get ����() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.����
    ���� = mvar����
End Property

Public Property Let ��������(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvar�������� = vData
End Property


Public Property Get ��������() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.����
    �������� = mvar��������
End Property

Public Property Let ����(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvar���� = vData
End Property


Public Property Get ����() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.����
    ���� = mvar����
End Property

Public Property Let ����ϵ��(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.����ϵ�� = 5
    mvar����ϵ�� = vData
End Property


Public Property Get ����ϵ��() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.����ϵ��
    ����ϵ�� = mvar����ϵ��
End Property

Public Property Let �����װ(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.����ϵ�� = 5
    mvar�����װ = vData
End Property

Public Property Get �����װ() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.����ϵ��
    �����װ = mvar�����װ
End Property

Public Property Let ������λ(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.������λ = 5
    mvar������λ = vData
End Property

Public Property Get ������λ() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.������λ
    ������λ = mvar������λ
End Property

Public Property Let ���ﵥλ(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���ﵥλ = 5
    mvar���ﵥλ = MidUni(vData, 1, 20)
End Property


Public Property Get ���ﵥλ() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.���ﵥλ
    ���ﵥλ = mvar���ﵥλ
End Property

Public Property Let ���㵥λ(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���ﵥλ = 5
    mvar���㵥λ = MidUni(vData, 1, 20)
End Property


Public Property Get ���㵥λ() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.���ﵥλ
    ���㵥λ = mvar���㵥λ
End Property


Public Property Let ���ϵ��(ByVal vData As Integer)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���ϵ�� = 5
    mvar���ϵ�� = vData
End Property


Public Property Get ���ϵ��() As Integer
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.���ϵ��
    ���ϵ�� = mvar���ϵ��
End Property



Public Property Let ʹ��״̬(ByVal vData As Integer)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ʹ��״̬ = 5
    mvarʹ��״̬ = vData
End Property


Public Property Get ʹ��״̬() As Integer
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.ʹ��״̬
    ʹ��״̬ = mvarʹ��״̬
End Property



Public Property Let ִ�з���(ByVal vData As Integer)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ִ�з��� = 5
    mvarִ�з��� = vData
End Property


Public Property Get ִ�з���() As Integer
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.ִ�з���
    ִ�з��� = mvarִ�з���
End Property



Public Property Let ���(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��� = 5
    mvar��� = MidUni(vData, 1, 40)
End Property


Public Property Get ���() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.���
    ��� = mvar���
End Property

Public Property Let ҩƷ����(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ҩƷ���� = 5
    mvarҩƷ���� = MidUni(vData, 1, 80)
End Property


Public Property Get ҩƷ����() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.ҩƷ����
    ҩƷ���� = mvarҩƷ����
End Property



Public Property Let ҩƷID(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ҩƷID = 5
    mvarҩƷID = MidUni(vData, 1, 80)
End Property


Public Property Get ҩƷID() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.ҩƷID
    ҩƷID = mvarҩƷID
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


Public Property Let ���ͺ�(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ҽ��ID = 5
    mvar���ͺ� = vData
End Property


Public Property Get ���ͺ�() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.ҽ��ID
    ���ͺ� = mvar���ͺ�
End Property

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

Public Property Let �Ǽ�ʱ��(ByVal vData As Date)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��� = 5
    mvar�Ǽ�ʱ�� = vData
End Property


Public Property Get �Ǽ�ʱ��() As Date
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.���
    �Ǽ�ʱ�� = mvar�Ǽ�ʱ��
End Property

Public Property Let ʹ��ժҪ(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ҩƷ���� = 5
    mvarʹ��ժҪ = MidUni(vData, 1, 200)
End Property


Public Property Get ʹ��ժҪ() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.ҩƷ����
    ʹ��ժҪ = mvarʹ��ժҪ
End Property

Public Property Let ������(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ҩƷ���� = 5
    mvar������ = vData
End Property


Public Property Get ������() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
' X.ҩƷ����
    ������ = mvar������
End Property