VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExpenseBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"Details"
Option Explicit 'Ҫ���������

Public Key As String

'��������ֵ�ľֲ�����
Private mvarNO As String '�ֲ�����
Private mvar����ID As Long '�ֲ�����
Private mvar��ҳID As Long '�ֲ�����
Private mvar����ID As Long '�ֲ�����
Private mvar����ID As Long '�ֲ�����
Private mvar���� As String '�ֲ�����
Private mvar�Ա� As String '�ֲ�����
Private mvar���� As String '�ֲ�����
Private mvar��ʶ�� As Double '�ֲ�����
Private mvar���� As String '�ֲ�����
Private mvar�ѱ� As String '�ֲ�����
Private mvar�����־ As Byte '�ֲ�����
Private mvar�Ӱ��־ As Byte '�ֲ�����
Private mvarӤ���� As Byte '�ֲ�����
Private mvar�ಡ�˵� As Boolean '�ֲ�����
Private mvar��������ID As Long '�ֲ�����
Private mvar������ As String '�ֲ�����
Private mvar������ As String '�ֲ�����
Private mvar����Ա���� As String '�ֲ�����
Private mvar����Ա��� As String '�ֲ�����
Private mvar����ʱ�� As Date '�ֲ�����
Private mvar�Ǽ�ʱ�� As Date '�ֲ�����
Private mvarDetails As BillDetails '�ֲ�����

Public Property Set Details(ByVal vData As BillDetails)
'������ָ�ɶ���ʱʹ�ã�λ�� Set ������ߡ�
'Syntax: Set x.Details = Form1
    Set mvarDetails = vData
End Property

Public Property Get Details() As BillDetails
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.Details
    Set Details = mvarDetails
End Property

Public Property Let �Ǽ�ʱ��(ByVal vData As Date)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.����ʱ�� = 5
    mvar�Ǽ�ʱ�� = vData
End Property

Public Property Get �Ǽ�ʱ��() As Date
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.����ʱ��
    �Ǽ�ʱ�� = mvar�Ǽ�ʱ��
End Property

Public Property Let ����ʱ��(ByVal vData As Date)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.����ʱ�� = 5
    mvar����ʱ�� = vData
End Property

Public Property Get ����ʱ��() As Date
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.����ʱ��
    ����ʱ�� = mvar����ʱ��
End Property

Public Property Let ����Ա���(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.����Ա��� = 5
    mvar����Ա��� = vData
End Property

Public Property Get ����Ա���() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.����Ա���
    ����Ա��� = mvar����Ա���
End Property

Public Property Let ����Ա����(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.����Ա���� = 5
    mvar����Ա���� = vData
End Property

Public Property Get ����Ա����() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.����Ա����
    ����Ա���� = mvar����Ա����
End Property

Public Property Let ������(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.������ = 5
    mvar������ = vData
End Property

Public Property Get ������() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.������
    ������ = mvar������
End Property

Public Property Let ������(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.������ = 5
    mvar������ = vData
End Property

Public Property Get ������() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.������
    ������ = mvar������
End Property

Public Property Let ��������ID(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��������ID = 5
    mvar��������ID = vData
End Property

Public Property Get ��������ID() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.��������ID
    ��������ID = mvar��������ID
End Property

Public Property Let �ಡ�˵�(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�ಡ�˵� = 5
    mvar�ಡ�˵� = vData
End Property

Public Property Get �ಡ�˵�() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.�ಡ�˵�
    �ಡ�˵� = mvar�ಡ�˵�
End Property

Public Property Let �Ӱ��־(ByVal vData As Byte)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�Ӱ��־ = 5
    mvar�Ӱ��־ = vData
End Property

Public Property Get �Ӱ��־() As Byte
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.�Ӱ��־
    �Ӱ��־ = mvar�Ӱ��־
End Property

Public Property Let �����־(ByVal vData As Byte)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�����־ = 5
    mvar�����־ = vData
End Property

Public Property Get �����־() As Byte
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.�����־
    �����־ = mvar�����־
End Property

Public Property Let �ѱ�(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�ѱ� = 5
    mvar�ѱ� = vData
End Property

Public Property Get �ѱ�() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.�ѱ�
    �ѱ� = mvar�ѱ�
End Property

Public Property Let ����(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvar���� = vData
End Property

Public Property Get ����() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.����
    ���� = mvar����
End Property

Public Property Let ��ʶ��(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��ʶ�� = 5
    mvar��ʶ�� = vData
End Property

Public Property Get ��ʶ��() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.��ʶ��
    ��ʶ�� = mvar��ʶ��
End Property

Public Property Let ����(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvar���� = vData
End Property

Public Property Get ����() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.����
    ���� = mvar����
End Property

Public Property Let �Ա�(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�Ա� = 5
    mvar�Ա� = vData
End Property

Public Property Get �Ա�() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.�Ա�
    �Ա� = mvar�Ա�
End Property

Public Property Let ����(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvar���� = vData
End Property

Public Property Get ����() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.����
    ���� = mvar����
End Property

Public Property Let ��ҳID(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��ҳID = 5
    mvar��ҳID = vData
End Property

Public Property Get ��ҳID() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.��ҳID
    ��ҳID = mvar��ҳID
End Property

Public Property Let ����ID(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.����ID = 5
    mvar����ID = vData
End Property

Public Property Get ����ID() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.����ID
    ����ID = mvar����ID
End Property

Public Property Let NO(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.NO = 5
    mvarNO = vData
End Property

Public Property Get NO() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.NO
    NO = mvarNO
End Property

Private Sub Class_Initialize()
    Set Me.Details = New BillDetails
End Sub

Private Sub Class_Terminate()
    Set Me.Details = Nothing
End Sub

Public Property Let ����ID(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.����ID = 5
    mvar����ID = vData
End Property

Public Property Get ����ID() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.����ID
    ����ID = mvar����ID
End Property

Public Property Let ����ID(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.����ID = 5
    mvar����ID = vData
End Property

Public Property Get ����ID() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.����ID
    ����ID = mvar����ID
End Property

Public Property Let Ӥ����(ByVal vData As Byte)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Ӥ���� = 5
    mvarӤ���� = vData
End Property

Public Property Get Ӥ����() As Byte
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.Ӥ����
    Ӥ���� = mvarӤ����
End Property
