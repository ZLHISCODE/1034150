VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOperationAdvice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit 'Ҫ���������

'��������ֵ�ľֲ�����
Private mvar����ID As Long '
Private mvar��ҳID As Long
Private mvar�Һŵ� As String
Private mvar������Դ As Long '��1-���2-סԺ,3-����,4-���
Private mvar������ĿID As Long '
Private mvar������ĿIDs As String '��λ��1;������1,������2|��λ��2;������1,������2
Private mvar������� As Long '0-���ڣ�1-���2-����
Private mvarִ�п���ID As Long '����ִ�п���
Private mvar��������ID As Long '����ִ�п���
Private mvar��ʼʱ�� As Date 'ҽ����ʼִ��ʱ��
Private mvar����ʱ�� As Date '���ŵ�����ʱ��
Private mvar���˿���ID As Long
Private mvar��������ID As Long
Private mvar����ҽ�� As String
Private mvar����ʱ�� As Date
Private mvar������־ As Long '0-��ͨ;1-����;2-��¼(��������Ч)
Private mvar���븽� As String '��ʽ="��Ŀ��1<Split2>0/1(�����)<Split2>Ҫ��ID<Split2>����<Split1>��
Private mvar�������IDs As String
Private mvar�޸�ҽ��ID As Long
Private mvar���� As String '
Private mvar������λ As String

'------------------------------------------------------------------------
Public Property Let ����ID(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��� = 5
    mvar����ID = vData
End Property

Public Property Get ����ID() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.���
    ����ID = mvar����ID
End Property

Public Property Let ��ҳID(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��� = 5
    mvar��ҳID = vData
End Property

Public Property Get ��ҳID() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.���
    ��ҳID = mvar��ҳID
End Property


Public Property Let �Һŵ�(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��� = 5
    mvar�Һŵ� = vData
End Property

Public Property Get �Һŵ�() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.���
    �Һŵ� = mvar�Һŵ�
End Property

Public Property Let ������Դ(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��� = 5
    mvar������Դ = vData
End Property

Public Property Get ������Դ() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.���
    ������Դ = mvar������Դ
End Property

 
Public Property Let ������ĿID(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��� = 5
    mvar������ĿID = vData
End Property

Public Property Get ������ĿID() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.���
    ������ĿID = mvar������ĿID
End Property

Public Property Let ������ĿIDs(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��� = 5
    mvar������ĿIDs = vData
End Property

Public Property Get ������ĿIDs() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.���
    ������ĿIDs = mvar������ĿIDs
End Property

Public Property Let ���븽�(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��� = 5
    mvar���븽� = vData
End Property

Public Property Get ���븽�() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.���
    ���븽� = mvar���븽�
End Property

Public Property Let �������(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��� = 5
    mvar������� = vData
End Property

Public Property Get �������() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.���
    ������� = mvar�������
End Property

Public Property Let ִ�п���ID(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��� = 5
    mvarִ�п���ID = vData
End Property

Public Property Get ִ�п���ID() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.���
    ִ�п���ID = mvarִ�п���ID
End Property

Public Property Let ��������ID(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��� = 5
    mvar��������ID = vData
End Property

Public Property Get ��������ID() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.���
    ��������ID = mvar��������ID
End Property


Public Property Let ���˿���ID(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��� = 5
    mvar���˿���ID = vData
End Property

Public Property Get ���˿���ID() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.���
    ���˿���ID = mvar���˿���ID
End Property

Public Property Let ��������ID(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��� = 5
    mvar��������ID = vData
End Property

Public Property Get ��������ID() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.���
    ��������ID = mvar��������ID
End Property

Public Property Let ������־(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��� = 5
    mvar������־ = vData
End Property

Public Property Get ������־() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.���
    ������־ = mvar������־
End Property

Public Property Let ��ʼʱ��(ByVal vData As Date)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��� = 5
    mvar��ʼʱ�� = vData
End Property

Public Property Get ��ʼʱ��() As Date
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.���
    ��ʼʱ�� = mvar��ʼʱ��
End Property

Public Property Let ����ʱ��(ByVal vData As Date)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��� = 5
    mvar����ʱ�� = vData
End Property

Public Property Get ����ʱ��() As Date
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.���
    ����ʱ�� = mvar����ʱ��
End Property


Public Property Let ����ʱ��(ByVal vData As Date)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��� = 5
    mvar����ʱ�� = vData
End Property

Public Property Get ����ʱ��() As Date
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.���
    ����ʱ�� = mvar����ʱ��
End Property


Public Property Let ����ҽ��(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��� = 5
    mvar����ҽ�� = vData
End Property

Public Property Get ����ҽ��() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.���
    ����ҽ�� = mvar����ҽ��
End Property

Public Property Let �������IDs(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��� = 5
    mvar�������IDs = vData
End Property

Public Property Get �������IDs() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.���
    �������IDs = mvar�������IDs
End Property

Public Property Let �޸�ҽ��ID(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��� = 5
    mvar�޸�ҽ��ID = vData
End Property

Public Property Get �޸�ҽ��ID() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.���
    �޸�ҽ��ID = mvar�޸�ҽ��ID
End Property


Public Property Let ����(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��� = 5
    mvar���� = vData
End Property

Public Property Get ����() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.���
    ���� = mvar����
End Property

Public Property Let ������λ(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��� = 5
    mvar������λ = vData
End Property

Public Property Get ������λ() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.���
    ������λ = mvar������λ
End Property
