VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SquareCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'��������ֵ�ľֲ�����
Private mvarobjSquareCard As Object '�ֲ�����
Private mvarobjCurCard As Card
Private mvarobjDefaultCard  As Card
Private mvarblnȱʡ�������� As Boolean '�ֲ�����
Private mvarintȱʡ���ų��� As Integer   '�ֲ�����
Private mvarbln��ȱʡ������   As Boolean

Public Property Let bln��ȱʡ������(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.bln��ȱʡ������ = 5
    mvarbln��ȱʡ������ = vData
End Property
Public Property Get bln��ȱʡ������() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.bln��ȱʡ������
    bln��ȱʡ������ = mvarbln��ȱʡ������
End Property

Public Property Set objDefaultCard(ByVal vData As Card)
'������ָ�ɶ���ʱʹ�ã�λ�� Set ������ߡ�
'Syntax: Set x.objDefaultCard = Form1
    Set mvarobjDefaultCard = vData
End Property


Public Property Get objDefaultCard() As Card
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.objDefaultCard
    Set objDefaultCard = mvarobjDefaultCard
End Property

Public Property Set objCurCard(ByVal vData As Card)
'������ָ�ɶ���ʱʹ�ã�λ�� Set ������ߡ�
'Syntax: Set x.objCurCard = Form1
    Set mvarobjCurCard = vData
End Property


Public Property Get objCurCard() As Card
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.objCurCard
    Set objCurCard = mvarobjCurCard
End Property
 
Public Property Let intȱʡ���ų���(ByVal vData As Integer)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.intҽ�ƿ����� = 5
    mvarintȱʡ���ų��� = vData
End Property


Public Property Get intȱʡ���ų���() As Integer
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.intҽ�ƿ�����
    intȱʡ���ų��� = mvarintȱʡ���ų���
End Property
 
Public Property Let blnȱʡ��������(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.blnȱʡ�������� = 5
    mvarblnȱʡ�������� = vData
End Property
Public Property Get blnȱʡ��������() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.blnȱʡ��������
    blnȱʡ�������� = mvarblnȱʡ��������
End Property
Public Property Set objSquareCard(ByVal vData As Object)
'������ָ�ɶ���ʱʹ�ã�λ�� Set ������ߡ�
'Syntax: Set x.objSquareCard = Form1
    Set mvarobjSquareCard = vData
End Property


Public Property Get objSquareCard() As Object
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.objSquareCard
    Set objSquareCard = mvarobjSquareCard
End Property