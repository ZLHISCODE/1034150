VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PatiFeeInfor"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'��������ֵ�ľֲ�����
Private mvarδ����� As Double '�ֲ�����
Private mvarԤ����� As Double '�ֲ�����
Private mvar������ As Double '�ֲ�����
Private mvarԤ����� As Double '�ֲ�����
Private mvarʣ��� As Double '�ֲ�����
Public Property Let ʣ���(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ʣ��� = 5
    mvarʣ��� = vData
End Property
Public Property Get ʣ���() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ʣ���
    ʣ��� = mvarʣ���
End Property



Public Property Let Ԥ�����(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Ԥ����� = 5
    mvarԤ����� = vData
End Property


Public Property Get Ԥ�����() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Ԥ�����
    Ԥ����� = mvarԤ�����
End Property



Public Property Let ������(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.������ = 5
    mvar������ = vData
End Property


Public Property Get ������() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.������
    ������ = mvar������
End Property



Public Property Let Ԥ�����(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Ԥ����� = 5
    mvarԤ����� = vData
End Property


Public Property Get Ԥ�����() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Ԥ�����
    Ԥ����� = mvarԤ�����
End Property



Public Property Let δ�����(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.δ����� = 5
    mvarδ����� = vData
End Property


Public Property Get δ�����() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.δ�����
    δ����� = mvarδ�����
End Property


