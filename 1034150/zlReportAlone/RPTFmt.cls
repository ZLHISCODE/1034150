VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RPTFmt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Public Key As String

'��������ֵ�ľֲ�����
Private mvar��� As Byte '�ֲ�����
Private mvar˵�� As String '�ֲ�����
Private mvarW As Long '�ֲ�����
Private mvarH As Long '�ֲ�����
Private mvarֽ�� As Integer '�ֲ�����
Private mvarֽ�� As Byte '�ֲ�����
Private mvar��ֽ̬�� As Boolean '�ֲ�����
Private mvarͼ�� As Long
Private mblnͣ�� As Boolean
Private mstrͣ��ԭ�� As String

Public Property Let ͣ��ԭ��(ByVal vData As String)
    mstrͣ��ԭ�� = vData
End Property

Public Property Get ͣ��ԭ��() As String
    ͣ��ԭ�� = mstrͣ��ԭ��
End Property

Public Property Let ͣ��(ByVal vData As Boolean)
    mblnͣ�� = vData
End Property

Public Property Get ͣ��() As Boolean
    ͣ�� = mblnͣ��
End Property

Public Property Let ͼ��(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ͼ�� = 5
    mvarͼ�� = vData
End Property

Public Property Get ͼ��() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ͼ��
    ͼ�� = mvarͼ��
End Property

Public Property Let ˵��(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.˵�� = 5
    mvar˵�� = vData
End Property

Public Property Get ˵��() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.˵��
    ˵�� = mvar˵��
End Property

Public Property Let ���(ByVal vData As Byte)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��� = 5
    mvar��� = vData
End Property

Public Property Get ���() As Byte
Attribute ���.VB_UserMemId = 0
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.���
    ��� = mvar���
End Property

Public Property Let ֽ��(ByVal vData As Byte)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ֽ�� = 5
    mvarֽ�� = vData
End Property

Public Property Get ֽ��() As Byte
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ֽ��
    ֽ�� = mvarֽ��
End Property

Public Property Let ֽ��(ByVal vData As Integer)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ֽ�� = 5
    mvarֽ�� = vData
End Property

Public Property Get ֽ��() As Integer
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ֽ��
    ֽ�� = mvarֽ��
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

Public Property Let ��ֽ̬��(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��ֽ̬�� = 5
    mvar��ֽ̬�� = vData
End Property

Public Property Get ��ֽ̬��() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.��ֽ̬��
    ��ֽ̬�� = mvar��ֽ̬��
End Property