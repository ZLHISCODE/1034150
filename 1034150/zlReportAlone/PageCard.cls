VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PageCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Key As String
'��������ֵ�ľֲ�����
Private mvarID As Integer '�ֲ�����
Private mvarX As Long '�ֲ�����
Private mvarY As Long '�ֲ�����
Private mvarW As Long '�ֲ�����
Private mvarH As Long '�ֲ�����
Private mvarRowB As Long '�ֲ�����
Private mvarRowE As Long '�ֲ�����
Private mvarCol As Long
Private mvarRow As Long
Private mColItem As Collection

Public Property Let RowE(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.RowE = 5
    mvarRowE = vData
End Property

Public Property Get RowE() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.RowE
    RowE = mvarRowE
End Property

Public Property Let RowB(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.RowB = 5
    mvarRowB = vData
End Property

Public Property Get RowB() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.RowB
    RowB = mvarRowB
End Property

Public Property Let Row(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.RowE = 5
    mvarRow = vData
End Property

Public Property Get Row() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.RowE
    Row = mvarRow
End Property

Public Property Let Col(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.RowB = 5
    mvarCol = vData
End Property

Public Property Get Col() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.RowB
    Col = mvarCol
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

Public Property Let Y(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Y = 5
    mvarY = vData
End Property

Public Property Get Y() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Y
    Y = mvarY
End Property

Public Property Let X(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.X = 5
    mvarX = vData
End Property

Public Property Get X() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.X
    X = mvarX
End Property

Public Property Let id(ByVal vData As Integer)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ID = 5
    mvarID = vData
End Property

Public Property Get id() As Integer
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ID
    id = mvarID
End Property

Public Property Let item(ByVal vData As Collection)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ID = 5
    Set mColItem = vData
End Property

Public Property Get item() As Collection
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ID
    Set item = mColItem
End Property