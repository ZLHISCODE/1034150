VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PageCell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit
Public Key As String
'��������ֵ�ľֲ�����
Private mvarID As Integer '�ֲ�����
Private mvarX As Long '�ֲ�����
Private mvarY As Long '�ֲ�����
Private mvarW As Long '�ֲ�����
Private mvarH As Long '�ֲ�����
Private mvarMaxH As Long '�ֲ�����
Private mvarFixW As Long '�ֲ�����
Private mvarFixH As Long '�ֲ�����
Private mvarRowB As Long '�ֲ�����
Private mvarRowE As Long '�ֲ�����
Private mvarVRowE As Long '�ֲ�����
Private mvarColB As Long '�ֲ�����
Private mvarColE As Long '�ֲ�����
Private mvarCopys As Byte '�ֲ�����

Public Property Let Copys(ByVal vData As Byte)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Copys = 5
    mvarCopys = vData
End Property

Public Property Get Copys() As Byte
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Copys
    Copys = mvarCopys
End Property

Public Property Let ColE(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ColE = 5
    mvarColE = vData
End Property

Public Property Get ColE() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ColE
    ColE = mvarColE
End Property

Public Property Let ColB(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ColB = 5
    mvarColB = vData
End Property

Public Property Get ColB() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ColB
    ColB = mvarColB
End Property

Public Property Let VRowE(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.RowE = 5
    mvarVRowE = vData
End Property

Public Property Get VRowE() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.RowE
    VRowE = mvarVRowE
End Property

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

Public Property Let MaxH(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.MaxH = 5
    mvarMaxH = vData
End Property

Public Property Get MaxH() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.MaxH
    MaxH = mvarMaxH
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

Public Property Let ID(ByVal vData As Integer)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ID = 5
    mvarID = vData
End Property

Public Property Get ID() As Integer
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ID
    ID = mvarID
End Property

Public Property Let FixH(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.H = 5
    mvarFixH = vData
End Property

Public Property Get FixH() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.H
    FixH = mvarFixH
End Property

Public Property Let FixW(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.W = 5
    mvarFixW = vData
End Property

Public Property Get FixW() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.W
    FixW = mvarFixW
End Property