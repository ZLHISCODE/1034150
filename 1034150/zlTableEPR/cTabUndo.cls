VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTabUndo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mvarKey As String   '��"K" & Row_Col��ɣ���ʾ����������
Private mvarType As Byte    '��Ԫ������
Private mvarEkey As String  'Ҫ��Key���������ǵ�Ҫ�أ�Ҳ�����ǻ�������Ҫ�ش�
Private mvarTxt As String   '�̶��ı�����ͨ�ı�����
Private mvarTkey As String  '�ı��� ��������ı���
Private mvarPkey As String  'ͼƬ�ؼ���
Private mvarPmkey As String 'ͼƬ��ǹؼ���
Private mvarRow As Long     '������
Private mvarCol As Long     '������
Private mvarOrigPic As StdPicture           'ԭʼͼƬ
Public Property Let Key(ByVal vData As String)
    mvarKey = vData
    Me.Row = Split(Mid(mvarKey, 2), "_")(0)
    Me.Col = Split(Mid(mvarKey, 2), "_")(1)
End Property

Public Property Get Key() As String
    Key = mvarKey
End Property

Public Property Let CT(ByVal vData As Byte)
    mvarType = vData
End Property

Public Property Get CT() As Byte
    CT = mvarType
End Property
Public Property Let Ekey(ByVal vData As String)
    mvarEkey = vData
End Property

Public Property Get Ekey() As String
    Ekey = mvarEkey
End Property
Public Property Let CTxt(ByVal vData As String)
    mvarTxt = vData
End Property

Public Property Get CTxt() As String
    CTxt = mvarTxt
End Property
Public Property Let Tkey(ByVal vData As String)
    mvarTkey = vData
End Property

Public Property Get Tkey() As String
    Tkey = mvarTkey
End Property
Public Property Let PKey(ByVal vData As String)
    mvarPkey = vData
End Property

Public Property Get PKey() As String
    PKey = mvarPkey
End Property
Public Property Let PmKey(ByVal vData As String)
    mvarPmkey = vData
End Property

Public Property Get PmKey() As String
    PmKey = mvarPmkey
End Property
Public Property Let Row(ByVal vData As Long)
    mvarRow = vData
End Property
Public Property Get Row() As Long
    Row = mvarRow
End Property
Public Property Let Col(ByVal vData As Long)
    mvarCol = vData
End Property
Public Property Get Col() As Long
    Col = mvarCol
End Property
Public Property Set OrigPic(ByRef vData As StdPicture)
    Set mvarOrigPic = vData
End Property

Public Property Let OrigPic(ByRef vData As StdPicture)
    Set mvarOrigPic = vData
End Property

Public Property Get OrigPic() As StdPicture
    Set OrigPic = mvarOrigPic
End Property

Private Sub Class_Initialize()
    Set mvarOrigPic = New StdPicture
End Sub

Private Sub Class_Terminate()
    Set mvarOrigPic = Nothing
End Sub