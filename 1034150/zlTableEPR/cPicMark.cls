VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTabPicMark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit
'ѡ��|����|����|�㼯|X1|Y1|X2|Y2|���ɫ|��䷽ʽ|����ɫ|����ɫ|����|�߿�|������
Private mvarKey As String           '�ؼ��֣�����ѡ���ɾ���ö���
Private mvar��ID As Long
Private mvarѡ�� As Boolean         '�Ƿ�ѡ�У�����ƽ�Ƹö���
Private mvar���� As Byte            '0-�ı�,1-����,2,����,3-����,4-�����,5-Բ(��Բ), 6-˳����
Private mvar���� As String          '�ı�������
Private mvar�㼯 As String          '���߻���εĵ㼯:2,3;34,25;...
Private mvarX1 As Long              '���Ͻ�����
Private mvarY1 As Long              '
Private mvarX2 As Long              '���½�����
Private mvarY2 As Long              '
Private mvar���ɫ As Long     '
Private mvar��䷽ʽ As Integer     '-1-�����,-2-ʵ��,0-ˮƽ��,1-��ֱ��,2-��б��,3-��б��,4-ˮƽ�ʹ�ֱ��,5-������
Private mvar����ɫ As Long     '
Private mvar����ɫ As Long     '
Private mvar���� As Byte            '0-ʵ��,1-����,2-����,3-�㻮��,4-˫�㻮��
Private mvar�߿� As Byte            '
Private mvar���� As String     '����,�ֺ�,��ɫ,0000����λ�ֱ��ʾ����,б��,����,ɾ����
Private mvar�������� As Boolean
Private mvar��ʼ�� As Long
Private mvar��ֹ�� As Long
Public Property Let Key(ByVal vData As String)
    mvarKey = vData
End Property

Public Property Get Key() As String
    Key = mvarKey
End Property
Public Property Let ��ID(ByVal vData As Long)
    mvar��ID = vData
End Property

Public Property Get ��ID() As Long
    ��ID = mvar��ID
End Property
Public Property Let ѡ��(ByVal vData As Boolean)
    mvarѡ�� = vData
End Property

Public Property Get ѡ��() As Boolean
    ѡ�� = mvarѡ��
End Property
Public Property Let ����(ByVal vData As Byte)
    mvar���� = vData
End Property

Public Property Get ����() As Byte
    ���� = mvar����
End Property
Public Property Let ����(ByVal vData As String)
    '�滻�����ַ������ڴ洢�����ַ��������ӷ�
    mvar���� = Replace(vData, "|", "��")
End Property

Public Property Get ����() As String
    ���� = mvar����
End Property
Public Property Let �㼯(ByVal vData As String)
    mvar�㼯 = vData
End Property

Public Property Get �㼯() As String
    �㼯 = mvar�㼯
End Property
Public Property Let X1(ByVal vData As Long)
    mvarX1 = vData
End Property

Public Property Get X1() As Long
    X1 = mvarX1
End Property
Public Property Let Y1(ByVal vData As Long)
    mvarY1 = vData
End Property

Public Property Get Y1() As Long
    Y1 = mvarY1
End Property
Public Property Let X2(ByVal vData As Long)
    mvarX2 = vData
End Property

Public Property Get X2() As Long
    X2 = mvarX2
End Property
Public Property Let Y2(ByVal vData As Long)
    mvarY2 = vData
End Property

Public Property Get Y2() As Long
    Y2 = mvarY2
End Property
Public Property Let ���ɫ(ByVal vData As Long)
    mvar���ɫ = vData
End Property

Public Property Get ���ɫ() As Long
    ���ɫ = mvar���ɫ
End Property

Public Property Let ��䷽ʽ(ByVal vData As Integer)
    mvar��䷽ʽ = vData
End Property

Public Property Get ��䷽ʽ() As Integer
    ��䷽ʽ = mvar��䷽ʽ
End Property
Public Property Let ����ɫ(ByVal vData As Long)
    mvar����ɫ = vData
End Property

Public Property Get ����ɫ() As Long
    ����ɫ = mvar����ɫ
End Property

Public Property Let ����ɫ(ByVal vData As Long)
    mvar����ɫ = vData
End Property

Public Property Get ����ɫ() As Long
    ����ɫ = mvar����ɫ
End Property
Public Property Let ����(ByVal vData As Byte)
    mvar���� = vData
End Property

Public Property Get ����() As Byte
    ���� = mvar����
End Property
Public Property Let �߿�(ByVal vData As Byte)
    mvar�߿� = vData
End Property

Public Property Get �߿�() As Byte
    �߿� = mvar�߿�
End Property

Public Property Let ����(ByVal vData As String)
    mvar���� = vData
End Property

Public Property Get ����() As String
    ���� = mvar����
End Property
Public Property Let ��������(ByVal vData As Boolean)
    mvar�������� = vData
End Property

Public Property Get ��������() As Boolean
    �������� = mvar��������
End Property
Public Property Let ��ʼ��(ByVal vData As Long)
    mvar��ʼ�� = vData
End Property

Public Property Get ��ʼ��() As Long
    ��ʼ�� = mvar��ʼ��
End Property
Public Property Let ��ֹ��(ByVal vData As Long)
    mvar��ֹ�� = vData
End Property

Public Property Get ��ֹ��() As Long
    ��ֹ�� = mvar��ֹ��
End Property
'################################################################################################################
'## ���ܣ�  ���Ʊ������
'##
'## ������  blnWithKeys :�Ƿ���ͬKeyֵһ����
'################################################################################################################
Public Sub Clone(cNode As cTabPicMark)
''ѡ��|����|����|�㼯|X1|Y1|X2|Y2|���ɫ|��䷽ʽ|����ɫ|����ɫ|����|�߿�|������
    With cNode
         .ѡ�� = Me.ѡ��
         .���� = Me.����
         .���� = Me.����
         .�㼯 = Me.�㼯
         .X1 = Me.X1
         .Y1 = Me.Y1
         .X2 = Me.X2
         .Y2 = Me.Y2
         .���ɫ = Me.���ɫ
         .��䷽ʽ = Me.��䷽ʽ
         .����ɫ = Me.����ɫ
         .���� = Me.����
         .�߿� = Me.�߿�
         .���� = Me.����
         .����ɫ = Me.����ɫ
         .�������� = Me.��������
         .��ʼ�� = Me.��ʼ��
         .��ֹ�� = Me.��ֹ��
    End With
End Sub

'################################################################################################################
'## ���ܣ�  ��ȡ��ǵ������ַ���
'################################################################################################################
Public Property Get ��������() As String
''ѡ��|����|����|�㼯|X1|Y1|X2|Y2|���ɫ|��䷽ʽ|����ɫ|����ɫ|����|�߿�|������
    Dim strR As String
    strR = strR & IIf(Me.ѡ��, "1", "0") & "|"
    strR = strR & Me.���� & "|"
    strR = strR & IIf(Me.���� = "", " ", Me.����) & "|"
    strR = strR & IIf(Me.�㼯 = "", " ", Me.�㼯) & "|"
    strR = strR & Me.X1 & "|"
    strR = strR & Me.Y1 & "|"
    strR = strR & Me.X2 & "|"
    strR = strR & Me.Y2 & "|"
    strR = strR & Me.���ɫ & "|"
    strR = strR & Me.��䷽ʽ & "|"
    strR = strR & Me.����ɫ & "|"
    strR = strR & Me.����ɫ & "|"
    strR = strR & Me.���� & "|"
    strR = strR & Me.�߿� & "|"
    strR = strR & IIf(Me.���� = "", " ", Me.����)
    �������� = strR
End Property

'################################################################################################################
'## ���ܣ�  ͨ��һ���ַ��������ñ������
'################################################################################################################
Public Property Let ��������(vData As String)
    On Error Resume Next
    Dim T As Variant
    T = Split(vData, "|")
    If UBound(T) > 0 Then
        Me.ѡ�� = IIf(T(0) = 1, True, False)
        Me.���� = T(1)
        Me.���� = Trim(T(2))
        Me.�㼯 = Trim(T(3))
        Me.X1 = T(4)
        Me.Y1 = T(5)
        Me.X2 = T(6)
        Me.Y2 = T(7)
        Me.���ɫ = T(8)
        Me.��䷽ʽ = T(9)
        Me.����ɫ = T(10)
        Me.����ɫ = T(11)
        Me.���� = T(12)
        Me.�߿� = T(13)
        Me.���� = Trim(T(14))
    End If
End Property
