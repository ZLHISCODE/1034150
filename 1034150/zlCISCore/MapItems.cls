VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MapItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"MapItem"
Attribute VB_Ext_KEY = "Member0" ,"MapItem"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'�ֲ����������漯��
Private mCol As Collection

Public Function Add(���� As Byte, ���� As String, ���� As String, �㼯 As String, X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, ���ɫ As Long, ��䷽ʽ As Integer, ����ɫ As Long, ���� As Byte, �߿� As Byte, Optional sKey As String, Optional Oper As Byte) As MapItem
    '�����¶���
    Dim objNewMember As MapItem
    Set objNewMember = New MapItem

    '���ô��뷽��������
    objNewMember.���� = ����
    
    objNewMember.���� = ����
    objNewMember.���� = ����
    
    objNewMember.�㼯 = �㼯
    objNewMember.X1 = X1
    objNewMember.Y1 = Y1
    objNewMember.X2 = X2
    objNewMember.Y2 = Y2
    
    objNewMember.���ɫ = ���ɫ
    objNewMember.��䷽ʽ = ��䷽ʽ
    objNewMember.����ɫ = ����ɫ
    objNewMember.�߿� = �߿�
    objNewMember.���� = ����
    
    objNewMember.Oper = Oper
    objNewMember.Key = sKey
    
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    '�����Ѵ����Ķ���
    Set Add = objNewMember
    Set objNewMember = Nothing
End Function

Public Property Get Item(vntIndexKey As Variant) As MapItem
Attribute Item.VB_UserMemId = 0
    '���ü����е�һ��Ԫ��ʱʹ�á�
    'vntIndexKey �������ϵ�������ؼ��֣�
    '����ΪʲôҪ����Ϊ Variant ��ԭ��
    '�﷨��Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    '���������е�Ԫ����ʱʹ�á��﷨��Debug.Print x.Count
    Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    'ɾ�������е�Ԫ��ʱʹ�á�
    'vntIndexKey ����������ؼ��֣�����ΪʲôҪ����Ϊ Variant ��ԭ��
    '�﷨��x.Remove(xyz)
    mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    '������������ For...Each �﷨ö�ٸü��ϡ�
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    '������󴴽�����
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    '����ֹ���ƻ�����
    Set mCol = Nothing
End Sub