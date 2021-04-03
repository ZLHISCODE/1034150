VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RPTRelations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'�ֲ����������漯��
Private mCol As Collection

Public Function Add(��������ID As Long, ������ As String, ����ֵ��Դ As String, ������������ As String) As RPTRelation
'KEy="_" & ������
    '�����¶���
    Dim objNewMember As RPTRelation
    Set objNewMember = New RPTRelation

    '���ô��뷽��������
    objNewMember.Key = "_" & ������
    objNewMember.��������ID = ��������ID
    objNewMember.������ = ������
    objNewMember.����ֵ��Դ = ����ֵ��Դ
    objNewMember.������������ = ������������
    
    mCol.Add objNewMember, objNewMember.Key

    '�����Ѵ����Ķ���
    Set Add = objNewMember
    Set objNewMember = Nothing
End Function

Public Property Get Item(vntIndexKey As Variant) As RPTRelation
    '���ü����е�һ��Ԫ��ʱʹ�á�
    'vntIndexKey �������ϵ�������ؼ��֣�
    '����ΪʲôҪ����Ϊ Variant ��ԭ��
    '�﷨��Set foo = x.Item(xyz) or Set foo = x.Item(5)
    On Error Resume Next
    Set Item = mCol(vntIndexKey)
    If Err.Number > 0 Then Set Item = New RPTRelation
End Property

Public Property Get count() As Long
    '���������е�Ԫ����ʱʹ�á��﷨��Debug.Print x.Count
    count = mCol.count
End Property

Public Sub Remove(vntIndexKey As Variant)
    'ɾ�������е�Ԫ��ʱʹ�á�
    'vntIndexKey ����������ؼ��֣�����ΪʲôҪ����Ϊ Variant ��ԭ��
    '�﷨��x.Remove(xyz)
    mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
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

