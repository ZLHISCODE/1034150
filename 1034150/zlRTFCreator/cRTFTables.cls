VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRTFTables"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"cRTFTable"
Attribute VB_Ext_KEY = "Member0" ,"cRTFTable"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'�ֲ����������漯��
Private mCol As Collection

Public Function LastTable() As cRTFTable
Attribute LastTable.VB_Description = "��ȡ������ӵı������"
    Set LastTable = mCol(mCol.Count)
End Function

Public Function Add(RowCount As Long, ColCount As Long, _
Optional Alignment As RowAlignTypeEnum = ral0_Ĭ�϶���, _
Optional sKey As String) As Long
Attribute Add.VB_Description = "����һ������"
    '�����¶���
    Dim objNewMember As cRTFTable
    Set objNewMember = New cRTFTable

    objNewMember.RowCount = RowCount
    objNewMember.ColCount = ColCount
    objNewMember.Alignment = Alignment
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    '�����Ѵ����Ķ���
    Add = mCol.Count
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As cRTFTable
Attribute Item.VB_UserMemId = 0
    '���ü����е�һ��Ԫ��ʱʹ�á�
    'vntIndexKey �������ϵ�������ؼ��֣�
    '����ΪʲôҪ����Ϊ Variant ��ԭ��
    '�﷨��Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
Attribute Count.VB_Description = "����������"
    '���������е�Ԫ����ʱʹ�á��﷨��Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
Attribute Remove.VB_Description = "ɾ��ָ������Ԫ�ض���"
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
'  Set mvarcRTFText = Nothing
    '����ֹ���ƻ�����
    Set mCol = Nothing
End Sub
