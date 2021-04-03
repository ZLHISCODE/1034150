VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BillDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"BillDetail"
Attribute VB_Ext_KEY = "Member0" ,"BillDetail"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'�ֲ����������漯��
Private mCol As Collection

Public Function Add(Detail As Detail, �շ�ϸĿID As Long, �շ����� As String, ��� As Integer, _
    �շ���� As String, ���㵥λ As String, ���� As Single, ��׼���� As Currency, ʵ�ս�� As Currency, _
    Ӧ�ս�� As Currency, ���ӱ�־ As Byte, ִ�в���ID As Long, InComes As BillInComes, _
    Optional ������Ŀ�� As Boolean, Optional ���մ���ID As Long, Optional ���ձ��� As String) As BillDetail
    
    '�����¶���
    Dim objNewMember As BillDetail
    Set objNewMember = New BillDetail


    '���ô��뷽��������
    objNewMember.Key = "R" & (��� - 1)
    Set objNewMember.Detail = Detail
    objNewMember.��� = ���
    objNewMember.�շ���� = �շ����
    objNewMember.�շ�ϸĿID = �շ�ϸĿID
    objNewMember.�շ����� = �շ�����
    objNewMember.���㵥λ = ���㵥λ
    objNewMember.���� = ����
    objNewMember.��׼���� = ��׼����
    objNewMember.ʵ�ս�� = ʵ�ս��
    objNewMember.Ӧ�ս�� = Ӧ�ս��
    objNewMember.���ӱ�־ = ���ӱ�־
    objNewMember.ִ�в���ID = ִ�в���ID
    objNewMember.������Ŀ�� = ������Ŀ��
    objNewMember.���մ���ID = ���մ���ID
    objNewMember.���ձ��� = ���ձ���

    Set objNewMember.InComes = InComes
    
    mCol.Add objNewMember, "R" & (��� - 1)
    
    '�����Ѵ����Ķ���
    Set Add = objNewMember
    Set objNewMember = Nothing
End Function

Public Function AddEmpty(��� As Integer) As BillDetail
    '�����¶���
    Dim objNewMember As BillDetail
    Set objNewMember = New BillDetail


    '���ô��뷽��������
    objNewMember.Key = "R" & (��� - 1)
    objNewMember.��� = ���
    
    mCol.Add objNewMember, "R" & (��� - 1)
    
    '�����Ѵ����Ķ���
    Set AddEmpty = objNewMember
    Set objNewMember = Nothing
End Function

Public Property Get Item(vntIndexKey As Variant) As BillDetail
Attribute Item.VB_UserMemId = 0
    '���ü����е�һ��Ԫ��ʱʹ�á�
    'vntIndexKey �������ϵ�������ؼ��֣�
    '����ΪʲôҪ����Ϊ Variant ��ԭ��
    '�﷨��Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    '���������е�Ԫ����ʱʹ�á��﷨��'debug.print x.Count
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