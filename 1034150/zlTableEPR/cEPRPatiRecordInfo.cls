VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTabPatiRecordInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'��¼��Ϣ
Public ID As Long                          '��¼ID
Public ������Դ As PatiFrom            '1-����;2-סԺ;3-����(������ڸ��ﲿ�Ž�����������);4-��첡��
Public ����ID As Long                      '
Public ��ҳID As Long                      'סԺ���˱�����д
Public Ӥ�� As Long                        '�����Ӥ������ʾʱ����������Ϊ**֮Ӥ
Public ����ID As Long                      '�������,��¼��д���β�����ҽ����ʿ���ڿ���
Public �������� As EPRDocType          '��Ӧ�ڲ����ļ�������
Public �ļ�ID As Long                      '�����Ӧ�Ĳ����ļ��Ѿ�ɾ��,���ʽ��ȱʡ������
Public �������� As String                  '
Public ������ As String
Public ����ʱ�� As Date
Public ������ As String
Public ����ʱ�� As Date
Public �鵵�� As String
Public �鵵���� As Date
Public ���ʱ�� As Date                    '��ǩ��ʱ��д,ƽ�������µİ汾����ʱ����д,��ʾ���ڽ�����
Public ���汾 As Integer                 '��Ŀ��ǩ���汾
Public ǩ������ As EPRSignLevel        '1-��д;2-����ҽʦ����;3-����ҽʦ���ġ�סԺ��������Ĳ���ֻ����д������״̬
Public ҽ��id As Long
Public Sub GetPatiRecordInfo(ByVal lngFileID As Long, ByVal blnMoved As Boolean)
Dim rsTemp As New ADODB.Recordset, lngTmp As Long

    On Error GoTo errHand
    gstrSQL = "Select ID, ������Դ, ����ID, ��ҳID, Ӥ��, ����ID, ��������, �ļ�ID, " & _
            " ��������, �鵵��, �鵵����, ������, ����ʱ��, ������, ����ʱ��, " & _
            " ���汾 , ǩ������ , ���ʱ�� From ���Ӳ�����¼ Where ID=[1] "
    If blnMoved Then gstrSQL = Replace(gstrSQL, "���Ӳ�����¼", "H���Ӳ�����¼")
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "cTableEPR", lngFileID)  '��ʱ��lngFileId��ʾ���˵��Ӳ�����ID
    With Me
        .ID = lngFileID
        If Not rsTemp.EOF Then
            .������Դ = Nvl(rsTemp("������Դ"), 1)
            .����ID = Nvl(rsTemp("����ID"), 0)
            .��ҳID = Nvl(rsTemp("��ҳID"), 0)
            .Ӥ�� = Nvl(rsTemp("Ӥ��"), 0)
            .����ID = Nvl(rsTemp("����ID"), 0)
            .�������� = Nvl(rsTemp("��������"), 0)
            .�ļ�ID = Nvl(rsTemp("�ļ�ID"), 0)
            .�������� = Nvl(rsTemp("��������"))
            .������ = Nvl(rsTemp("������"))
            .����ʱ�� = Nvl(rsTemp("����ʱ��"), 0)
            .������ = Nvl(rsTemp("������"))
            .����ʱ�� = Nvl(rsTemp("����ʱ��"), 0)
            .�鵵�� = Nvl(rsTemp("�鵵��"))
            .�鵵���� = Nvl(rsTemp("�鵵����"), 0)
            .���ʱ�� = Nvl(rsTemp("���ʱ��"), 0)
            .���汾 = Nvl(rsTemp("���汾"), 1)
            lngTmp = Nvl(rsTemp("ǩ������"), 0)     '��¼���α༭�ĳ�ʼ�汾״̬
            If lngTmp >= 4 Then
                .ǩ������ = TabSL_����
            ElseIf lngTmp >= 2 Then
                .ǩ������ = TabSL_����
            ElseIf lngTmp >= 1 Then
                .ǩ������ = TabSL_����
            Else
                .ǩ������ = TabSL_�հ�
            End If
        End If
    End With
    Exit Sub
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Function GetSaveSQL(arrSQL As Variant) As Boolean
    On Error GoTo errHand
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "Zl_���Ӳ�����¼_Apply(" & Me.ID & "," & Me.������Դ & "," & Me.����ID & "," & Me.��ҳID & _
                            "," & Me.Ӥ�� & "," & IIf(Me.����ID = 0, UserInfo.����ID, Me.����ID) & "," & Me.�������� & "," & Me.�ļ�ID & ",'" & Me.�������� & _
                            "','" & Me.������ & "','" & UserInfo.���� & "'," & Me.ҽ��id & ")"
    GetSaveSQL = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
