VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTabDemoInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'�����ļ�ʾ��
Public ID As Long                           'ʾ��ID
Public �ļ�ID As Long                       'ԭʼ�ļ�����ID
Public ��� As String                       'ʾ�����
Public ���� As String                       'ʾ������
Public ���� As Integer                      'ʾ������:0-����;1-Ƭ��
Public ͨ�ü� As Long                       '0-ȫԺͨ��;1-����ͨ��;2-����ʹ��
Public ����ID As Long                       '��¼����ID
Public ��ԱID As Long                       '��¼��ԱID
Public Sub GetDemoInfo(ByVal lngFileID As Long)
Dim rsTemp As New ADODB.Recordset

    On Error GoTo errHand
    gstrSQL = "Select ID, �ļ�ID, ���, ����, ����, ˵��, ͨ�ü�, ����ID, ��ԱID From ��������Ŀ¼ where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "cTableEPR", lngFileID)  '��ʱ��lngFileId��ʾ�ļ�ʾ����ID
    With Me
        .ID = lngFileID
        If Not rsTemp.EOF Then
            .�ļ�ID = Nvl(rsTemp!�ļ�ID, 0)
            .��� = Nvl(rsTemp!���, 0)
            .���� = Nvl(rsTemp!����)
            .���� = Nvl(rsTemp!����, 0)
            .ͨ�ü� = Nvl(rsTemp!ͨ�ü�, 0)
            .����ID = Nvl(rsTemp!����ID, 0)
            .��ԱID = Nvl(rsTemp!��ԱID, 0)
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
    arrSQL(UBound(arrSQL)) = "Zl_������������_Commit(" & Me.ID & ")"
    
    GetSaveSQL = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
