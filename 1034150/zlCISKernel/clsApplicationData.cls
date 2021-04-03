VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsApplicationData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public blnIsModify As Boolean               '�Ƿ��޸�
Public strApplicationPageName As String         '���뵥����
Public lngApplicationPageId As Long             '���뵥Id
Public strRequestTime As String             '����ʱ��

Public lngUpdateAppNo As Long               '������ţ�ֻ�и���ʱ�������������
Public lngUpdateAdviceId As Long            'ҽ��ID��ֻ�и���ʱ������ҽ��ID
Public blnAllowUpdate As Boolean            '�Ƿ��������£�У����ҽ������������


Public blnIsPriority As Boolean             '�Ƿ����
Public blnIsAdditionalRec As Boolean        '�Ƿ�¼
Public lngProjectId As Long                 '������ĿId
Public lngExeType As Long                   'ִ������
Public strStartExeTime As String            '��ʼִ��ʱ��
Public lngExeRoomId As Long                 'ִ�п���
Public strExeRoomName As String             'ִ�п�������
Public lngExeRoomType As Long               'ִ�п�������
Public strPartMethod As String              '��λ����   ��ʽΪ����λ��1;������1,������2|��λ��2;������1,������2
Public strRequestAffix As String            '���븽��   ��ʽΪ������1:��������|����2:��������
Public strRequestAffixCfg As String         '���븽������� ��ʽΪ������1��������У�Ҫ��ID
Public strDiagnoseId As String              '�ٴ����Id,���Idֱ���ö��š�,���ָ�
Public lngRequestRoomId As Long             '�������Id
Public strAbstract As String                'ժҪ���� gclsInsure.GetItemInfo ������ã��ύ����ʱ���ⲿ��ֵ


Private Sub Class_Initialize()
'����Ĭ��ֵ
    blnAllowUpdate = True
    blnIsModify = False
End Sub

Public Function Get���뵥��Ϣ(ByVal lng��Ŀid As Long, ByVal intӦ�ó��� As Integer) As String
'���ܣ���ȡ���뵥ID���뵥����
'���أ���ʽ ���뵥ID<Split>���뵥����
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    On Error GoTo errH
    strSQL = "Select b.Id, b.���� From ��������Ӧ�� A, �����ļ��б� B Where a.������Ŀid =[1] And a.�����ļ�id = b.Id And b.���� = 7 and a.Ӧ�ó���=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "clsApplicationData", lng��Ŀid, intӦ�ó���)
    If Not rsTmp.EOF Then
        Get���뵥��Ϣ = rsTmp!ID & "<Split>" & rsTmp!����
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Get���븽��Ŀ����(ByVal lng�ļ�ID As Long) As String
'���ܣ���ȡ���뵥����������Ϣ
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String
    Dim i As Long
    
    On Error GoTo errH
    strSQL = "select a.�ļ�ID,a.��Ŀ,a.����,a.����,Ҫ��Id,b.������ as Ҫ����, a.����,a.ֻ��  " & _
            " from �������ݸ��� a, ����������Ŀ b  " & _
            " where a.Ҫ��id=b.id(+) and a.�ļ�Id=[1] order by ���� "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "clsApplicationData", lng�ļ�ID)
    For i = 1 To rsTmp.RecordCount
        strTmp = IIF(strTmp = "", "", strTmp & "|") & rsTmp!��Ŀ & ":" & Val(rsTmp!���� & "") & "," & Val(rsTmp!���� & "") & "," & Nvl(rsTmp!Ҫ��ID) & ","
        rsTmp.MoveNext
    Next
    Get���븽��Ŀ���� = strTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function MakePacsData(ByVal lngNo As Long, ByRef objAppPages() As clsApplicationData, Optional ByVal blnMore As Boolean) As ADODB.Recordset
'���ܣ�ͨ����̨��ѯ������֯PACS���뵥����
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim objTmp As clsApplicationData
    Dim strTmp As String
    Dim intӦ�ó��� As Integer
    Dim lng��������ID As Long
    Dim i As Long
    Dim lng���ID As Long
    Dim lng����ID As Long
    Dim rsMore As ADODB.Recordset
    Dim rsOther As ADODB.Recordset
    
    On Error GoTo errH
    If lngNo = 0 Then Exit Function
    strSQL = "select a.id,a.��ҳID,a.�Һŵ�,a.��������ID,a.����ID from ����ҽ����¼ a where a.�������=[1] and a.���ID is null order by a.���"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "clsApplicationData", lngNo)
    
    With rsTmp
        If Not IsNull(!�Һŵ�) Then
            intӦ�ó��� = 1
        Else
            intӦ�ó��� = 2
        End If
        
        If blnMore Then
            strSQL = "select id as ����ID,id as ����ID,id as ��ҳID,id as �Һ�ID,id as ��������ID,id as Ӧ�ó��� from ���ű� where id=0"
            Set rsMore = zlDatabase.OpenSQLRecord(strSQL, "clsApplicationData")
            Set rsMore = zlDatabase.CopyNewRec(rsMore, True)
            
            If intӦ�ó��� = 1 Then
                strSQL = "select a.id from ���˹Һż�¼ a where a.No=[1]"
                Set rsOther = zlDatabase.OpenSQLRecord(strSQL, "clsApplicationData", !�Һŵ� & "")
                lng����ID = Val(rsOther!ID & "")
            End If
            rsMore.AddNew Array("����ID", "����ID", "��ҳID", "Ӧ�ó���", "��������ID"), Array(Val(!����ID & ""), lng����ID, Val(!��ҳID & ""), intӦ�ó���, Val(!��������ID & ""))
        End If
        
        For i = 1 To .RecordCount
            Set objTmp = New clsApplicationData
            Set objTmp = ReadAdviceInf(Val(!ID & ""), intӦ�ó���)
            ReDim Preserve objAppPages(i - 1)
            Set objAppPages(i - 1) = objTmp
            .MoveNext
        Next
    End With
    Set MakePacsData = rsMore
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ReadAdviceInf(ByVal lngAdviceID As Long, ByVal intӦ�ó��� As Integer) As clsApplicationData
'��ȡ��ǰҽ����Ϣ
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim strCurPart As String
    Dim strCurMethod As String
    Dim strTemp As String
    Dim objReturnAppData As New clsApplicationData
    Dim str������� As String
    Dim str���Id As String
    
    Set objReturnAppData = Nothing
    
    '��ȡҽ�������Ϣ
    strSQL = "select a.id,a.���Id,a.Ӥ��,a.������ĿId,a.�걾��λ,a.��鷽��,a.�շ�ϸĿId,a.ҽ��״̬,a.ִ�б��,a.����ʱ��,a.��������Id," & _
            "a.ִ�п���ID,b.����,a.��ʼִ��ʱ��,a.������־,a.ҽ��״̬,a.������� " & _
            " from ����ҽ����¼ a, ���ű� b" & _
            " where a.ִ�п���Id=b.Id(+) and (a.Id=[1] or a.���Id=[1]) order by a.�걾��λ desc, a.��鷽�� desc"
            
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯҽ����λ����", lngAdviceID)
    
    If rsData.RecordCount <= 0 Then Exit Function
    
    objReturnAppData.lngUpdateAdviceId = Val(Nvl(rsData!ID))
    objReturnAppData.lngUpdateAppNo = Val(Nvl(rsData!�������))
    objReturnAppData.blnAllowUpdate = Decode(Val(Nvl(rsData!ҽ��״̬)), -1, True, 1, True, 2, True, False)
    objReturnAppData.blnIsModify = False
    
    objReturnAppData.lngProjectId = Val(Nvl(rsData!������ĿID))
    objReturnAppData.lngExeType = Val(Nvl(rsData!ִ�б��))
    objReturnAppData.lngExeRoomId = Val(Nvl(rsData!ִ�п���ID))
    objReturnAppData.strExeRoomName = Nvl(rsData!����)
    objReturnAppData.strStartExeTime = Nvl(rsData!��ʼִ��ʱ��)
    objReturnAppData.blnIsPriority = IIF(Val(Nvl(rsData!������־)) = 0, False, True)
    objReturnAppData.lngRequestRoomId = Val(Nvl(rsData!��������ID))
    objReturnAppData.strRequestTime = Nvl(rsData!����ʱ��)
    
    str���Id = GetAdviceDiag(Val(Nvl(rsData!ID)), str�������)
    objReturnAppData.strDiagnoseId = str���Id
    
    '��ȡ��λ����
    strCurPart = ""
    strCurMethod = ""
    strTemp = ""      '��λ��1;������1,������2|��λ��2;������1,������2|...'<vbTab>0-����/1-����/2-����  ���ﲻ��Ҫ�ں�������ִ�б��
    
    While Not rsData.EOF
        If Nvl(rsData!�걾��λ) <> strCurPart And strCurPart <> "" Then
            strTemp = strTemp & "|" & strCurPart & ";" & strCurMethod
            strCurMethod = ""
        End If
        
        If Nvl(rsData!��鷽��) <> "" Then
            strCurMethod = strCurMethod & IIF(strCurMethod <> "", "," & Nvl(rsData!��鷽��), Nvl(rsData!��鷽��))
        End If
        
        strCurPart = Nvl(rsData!�걾��λ)
        
        Call rsData.MoveNext
    Wend
    
    If strCurPart <> "" Then
        strTemp = strTemp & "|" & strCurPart & ";" & strCurMethod
    End If
    
    If strTemp <> "" Then strTemp = Mid(strTemp, 2)   '��ȡ","֮��Ĳ�λ�������
    
    objReturnAppData.strPartMethod = strTemp
    
    '��ȡ��Ӧ�����뵥Id
    strSQL = "select a.�����ļ�Id, b.���� from ��������Ӧ�� a, �����ļ��б� b where a.�����ļ�Id=b.Id and ������ĿId=[1] and Ӧ�ó���=[2]"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯҽ����Ӧ���뵥", objReturnAppData.lngProjectId, intӦ�ó���)
    
    If rsData.RecordCount > 0 Then
        objReturnAppData.lngApplicationPageId = Val(Nvl(rsData!�����ļ�ID))
        objReturnAppData.strApplicationPageName = Nvl(rsData!����) & "����"
    End If
    
    '��ȡ���븽��
    strSQL = "select ��Ŀ,���� from ����ҽ������ where ҽ��Id=[1]"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯҽ�����븽��", lngAdviceID)
    
    strTemp = ""
    If rsData.RecordCount > 0 Then
        While Not rsData.EOF
            If strTemp <> "" Then strTemp = strTemp & "|"
            
            strTemp = strTemp & Nvl(rsData!��Ŀ) & ":" & Nvl(rsData!����)
            
            Call rsData.MoveNext
        Wend
    End If
    
    objReturnAppData.strRequestAffix = strTemp
    
    Set ReadAdviceInf = objReturnAppData
End Function
