VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCommunity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function Initialize(cnMain As ADODB.Connection) As Boolean
'���ܣ���ʼ�����������ӿڲ��������漰����������ӿڵ���(��Ϊ�Ƕ�̬��ʼ��)
'������cnMain=���ݿ����Ӷ���
'���أ���������õ������ӿڣ��򷵻�True�������ʾû�������κ������ӿ�
    Dim strSQL As String
    
    Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
        
    '��ȡ����Ŀ¼����(���е�)
    On Error GoTo errH
    strSQL = "Select * From ����Ŀ¼ Order by ���"
    Set grsCommunity = zlDatabase.OpenSQLRecord(strSQL, "Initialize")
    If grsCommunity.EOF Then Exit Function
    grsCommunity.Filter = "����=1"
    If grsCommunity.EOF Then Exit Function
    
    Initialize = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Terminate() As Boolean
'���ܣ���ֹ�����ӿڣ���HIS�����������˳�ʱ����
'���أ�ִ�гɹ�/ʧ��
    Dim blnReturn As Boolean, i As Integer
    
    blnReturn = True
    
    '���ù��ľ��������ӿڶ�������ֹ
    For i = 1 To gcolCommunity.Count
        blnReturn = blnReturn And gcolCommunity(i).Terminate
    Next
    
    If Not grsCommunity Is Nothing Then
        If grsCommunity.State = 1 Then grsCommunity.Close
        Set grsCommunity = Nothing
    End If
    Set gobjCommunity = Nothing
    Set gcolCommunity = Nothing
    
    Terminate = blnReturn
End Function

Public Function Setup(ByVal int���� As Integer) As Boolean
'���ܣ���ָ���������ӿڲ�����������
    Set gobjCommunity = GetCommunity(int����)
    
    If Not gobjCommunity Is Nothing Then
        On Error GoTo errH
        Setup = gobjCommunity.Setup(int����)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Identify(ByVal lngSys As Long, ByVal lngModual As Long, _
    int���� As Integer, str������ As String, colInfo As Collection, _
    Optional lng����ID As Long, Optional ByVal lng����ID As Long) As Boolean
'���ܣ�����������֤����֤ʧ�ܷ���False(����ƽ̨�ӿڶ�����������Ϣ��״̬���ж�)
'������
'     lngSys,lngModual(In)�����øýӿڵ�ϵͳ��ģ��ţ���100-��׼��,1111-�Һš�
'     int����(Out)���ò����������������(��1-�ɶ���������)
'         ���ϵͳ��ֻ������һ�������������ֱ��ȡ�ã���������˶�����򵯳�������Ϊȷ����ǰ����������������
'         �ھ���������ӿ��У��ú�������Ϊ����(In)��
'     str������(Out)���ò����ڸ������е�Ψһ��ţ���ABC123
'     colInfo(Out)���������˸�����Ϣ��Ŀ���ݵļ��ϣ�KeyΪ"_�ֶ���"����"_����"��ʧ��ʱ���ؿռ�(��Nothing)
'     lng����ID(In/Out)����ѡ�����������ȷ�����˻��߲���������֤ʱ���롣����ΪOut������
'     lng����ID(In)����ѡ����������������֤ʱʹ�ã���Ӧ"�Һ�ID"��"��ҳID"����ͨ������ģ�鳡��ȷ����
'���أ������֤�ɹ��򷵻�True�����Out����

    If int���� = 0 Then
        int���� = frmSelect.ShowMe
        If int���� = 0 Then Exit Function 'ȡ��ѡ�������ԭ��
    End If
    
    Set gobjCommunity = GetCommunity(int����)
    If Not gobjCommunity Is Nothing Then
        On Error GoTo errH
        Set colInfo = New Collection
        Identify = gobjCommunity.Identify(lngSys, lngModual, int����, str������, colInfo, lng����ID, lng����ID)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetCommunityFunc(ByVal lngSys As Long, ByVal lngModual As Long, ByVal int���� As Integer) As String
'���ܣ������ض��������ض�ģ���еĿɵ��ù��ܣ����ڵ���ģ��������˵���ʾ
'���������Ǵ������
'      lngSys,lngModual=ϵͳ�ţ�ģ���
'      int����=�ض�����������ҽԺ�ɶԶ�������Ĳ��˾���ʱ���������ֲ�ͬ�����Ĺ���
'���أ�����������="���ܺ�1,������1,����˵��1;���ܺ�2,������2,����˵��2;..."
'      ���ܺ�Ϊ>0�����������Բ�����������˵����Ϊ�գ�������ϼ�������������͹���˵���в��ܰ�����ǵļ����
'      ���ĳ��������Ҫ�ڲ��˽���ʱ�Զ�����(Ҳ���ֹ�����),����ǰ�����"Auto:"ǰ׺����"Auto:1,�����鿴,�鿴��������;...",���ֹ���ֻ������һ��
'      ���ؿձ�ʾ���ṩ����
    
    Set gobjCommunity = GetCommunity(int����)
    If Not gobjCommunity Is Nothing Then
        On Error GoTo errH
        GetCommunityFunc = gobjCommunity.GetCommunityFunc(lngSys, lngModual, int����)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CommunityFunc(ByVal lngSys As Long, ByVal lngModual As Long, ByVal int���ܺ� As Integer, ByVal int���� As Integer, ByVal str������ As String, ByVal lng����ID As Long, ByVal lng����ID As Long) As Boolean
'���ܣ��������ܵ��ã�����ƽ̨�ӿ����ṩ����������ҵ���ܡ�
'���������Ǵ������
'      int���ܺţ���"GetCommunityFunc"���ص�Ԥ����һ��������У��ӿ��и����������ٵ���ƽ̨�ӿڵ���ع��ܡ�
'      int������str�����š�lng����ID��lng����ID����ǰ���˵ľ����ʶ�����о���ID��Ӧ"�Һ�ID"��"��ҳID"����ͨ������ģ�鳡��ȷ��
    
    Set gobjCommunity = GetCommunity(int����)
    If Not gobjCommunity Is Nothing Then
        On Error GoTo errH
        CommunityFunc = gobjCommunity.CommunityFunc(lngSys, lngModual, int���ܺ�, int����, str������, lng����ID, lng����ID)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function UpdateInfo(ByVal lngSys As Long, ByVal lngModual As Long, ByVal int���� As Integer, ByVal str������ As String, ByVal lng����ID As Long, ByVal lng����ID As Long) As Boolean
'���ܣ���������ͬ����HIS�е�����������Ϣ�����仯ʱ������ƽ̨�ӿڶ����ĵ���������ݽ���ͬ����
'���������Ǵ������
'˵�����ýӿ�Ӧ��HIS��Ӧ���ܵ�ͬһ�����ݿ������е��ã��������ʧ�ܣ���HIS�л�������
    Set gobjCommunity = GetCommunity(int����)
    If Not gobjCommunity Is Nothing Then
        On Error GoTo errH
        UpdateInfo = gobjCommunity.UpdateInfo(lngSys, lngModual, int����, str������, lng����ID, lng����ID)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ClinicSubmit(ByVal lngSys As Long, ByVal lngModual As Long, ByVal int���� As Integer, ByVal str������ As String, ByVal lng����ID As Long, ByVal lng����ID As Long) As Boolean
'���ܣ�������Ϣ�ύ������������ɱ��ξ���ʱ(����ҽ��վ��ɾ���)�����ýӿڽ����ξ�������е�ҵ��������Ϣ�ϴ������ĵ����⡣
'���������Ǵ������
'˵�����ýӿ�Ӧ��HIS��Ӧ���ܵ�ͬһ�����ݿ������е��ã��������ʧ�ܣ���HIS�л�������
    Set gobjCommunity = GetCommunity(int����)
    If Not gobjCommunity Is Nothing Then
        On Error GoTo errH
        ClinicSubmit = gobjCommunity.ClinicSubmit(lngSys, lngModual, int����, str������, lng����ID, lng����ID)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function