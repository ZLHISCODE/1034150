VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsQueueMsgCenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const M_STR_ROOT_NODE As String = "message"

Private WithEvents mobjMipModuleEvent As clsMipModuleEvent
Attribute mobjMipModuleEvent.VB_VarHelpID = -1
Private mobjMsgProcess As Object    'zl9ComLib.clsMipModule
Private mobjXml As Object           'zl9ComLib.clsXML

Private mobjComLib As Object        'zl9ComLib.clsComLib

Private mblnMsgOpened As Boolean    '�Ƿ����Ϣ����
Private mlngBusineesType As Long    'ҵ������
Private mlngModule As Long          'ģ���
Private mstrComputerName As String  '�������
Private mblnDebugMsg As Boolean     '������Ϣ����

Public Event OnRecevieMsg(ByVal strMsgItemIdentity As String, ByVal strXmlContext As String, rsData As ADODB.Recordset)
Public Event OnOpenModule(ByVal lngSystem As Long, ByVal lngModule As Long)
Public Event OnOpenLink(ByVal strMsgKey As String, ByVal strLinkPara As String)

Private Type TGUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

'����GUID
Private Declare Function CoCreateGuid Lib "ole32.dll" (pguid As TGUID) As Long

Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Function GetSendTag() As String
'��ȡ���ͱ��
    Dim strXml As String
    
    
    strXml = "<send_computer>" & mstrComputerName & "</send_computer>" & vbCrLf & _
             "<send_program>" & mlngModule & "</send_program>"
             
    GetSendTag = strXml
End Function

Private Function SendMsg(ByVal lngQueueId As Long, ByVal strMsgTag As String) As Boolean
'������Ϣ
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    Dim lngStartTime As Long
    
    SendMsg = False
    
    If Not mblnMsgOpened Then Exit Function
    
    strSql = "select zl_�Ŷ���Ϣ_XML���ݻ�ȡ([1], [2]) as XML from dual"
    Set rsData = mobjComLib.zlDatabase.OpenSQLRecord(strSql, "��ѯ�Ŷ���Ϣ�����ı�", lngQueueId, strMsgTag)
    
    If rsData.RecordCount <= 0 Then Exit Function
    
    If mblnDebugMsg Then
        lngStartTime = GetTickCount
    End If
    
    SendMsg = mobjMsgProcess.CommitMessage(strMsgTag, Nvl(rsData!XML) & GetSendTag)
    
    If mblnDebugMsg Then
        OutputDebugString "��Ϣ [" & strMsgTag & "] ����ʱ��Ϊ" & GetTickCount - lngStartTime & "���롣"
    End If
    
    If SendMsg = False Then
        Debug.Print "��Ϣ" & strMsgTag & "����ʧ�ܡ�"
    End If
End Function

Public Function SendMsg_LineQueue(ByVal lngQueueId As Long) As Boolean
'�����Ŷ���Ϣ
    SendMsg_LineQueue = SendMsg(lngQueueId, G_STR_MSG_QUEUE_001)
End Function

Public Function SendMsg_CompleteQueue(ByVal lngQueueId As Long) As Boolean
'���������Ϣ
    SendMsg_CompleteQueue = SendMsg(lngQueueId, G_STR_MSG_QUEUE_002)
End Function

Public Function SendMsg_StateSync(ByVal lngQueueId As Long) As Boolean
'״̬ͬ����Ϣ
    SendMsg_StateSync = SendMsg(lngQueueId, G_STR_MSG_QUEUE_003)
End Function


Public Function SendMsg_VoicePlay(ByVal lngQueueId As Long) As Boolean
'����������Ϣ
    SendMsg_VoicePlay = SendMsg(lngQueueId, G_STR_MSG_QUEUE_004)
End Function

Public Sub ConfigMsgBusinessType(ByVal lngBusinessType As Long)
'��ʼ����Ϣ��������
    mlngBusineesType = lngBusinessType
End Sub

Public Sub OpenMsgCenter(ByVal lngSys As Long, ByVal lngModule As Long, _
    ByVal lngBusineesType As Long, Optional ByVal strPrivs As String = "")
'����Ϣ����
    Dim blnInit As Boolean
    Dim blnAdd As Boolean
    
    mlngBusineesType = lngBusineesType
    mlngModule = lngModule
    
    Set mobjMsgProcess = CreateObject("zl9ComLib.clsMipModule")

    blnInit = mobjMsgProcess.InitMessage(lngSys, lngModule, strPrivs)
    blnAdd = gobjComLib.AddMipModule(mobjMsgProcess)
    
    Set mobjMipModuleEvent = New clsMipModuleEvent
    Set mobjMsgProcess.receiveobject = mobjMipModuleEvent
    
    mstrComputerName = AnalyseComputer & "(" & GetGUID & ")"
    
    If blnInit = True And blnAdd = True Then
        mblnMsgOpened = True
    Else
        mblnMsgOpened = False
    End If
    
    mblnDebugMsg = Val(GetSetting("ZLSOFT", "����ģ��", "��Ϣ����", 0))
End Sub

Public Function GetGUID() As String
'��ȡGUID��
    Dim udtGuid As TGUID
    
    If (CoCreateGuid(udtGuid) = 0) Then
        GetGUID = _
        String(8 - Len(Hex$(udtGuid.Data1)), "0") & Hex$(udtGuid.Data1) & _
        String(4 - Len(Hex$(udtGuid.Data2)), "0") & Hex$(udtGuid.Data2) & _
        String(4 - Len(Hex$(udtGuid.Data3)), "0") & Hex$(udtGuid.Data3) & _
        IIf((udtGuid.Data4(0) < &H10), "0", "") & Hex$(udtGuid.Data4(0)) & _
        IIf((udtGuid.Data4(1) < &H10), "0", "") & Hex$(udtGuid.Data4(1)) & _
        IIf((udtGuid.Data4(2) < &H10), "0", "") & Hex$(udtGuid.Data4(2)) & _
        IIf((udtGuid.Data4(3) < &H10), "0", "") & Hex$(udtGuid.Data4(3)) & _
        IIf((udtGuid.Data4(4) < &H10), "0", "") & Hex$(udtGuid.Data4(4)) & _
        IIf((udtGuid.Data4(5) < &H10), "0", "") & Hex$(udtGuid.Data4(5)) & _
        IIf((udtGuid.Data4(6) < &H10), "0", "") & Hex$(udtGuid.Data4(6)) & _
        IIf((udtGuid.Data4(7) < &H10), "0", "") & Hex$(udtGuid.Data4(7))
    End If
End Function

Public Sub CloseMsgCenter()
'�ر���Ϣ����
    If Not mblnMsgOpened Then Exit Sub
    
    Call gobjComLib.DelMipModule(mobjMsgProcess)
    
    Call mobjMsgProcess.CloseMessage
    Set mobjMsgProcess = Nothing
    Set mobjMipModuleEvent = Nothing
End Sub

Private Sub Class_Initialize()
'���ʼ��
    mblnMsgOpened = False
    
    Set mobjXml = CreateObject("zl9ComLib.clsXML")
End Sub

Public Sub setComLib(objComLib As Object)
    Set mobjComLib = objComLib
End Sub

Private Sub Class_Terminate()
'�ͷ�xml��������
    Set mobjXml = Nothing
    Set mobjComLib = Nothing
End Sub

Private Sub mobjMsgProcess_OpenLink(ByVal strMsgKey As String, ByVal strLinkPara As String)
'��ʾ�����Ӵ��¼�
On Error Resume Next
    RaiseEvent OnOpenLink(strMsgKey, strLinkPara)
End Sub

Private Sub mobjMsgProcess_OpenModule(ByVal lngSystem As Long, ByVal lngModule As Long)
'ģ����¼�
On Error Resume Next
    RaiseEvent OnOpenModule(lngSystem, lngModule)
End Sub

Private Sub mobjMsgProcess_ReceiveMessage(ByVal strMsgItemIdentity As String, ByVal strMsgContent As String)
'��Ϣ�����¼�
On Error GoTo errHandle
    Dim rsData As ADODB.Recordset
    Dim strValue As String
    
    '�����Ƿ��Ŷ���ص���Ϣ��������ǣ����˳�����
    If strMsgItemIdentity <> G_STR_MSG_QUEUE_001 _
        And strMsgItemIdentity <> G_STR_MSG_QUEUE_002 _
        And strMsgItemIdentity <> G_STR_MSG_QUEUE_003 _
        And strMsgItemIdentity <> G_STR_MSG_QUEUE_004 Then
        Exit Sub
    End If
    
    Call mobjXml.OpenXMLDocument(strMsgContent)
    
    '�ж�ҵ�������Ƿ���ͬ�������Ϣ�е�ҵ�������뵱ǰ�����ҵ�����Ͳ�ͬ���򲻽�����Ϣ����
    If Not mobjXml.GetSingleNodeValue("business_type", strValue) Then
        Debug.Print "��ȡ�Ŷ�ҵ������ʱʧ�ܣ���ֹ��Ϣ������"
        Exit Sub
    End If
    
    If Val(strValue) <> mlngBusineesType Then
        Debug.Print "��ȡ��Ϣ��ҵ�������뵱ǰҵ�����Ͳ�ͬ��������Ϣ������"
        Exit Sub
    End If
    
    '��ͬ���󷢹�������Ϣ�����д���
    If Not mobjXml.GetSingleNodeValue("send_instance", strValue) Then
        Debug.Print "��ȡ�Ŷ�ҵ����Դվ��ʱʧ�ܣ���ֹ��Ϣ������"
        Exit Sub
    End If

    If strValue = mobjMsgProcess.Instance Then
        '��ȡ��Ϣ�Ķ����뵱ǰ������ͬ��������Ϣ����
        Debug.Print "��ȡ��Ϣ�Ķ����뵱ǰ������ͬ��������Ϣ������"
        Exit Sub
    End If
    
    If Not mobjXml.GetAllNodeRecord(rsData) Then
        Debug.Print "����GetAllNodeRecord����ʧ�ܣ���ֹ��Ϣ������"
        Exit Sub
    End If
    
    If rsData.RecordCount <= 0 Then
        Debug.Print "δ��ȡ����Ч�Ľڵ����ݣ���ֹ��Ϣ������"
        Exit Sub
    End If
    
    '�����¼�������Ϣ
    RaiseEvent OnRecevieMsg(strMsgItemIdentity, strMsgContent, rsData)
    
Exit Sub
errHandle:
    Debug.Print Err.Description
End Sub

Private Sub mclsMipModuleEvent_ResponseEvent(ByVal strEventName As String, ByVal objCollEventPara As Collection)
    Select Case strEventName
        Case "ReceiveMessage"
            Call mobjMsgProcess_ReceiveMessage(objCollEventPara(1), objCollEventPara(2))
        Case "OpenModule"
            Call mobjMsgProcess_OpenModule(objCollEventPara(1), objCollEventPara(2))
        Case "OpenLink"
            Call mobjMsgProcess_OpenLink(objCollEventPara(1), objCollEventPara(2))
        
    End Select
End Sub