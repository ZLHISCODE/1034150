VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsQueueOperation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'����״̬
Public Enum TQueueState
    qsPlace = -1        'ռλ�У����������ʱ��Ϊռλ״̬
    qsQueueing = 0      '�Ŷ��У�ִ�п�ʼ�ŶӺ󣬽����״̬
    qsCalling = 1       '�����У���ǰ���ں��е�����
    qsAbstain = 2       '�����ţ��������е�����
    qsPause = 3         '����ͣ����ͣ���е�����
    qsComplete = 4      '����ɣ���ɺ��е�����
    qs_BroadCast = 5    '�������㲥
    qs_ReDiagnose = 6   '����������
    qsCalled = 7        '�Ѻ��У���ִ�к��е����ǵ�ǰ���е�����
    qsDiagnose = 8      '������
    qsWaitCall = 9      '������
End Enum

Private mobjComLib As Object                'zl9ComLib.clsComLib
Private mobjVoice As Object                 '�������ж���

Private mstrVoiceType           As String   '΢��������������
Private mblnIsPlayHintSound     As Boolean  '����ǰ������ʾ��
Private mlngPlayCount           As Long     '���д��� Ĭ��Ϊ1
Private mlngPlaySpeed           As Long     '���������ٶ�
Private mlngPlayTimeLength      As Long     '����ʱ�䳤�ȣ�Ĭ��10��
Private mstrPlayStation         As String   '����������վ����,Ĭ�ϵ�ǰվ��
Private mstrLocalStation        As String   '����վ����
Private mlngWorkType            As Long     'ҵ������
Private mstrCusVoiceScript      As String   '�Զ����������Žű�
Private mblnUseVbsPlay          As Boolean
Private mblnAbortState          As Boolean

'�������ŵĺ���
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private Function StartCall(ByVal strCallingContext As String) As Boolean
'lngQueueId������id
'strCallingContext:��������

'���ܣ���ɺ���
    Dim dtStart                   As Date
    Dim intCurCallCount         As Long
    
    StartCall = False
    If strCallingContext = "" Then Exit Function
                        
    intCurCallCount = 0
    While (intCurCallCount < mlngPlayCount)
        '�����ǿ�н�����������ֱ���˳�
        If mblnAbortState Then
            StartCall = True
            Exit Function
        End If
            
        dtStart = Timer
        
        '΢����������
        Call MSSoundPlay(strCallingContext, mlngPlaySpeed)
        
        intCurCallCount = intCurCallCount + 1
                                        
        Do While Timer < dtStart + mlngPlayTimeLength
            Call Sleep(15)
            
            If mblnAbortState Then
                StartCall = True
                Exit Function
            End If
            
            DoEvents
        Loop
    Wend
    
    StartCall = True
End Function

Public Sub StopVoice()
'blnWaitFinish���Ƿ�ȴ���ֹ���
'���ܣ���ֹ����
    If mblnAbortState = True Then Exit Sub
    
    mblnAbortState = True
    
    'ֹͣ΢����������
    If Not (mobjVoice Is Nothing) Then
        Call mobjVoice.Speak("", 2)
    End If
End Sub

Private Sub PlayHintSound()
'������ʾ����
    On Error Resume Next
    
    '�÷�����Ҫ������΢���������ܽ��в���
    If mobjVoice Is Nothing Then
        Set mobjVoice = CreateObject("SAPI.SpVoice")
    End If
    
    Set mobjVoice.Voice = mobjVoice.GetVoices("Name=" & mstrVoiceType).Item(0)
    mobjVoice.Speak "DingDong.wav", 4
End Sub

Private Sub MSSoundPlay(ByVal strConnetxt As String, ByVal lngSoundSpeed As Long)
'strConnetxt:����Ϊ�������ַ���
'lngSoundSpeed:�����ٶ�

'���ܣ������ı�����
    On Error Resume Next
    
    '�÷�����Ҫ������΢���������ܽ��в���
    If mobjVoice Is Nothing Then
        Set mobjVoice = CreateObject("SAPI.SpVoice")
    End If
    
    Set mobjVoice.Voice = mobjVoice.GetVoices("Name=" & mstrVoiceType).Item(0)
    mobjVoice.Rate = lngSoundSpeed   '�ٶ�:-10,10  0
    mobjVoice.Volume = 100 '����:0,100   100
    mobjVoice.Speak strConnetxt, 1
End Sub

Public Function GetQueueState(ByVal lngQueueId As Long) As Long
'lngQueueID: �Ŷ�ID
'-1-ռλ�У�0-�Ŷ��У�1-�����У�2-�����ţ�3-����ͣ��4-��ɾ���,5-���ã�6-���ã�7-�Ѻ��У�
'���ܣ���ȡ�Ŷ�״̬��ʧ�ܷ���-2
On Error GoTo ErrorHand
    Dim strSql              As String
    Dim rsQueue             As ADODB.Recordset
        
    GetQueueState = -2
    
    strSql = "select �Ŷ�״̬ from �ŶӽкŶ��� where id=[1] and ҵ������=[2]"
    Set rsQueue = mobjComLib.zlDatabase.OpenSQLRecord(strSql, "�Ŷ�״̬", lngQueueId, mlngWorkType)
    
    If rsQueue.RecordCount > 0 Then GetQueueState = Nvl(rsQueue!�Ŷ�״̬)
    
    Exit Function
ErrorHand:
    GetQueueState = -2
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function

Public Sub DebugMsg(ByVal strClassName, strMethod, strParameter As String, Optional strExceptionMsg As String = "")
    If False Then
        Call OutputDebugString(Now & ">> [���ù��̣�" & strClassName & "." & strMethod & "]  [�������ݣ�" & strParameter & "]  " & _
                                IIf(Trim(strExceptionMsg) <> "", "[�쳣��Ϣ��" & strExceptionMsg & "]", ""))
    End If
End Sub

Public Sub DelVoiceData(ByVal lngVoiceId As Long)
'ɾ����������
    Dim strSql As String
    
    strSql = "ZL_�Ŷ���������_DELETE(" & lngVoiceId & ")"
    mobjComLib.zlDatabase.ExecuteProcedure strSql, "�����������"
End Sub

Public Function PlayCustomVoice(ByVal lngVoiceId As Long, ByVal blnAutoDelVoice As Boolean, _
                                Optional ByVal strVoiceContext As String = "") As Boolean
'�����Զ������������
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    Dim strCurVoice As String
    
    PlayCustomVoice = False
    
    strCurVoice = strVoiceContext
    If Trim(strVoiceContext) = "" Then
        strSql = "select �������� from �Ŷ���������  where ID=[1]"
        Set rsData = mobjComLib.zlDatabase.OpenSQLRecord(strSql, "��ѯ��������", lngVoiceId)
        
        If rsData.RecordCount <= 0 Then Exit Function
        
        strCurVoice = Nvl(rsData!��������)
    End If
    
    '�����Զ�����к�Ͳ�����������
    If mblnUseVbsPlay = True Then
        If Trim(mstrCusVoiceScript) = "" Then
            Call DebugMsg("clsQueueOperation", "PlayCustomVoice", "mstrCusVoiceScript", "�Զ�����нű�Ϊ�ա�")
            Exit Function
        End If
        
        Call DoScriptCall(lngVoiceId, strCurVoice)
    Else
        If StartCall(strCurVoice) = False Then
            Exit Function
        End If
    End If
    
    '���гɹ���ɾ�����й�������
    If blnAutoDelVoice = True Then
        Call DelVoiceData(lngVoiceId)
    End If
End Function

Public Function PlayQueueVoice(ByVal objMsgCenter As Object, ByVal lngVoiceId As Long, ByVal lngQueueId As Long, ByVal blnAutoDelVoice As Boolean, _
                            Optional ByVal strVoiceContext As String = "") As Boolean
'����ָ�����е���������
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    Dim strCurVoice As String
    
    PlayQueueVoice = False
    
    strCurVoice = strVoiceContext
    If Trim(strVoiceContext) = "" Then
        strSql = "select �������� from �Ŷ���������  where ����ID=[1]"
        Set rsData = mobjComLib.zlDatabase.OpenSQLRecord(strSql, "��ѯ��������", lngQueueId)
        
        If rsData.RecordCount <= 0 Then Exit Function
        
        strCurVoice = Nvl(rsData!��������)
    End If
        
    '����ö��в����ڴ�����״̬���򲻽��к��д���
    If lngQueueId <= 0 Then Exit Function
    
    '���º���״̬Ϊ������
    If GetQueueState(lngQueueId) = qsWaitCall Then
        strSql = "Zl_�ŶӽкŶ���_״̬����(" & lngQueueId & "," & qsCalling & ")"
        Call mobjComLib.zlDatabase.ExecuteProcedure(strSql, "���º���״̬")
        
        If Not objMsgCenter Is Nothing Then Call objMsgCenter.SendMsg_StateSync(lngQueueId)
    End If
    
    '���ź���ǰ����ʾ��
    If mblnIsPlayHintSound Then Call PlayHintSound
    
    '�����Զ�����к�Ͳ�����������
    If mblnUseVbsPlay = True Then
        If Trim(mstrCusVoiceScript) = "" Then
            Call DebugMsg("clsQueueOperation", "PlayQueueVoice", "mstrCusVoiceScript", "�Զ�����нű�Ϊ�ա�")
            Exit Function
        End If
        
        Call DoScriptCall(lngVoiceId, strCurVoice)
    Else
        If StartCall(strCurVoice) = False Then
            Exit Function
        End If
    End If
    
    '���º���״̬Ϊ�Ѻ���,��Ҫ�ж�״̬�Ƿ��Ѿ������ı䣬����ı䣬�򲻽���״̬����
    If GetQueueState(lngQueueId) = qsCalling Then
        strSql = "Zl_�ŶӽкŶ���_״̬����(" & lngQueueId & "," & qsCalled & ")"
        Call mobjComLib.zlDatabase.ExecuteProcedure(strSql, "���º���״̬")
        
        If Not objMsgCenter Is Nothing Then Call objMsgCenter.SendMsg_StateSync(lngQueueId)
    End If
    
    '���гɹ���ɾ�����й�������
    If blnAutoDelVoice = True Then
        Call DelVoiceData(lngVoiceId)
    End If
    
    PlayQueueVoice = True
End Function

Private Sub DoScriptCall(ByVal lngCallId As Long, ByVal strCallContext As String)
'����vbs�ű�ʵ���Զ������
'lngCallId:������id
'strCallContext: ��������

    Dim lngIndex As Long
    Dim objCall As Object
    Dim strMethod As String
    Dim start As Long

    If mobjVoice Is Nothing Then
        Set mobjVoice = CreateObject("SAPI.SpVoice")
    End If

    '�����ű�ִ�ж���
    Set objCall = CreateObject("ScriptControl")
    objCall.timeout = 60000

    strMethod = mstrCusVoiceScript
    lngIndex = InStr(UCase(strMethod), "FUNCTION")

    If lngIndex <= 0 Then
        lngIndex = InStr(UCase(strMethod), "SUB")
        '���ù������Ŀ�ʼλ��
        lngIndex = lngIndex + 4
    Else
        lngIndex = lngIndex + 9
    End If

    strMethod = Mid(strMethod, lngIndex, InStr(strMethod, "(") - lngIndex)

    objCall.Language = "vbscript"

    Call objCall.AddObject("Me", Me, True)
    Call objCall.AddObject("SpVoice", mobjVoice, True)
'    Call objCall.AddObject("zlDatabase", zlDatabase, True)
'    Call objCall.AddObject("zlCommFun", zlCommFun, True)
'    Call objCall.AddObject("zlControl", zlControl, True)

    Call objCall.AddCode(mstrCusVoiceScript)

    start = Timer
    
    Call objCall.Run(Trim(strMethod), lngCallId, strCallContext)

    Do While Timer < start + mlngPlayTimeLength
        Call Sleep(5)

        If mblnAbortState Then Exit Sub

        DoEvents
    Loop
End Sub

Private Function AnalyseComputer() As String
'��ȡ���ؼ��������
    Dim strComputer As String * 256
    Call GetComputerName(strComputer, 255)
    AnalyseComputer = strComputer
    AnalyseComputer = Trim(Replace(AnalyseComputer, Chr(0), ""))
End Function

Private Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'���ܣ��൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
    Nvl = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

'��ʼ��������ֵ
Private Sub Class_Initialize()
    
    '����Ĭ������
    mstrVoiceType = ""
    mblnIsPlayHintSound = False
    mlngPlayCount = 1
    mlngPlaySpeed = 10
    mlngPlayTimeLength = 10
    mstrCusVoiceScript = ""
    mblnUseVbsPlay = False
    mblnAbortState = False
End Sub

Public Sub setComLib(objComLib As Object)
    Set mobjComLib = objComLib
End Sub

Private Sub Class_Terminate()
On Error Resume Next
    Call StopVoice
    Set mobjVoice = Nothing
    Set mobjComLib = Nothing
Err.Clear
End Sub

'���з�ʽ
Public Property Get VoiceType() As String
    VoiceType = mstrVoiceType
End Property

Public Property Let VoiceType(ByVal value As String)
    mstrVoiceType = value
End Property

'���д���
Public Property Get PlayCount() As Integer
    PlayCount = mlngPlayCount
End Property

Public Property Let PlayCount(ByVal value As Integer)
    mlngPlayCount = value
End Property

'���������ٶ�
Public Property Get PlaySpeed() As Long
    PlaySpeed = mlngPlaySpeed
End Property

Public Property Let PlaySpeed(ByVal value As Long)
    mlngPlaySpeed = value
End Property

'���в���ʱ�䳤��
Public Property Get PlayTimeLength() As Long
    PlayTimeLength = mlngPlayTimeLength
End Property

Public Property Let PlayTimeLength(ByVal value As Long)
    mlngPlayTimeLength = value
End Property

'����վ����
Public Property Get LocalStation() As String
    LocalStation = mstrLocalStation
End Property

Public Property Let LocalStation(ByVal value As String)
    mstrLocalStation = value
End Property

'������վ����
Public Property Get PlayStation() As String
    PlayStation = mstrPlayStation
End Property

Public Property Let PlayStation(ByVal value As String)
    mstrPlayStation = value
End Property

'�Զ����������Žű�
Public Property Get CusVoiceScript() As String
    CusVoiceScript = mstrCusVoiceScript
End Property

Public Property Let CusVoiceScript(ByVal value As String)
    mstrCusVoiceScript = value
End Property

'ʹ��vbScript�ű�������������
Public Property Get UseVbsPlay() As Boolean
    UseVbsPlay = mblnUseVbsPlay
End Property

Public Property Let UseVbsPlay(ByVal value As Boolean)
    mblnUseVbsPlay = value
End Property

'����ǰ������ʾ��
Public Property Get IsPlayHintSound() As Boolean
    IsPlayHintSound = mblnIsPlayHintSound
End Property

Public Property Let IsPlayHintSound(value As Boolean)
    mblnIsPlayHintSound = value
End Property

'ҵ������
Public Property Get BusinessType() As Long
    BusinessType = mlngWorkType
End Property

Public Property Let BusinessType(value As Long)
    mlngWorkType = value
End Property