VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDataQueue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mobjMsgQueue() As Object     '��Ϣ���ݶ���
Private mintQueueIndex As Integer    '��¼�����е�һ����Ϣ������

Public Function MsgInQueue(objMsg As Object) As Boolean
'------------------------------------------------
'���ܣ���Ϣ��ӣ��Ⱥ򱻴���
'������ strMsg ������Ҫ��ӵ���Ϣ���ݣ���Ϣ�����ɴ洢��ַ���洢�û������洢���룬�洢Ŀ¼������Ŀ¼���洢�������
'���أ�True������ӳɹ���False�������ʧ��
'-----------------------------------------------
    Dim intCount As Integer
    
    On Error GoTo err

    MsgInQueue = False
    
    '������Ϣ���
    intCount = UBound(mobjMsgQueue) + 1
frmMain.BUGEX "MsgInQueue intCount=" & intCount & "    " & objMsg.SubDir
    ReDim Preserve mobjMsgQueue(intCount) As Object
    
    Set mobjMsgQueue(intCount) = objMsg
        
    MsgInQueue = True
    
    Exit Function
err:
    '�������˳�����ʱ��������
End Function

Public Function MsgOutQueue() As Object
'------------------------------------------------
'���ܣ���Ϣ���ӣ����ù��̸��������ӵ���Ϣ
'������
'���أ����س��ӵ���Ϣ����
'-----------------------------------------------
    Dim intCount As Integer
    Dim objMsg As Object
    
    On Error GoTo err
    
    Set MsgOutQueue = Nothing        '��ʼ��Ϊ����Ϣ

    '��Ϣ���Ӵ���
    intCount = UBound(mobjMsgQueue)

    If intCount = 0 Then Exit Function      '����Ϊ�գ����ó���
    
    '�Ӷ�������ȡ��Ϣ
    Set objMsg = mobjMsgQueue(mintQueueIndex)
frmMain.BUGEX "mintQueueIndex= " & mintQueueIndex & "    intCount= " & intCount & "    " & objMsg.SubDir
    '��������ָ��
    mintQueueIndex = mintQueueIndex + 1
    
    If mintQueueIndex > intCount Then
        '�����ǰȡ�������Ƕ����е����һ����Ϣ���������Ϣ����
        ReDim Preserve mobjMsgQueue(0) As Object

        mintQueueIndex = 1
    End If
    
    Set MsgOutQueue = objMsg
    
    Exit Function
err:
    '�������˳�����ʱ��������
frmMain.BUGEX "MsgOutQueue Err " & err.Description
End Function

Public Function QueueCount() As Integer
    QueueCount = UBound(mobjMsgQueue)
End Function

Private Sub Class_Initialize()
    ReDim mobjMsgQueue(0) As Object
    mintQueueIndex = 1
End Sub