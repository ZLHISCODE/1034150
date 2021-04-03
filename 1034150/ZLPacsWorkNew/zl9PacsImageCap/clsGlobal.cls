VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGlobal"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public lngTestCount As Long



Public Sub InitActiveVideo()
'ʹ�䴥��Init�¼�...

End Sub

''��ȡ���������
'Public Function GetKeyAliasEx(ByVal lngVirtualKey As Long) As String
'    GetKeyAliasEx = mdlPublic.GetKeyAliasEx(lngVirtualKey)
'End Function

''��ȡ��ϼ�����
'Public Function GetKeyAlias(ByVal lngKeyCode As Integer, ByVal lngShift As Integer) As String
'    GetKeyAlias = mdlPublic.GetKeyAlias(lngKeyCode, lngShift)
'End Function


Public Function GetInstanceCount() As Boolean
'��ȡ�Ѿ�������ʵ������
    GetInstanceCount = glngInstanceCount
End Function


Public Sub ForceCreate()
'ǿ�����´�����Դ
    Call CreateGlobalResource
End Sub


Public Sub ForceFree()
'ǿ���ͷ���Դ
    Call FreeGlobalResource
End Sub


Private Sub CreateGlobalResource()
'��ʼ����Ƶ�ɼ�����
    
BUGEX "Global>CreateResource 1", True

    gblnOpenDebug = GetSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "IsOpenDebug", 0)
    Call SaveSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "IsOpenDebug", IIf(gblnOpenDebug, 1, 0))
    
    If gblnOpenDebug Then
BUGEX "Global>CreateResource Open Debug Successed", True
    Else
BUGEX "Global>CreateResource Open Debug Failed", True
    End If
        
    '����������Դ
    If gobjCapturePar Is Nothing Then
        BUGEX "Global>CreateResource gobjCapturePar", True
        Set gobjCapturePar = New clsCaptureParameter
    End If
    
'BUGEX "CreateResource 2", True
'    If gobjComLib Is Nothing Then
'        BUGEX "CreateResource gobjComLib", True
'
'        Set gobjComLib = CreateObject("zl9ComLib.clsComLib")  'New zl9ComLib.clsComLib  '
'    End If
    

BUGEX "Global>CreateResource 2", True
    If gobjVideo Is Nothing Then
        BUGEX "Global>CreateResource gobjVideo", True
        Set gobjVideo = New frmWork_Video
    End If
    
BUGEX "Global>CreateResource 3", True
    If gobjNotifyEvent Is Nothing Then
        BUGEX "Global>CreateResource gobjNotifyEvent", True
        Set gobjNotifyEvent = New clsNotifyEvent
    End If

BUGEX "Global>CreateResource End.", True

End Sub


Private Sub FreeGlobalResource()
'�ͷ���Ƶ�ɼ���Դ
        
BUGEX "Global>FreeGlobalResource 1", True
        '�ͷ��¼�֪ͨ����
        If Not gobjNotifyEvent Is Nothing Then
            Set gobjNotifyEvent = Nothing
        End If
                
BUGEX "Global>FreeGlobalResource 2", True
        'ֻ��ʵ������Ϊһ��ʱ�򣬲��ͷ�ȫ�ֹ�����Դ
        If Not gobjVideo Is Nothing Then
            Unload gobjVideo
            Set gobjVideo = Nothing
        End If
        
BUGEX "Global>FreeGlobalResource 3", True
        If Not gobjZOrder Is Nothing Then
            Call gobjZOrder.RemoveAll
            Set gobjZOrder = Nothing
        End If
        
BUGEX "Global>FreeGlobalResource 4", True

        '���ò������ö���
        If Not gobjCapturePar Is Nothing Then
            Set gobjCapturePar = Nothing
        End If
        
BUGEX "Global>FreeGlobalResource 5", True
        '�ͷ�oracle����
        If Not gcnVideoOracle Is Nothing Then
            Set gcnVideoOracle = Nothing
        End If
    
BUGEX "Global>FreeGlobalResource End.", True
End Sub



Private Sub Class_Initialize()
'�򿪵�������
BUGEX "Global Initialize 1", True
    Call CreateGlobalResource
    
BUGEX "Global Initialize End", True
End Sub

Private Sub Class_Terminate()
'��֤ȫ�ֶ����Ƿ�ж��
BUGEX "Global Terminate 1", True

    Call FreeGlobalResource
    
BUGEX "Global Terminate End", True
End Sub