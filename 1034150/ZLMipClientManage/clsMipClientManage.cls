VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMipClientManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'######################################################################################################################

Public Version As String

Private mfrmActiveForm As Object
Private WithEvents mfrmTable As frmTable
Attribute mfrmTable.VB_VarHelpID = -1
Private WithEvents mfrmItem As frmItem
Attribute mfrmItem.VB_VarHelpID = -1
Private WithEvents mfrmRunControl As frmRunControl
Attribute mfrmRunControl.VB_VarHelpID = -1
'Private WithEvents mfrmSendLog As frmSendLog
'Private WithEvents mfrmReceiveLog As frmReceiveLog
Private WithEvents mfrmMessageMonitor As frmMessageMonitor
Attribute mfrmMessageMonitor.VB_VarHelpID = -1
Private mobjCommunicateProxy As zlMipClient.CommunicateProxy
Private mstrIP As String
Private mintCount As Integer

Public Event AfterClose(ByVal lngModual As Long)
Public Event AfterLoad(ByVal intIndex As Integer, ByVal strContent As String)
Public Event AfterCommunicateChange(ByVal strTitle As String, ByVal strNumber As String)

Public Function CommunicateProxyLogin(ByVal strIp As String, ByVal strPort As String, ByVal strUser As String, ByVal strPwd As String) As Boolean
    '���ӵ�������
    '---------------------------------------------------------------------------------------------------------------------------
    
    On Error GoTo errHand
    
    If mobjCommunicateProxy Is Nothing Then
        Set mobjCommunicateProxy = New zlMipClient.CommunicateProxy
        Call mobjCommunicateProxy.Init(strIp, Val(strPort))
    End If
    
    If mobjCommunicateProxy Is Nothing Then Exit Function
    
    Call mobjCommunicateProxy.Login(strUser, strPwd, "", , True)
    If mobjCommunicateProxy.IsConnected = False Then Exit Function
    
    mstrIP = strIp
    CommunicateProxyLogin = True
    Exit Function
errHand:
    MsgBox Err.Description, vbInformation + vbOKOnly, "��Ϣ��ʾ"
End Function

Public Function CommunicateProxyInstall(ByVal strFilePath As String, Optional ByRef rsErr As ADODB.Recordset) As Boolean
    '��������
    
    On Error GoTo errHand
    
    Dim strParam() As String
    Dim bytAttachData() As Byte
    Dim bytEmpty() As Byte
    Dim strRemoteFile As String
    Dim strLock As String
    Dim strResult As String
    Dim objXML As New clsMipXML
    Dim rs As New ADODB.Recordset
    Dim strError As String
    Dim blnRet As Boolean
    ReDim strParam(1) As String
    Dim intCount As Integer
    Dim intLoop As Integer
    
    If OpenSQLiteFile(strFilePath, bytAttachData) Then
    
        '�ϴ��ļ�
        strRemoteFile = mobjCommunicateProxy.InvokeCustomService("UploadImportFile", strParam, bytAttachData)
        
        RaiseEvent AfterCommunicateChange("��������...", 100)
        '����
        ReDim strParam(1)
        strParam(0) = mstrIP
        strParam(1) = "30"
        strLock = mobjCommunicateProxy.InvokeCustomService("BeginImport", strParam, bytEmpty)
        
        RaiseEvent AfterCommunicateChange("�������÷���������Ŀ¼�µ��ļ�����...", 100)
        '���÷���������Ŀ¼�µ��ļ�����
        ReDim strParam(1)
        strParam(0) = strRemoteFile
        strParam(1) = strLock
        Call mobjCommunicateProxy.InvokeCustomService("SetImportFile", strParam, bytEmpty)
        
        RaiseEvent AfterCommunicateChange("���ڻ�ȡ�����ļ�������Ϣ...", 100)
        '��ȡ�����ļ��е�������Ϣ
        ReDim strParam(0)
        strParam(0) = strLock
        strResult = mobjCommunicateProxy.InvokeCustomService("GetImportFlowInfo", strParam, bytEmpty)
        If objXML.OpenXMLDocument(strResult) Then
            If objXML.GetMultiNodeRecord("UpdateFlow/ID", rs) Then
                
            End If
            Call objXML.CloseXMLDocument
        End If
        
        RaiseEvent AfterCommunicateChange("����ֹͣ����...", 100)

        'ֹͣ����
        ReDim strParam(0)
        If rs.RecordCount > 0 Then
            rs.MoveFirst
            On Error Resume Next
            For intLoop = 0 To rs.RecordCount - 1
                RaiseEvent AfterCommunicateChange("����ֹͣ����...", CStr(CInt(100 * (intLoop + 1) / (rs.RecordCount))))
            
                strParam(0) = rs("node_value").Value
                Call mobjCommunicateProxy.InvokeCustomService("StopFlows", strParam, bytEmpty)
                If Err.Number > 0 Then
                    If rsErr.State = adStateClosed Then
                        rsErr.Open
                    End If
                    mintCount = mintCount + 1
                    rsErr.AddNew
                    rsErr("���").Value = mintCount
                    rsErr("����").Value = Err.Description
                    Err.Clear
                End If
                DoEvents
                rs.MoveNext
            Next
            
            On Error GoTo errHand
        End If

        RaiseEvent AfterCommunicateChange("������֤...", 100)
        '��֤
        ReDim strParam(0)
        strParam(0) = strLock
        strError = mobjCommunicateProxy.InvokeCustomService("ValidateData", strParam, bytEmpty)
        
        RaiseEvent AfterCommunicateChange("���ڵ���...", 100)
        '����
        ReDim strParam(0)
        strParam(0) = strLock
        Call mobjCommunicateProxy.InvokeCustomService("ImportData", strParam, bytEmpty)
        
        '���»�ȡ�����ļ��е�������Ϣ
        ReDim strParam(0)
        strParam(0) = strLock
        strResult = mobjCommunicateProxy.InvokeCustomService("GetImportFlowInfo", strParam, bytEmpty)
        If objXML.OpenXMLDocument(strResult) Then
            If objXML.GetMultiNodeRecord("UpdateFlow/ID", rs) Then
                
            End If
            Call objXML.CloseXMLDocument
        End If
        
        '����
        ReDim strParam(0)
        strParam(0) = strLock
        blnRet = mobjCommunicateProxy.InvokeCustomService("EndImport", strParam, bytEmpty)
        
        'ɾ��
        ReDim strParam(0)
        strParam(0) = strRemoteFile
        Call mobjCommunicateProxy.InvokeCustomService("DeleteImportFile", strParam, bytEmpty)
        
        RaiseEvent AfterCommunicateChange("������������...", 100)

        '��������
        ReDim strParam(0)
        If rs.RecordCount > 0 Then
            rs.MoveFirst
            On Error Resume Next
            For intLoop = 0 To rs.RecordCount - 1
                RaiseEvent AfterCommunicateChange("������������...", CStr(CInt(100 * (intLoop + 1) / (rs.RecordCount))))
            
                strParam(0) = rs("node_value").Value
                Call mobjCommunicateProxy.InvokeCustomService("StartFlows", strParam, bytEmpty)
                If Err.Number > 0 Then
                    If rsErr.State = adStateClosed Then
                        rsErr.Open
                    End If
                    mintCount = mintCount + 1
                    rsErr.AddNew
                    rsErr("���").Value = mintCount
                    rsErr("����").Value = Err.Description
                    Err.Clear
                End If
                DoEvents
                rs.MoveNext
            Next
            On Error GoTo errHand
        End If
    End If
    
    CommunicateProxyInstall = True
    
    Exit Function
errHand:
    MsgBox Err.Description, vbInformation + vbOKOnly, "��Ϣ��ʾ"
End Function

Private Function OpenSQLiteFile(strFile As String, ByRef bytData() As Byte) As Boolean
    '-------------------------------------------------------------
    '���ܣ���ָ����ͼ���ļ����浽ָ���ļ�¼���ֶ���
    '������
    '       strFile��ͼ���ļ�����
    '       rsTable��ͼ�δ洢��¼��
    '       strField��ͼ���ֶ�
    '���أ�
    '-------------------------------------------------------------
    Const conChunkSize As Integer = 10240
    Dim lngFileSize As Long, lngCurSize As Long, lngModSize As Long
    Dim intBolcks As Integer, FileNum, j
    Dim aryChunk() As Byte
    
    On Error GoTo errH
    
    FileNum = FreeFile
    Open strFile For Binary Access Read As FileNum
    lngFileSize = LOF(FileNum)
    ReDim bytData(lngFileSize - 1) As Byte
    Get FileNum, , bytData()
    
    Close FileNum
    OpenSQLiteFile = True
    Exit Function

errH:
    Close FileNum
    OpenSQLiteFile = False

End Function

Public Function CommunicateProxyLogout()
    '�Ͽ�������
    '---------------------------------------------------------------------------------------------------------------------------
    If Not (mobjCommunicateProxy Is Nothing) Then
        Call mobjCommunicateProxy.Logout
        Set mobjCommunicateProxy = Nothing
    End If
    
End Function

Public Function GetForm(ByVal lngModual As Long)
    
    ParamInfo.ģ��� = lngModual
    
'    ParamInfo.ϵͳ�� = lngSys

    Select Case lngModual
    Case 1001  'ҵ����Ϣ����
        If mfrmTable Is Nothing Then Set mfrmTable = New frmTable
        Set GetForm = mfrmTable
    Case 1002  '��Ϣ��Ŀ����
        If mfrmItem Is Nothing Then Set mfrmItem = New frmItem
        Set GetForm = mfrmItem
    Case 1003  '��Ϣվ�����
        If mfrmRunControl Is Nothing Then Set mfrmRunControl = New frmRunControl
        Set GetForm = mfrmRunControl
    Case 1004  '��Ϣ�շ����
        If mfrmMessageMonitor Is Nothing Then Set mfrmMessageMonitor = New frmMessageMonitor
        Set GetForm = mfrmMessageMonitor
    End Select
End Function

Public Function ActiveForm()
    On Error Resume Next
    DoEvents
    mfrmActiveForm.SetFocus
End Function

Public Sub ShowForm(ByVal lngModul As Long, objDataOracle As zlDataOracle.clsDataOracle, frmMain As Object, ByVal strDBUser As String, Optional ByVal blnShowInTaskBar As Boolean)
    '******************************************************************************************************************
    '���ܣ� ����������ָ�����ܣ�����ִ����س���
    '������
    '   lngModul:��Ҫִ�еĹ������
    '   cnMain:����������ݿ�����
    '   frmMain:������
    '   strDBUser:��ǰ��¼���ݿ��û���
    '******************************************************************************************************************
    
    Set gclsDataOracle = objDataOracle
    
    ParamInfo.ϵͳ���� = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
    ParamInfo.��Ʒ���� = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒ����", "����")

    Set gclsBusiness = New clsBusiness
    Call gclsBusiness.InitBusiness(gclsDataOracle, ParamInfo.ϵͳ����, "ZLTOOLS")
    
    UserInfo.�û��� = strDBUser
    
'    If Not GetUserInfo Then
'        MsgBox "��ǰ�û�δ���ö�Ӧ����Ա��Ϣ,����ϵͳ����Ա��ϵ,�ȵ��û���Ȩ���������ã�", vbInformation, ParamInfo.ϵͳ����
'        Exit Sub
'    End If

    Set gfrmMain = frmMain

'    UserInfo.��λ���� = zlComLib.zlRegInfo("��λ����")
    
    zlComLib.SetDbUser UserInfo.�û���
    
    '------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHand
    
    Call InitSysPara
    
    Select Case lngModul
    '------------------------------------------------------------------------------------------------------------------
    Case 1001               'ҵ����Ϣ����
        Call mfrmTable.ShowForm
        DoEvents
        mfrmTable.SetFocus
        Set mfrmActiveForm = mfrmTable
    '------------------------------------------------------------------------------------------------------------------
    Case 1002               '��Ϣ��Ŀ����
        Call mfrmItem.ShowForm
        DoEvents
        mfrmItem.SetFocus
        Set mfrmActiveForm = mfrmItem
    '------------------------------------------------------------------------------------------------------------------
    Case 1003               '���п���
        Call mfrmRunControl.ShowForm
        DoEvents
        mfrmRunControl.SetFocus
        Set mfrmActiveForm = mfrmRunControl
    '------------------------------------------------------------------------------------------------------------------
    Case 1004           '��Ϣ���
        Call mfrmMessageMonitor.ShowForm
        DoEvents
        mfrmMessageMonitor.SetFocus
        Set mfrmActiveForm = mfrmMessageMonitor
    End Select
    
    Exit Sub
    
errHand:

    If zlComLib.ErrCenter = 1 Then
        Resume
    End If
    Call zlComLib.SaveErrLog
End Sub

Public Sub UnloadForm()
    
    If Not (mfrmTable Is Nothing) Then
        Unload mfrmTable
        Set mfrmTable = Nothing
    End If
    
    If Not (mfrmItem Is Nothing) Then
        Unload mfrmItem
        Set mfrmItem = Nothing
    End If
    
    If Not (mfrmRunControl Is Nothing) Then
        Unload mfrmRunControl
        Set mfrmRunControl = Nothing
    End If
        
    If Not (mfrmMessageMonitor Is Nothing) Then
        Unload mfrmMessageMonitor
        Set mfrmMessageMonitor = Nothing
    End If
End Sub

Public Function CloseWindows() As Boolean
    '******************************************************************************************************************
    '����:�ر������Ӵ���
    '******************************************************************************************************************
    Dim frmThis As Form
    
    On Error Resume Next
    
    For Each frmThis In Forms
        Unload frmThis
    Next
    CloseWindows = (Forms.Count = 0)
    
End Function

Private Sub Class_Initialize()
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeBlue
        
    Version = App.Major & "." & App.Minor & "." & App.Revision
    
    Set gfrmPubResource = New frmPubResource
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    
'    If Not (gclsMsgSystem Is Nothing) Then Set gclsMsgSystem = Nothing
    If Not (gclsBusiness Is Nothing) Then Set gclsBusiness = Nothing
    If Not (gfrmPubResource Is Nothing) Then
        Unload gfrmPubResource
        Set gfrmPubResource = Nothing
    End If
    If Not (mobjCommunicateProxy Is Nothing) Then Set mobjCommunicateProxy = Nothing
End Sub

Private Sub mfrmMessageMonitor_AfterClose(ByVal lngModual As Long)
    RaiseEvent AfterClose(lngModual)
End Sub

Private Sub mfrmMessageMonitor_AfterLoad(ByVal intIndex As Integer, ByVal strContent As String)
    RaiseEvent AfterLoad(intIndex, strContent)
End Sub

Private Sub mfrmTable_AfterClose(ByVal lngModual As Long)
    RaiseEvent AfterClose(lngModual)
End Sub

Private Sub mfrmTable_AfterLoad(ByVal intIndex As Integer, ByVal strContent As String)
    RaiseEvent AfterLoad(intIndex, strContent)
End Sub

Private Sub mfrmItem_AfterClose(ByVal lngModual As Long)
    RaiseEvent AfterClose(lngModual)
End Sub

Private Sub mfrmItem_AfterLoad(ByVal intIndex As Integer, ByVal strContent As String)
    RaiseEvent AfterLoad(intIndex, strContent)
End Sub

Private Sub mfrmRunControl_AfterClose(ByVal lngModual As Long)
    RaiseEvent AfterClose(lngModual)
End Sub

Private Sub mfrmRunControl_AfterLoad(ByVal intIndex As Integer, ByVal strContent As String)
    RaiseEvent AfterLoad(intIndex, strContent)
End Sub

