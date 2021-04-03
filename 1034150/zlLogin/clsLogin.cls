VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long

Public Function Login(Optional ByVal intType As Integer, Optional strCommand As String, Optional ByVal blnInitCom As Boolean = True, Optional ByVal blnToolsCheck As Boolean, _
                        Optional ByVal strExeFile As String, Optional ByVal lnghIns As Long) As ADODB.Connection
'���ܣ���¼��ʵ�֣��ɹ��������ӣ����ɹ�������Noting
'          intType=0-��ͨҵ����ã���ʱ��鰲װ����������������⣬չʾ�޸�����
'                        1-��ͨ�����࣬չʾ����������
'          strCommand=����������������
'          blnInitCom=�Ƿ��ʼ��ZLHIS��ϵ�Ĺ�������ZL9Comlib,��ʼ������ܽ���ע����Ϣ�Ķ�ȡ
'          blnToolsCheck=����Ƿ�װ���������Լ��Ƿ���ϵͳ�������Լ�DBA
'          strExeFile=���øò�����EXE
'          lnghIns=��ǰ���̾��,�ò�����10.35.120���������Ѿ�����
'˵������������(�����ȼ��Ӹߵ����жϣ�
'         1��ZLUA������(����"&"������",",��Ҫ���AppSoft\ZLUA.ini�ļ�),�����г���
'         2�������¼(����",",��Ҫ���AppSoft\Portal.ini�ļ�)
'         3��������¼(����"=")��ʽΪ��USER=�û��� PASS=���� SERVER=������ ONLYONE=0\1
'         4���û���/����@������  ��ʱ����Ϊ���ݿ�����
'         5���û��� ���� �˵�����  ��ʱ�û���������"=",�����������Ž���Ϊ�˵�����
'         6���˵�����
'         7�����뿪ͷ�������Զ���������������
    Dim blnShowFlash As Boolean
    Dim objFSO As New FileSystemObject
    Dim blnServiceCall As Boolean

    gstrExeFile = strExeFile
    If strExeFile <> "" Then
        gstrSetupPath = objFSO.GetParentFolderName(strExeFile)
    Else
        gstrSetupPath = objFSO.GetParentFolderName(App.Path)
    End If
    
    If IsDesinMode Then
        blnServiceCall = gobjFile.FileExists("C:APPSOFT\Public\zlServiceCall.dll")
    Else
        blnServiceCall = gobjFile.FileExists(gstrSetupPath & "\Public\zlServiceCall.dll")
    End If
    
    On Error GoTo 0
    glnghInstance = lnghIns
    Call ClearVars
    gintCallType = intType
    If strCommand Like "ZLSV*:*" Then
        gstrCommand = Sm4DecryptEcb(strCommand)
    Else
        gstrCommand = strCommand
    End If
    '���������л�ȡ�˵�������
    Me.MenuGroup = GetMenuGroup(gstrCommand)
 
    Call SetAppBusyState
    'Ϊʵ��XP�������ʾ����ǰ����ִ�иú���
    Call InitCommonControls
    If blnServiceCall = False Then
        If InStr(gstrCommand, "=") <= 0 Then Load frmSplash
        blnShowFlash = ShowSplash(1)
    End If
    '�������ܲ���
    If Not CreateRegister Then
        If blnServiceCall = False Then Unload frmSplash
        Exit Function
    End If
    '�û�ע��
    If InStr(gstrCommand, ",") > 0 Or InStr(gstrCommand, "=") > 0 Or InStr(gstrCommand, "&") > 0 Then
        If blnServiceCall Then
            If Not frmUserLogin.Docmd(gstrCommand) Then
                '������������������û�е�¼�ɹ������Զ��˳�
                If gblnExitApp Or Me.HelperMainType <> 0 Then Exit Function
                frmUserLogin.ShowMe
            End If
        Else
            If Not frmUserLoginOld.Docmd(gstrCommand) Then
                '������������������û�е�¼�ɹ������Զ��˳�
                If gblnExitApp Or Me.HelperMainType <> 0 Then Unload frmSplash: Exit Function
                frmUserLoginOld.ShowMe
            End If
        End If
    Else
        If blnServiceCall Then
            frmUserLogin.ShowMe
        Else
            frmUserLoginOld.ShowMe
        End If
    End If
    
    '��¼ʱȡ��
    If gcnOracle Is Nothing Then
        If blnServiceCall = False Then Unload frmSplash
        Exit Function
    ElseIf gcnOracle.State <> adStateOpen Then
        If blnServiceCall = False Then Unload frmSplash
        Exit Function
    End If
    
    '�ռ���ص��ֶΣ������ԣ�
    Call CollectTableField("zlUpgradeServer.FTP����ģʽ")
    
    '���ڻ�û�г�ʼ���������������Ӷ��󣬸ú����в���ʹ�ù��������е����ݿ���ʷ���
    If CheckAllowByTerminal = False Then
        If blnServiceCall = False Then Unload frmSplash
        Exit Function
    End If
    Call StartInstall
    '��ʼ����������
    If blnInitCom Then
        If zlGetComLib Then
            gobjComlib.SetNodeNo Me.NodeNo
            gobjComlib.SetNodeName Me.NodeName
            Call gobjComlib.InitCommon(gcnOracle)
        Else
            MsgBox "����ZL9Comlibʧ�ܡ������Ƿ���ȷע�ᣡ"
            If blnServiceCall = False Then Unload frmSplash
            Exit Function
        End If
        If gcnOracle.State <> adStateOpen Then
            If blnServiceCall = False Then Unload frmSplash
            Exit Function
        End If
        If Not SaveRegInfo Then
            If blnServiceCall = False Then Unload frmSplash
            Exit Function
        End If
    End If
    If blnToolsCheck Then
        '����������û����Ҫ����Ƿ�ϵͳ�����ߵ�¼�Լ��Ƿ񴴽��˹�������,
        '����DBAδ��������������ֱ���˳���DBA�û���δ�����������߷��ڵ�¼֮�⴦��
        If Not CheckSysState Then
            If blnServiceCall = False Then Unload frmSplash
            Exit Function
        End If
    End If
    '-------------------------------------------------------------
    '��鱾����װ����
    '-------------------------------------------------------------
    If intType = 0 Then '
        If TestComponent = False Then
            MsgBox "��û�в����κ�ϵͳ��Ȩ�ޣ��������˳���", vbInformation, gstrSysName
            If blnServiceCall = False Then Unload frmSplash
            Exit Function
        End If
        '-------------------------------------------------------------
        '��������ѡ����
        '-------------------------------------------------------------
        With FrmAccoutChoose
            Me.Systems = .Show_me
            If .BlnSelect = False Then
                Exit Function
            End If
            If Me.Systems = "" Then
                MsgBox "��û�в����κ�ϵͳ��Ȩ�ޣ��������˳���", vbInformation, gstrSysName
                If blnServiceCall = False Then Unload frmSplash
                Exit Function
            End If
        End With
    End If
    If blnServiceCall = False Then Unload frmSplash
    Set Login = gcnOracle
End Function

Public Function UpdateClient(Optional ByVal blnForceUpdate As Boolean) As Boolean
    gblnTimer = True
    UpdateClient = UpdateZLHIS(ComputerName, True, blnForceUpdate)
    gblnTimer = False
End Function
'���ԣ���ȡ������û���
Public Property Get InputUser() As String
    InputUser = gstrInputUser
End Property

Friend Property Let InputUser(ByVal strUser As String)
    gstrInputUser = strUser
End Property
'���ԣ���ȡ���������
Public Property Get InputPwd() As String
    InputPwd = gstrInputPwd
End Property

Friend Property Let InputPwd(ByVal strPWD As String)
    gstrInputPwd = strPWD
End Property

'���ԣ���ȡ����ķ�����
Public Property Get ServerName() As String
    ServerName = gstrServerName
End Property

Friend Property Let ServerName(ByVal strServerName As String)
    gstrServerName = strServerName
End Property

'���ԣ���ȡ������û���Ӧ�����ݿ��û�
Public Property Get DBUser() As String
    DBUser = gstrDBUser
End Property

Friend Property Let DBUser(ByVal strDBUser As String)
    gstrDBUser = strDBUser
End Property

'���ԣ���ȡ�Ƿ�������Ҫת��
Public Property Get IsTransPwd() As Boolean
    IsTransPwd = gblnTransPwd
End Property

Friend Property Let IsTransPwd(ByVal blnTransPwd As Boolean)
    gblnTransPwd = blnTransPwd
End Property

'���ԣ���ȡ�Ƿ�ϵͳ������
Public Property Get IsSysOwner() As Boolean
    IsSysOwner = gblnSysOwner
End Property

Friend Property Let IsSysOwner(ByVal blnSysOwner As Boolean)
    gblnSysOwner = blnSysOwner
End Property

'���ԣ���ȡSelSystems
Public Property Get Systems() As String
    Systems = gstrSystems
End Property

Friend Property Let Systems(ByRef strSystems As String)
    gstrSystems = strSystems
End Property

'���ԣ��Ƿ�Cancelȡ��
Public Property Get IsCancel() As Boolean
    IsCancel = gblnCancel
End Property

Friend Property Let IsCancel(ByVal blnCancel As Boolean)
    gblnCancel = blnCancel
End Property
'���ԣ��˵�������
Public Property Get MenuGroup() As String
    MenuGroup = gstrMenuGroup
End Property

Friend Property Let MenuGroup(ByRef strMenuGroup As String)
    gstrMenuGroup = strMenuGroup
End Property
'���ԣ���¼�������ڵĲ�������
Public Property Get DeptName() As String
    DeptName = gstrDeptNameTerminal
End Property

Friend Property Let DeptName(ByRef strDeptName As String)
    gstrDeptNameTerminal = strDeptName
End Property

'���ԣ���¼����վ����
Public Property Get Station() As String
    Station = gstrStation
End Property

Friend Property Let Station(ByRef strStation As String)
    gstrStation = strStation
End Property
'���ԣ���¼վ����
Public Property Get NodeNo() As String
    NodeNo = gstrNodeNo
End Property

Friend Property Let NodeNo(ByRef strNodeNo As String)
    gstrNodeNo = strNodeNo
End Property

'���ԣ���¼վ������
Public Property Get NodeName() As String
    NodeName = gstrNodeName
End Property

Friend Property Let NodeName(ByRef strNodeName As String)
    gstrNodeName = strNodeName
End Property

'���ԣ���ȡEMR����
Public Property Get EMR() As Object
    Set EMR = gobjEmr
End Property
'���ԣ�����EMR����
Public Property Set EMR(ByRef objEMR As Object)
    Set gobjEmr = objEMR
End Property
'���ԣ�EMR����֤�û�
Public Property Get EMRUser() As String
    EMRUser = gstrEMRUser
End Property

Friend Property Let EMRUser(ByRef strEMRUser As String)
    gstrEMRUser = strEMRUser
End Property
'���ԣ�EMR����֤���룬�������ݿ���������[DBPASSWORD]ǰ׺
Public Property Get EMRPwd() As String
    EMRPwd = gstrEMRPwd
End Property

Friend Property Let EMRPwd(ByRef strEMRPwd As String)
    gstrEMRPwd = strEMRPwd
End Property

'���ԣ�EMRUser��EMRPwd�Ƿ����ڴ����û����統LIS������ʱ��ʹ��ZLHIS�˻�����ʼ��EMR,��ʱ������ΪTRUE
Public Property Get IsEMRProxy() As Boolean
    IsEMRProxy = gblnEMRProxy
End Property

Friend Property Let IsEMRProxy(ByRef blnEMRProxy As Boolean)
    gblnEMRProxy = blnEMRProxy
End Property

'���ԣ�ʹ�ó��ϣ��ⲿ���ó������ø����ԣ���ҪӦ�������LIS������¼
Public Property Get UsageOccasion() As String
    UsageOccasion = gstrUsageOccasion
End Property

Public Property Let UsageOccasion(ByRef strUsageOccasion As String)
    gstrUsageOccasion = strUsageOccasion
End Property
'�������ֵ�������
Public Property Get HelperMainType() As Long
    HelperMainType = glngHelperMainType
End Property

Public Property Let HelperMainType(ByRef lngHelperMainType As Long)
    glngHelperMainType = lngHelperMainType
End Property
'�������ֵ���ʱ���д���ID
Public Property Get ParallelID() As Long
    ParallelID = glngParallelID
End Property

Public Property Let ParallelID(ByRef lngParallelID As Long)
    glngParallelID = lngParallelID
End Property
'���ԣ���ȡע�Ჿ������
Public Property Get Register() As Object
    Set Register = gobjRegister
End Property

'���ԣ���ȡ������
Public Property Get Command() As String
    Command = gstrCommand
End Property

Private Sub Class_Initialize()
    Set gclsLogin = Me
    glngInstanceCount = glngInstanceCount + 1
End Sub

Public Sub ClearVars(Optional blnClearObject As Boolean)
    If blnClearObject Then
        If Not gobjRegister Is Nothing Then Set gobjRegister = Nothing
        If Not gobjEmr Is Nothing Then
            gobjEmr.Dispose
            Set gobjEmr = Nothing
            gblnEMRProxy = False
            gstrEMRPwd = ""
            gstrEMRUser = ""
        End If
    End If
    
    Set gcolTableField = Nothing
    gstrUserName = ""
    gstrUserID = ""
    gstrDeptID = ""
    gstrDeptName = ""
    gstrDeptNameTerminal = ""
    gstrIP = ""
    gstrSessionID = ""
    
    gstrInputPwd = ""
    gstrServerName = ""
    gstrDBUser = ""
    gblnTransPwd = False
    gblnSysOwner = False
    gstrSystems = ""
    gstrDeptName = ""
    gstrStation = ""
    gstrNodeNo = ""
    gstrNodeName = ""
    gblnExitApp = False
    gblnCancel = False
    gstrUsageOccasion = ""
End Sub

Private Sub Class_Terminate()
    glngInstanceCount = glngInstanceCount - 1
    If glngInstanceCount = 0 Then
        Call ClearVars(True)
    End If
End Sub
