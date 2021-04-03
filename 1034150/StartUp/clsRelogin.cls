VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRelogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mobjEmr As Object           'EMR�°���Ӳ���
Private mstrUserName As String      'InputUser����
Private mstrInputPwd As String      'InputPwd����
Private mstrServerName As String    'ServerName����
Private mstrDBUser As String        'DBUser����
Private mblnTransPwd As Boolean     'blnTransPwd����
Private mstrMenuGroup As String     '�˵���
Private mblnSysOwner As Boolean     '�Ƿ�ϵͳ������
Private mstrSystems As String       '������ѡ���ϵͳ
Private mstrDeptName As String      '��¼����
Private mstrStation As String       '��¼����վ����
'���ԣ���ȡ������û���
Public Property Get InputUser() As String
    InputUser = mstrUserName
End Property

Friend Property Let InputUser(ByVal strUserName As String)
    mstrUserName = strUserName
End Property
'���ԣ���ȡ���������
Public Property Get InputPwd() As String
    InputPwd = mstrInputPwd
End Property

Friend Property Let InputPwd(ByVal strInputPwd As String)
    mstrInputPwd = strInputPwd
End Property
'���ԣ���ȡ����ķ�����
Public Property Get ServerName() As String
    ServerName = mstrServerName
End Property

Friend Property Let ServerName(ByVal strServerName As String)
    mstrServerName = strServerName
End Property
'���ԣ���ȡ�Ƿ�������Ҫת��
Public Property Get IsTransPwd() As Boolean
    IsTransPwd = mblnTransPwd
End Property

Friend Property Let IsTransPwd(ByVal blnTransPwd As Boolean)
    mblnTransPwd = blnTransPwd
End Property
'���ԣ���ȡ������û���Ӧ�����ݿ��û�
Public Property Get DBUser() As String
    DBUser = mstrDBUser
End Property

Friend Property Let DBUser(ByVal strDBUser As String)
    mstrDBUser = strDBUser
End Property
'���ԣ���ȡ�Ƿ�ϵͳ������
Public Property Get IsSysOwner() As Boolean
    IsSysOwner = mblnSysOwner
End Property

Friend Property Let IsSysOwner(ByVal blnSysOwner As Boolean)
    mblnSysOwner = blnSysOwner
End Property
'���ԣ���ȡ������ѡ���Systems
Public Property Get Systems() As String
    Systems = mstrSystems
End Property

Friend Property Let Systems(ByVal strSystems As String)
    mstrSystems = strSystems
End Property
'���ԣ��˵�������
Public Property Get MenuGroup() As String
    MenuGroup = mstrMenuGroup
End Property

Friend Property Let MenuGroup(ByRef strMenuGroup As String)
    mstrMenuGroup = strMenuGroup
End Property
'���ԣ���ȡ��¼����
Public Property Get DeptName() As String
    DeptName = mstrDeptName
End Property

Friend Property Let DeptName(ByVal strDeptName As String)
    mstrDeptName = strDeptName
End Property
'���ԣ���¼վ����
Public Property Get NodeNo() As String
    NodeNo = gstrNodeNo
End Property

Friend Property Let NodeNo(ByRef strNodeNo As String)
    gstrNodeNo = strNodeNo
End Property
'���ԣ���¼����վ����
Public Property Get Station() As String
    Station = mstrStation
End Property

Friend Property Let Station(ByRef strStation As String)
    mstrStation = strStation
End Property
'���ԣ���ȡEMR����
Public Property Get EMR() As Object
    Set EMR = mobjEmr
End Property
'���ԣ�����EMR����
Public Property Set EMR(ByRef objEMR As Object)
    Set mobjEmr = objEMR
End Property

'ж�ز���ȫ�ֶ���
Public Sub Dispose()
    If Not mobjEmr Is Nothing Then
        Call mobjEmr.Dispose
        Set mobjEmr = Nothing
    End If
    If Not gobjWait Is Nothing Then
        Unload gobjWait
        Set gobjWait = Nothing
    End If
End Sub

Public Sub Relogin(ByVal FrmMainObj As Object)
    Unload FrmMainObj
    Call Main
End Sub

Public Function UpdateClient(Optional ByVal blnForceUpdate As Boolean) As Boolean
    gblnTimer = True
    UpdateClient = UpdateZLHIS(ComputerName, True, blnForceUpdate)
    gblnTimer = False
End Function