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
'功能：登录类实现，成功返回连接，不成功，返回Noting
'          intType=0-普通业务调用，此时检查安装部件，检查帐套问题，展示修改密码
'                        1-普通工具类，展示服务器配置
'          strCommand=程序启动的命令行
'          blnInitCom=是否初始化ZLHIS体系的公共部件ZL9Comlib,初始化后才能进行注册信息的读取
'          blnToolsCheck=检查是否安装管理工具以及是否是系统所有者以及DBA
'          strExeFile=调用该部件的EXE
'          lnghIns=当前进程句柄,该参数在10.35.120独立升级已经弃用
'说明：命令种类(按优先级从高到底判断）
'         1、ZLUA命令行(包含"&"不包含",",需要配合AppSoft\ZLUA.ini文件),命令行长度
'         2、单点登录(包含",",需要配合AppSoft\Portal.ini文件)
'         3、三方登录(包含"=")格式为：USER=用户名 PASS=密码 SERVER=服务器 ONLYONE=0\1
'         4、用户名/密码@服务器  此时密码为数据库密码
'         5、用户名 密码 菜单组名  此时用户名不包含"=",第三个参数才解析为菜单组名
'         6、菜单组名
'         7、密码开头，用于自动升级和升级助手
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
    '从命令行中获取菜单组名称
    Me.MenuGroup = GetMenuGroup(gstrCommand)
 
    Call SetAppBusyState
    '为实现XP风格，在显示窗体前必须执行该函数
    Call InitCommonControls
    If blnServiceCall = False Then
        If InStr(gstrCommand, "=") <= 0 Then Load frmSplash
        blnShowFlash = ShowSplash(1)
    End If
    '创建加密部件
    If Not CreateRegister Then
        If blnServiceCall = False Then Unload frmSplash
        Exit Function
    End If
    '用户注册
    If InStr(gstrCommand, ",") > 0 Or InStr(gstrCommand, "=") > 0 Or InStr(gstrCommand, "&") > 0 Then
        If blnServiceCall Then
            If Not frmUserLogin.Docmd(gstrCommand) Then
                '升级助手任务处理，若没有登录成功，则自动退出
                If gblnExitApp Or Me.HelperMainType <> 0 Then Exit Function
                frmUserLogin.ShowMe
            End If
        Else
            If Not frmUserLoginOld.Docmd(gstrCommand) Then
                '升级助手任务处理，若没有登录成功，则自动退出
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
    
    '登录时取消
    If gcnOracle Is Nothing Then
        If blnServiceCall = False Then Unload frmSplash
        Exit Function
    ElseIf gcnOracle.State <> adStateOpen Then
        If blnServiceCall = False Then Unload frmSplash
        Exit Function
    End If
    
    '收集相关的字段（兼容性）
    Call CollectTableField("zlUpgradeServer.FTP连接模式")
    
    '由于还没有初始化公共部件的连接对象，该函数中不能使用公共部件中的数据库访问方法
    If CheckAllowByTerminal = False Then
        If blnServiceCall = False Then Unload frmSplash
        Exit Function
    End If
    Call StartInstall
    '初始化公共部件
    If blnInitCom Then
        If zlGetComLib Then
            gobjComlib.SetNodeNo Me.NodeNo
            gobjComlib.SetNodeName Me.NodeName
            Call gobjComlib.InitCommon(gcnOracle)
        Else
            MsgBox "创建ZL9Comlib失败。请检查是否正确注册！"
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
        '管理工具类没有需要检查是否系统所有者登录以及是否创建了管理工具,
        '若非DBA未创建管理工具则直接退出。DBA用户且未创建管理工具放在登录之外处理
        If Not CheckSysState Then
            If blnServiceCall = False Then Unload frmSplash
            Exit Function
        End If
    End If
    '-------------------------------------------------------------
    '检查本机安装部件
    '-------------------------------------------------------------
    If intType = 0 Then '
        If TestComponent = False Then
            MsgBox "您没有操作任何系统的权限，程序被迫退出！", vbInformation, gstrSysName
            If blnServiceCall = False Then Unload frmSplash
            Exit Function
        End If
        '-------------------------------------------------------------
        '调用帐套选择窗体
        '-------------------------------------------------------------
        With FrmAccoutChoose
            Me.Systems = .Show_me
            If .BlnSelect = False Then
                Exit Function
            End If
            If Me.Systems = "" Then
                MsgBox "您没有操作任何系统的权限，程序被迫退出！", vbInformation, gstrSysName
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
'属性，获取输入的用户名
Public Property Get InputUser() As String
    InputUser = gstrInputUser
End Property

Friend Property Let InputUser(ByVal strUser As String)
    gstrInputUser = strUser
End Property
'属性，获取输入的密码
Public Property Get InputPwd() As String
    InputPwd = gstrInputPwd
End Property

Friend Property Let InputPwd(ByVal strPWD As String)
    gstrInputPwd = strPWD
End Property

'属性，获取输入的服务器
Public Property Get ServerName() As String
    ServerName = gstrServerName
End Property

Friend Property Let ServerName(ByVal strServerName As String)
    gstrServerName = strServerName
End Property

'属性，获取输入的用户对应的数据库用户
Public Property Get DBUser() As String
    DBUser = gstrDBUser
End Property

Friend Property Let DBUser(ByVal strDBUser As String)
    gstrDBUser = strDBUser
End Property

'属性，获取是否密码需要转换
Public Property Get IsTransPwd() As Boolean
    IsTransPwd = gblnTransPwd
End Property

Friend Property Let IsTransPwd(ByVal blnTransPwd As Boolean)
    gblnTransPwd = blnTransPwd
End Property

'属性，获取是否系统所有者
Public Property Get IsSysOwner() As Boolean
    IsSysOwner = gblnSysOwner
End Property

Friend Property Let IsSysOwner(ByVal blnSysOwner As Boolean)
    gblnSysOwner = blnSysOwner
End Property

'属性，获取SelSystems
Public Property Get Systems() As String
    Systems = gstrSystems
End Property

Friend Property Let Systems(ByRef strSystems As String)
    gstrSystems = strSystems
End Property

'属性，是否Cancel取消
Public Property Get IsCancel() As Boolean
    IsCancel = gblnCancel
End Property

Friend Property Let IsCancel(ByVal blnCancel As Boolean)
    gblnCancel = blnCancel
End Property
'属性，菜单组名称
Public Property Get MenuGroup() As String
    MenuGroup = gstrMenuGroup
End Property

Friend Property Let MenuGroup(ByRef strMenuGroup As String)
    gstrMenuGroup = strMenuGroup
End Property
'属性，登录机器所在的部门名称
Public Property Get DeptName() As String
    DeptName = gstrDeptNameTerminal
End Property

Friend Property Let DeptName(ByRef strDeptName As String)
    gstrDeptNameTerminal = strDeptName
End Property

'属性，登录工作站名称
Public Property Get Station() As String
    Station = gstrStation
End Property

Friend Property Let Station(ByRef strStation As String)
    gstrStation = strStation
End Property
'属性，登录站点编号
Public Property Get NodeNo() As String
    NodeNo = gstrNodeNo
End Property

Friend Property Let NodeNo(ByRef strNodeNo As String)
    gstrNodeNo = strNodeNo
End Property

'属性，登录站点名称
Public Property Get NodeName() As String
    NodeName = gstrNodeName
End Property

Friend Property Let NodeName(ByRef strNodeName As String)
    gstrNodeName = strNodeName
End Property

'属性，获取EMR对象
Public Property Get EMR() As Object
    Set EMR = gobjEmr
End Property
'属性，设置EMR对象
Public Property Set EMR(ByRef objEMR As Object)
    Set gobjEmr = objEMR
End Property
'属性，EMR的验证用户
Public Property Get EMRUser() As String
    EMRUser = gstrEMRUser
End Property

Friend Property Let EMRUser(ByRef strEMRUser As String)
    gstrEMRUser = strEMRUser
End Property
'属性，EMR的验证密码，若是数据库密码增加[DBPASSWORD]前缀
Public Property Get EMRPwd() As String
    EMRPwd = gstrEMRPwd
End Property

Friend Property Let EMRPwd(ByRef strEMRPwd As String)
    gstrEMRPwd = strEMRPwd
End Property

'属性，EMRUser、EMRPwd是否属于代理用户。如当LIS单独库时，使用ZLHIS账户来初始化EMR,此时该属性为TRUE
Public Property Get IsEMRProxy() As Boolean
    IsEMRProxy = gblnEMRProxy
End Property

Friend Property Let IsEMRProxy(ByRef blnEMRProxy As Boolean)
    gblnEMRProxy = blnEMRProxy
End Property

'属性，使用场合，外部调用程序设置该属性，主要应用于体检LIS单独登录
Public Property Get UsageOccasion() As String
    UsageOccasion = gstrUsageOccasion
End Property

Public Property Let UsageOccasion(ByRef strUsageOccasion As String)
    gstrUsageOccasion = strUsageOccasion
End Property
'升级助手调用类型
Public Property Get HelperMainType() As Long
    HelperMainType = glngHelperMainType
End Property

Public Property Let HelperMainType(ByRef lngHelperMainType As Long)
    glngHelperMainType = lngHelperMainType
End Property
'升级助手调用时并行处理ID
Public Property Get ParallelID() As Long
    ParallelID = glngParallelID
End Property

Public Property Let ParallelID(ByRef lngParallelID As Long)
    glngParallelID = lngParallelID
End Property
'属性，获取注册部件对象
Public Property Get Register() As Object
    Set Register = gobjRegister
End Property

'属性，获取命令行
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

