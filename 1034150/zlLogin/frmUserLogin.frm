VERSION 5.00
Begin VB.Form frmUserLogin 
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "操作员登录"
   ClientHeight    =   6225
   ClientLeft      =   -15
   ClientTop       =   -45
   ClientWidth     =   9180
   Icon            =   "frmUserLogin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmUserLogin.frx":1CFA
   ScaleHeight     =   6225
   ScaleWidth      =   9180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picUp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   8520
      Picture         =   "frmUserLogin.frx":EB61
      ScaleHeight     =   240
      ScaleWidth      =   360
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3120
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox picDown 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   4380
      Picture         =   "frmUserLogin.frx":F24B
      ScaleHeight     =   360
      ScaleWidth      =   360
      TabIndex        =   15
      Top             =   3270
      Width           =   360
   End
   Begin VB.TextBox txtServer 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   290
      Left            =   2640
      TabIndex        =   2
      Text            =   "zlhishis"
      Top             =   3350
      Width           =   1740
   End
   Begin VB.PictureBox picCon 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   2280
      Picture         =   "frmUserLogin.frx":F935
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   3420
      Width           =   240
   End
   Begin VB.PictureBox picPWD 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   2280
      Picture         =   "frmUserLogin.frx":10337
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2975
      Width           =   240
   End
   Begin VB.PictureBox picUser 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   2280
      Picture         =   "frmUserLogin.frx":10D39
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2520
      Width           =   240
   End
   Begin VB.PictureBox picLogin 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   2280
      ScaleHeight     =   405
      ScaleWidth      =   2415
      TabIndex        =   4
      Top             =   3840
      Width           =   2415
      Begin VB.Label lblLogin 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "登 录"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   900
         TabIndex        =   3
         Top             =   90
         Width           =   615
      End
   End
   Begin VB.PictureBox picSet 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   4800
      Picture         =   "frmUserLogin.frx":1173B
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   7
      Top             =   3405
      Width           =   240
   End
   Begin VB.PictureBox picModify 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   4440
      Picture         =   "frmUserLogin.frx":1213D
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   6
      Top             =   2930
      Width           =   240
   End
   Begin VB.PictureBox picHos 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2385
      Left            =   1440
      ScaleHeight     =   2385
      ScaleWidth      =   4740
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   4740
   End
   Begin VB.ComboBox cboServer 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   2640
      Sorted          =   -1  'True
      TabIndex        =   11
      Text            =   "cboServer"
      Top             =   3345
      Width           =   2040
   End
   Begin VB.TextBox txtPassWord 
      BorderStyle     =   0  'None
      Height          =   275
      IMEMode         =   3  'DISABLE
      Left            =   2640
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2930
      Width           =   1800
   End
   Begin VB.TextBox txtUser 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   275
      Left            =   2640
      TabIndex        =   0
      Text            =   "zlhishis"
      Top             =   2480
      Width           =   1800
   End
   Begin VB.Line Line11 
      BorderColor     =   &H00E0E0E0&
      X1              =   2280
      X2              =   4695
      Y1              =   3825
      Y2              =   3825
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00E0E0E0&
      X1              =   4695
      X2              =   4695
      Y1              =   3840
      Y2              =   4245
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00E0E0E0&
      X1              =   2265
      X2              =   2265
      Y1              =   3840
      Y2              =   4245
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00E0E0E0&
      X1              =   2280
      X2              =   4710
      Y1              =   4250
      Y2              =   4250
   End
   Begin VB.Image imgCancel 
      Height          =   360
      Left            =   6120
      Picture         =   "frmUserLogin.frx":12B3F
      Top             =   0
      Width           =   360
   End
   Begin VB.Line Line7 
      BorderColor     =   &H8000000A&
      X1              =   2640
      X2              =   4680
      Y1              =   3645
      Y2              =   3645
   End
   Begin VB.Line Line6 
      BorderColor     =   &H8000000A&
      X1              =   2640
      X2              =   4680
      Y1              =   3200
      Y2              =   3200
   End
   Begin VB.Line Line5 
      BorderColor     =   &H8000000A&
      X1              =   2640
      X2              =   4680
      Y1              =   2765
      Y2              =   2765
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   6495
      Y1              =   4420
      Y2              =   4420
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      X1              =   6480
      X2              =   6480
      Y1              =   0
      Y2              =   4430
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   6480
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   0
      Y1              =   -240
      Y2              =   4430
   End
   Begin VB.Image ImgIndicate 
      Appearance      =   0  'Flat
      Height          =   780
      Left            =   120
      Picture         =   "frmUserLogin.frx":13229
      Top             =   3240
      Width           =   780
   End
   Begin VB.Label LblProductName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "#"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   1440
      TabIndex        =   9
      Top             =   1200
      Width           =   4650
   End
   Begin VB.Image imgPic 
      Height          =   2745
      Left            =   120
      Picture         =   "frmUserLogin.frx":137B5
      Top             =   240
      Width           =   1260
   End
   Begin VB.Label lbltag 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "#"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   4440
      TabIndex        =   8
      Top             =   1680
      Width           =   195
   End
End
Attribute VB_Name = "frmUserLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'命令行格式：

'zlhis.exe 菜单
'zlhis.exe 用户名/密码        此种情况不需要进行密码转换
'zlhis.exe 用户名 密码
'zlhis.exe 用户名 密码 菜单
Private mblnFirst As Boolean  '为True表示已经正常显示出
Private mintTimes As Integer  '登录重试次数
Private mbln转换 As Boolean     '表示传入的密码是否为数据库密码，是否不需要再转换
Private mcolServer As New Collection  '保存服务器串列表
Private mblnAccess As Boolean  '为True外部调用ZLHIS成功
Private mblnUAAddUser As Boolean

Private mobjHttp As New XMLHTTP
Private mstrPostData As String
Private mstr断言 As String
Private mstrUserURL As String
Private mstrSamlAssertion As String
Private mstrError As String
Private mblnZLUA As Boolean
Private mstrAppID As String
Private mstrZLUAUser As String
Private mblnOk          As Boolean
Private Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long
'获取当前用户系统所选语言id
Private Declare Function GetUserDefaultUILanguage Lib "kernel32.dll" () As Long
'获取区域设置ID
Private Declare Function GetThreadLocale Lib "kernel32.dll" () As Long
'无边框时窗体拖动
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

Private Function GetUserInfor() As ADODB.Recordset

    Dim strSQL As String
 
    strSQL = "Select b.id as 人员ID,d.ID as 部门ID,d.名称 As 部门, b.姓名," & vbNewLine & _
            " Sys_Context('USERENV', 'IP_ADDRESS') Ip,sys_context('USERENV', 'SESSIONID') SESSIONID " & vbNewLine & _
            "From 上机人员表 A, 人员表 B, 部门人员 C, 部门表 D" & vbNewLine & _
            "Where a.用户名 = [1] And a.人员id = b.Id And b.Id = c.人员id And c.部门id = d.Id And c.缺省 = 1"
    On Error GoTo errH
    Set GetUserInfor = OpenSQLRecord(strSQL, Me.Caption, gstrDBUser)

    Exit Function
errH:
    MsgBox Err.Description & vbCrLf & strSQL, vbExclamation, "读取登录人员信息"
    Set GetUserInfor = New ADODB.Recordset
End Function

Private Sub cmdOK_Click()
    Dim strNote             As String
    Dim strUserName         As String
    Dim strServerName       As String
    Dim strPassword         As String
    Dim blnTransPassword    As Boolean
    Dim strError            As String
    Dim strInfo             As String
    
    Dim strServer As String
    Dim objService          As New clsService
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    mintTimes = mintTimes + 1
    SetConState False
    strInfo = CheckInput(strUserName, strPassword, strServerName)
    If LenB(strInfo) <> 0 Then
        SetConState
        GoTo errH
    End If
    
    If UCase(strUserName) = "SYS" Or UCase(strUserName) = "SYSTEM" Then
        blnTransPassword = False
    Else
        blnTransPassword = mbln转换
    End If
    If glngHelperMainType = 0 Then
        '2052代表简体中文，区域语言不是简体中文，无法登陆
        If GetThreadLocale <> 2052 Then
            strInfo = "本机系统语言不是简体中文，无法登陆导航台；" & vbCrLf & "请修改区域语言后重启电脑再登录！"
'            MsgBox "本机系统语言不是简体中文，无法登陆导航台；" & vbCrLf & "请修改区域语言后重启电脑再登录！"
            SetConState
            GoTo errH
        End If
    End If
    Set gcnOracle = gobjRegister.GetConnection(strServerName, strUserName, strPassword, blnTransPassword, , strError)
    'ora-28002:密码还有多少天过期，不会返回，因此，必须CheckPwdExpiry来辅助提示密码过期
    If gcnOracle.State = adStateClosed Then
        'zlRegister中已进行错误转换
        
        strInfo = strError
        txtPassWord.Text = ""
        mblnAccess = False
        If mblnZLUA = True Then mblnUAAddUser = True
        On Error Resume Next
        txtPassWord.SetFocus
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo errH
        SetConState
        GoTo errH
    Else
        '记录当前用户信息
        gclsLogin.DBUser = UCase(strUserName)
        Set rsTmp = GetUserInfor
        If Not rsTmp.EOF Then
            gstrUserName = rsTmp!姓名
            gstrUserID = rsTmp!人员id
            gstrDeptID = rsTmp!部门id
            gstrDeptName = rsTmp!部门
            gstrIP = rsTmp!IP & ""
            gstrSessionID = rsTmp!SESSIONID
        End If
        
        
        '检查是否启用应用信息跟踪包
        Set rsTmp = GetZLOptions(33)
        If Not rsTmp.EOF Then
            If Val("" & rsTmp!参数值) = 1 Then
                strInfo = "DEPT:" & gstrDeptName & ",UNAME:" & gstrUserName & ",USER:" & gclsLogin.DBUser & ",IP:" & gstrIP
                Call ExecuteProcedure("dbms_application_info.set_client_info('" & strInfo & "')", Me.Caption)
            End If
        End If
    
         '连接成功后就初始化连接对象到zlRegister，以便给病历部件获取连接对象
        Call gobjRegister.zlRegInit(gcnOracle)
    
        If blnTransPassword Then
            On Error Resume Next
            Call ExecuteProcedure("Zlpassword_Update('" & UCase(strUserName) & "','" & Sm4EncryptEcb(UCase(strPassword), GetGeneralAccountKey(G_PASSWORD_KEY)) & "')", Me.Caption)
            If Err.Number <> 0 Then Err.Clear
            On Error GoTo errH
        End If
        '检查是否切换了服务器
        If glngHelperMainType = 0 Then
            strServer = GetSetting("ZLSOFT", "注册信息\登陆信息", "SERVER")
            If UCase(strServer) <> UCase(txtServer.Text) Then
                ClearComponent
            End If
        End If
        If strUserName = strPassword Then
            strInfo = "登录用户名和密码相同，不符合系统安全要求，请您立即修改密码。"
            If gintCallType = 0 Then '现实修改按钮
                picModify_Click
                SetConState
            End If
            GoTo errH
        End If
        '检查密码复杂度是否符合要求
        strInfo = CheckPWDComplex(gcnOracle, strPassword)
        If LenB(strInfo) <> 0 Then
            If gintCallType = 0 Then '现实修改按钮
                picModify_Click
                SetConState
            End If
            GoTo errH
        End If
        
        '是否密码过期提醒
        If CheckPwdExpiry = True Then
            If gintCallType = 0 Then '现实修改按钮
                picModify_Click
                SetConState
            End If
            Exit Sub
        End If
    End If
        
    If CheckUserExpiry = False Then
        txtUser.Text = ""
        txtPassWord.Text = ""
        On Error Resume Next
        txtUser.SetFocus
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo errH
        SetConState
        Exit Sub
    End If
      
    '启动SQL Trace
    '-----------------------------------------------
    strNote = SetSQLTrace(strServerName, strUserName)
    If strNote <> "" Then
        MsgBox "已启动SQL Trace功能!" & vbCrLf & "跟踪结果文件:" & strNote & vbCrLf & _
                "存放在Oracle服务器udump目录下,超过100M后将停止写入.", vbInformation, "提示"
    End If
    If UCase(strServerName) = "RBO" Then
        SetRunWithRBO
    End If
    '接口调用，放到Trace启动之后
    '-----------------------------------------------
    '1.中联单点登录添加ZLUA账户
    If mblnUAAddUser = True And mstrUserURL <> "" Then
        mstr断言 = SoapEnvelope("AddUserAppInfo", mstrZLUAUser, mstrAppID, txtUser.Text & "/" & txtPassWord.Text & "@" & txtServer.Text, mstrSamlAssertion)
        Call PostData(mstrUserURL, "AddUserAppInfo", mstr断言, 5)
        mblnUAAddUser = False
    End If
    
    '2.新版病历、自动升级程序、导航台，需要的用户名及密码(用户输入的密码，zlbrw部件中会使用)
    gclsLogin.InputUser = strUserName
    gclsLogin.InputPwd = strPassword
    gclsLogin.ServerName = strServerName
    gclsLogin.IsTransPwd = blnTransPassword
    '修改注册表
    If strUserName = "ZLUA" Then
        'ZLUA登录，则不保存ZLUA
    Else
        SaveSetting "ZLSOFT", "注册信息\登陆信息", "USER", strUserName
        SaveSetting "ZLSOFT", "注册信息\登陆信息", "SERVER", strServerName
        SaveSetting "ZLSOFT", "注册信息\登陆信息", "服务器来源", cboServer.Tag
    End If
    If glngHelperMainType = 0 Then
        If Val(GetSetting("ZLSOFT", "公共模块\升级助手", "非首次启动服务", "0")) = 0 Then
            '启动服务
            If Not objService.IsInstalled("ZLHelperService") Then
                If gobjFile.FileExists(gstrSetupPath & "\ZLHelperService.exe") Then
                    Call objService.Install("ZLHelperService", "ZLSOFT Upgrade Helper Service", "中联升级助手服务", gstrSetupPath & "\ZLHelperService.exe")
                End If
            End If
            If objService.IsInstalled("ZLHelperService") Then
    '            Call objService.AutoRun("ZLHelperService")
                If Not objService.IsRunning("ZLHelperService") Then
                    Call objService.Start("ZLHelperService")
                End If
            End If
            SaveSetting "ZLSOFT", "公共模块\升级助手", "非首次启动服务", 1
        End If
    End If
    mblnAccess = True
    mblnOk = True
    Unload Me
    Exit Sub
errH:
    If mintTimes > 3 Then
        MsgBox "超过三次登录失败，系统将自动退出", vbInformation, gstrSysName
        cmdCancel_Click
    Else
        If LenB(strInfo) = 0 Then
            strInfo = Err.Description
        End If
        MsgBox strInfo, vbInformation, gstrSysName
        SetConState
    End If
End Sub

Private Function CheckUserExpiry() As Boolean
'功能:检查该账户是否已经过期
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim blnStop As Boolean, bln医生 As Boolean
    Dim str撤档时间 As String
    
    On Error Resume Next
    strSQL = "Select Nvl(b.撤档时间, To_Date('3000-01-01', 'YYYY-MM-DD')) 撤档时间" & vbNewLine & _
            "From 人员表 b" & vbNewLine & _
            "Where (Nvl(b.帐号到期时间, To_Date('3000-01-01', 'YYYY-MM-DD')) < Sysdate Or" & vbNewLine & _
            "      Nvl(b.撤档时间, To_Date('3000-01-01', 'YYYY-MM-DD')) < Sysdate) And b.Id = [1]"
    Set rsTemp = OpenSQLRecord(strSQL, "帐号到期检查", gstrUserID)
    If Not rsTemp Is Nothing Then
        Err.Clear
        On Error GoTo errH
        If rsTemp.RecordCount > 0 Then
            If CDate(rsTemp!撤档时间) <> CDate("3000-1-1") Then
                MsgBox gstrDBUser & "用户对应的人员已撤档，将自动停用该用户！", vbInformation, gstrSysName
                blnStop = True
            Else
                MsgBox gstrDBUser & "用户已到期，将自动停用该用户！", vbInformation, gstrSysName
                blnStop = True
            End If
        End If
    End If
    If blnStop Then
        strSQL = "Select 1 From 人员性质说明 B Where b.人员id = [1] And b.人员性质 = '医生'"
        Set rsTemp = OpenSQLRecord(strSQL, "检查人员性质", gstrUserID)
        If Not rsTemp Is Nothing Then bln医生 = Not rsTemp.EOF
        
        str撤档时间 = Format(Currentdate, "YYYY-MM-DD hh:mm:ss")
        
        On Error Resume Next
        strSQL = "Zl_人员表_停用自己(To_Date('" & str撤档时间 & "','YYYY-MM-DD HH24:MI:SS'))"
        Call ExecuteProcedure(strSQL, Me.Caption)
        If Err.Number = 0 Then
            If bln医生 Then
                '费用域：更新挂号安排，有其他系统独立安装的情况，故错误需忽略
                Call zlExseSvr_UpdRgstArrangeMent(2, gstrUserID, str撤档时间)
            End If
        Else
            Err.Clear: On Error GoTo errH
            strSQL = "Zl_人员表_停用自己"
            Call ExecuteProcedure(strSQL, Me.Caption)
        End If
        Exit Function
    End If
    CheckUserExpiry = True
    Exit Function
errH:
    MsgBox Err.Description, vbInformation, gstrSysName
    SetConState
End Function

Private Sub SetRunWithRBO()
'功能：当前会话以RBO优化器模式运行SQL语句
    Dim strSQL As String
    strSQL = "alter session set optimizer_mode=rule"
    On Error Resume Next
    gcnOracle.Execute strSQL
    If Err.Number = 0 Then
        MsgBox "已设置当前会话以RBO优化器模式运行！", vbInformation, gstrSysName
    End If
End Sub

Private Function GetTrcFile(ByVal strUserName As String) As String
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    Dim strFile As String
        
    On Error Resume Next
    strFile = "ZL_" & strUserName
    strSQL = "alter session set tracefile_identifier='" & strFile & "'"
    gcnOracle.Execute strSQL
    If Err.Number <> 0 Then     '发生错误,说明设置traceid失败,保存默认的traceFile名称
        strFile = ""
        Exit Function
    End If
    
    strSQL = "Select Lower(Sys_Context('userenv', 'instance_name')) || '_ora_' || p.Spid || '" & "_" & strFile & ".trc' As Trace_File" & vbNewLine & _
                    "From V$session S, V$process P" & vbNewLine & _
                    "Where s.Paddr = p.Addr And s.Sid = Userenv('sid') And s.Audsid = [1]"
    Set rsTmp = OpenSQLRecord(strSQL, "获取TraceFile名称", gstrSessionID)
    
    If rsTmp.RecordCount > 0 Then
        GetTrcFile = rsTmp!Trace_File
    End If
    
End Function

Private Function SetSQLTrace(ByVal strServerName As String, ByVal strUserName As String) As String
'功能:调用100046事件启动SQL Trace功能
'返回:Trc文件名
    Dim strSQL As String, strLevel As String, strFile As String
    Dim rsTmp As ADODB.Recordset
    
    strServerName = UCase(strServerName)
    
    If strServerName Like "SQLTRACE*" Then
        On Error Resume Next
        strSQL = "alter session set timed_statistics=true"
        gcnOracle.Execute strSQL
        strSQL = "alter session set max_dump_file_size='100M'"
        gcnOracle.Execute strSQL
        Err.Clear
        
        '设置Trc文件别名
        strFile = GetTrcFile(strUserName)
        
        strLevel = "12"
        If Replace(strServerName, "SQLTRACE", "") = "4" Then
            strLevel = "4"
        ElseIf Replace(strServerName, "SQLTRACE", "") = "8" Then
            strLevel = "8"
        ElseIf Replace(strServerName, "SQLTRACE", "") = "12" Then
            strLevel = "12"
        End If
        strSQL = "alter session set events '10046 trace name context forever ,level " & strLevel & "'"
        gcnOracle.Execute strSQL
        If Err.Number = 0 Then
            SetSQLTrace = strFile
            
            strSQL = "Select 1 From zlreginfo Where 项目=[1]"
            Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, "TRACE文件")
            
            If rsTmp.RecordCount > 0 Then
                strSQL = "Update zlreginfo Set 内容 ='" & strFile & "' Where 项目='TRACE文件'"
            Else
                strSQL = "Insert Into zlreginfo (项目,内容) Values ('TRACE文件','" & strFile & "')"
            End If
            gcnOracle.Execute strSQL

        End If
    End If
End Function

Private Sub cmdCancel_Click()
    Set gobjRegister = Nothing
    gclsLogin.IsCancel = True
    '密码不符合规则，修改密码点取消，此时gcnOracle不为nothing
    If Not gcnOracle Is Nothing Then
        If gcnOracle.State = adStateOpen Then
            gcnOracle.Close
        End If
    End If
    Unload Me
End Sub

Private Sub ModifyPWD()
    Dim strUserName As String
    Dim strPassword As String
    Dim strServerName As String
    Dim strNote As String
    
    On Error GoTo InputError
    '------检验用户是否oracle合法用户----------------
    strUserName = Trim(txtUser.Text)
    strPassword = Trim(txtPassWord.Text)
    strServerName = Trim(txtServer.Text)
    
    '有效字符串效验
    If Len(Trim(txtUser.Text)) = 0 Then
        strNote = "请输入用户名"
        txtUser.SetFocus
        GoTo InputError
    End If
    
    If Len(strUserName) <> 1 Then
        If Mid(strUserName, 1, 1) = "/" Or Mid(strUserName, 1, 1) = "@" Or Mid(strUserName, Len(strUserName) - 1, 1) = "/" Or Mid(strUserName, Len(strUserName) - 1, 1) = "@" Then
            txtUser.SetFocus
            strNote = "用户名错误"
            SetConState
            Exit Sub
        End If
    End If
    
    If Trim(strPassword) <> "" And Len(strPassword) <> 1 Then
        If Mid(strPassword, Len(strPassword) - 1, 1) = "/" Or Mid(strPassword, Len(strPassword) - 1, 1) = "@" Or Mid(strPassword, 1, 1) = "/" Or Mid(strPassword, 1, 1) = "@" Then
            If txtPassWord.Enabled Then txtPassWord.SetFocus
            strNote = "密码错误"
            GoTo InputError
        End If
    End If
    If Trim(strServerName) <> "" Then
        If Mid(strServerName, Len(strServerName) - 1, 1) = "/" Or Mid(strServerName, Len(strServerName) - 1, 1) = "@" Or Mid(strServerName, 1, 1) = "/" Or Mid(strServerName, 1, 1) = "@" Then
            strNote = "主机连接串错误"
            cboServer.SetFocus
            GoTo InputError
        End If
    End If
    
    '分离字符串
    Dim intPos As Integer
    intPos = InStr(strUserName, "@")
    If intPos > 0 Then
        strServerName = Mid(strUserName, intPos + 1)
        strUserName = Mid(strUserName, 1, intPos - 1)
    End If
    
    intPos = InStr(strUserName, "/")
    If intPos > 0 Then
        strPassword = Mid(strUserName, intPos + 1)
        strUserName = Mid(strUserName, 1, intPos - 1)
    End If
    
    intPos = InStr(strPassword, "@")
    If intPos > 0 Then
        strServerName = Mid(strPassword, intPos + 1)
        strPassword = Mid(strPassword, 1, intPos - 1)
    End If
    
    If FrmChangePass.ShowMe(Me, strUserName, strPassword, strServerName, mbln转换) Then
        txtPassWord.Text = strPassword
        cboServer.Text = strServerName
        txtServer.Text = strServerName
        If lblLogin.Enabled And picLogin.Enabled Then Call cmdOK_Click
    Else
        txtPassWord.SetFocus
    End If
    Exit Sub
InputError:
    If strNote <> "" Then
        MsgBox strNote, vbInformation, gstrSysName
    Else
        MsgBox Err.Description, vbInformation, gstrSysName
    End If
End Sub

Private Sub SetSourse()
    Dim strPath As String   'Oracle安装目录
    Dim strCommond As String, strError As String
    
    strPath = picSet.Tag
    If strPath = "" Then
        MsgBox "本机的Oracle是否正常安装，请检查。" & vbCrLf & strError, vbInformation, "提示"
        Exit Sub
    End If
    
    '执行Oracle 8 的Net Easy配置的程序
    strCommond = strPath & "\BIN\N8SW.EXE"
    If ExecuteCommand(strCommond) = True Then
        '已经成功
        Exit Sub
    End If
    
    '执行Oracle 8i,9i,10g,11g的Net Easy配置的程序
    strCommond = strPath & "\BIN\launch.exe """ & strPath & "\network\tools"" " & strPath & "\network\tools\netca.cl"
    If ExecuteCommand(strCommond) = True Then
        '已经成功
        Exit Sub
    End If
End Sub

Private Sub cboServer_Click()
    txtServer.Text = cboServer.Text
End Sub

Private Sub Form_Activate()
    Dim LngStyle As Long
    
    If mblnFirst = False Then
        
        If InStr(gstrCommand, "=") <= 0 And InStr(gstrCommand, "&") <= 0 Then
            '设置当前窗口在任务栏显示
            LngStyle = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
            LngStyle = LngStyle Or WinStyle
            Call SetWindowLong(Me.hwnd, GWL_EXSTYLE, LngStyle)
            
            ShowWindow Me.hwnd, 0 '先隐藏
            ShowWindow Me.hwnd, 1 '再显示
        
            If Trim(txtUser.Text) = "" Then
                txtUser.SetFocus
            Else
                txtPassWord.SetFocus
            End If
        End If
        
        mblnFirst = True
        If Trim(txtUser.Text) <> "" And Trim(txtPassWord.Text) <> "" Then Call cmdOK_Click
    End If
    If InStr(gstrCommand, "=") > 0 And InStr(gstrCommand, "&") = 0 Then Me.Hide
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Me.ActiveControl.Name = "txtPassWord" Then
            Call cmdOK_Click
        Else
            SendKeys "{Tab}"
        End If
    ElseIf KeyCode = vbKeyEscape Then
        Call imgCancel_Click
    End If
End Sub

Private Sub Form_Load()
    Dim ArrCommand
    
    picLogin.BackColor = RGB(0, 191, 255)
    lblLogin.BackColor = RGB(0, 191, 255)
    Me.AutoRedraw = True
    
    Call ShowSplash
    
    Call InitFaceType
    Call LoadServer
    
    On Error GoTo errH
    txtUser.Text = GetSetting(appName:="ZLSOFT", Section:="注册信息\登陆信息", key:="USER", Default:="")
    cboServer.Text = GetSetting(appName:="ZLSOFT", Section:="注册信息\登陆信息", key:="SERVER", Default:="")
    txtServer.Text = GetSetting(appName:="ZLSOFT", Section:="注册信息\登陆信息", key:="SERVER", Default:="")
    
    Call ApplyOEM_Picture(Me, "Icon")
    
    If InStr(gstrCommand, "=") > 0 And InStr(gstrCommand, "&") = 0 Then
        Me.Hide
    Else
        '不加这一句的话，由于已显示frmSplash窗体，在开启输入法的情况下，启动源程序，不会显示登录窗口，VB只能异常终止退出
        SetActiveWindow Me.hwnd
    End If
        
    '如果命令行参数中有用户名及密码，则填充并执行
    If gstrCommand <> "" And InStr(gstrCommand, "&") = 0 Then
        ArrCommand = Split(gstrCommand, " ")
        If UBound(ArrCommand) >= 1 Then
            If InStr(ArrCommand(0), "=") <= 0 Then
                Me.txtUser.Text = ArrCommand(0)
                Me.txtPassWord.Text = ArrCommand(1)
            End If
        ElseIf UBound(ArrCommand) = 0 Then
            '如果含有/，表示同时输入了用户名与密码，而且密码不需要进行转换
            If InStr(1, ArrCommand(0), "/") <> 0 And InStr(1, ArrCommand(0), ",") = 0 Then
                Me.txtUser.Text = Split(ArrCommand(0), "/")(0)
                Me.txtPassWord.Text = Split(ArrCommand(0), "/")(1)
                mbln转换 = False
            End If
        End If
    End If
    HookDefend txtPassWord.hwnd
    Me.Width = 6495
    Me.Height = 4440
    Exit Sub
errH:
    If CStr(gstrCommand) <> "" Then MsgBox CStr(Erl()) & "行出现错误，请手动登录！" & vbNewLine & Err.Description, vbQuestion
End Sub

Private Sub GetFocus(ByVal TxtBox As TextBox)
    With TxtBox
        If Trim(TxtBox.Text) = "" Then Exit Sub
        .SelStart = 0
        .SelLength = Len(TxtBox.Text)
    End With
End Sub

Private Sub cboServer_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then
        '回车键另行处理
        If KeyAscii <> vbKeyBack Then
            Call AppendText(KeyAscii)
        End If
    End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        ReleaseCapture
        SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    End If
End Sub

Private Sub Form_Resize()
    Me.PaintPicture Me.Picture, 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '密码不符合规则，修改密码点X，此时gcnOracle不为nothing
    If Not mblnOk Then
        If Not gcnOracle Is Nothing Then
            If gcnOracle.State = adStateOpen Then
                gcnOracle.Close
            End If
        End If
    End If
    Set mobjHttp = Nothing
    Set mcolServer = Nothing
End Sub

Private Sub imgCancel_Click()
    Set gobjRegister = Nothing
    gclsLogin.IsCancel = True
    '密码不符合规则，修改密码点取消，此时gcnOracle不为nothing
    If Not gcnOracle Is Nothing Then
        If gcnOracle.State = adStateOpen Then
            gcnOracle.Close
        End If
    End If
    Unload Me
End Sub

Private Sub imgCancel_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        imgCancel.Left = imgCancel.Left - 10
        imgCancel.Top = imgCancel.Top + 10
    End If
End Sub

Private Sub imgCancel_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        imgCancel.Left = imgCancel.Left + 10
        imgCancel.Top = imgCancel.Top - 10
    End If
End Sub

Private Sub ImgIndicate_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        ReleaseCapture
        SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    End If
End Sub

Private Sub imgPic_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        ReleaseCapture
        SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    End If
End Sub

Private Sub lblLogin_Click()
    If lblLogin.Enabled Then Call cmdOK_Click
End Sub

Private Sub lblLogin_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        lblLogin.BackColor = RGB(0, 180, 255)
        picLogin.BackColor = RGB(0, 180, 255)
        lblLogin.Left = lblLogin.Left - 10
        lblLogin.Top = lblLogin.Top + 10
    End If
End Sub

Private Sub lblLogin_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Button = vbLeftButton Then
        lblLogin.BackColor = RGB(0, 191, 255)
        picLogin.BackColor = RGB(0, 191, 255)
        lblLogin.Left = lblLogin.Left + 10
        lblLogin.Top = lblLogin.Top - 10
    End If
End Sub

Private Sub LblProductName_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        ReleaseCapture
        SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    End If
End Sub

Private Sub lbltag_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        ReleaseCapture
        SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    End If
End Sub

Private Sub picHos_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        ReleaseCapture
        SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    End If
End Sub

Private Sub picLogin_Click()
    If picLogin.Enabled Then Call cmdOK_Click
End Sub

Private Sub picLogin_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        lblLogin.BackColor = RGB(0, 180, 255)
        picLogin.BackColor = RGB(0, 180, 255)
        lblLogin.Left = lblLogin.Left - 10
        lblLogin.Top = lblLogin.Top + 10
    End If
End Sub

Private Sub picLogin_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        lblLogin.BackColor = RGB(0, 191, 255)
        picLogin.BackColor = RGB(0, 191, 255)
        lblLogin.Left = lblLogin.Left + 10
        lblLogin.Top = lblLogin.Top - 10
    End If
End Sub

Private Sub picModify_Click()
    If picModify.Enabled Then Call ModifyPWD
End Sub

Private Sub picModify_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        picModify.Left = picModify.Left - 10
        picModify.Top = picModify.Top + 10
    End If
End Sub

Private Sub picModify_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        picModify.Left = picModify.Left + 10
        picModify.Top = picModify.Top - 10
    End If
End Sub

Private Sub picSet_Click()
    If picSet.Enabled Then Call SetSourse
End Sub

Private Sub picDown_Click()

    cboServer.SetFocus
    SendMessage cboServer.hwnd, &H14F, 1, ByVal 0&
End Sub

Private Sub picSet_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        picSet.Left = picSet.Left - 10
        picSet.Top = picSet.Top + 10
    End If
End Sub

Private Sub picSet_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        picSet.Left = picSet.Left + 10
        picSet.Top = picSet.Top - 10
    End If
End Sub

Private Sub txtServer_GotFocus()
    If Me.ActiveControl Is txtServer Then
        OpenIme (False)
        If Trim(txtServer.Text) <> "" Then
            With txtServer
                .SelStart = 0
                .SelLength = Len(.Text)
            End With
        End If
    End If
End Sub

Private Sub txtServer_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then
        '回车键另行处理
        If KeyAscii <> vbKeyBack Then
            Call AppendText(KeyAscii)
        End If
    End If
End Sub

Private Sub txtUser_Change()
    If Not mblnFirst Then Exit Sub
End Sub

Private Sub txtUser_GotFocus()
    If Me.ActiveControl Is txtUser Then
        OpenIme (False)
        GetFocus txtUser
    End If
End Sub

Private Sub txtPassWord_GotFocus()
    GetFocus txtPassWord
End Sub

Private Sub SetConState(Optional ByVal BlnState As Boolean = True)
    picModify.Enabled = BlnState
    lblLogin.Enabled = BlnState
    picLogin.Enabled = BlnState
End Sub

Private Sub LoadServer()
'功能：读出本地的服务器列表
    Dim strPath As String, strFile As String, lngFile As Integer
    Dim strLine As String, lngPos As Long
    Dim strServer As String, strComputer As String, strSID As String
    Dim arrTmp As Variant
    Dim rsOraHome As ADODB.Recordset
    Dim intVersion As Integer, intTimes As Integer, intServer As Integer
    Dim i As Long, blnRead As Boolean
    Dim lngBeforeNum As Long, lngAfterNum As Long
    Dim lngFirstPos As Long, lngLastPos As Long
    Dim strChr As String, arrSer() As String

    cboServer.Clear
    '先从环境变量Tns_Admin中获取tnsnames.ora文件路径,如果没有找到,再去匹配注册表
    strPath = Environ("TNS_ADMIN")
    If strPath <> "" Then
        strFile = strPath & "\tnsnames.ora" 'Oracle 8i以上
        If Dir(strFile) = "" Then
            strFile = strPath & "NET80\ADMIN\tnsnames.ora" 'Oracle 8
        End If
        If Not gobjFile.FileExists(strFile) Then strFile = ""
    End If
    
    If strFile = "" Then
        Set rsOraHome = New ADODB.Recordset
        With rsOraHome
            .Fields.Append "Name", adVarChar, 256 'Name
            .Fields.Append "VerSion", adInteger  '版本
            .Fields.Append "Times", adInteger '第几次安装
            .Fields.Append "Server", adInteger '1-服务器,2-客户端
            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            .LockType = adLockOptimistic
            .Open
            '1:读取64位下32目录会自动定位到SOFTWARE\Wow6432Node\Oracle 2：读取32位下32位目录
            arrTmp = GetAllSubKey("HKEY_LOCAL_MACHINE\SOFTWARE\Oracle")
            If TypeName(arrTmp) = "Empty" Then
                If Is64bit Then
                    cboServer.ToolTipText = "没有找到注册表项HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Oracle！"
                Else
                    cboServer.ToolTipText = "没有找到注册表项HKEY_LOCAL_MACHINE\SOFTWARE\Oracle！"
                End If
            Else
                For i = LBound(arrTmp) To UBound(arrTmp)
                    If UCase(arrTmp(i)) Like "KEY_ORA*HOME*" Then
                        intVersion = 0: intTimes = 0:  intServer = 1
                        If GetOraInfoByRegKey(arrTmp(i), intVersion, intTimes, intServer) Then
                            .AddNew Array("Name", "VerSion", "Times", "Server"), Array("\" & arrTmp(i), intVersion, intTimes, intServer)
                            .Update
                        End If
                    End If
                Next
                If UBound(arrTmp) <> -1 Then ''顶级目录可能有Oracle_Home信息，默认读取这个
                    .AddNew Array("Name", "VerSion", "Times", "Server"), Array("", 0, 0, 1): .Update
                End If
                .Sort = "VerSion Desc,Times Desc,Server"
                Do While Not .EOF
                    strPath = ""
                    blnRead = Not GetRegValue("HKEY_LOCAL_MACHINE\SOFTWARE\Oracle" & !Name, "ORACLE_HOME", strPath)
                    blnRead = blnRead Or strPath = "" And !Name & "" = ""
                    If blnRead Then
                        Call GetRegValue("HKEY_LOCAL_MACHINE\SOFTWARE\Oracle", "ORA_CRS_HOME", strPath)
                    End If
                    If strPath <> "" Then
                        picSet.Tag = strPath '缓存OracleHome路径
                        strFile = strPath & "\network\ADMIN\tnsnames.ora" 'Oracle 8i以上
                        If Dir(strFile) <> "" Then Exit Do
                        strFile = strPath & "\NET80\ADMIN\tnsnames.ora" 'Oracle 8
                        If Dir(strFile) <> "" Then Exit Do
                    End If
                    strFile = ""
                    .MoveNext
                Loop
            End If
        End With
    End If
    If strFile = "" Then Exit Sub
    
    cboServer.ToolTipText = "服务器列表来源:" & strFile
    cboServer.Tag = strFile
    lngFile = FreeFile()
    Open strFile For Input Access Read As lngFile
    Set mcolServer = Nothing
    
    '获取tnsnames.ora文件中的内容
    Do Until EOF(lngFile)
        Input #lngFile, strLine
        strLine = ConvertStr(strLine)
         If strLine <> "" And Left(strLine, 1) <> "#" Then  '空行和注释行不取
            strServer = strServer & strLine
         End If
    Loop
    
    lngPos = 1
    Do While lngPos <> Len(strServer)   '循环每一个字符
        lngPos = lngPos + 1
        strChr = Mid(strServer, lngPos, 1)
            
        If strChr = "(" Then
            If lngFirstPos = 0 Then
                lngFirstPos = lngPos    '取第一个正括号的位置作为括号的开始位置
            End If
            
            lngBeforeNum = lngBeforeNum + 1
        ElseIf strChr = ")" Then
            lngAfterNum = lngAfterNum + 1
        End If
        
        '当正括号( 和范括号 )的个数相等,说明前后括号相匹配,可以删掉括号中的内容
        If lngBeforeNum = lngAfterNum And lngBeforeNum <> 0 Then
            lngLastPos = lngPos '取最后一个位置作为反括号的终止位置
            strServer = Replace(strServer, Mid(strServer, lngFirstPos, lngLastPos - lngFirstPos + 1), "")   '去掉括号中间的内容
            lngPos = 1
            lngBeforeNum = 0: lngAfterNum = 0
            lngFirstPos = 0: lngLastPos = 0
        End If
    Loop
    Close #lngFile
    
    If InStr(1, strServer, "(") > 0 Or InStr(1, strServer, ")") = 0 Then '
        arrSer = Split(strServer, "=")
        For i = 0 To UBound(arrSer)
            If arrSer(i) <> "" Then
                mcolServer.Add Array(arrSer(i), strComputer, strSID)
                cboServer.AddItem arrSer(i)
            End If
        Next
    End If
End Sub
Private Function GetOraInfoByRegKey(ByVal strOraHome As String, ByRef intVer As Integer, ByRef intTimes As Integer, ByRef intServer As Integer) As Boolean
'功能:通过OracleHome键获取Oracle信息
    Dim arrTmp As Variant
    Dim i As Long, blnRetrun As Boolean
    'KEY_OraDb11g_home1_32bit
    'Key_Ora*版本Home_32Bit
    'Key_Ora*版本_Home*
    arrTmp = Split(UCase(strOraHome), "_")
    For i = 1 To UBound(arrTmp)
        If arrTmp(i) Like "HOME*" Then
            intTimes = ValEx(arrTmp(2))
            blnRetrun = True
        ElseIf arrTmp(i) Like "*HOME*" Then
            intTimes = Val(Mid(arrTmp(1), InStr(UCase(arrTmp(1)), "HOME") + 4))
            blnRetrun = True
        End If
        If arrTmp(i) Like "ORADB*" Then
            intVer = ValEx(Mid(arrTmp(1), 6))
            intServer = 1
            blnRetrun = True
        ElseIf arrTmp(i) Like "ORACLIENT*" Then
            intVer = ValEx(Mid(arrTmp(1), 10))
            intServer = 2
            blnRetrun = True
        ElseIf arrTmp(i) Like "*CLIENT*" Then
            intServer = 2
            intVer = ValEx(arrTmp(i))
            blnRetrun = True
        End If
    Next
    GetOraInfoByRegKey = blnRetrun
End Function

Private Sub AppendText(KeyAscii As Integer)
'功能：向TextBox控件的Text追加内容，并根据当前Text的值在列表中检索可用的完整项目
'参数：KeyAscii    当前的按键
    Dim strTemp As String
    Dim strInput As String
    Dim lngStart As Long
    Dim varItem As Variant
    
    '首先当前用户输入的字符
    If KeyAscii < 0 Or InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789.", UCase(Chr(KeyAscii))) > 0 Then
        '输入字符只能是数字、英文和汉字
        strInput = Chr(KeyAscii)
        KeyAscii = 0
    End If
    
    With txtServer
        '记录上次的插入点位置
        lngStart = .SelStart + IIf(strInput <> "", 1, 0)
        '接着得到用户击键完成后文本框中出现的内容
        strInput = Mid(.Text, 1, .SelStart) & strInput & Mid(.Text, .SelStart + .SelLength + 1)
    End With
    '根据假想的内容得到可能的列表项
    strTemp = ""
    For Each varItem In mcolServer
        If UCase(varItem(0)) Like UCase(strInput & "*") Then
            strTemp = varItem(0)
        End If
    Next
    If strTemp <> "" Then
        cboServer.Text = strTemp
        txtServer = strTemp
        txtServer.SelStart = Len(strInput)
        txtServer.SelLength = 100
    Else
        cboServer.Text = strInput
        cboServer.SelStart = lngStart
        txtServer.Text = strInput
        txtServer.SelStart = lngStart
    End If

End Sub

Private Sub ClearComponent()
'功能：--清空注册表[本机部件]--因为不同的数据库可能使用的系统和版本不同
    If mblnFirst = True Then '启动时对控件的赋值不考虑在内
        SaveSetting "ZLSOFT", "注册信息", "本机部件", ""
    End If
End Sub

Private Function ReadINIToRec(ByVal strFile As String) As ADODB.Recordset
'功能：将指定INI配置文件的内容读取到记录集中
'返回：Nothing或包含"项目,内容"的记录集,其中同一项目可能有多行内容
    Dim rsTmp As New ADODB.Recordset
    Dim objINI As TextStream
    
    Dim strItem As String, strText As String
    Dim strLine As String
            
    rsTmp.Fields.Append "项目", adVarChar, 200
    rsTmp.Fields.Append "内容", adVarChar, 200
    rsTmp.CursorLocation = adUseClient
    rsTmp.LockType = adLockOptimistic
    rsTmp.CursorType = adOpenStatic
    rsTmp.Open
    
    Set objINI = gobjFile.OpenTextFile(strFile, ForReading)
    Do While Not objINI.AtEndOfStream
        strLine = Replace(objINI.ReadLine, vbTab, " ")
        strItem = Trim(Mid(strLine, InStr(strLine, "[") + 1, InStr(strLine, "]") - InStr(strLine, "[") - 1))
        strText = Trim(Mid(strLine, InStr(strLine, "]") + 1))
        If strItem <> "" And strText <> "" Then
            rsTmp.AddNew
            rsTmp!项目 = strItem
            rsTmp!内容 = strText
            rsTmp.Update
        End If
    Loop
    
    objINI.Close
    
    If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
    
    Set ReadINIToRec = rsTmp
End Function


Private Function SoapEnvelope(ByVal strMethod As String, ByVal parm1 As String, ByVal parm2 As String, ByVal parm3 As String, ByVal samlAssertion As String) As String
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strEnvelope As String
    
    SoapEnvelope = strEnvelope

    On Error GoTo Errhand
    
    strEnvelope = ""
    
    strEnvelope = strEnvelope & "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:Item=""http://tempuri.org/"">"
    
    If samlAssertion <> "" Then
        strEnvelope = strEnvelope & "<soapenv:Header>"
        strEnvelope = strEnvelope & "<wsse:Security xmlns:wsu=""http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd"" xmlns:wsse=""http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd"">"
        strEnvelope = strEnvelope & samlAssertion
        strEnvelope = strEnvelope & "</wsse:Security>"
        strEnvelope = strEnvelope & "</soapenv:Header>"
    End If
    
    strEnvelope = strEnvelope & "<soapenv:Body>"
    strEnvelope = strEnvelope & "<Item:" & strMethod & ">"
    Select Case strMethod
    Case "GetSAMLResponseByArtifact"
        strEnvelope = strEnvelope & "<Item:artifact>" & parm1 & "</Item:artifact>"
    Case "AddUserAppInfo"
        strEnvelope = strEnvelope & "<Item:account>" & parm1 & "</Item:account>"
        strEnvelope = strEnvelope & "<Item:appID>" & parm2 & "</Item:appID>"
        strEnvelope = strEnvelope & "<Item:appInfo>" & parm3 & "</Item:appInfo>"
    End Select
    strEnvelope = strEnvelope & "</Item:" & strMethod & ">"
    strEnvelope = strEnvelope & "</soapenv:Body>"
    strEnvelope = strEnvelope & "</soapenv:Envelope>"
    
    
    SoapEnvelope = strEnvelope
   
    Exit Function
Errhand:
    
End Function

Private Function PostData(ByVal strPostURL As String, _
                        ByVal strMethod As String, _
                        ByVal strPostContent As String, _
                        Optional ByVal intSendWaitTime As Integer = 30) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim lngWaitTimeOut As Long
    Dim oXmlDoc As Object
    Dim strPostCookie As String
    
    On Error GoTo Errhand
        
    If UCase(Left(strPostURL, 4)) <> "HTTP" Then strPostURL = "http://" & strPostURL
    strPostCookie = "ASPSESSIONIDAQACTAQB=HKFHJOPDOMAIKGMPGBJJDKLJ;"
    
    strPostCookie = Replace(strPostCookie, Chr(32), "%20")
    With mobjHttp
        Call .Open("POST", strPostURL, True)
        Select Case strMethod
        Case "GetSAMLResponseByArtifact"
            Call .setRequestHeader("SOAPAction", "http://tempuri.org/ISSOService/GetSAMLResponseByArtifact")
        Case "AddUserAppInfo"
            Call .setRequestHeader("SOAPAction", "http://tempuri.org/IAccountService/AddUserAppInfo")
        End Select
        Call .setRequestHeader("Content-Length", LenB(strPostContent))
        Call .setRequestHeader("Content-Type", "text/xml; charset=utf-8")
        Call .send(strPostContent)
    End With
    lngWaitTimeOut = 0
'    lngSecondNumber = 30 '超时多少秒
    Do
        DoEvents
        Call Wait(10)
        lngWaitTimeOut = lngWaitTimeOut + 1
    Loop Until (mobjHttp.readyState = 4 Or lngWaitTimeOut >= 100 * intSendWaitTime)
    
    If mobjHttp.readyState = 4 Then
        Set oXmlDoc = CreateObject("MSXML2.DOMDocument")

        oXmlDoc.Load mobjHttp.ResponseXML
        If oXmlDoc.xml = "" Then
            mstrError = mobjHttp.responseText
            PostData = False
        Else
            mstrPostData = oXmlDoc.xml
            PostData = True
        End If
    Else
        mstrError = mobjHttp.responseText
        PostData = False
    End If
    Exit Function
    
Errhand:
    mstrError = Err.Description
End Function


Private Sub Wait(tt)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim t, t1, t2, i
    t = tt
    If t > 10 Then
        t1 = Int(t / 10)
        t2 = t - t1 * 10
        For i = 1 To t1
            Call OSWait(10)
            DoEvents
        Next i
        If t2 > 0 Then Call OSWait(t2)
    Else
        If t > 0 Then Call OSWait(t)
    End If
End Sub

Private Sub ClearValues()
    '清理变量
    mblnFirst = False
    mintTimes = 1
    mbln转换 = True
    mblnAccess = False
    mblnUAAddUser = False
    
    mstrPostData = ""
    mstr断言 = ""
    mstrUserURL = ""
    mstrSamlAssertion = ""
    mstrError = ""
    mblnZLUA = False
    mstrAppID = ""
    mstrZLUAUser = ""
    mblnOk = False
End Sub

Public Function ShowMe() As Boolean
    '清理变量
    Call ClearValues
    Me.Show vbModal
End Function

Public Function Docmd(ByVal strCmd As String) As Boolean
    Dim ArrCommand
    Dim ArrCommandPortal
    Dim objSoap As Object
    Dim objDoc As Object
    Dim rsIni As ADODB.Recordset
    Dim strIp As String
    Dim strList As String
    Dim strResult As String
    Dim i As Integer
    Dim strPortURL As String
    Dim ResponseXML As Object
    Dim ResponseNode As Object
    Dim strArtifact助诊符 As String
    Dim strStatus As String
    Dim strSoapPost As String
    Dim strErr As String
    Dim strAppStart As String
    On Error GoTo Errhand
    '清理变量
    Call ClearValues
    'ZLUA登录
    strAppStart = gobjFile.GetParentFolderName(App.Path)
    If Len(strCmd) > 0 And InStr(strCmd, ",") = 0 And InStr(gstrCommand, "&") > 0 Then
        
        If Not gobjFile.FileExists(strAppStart & "\" & "ZLUA.ini") Then
            MsgBox "未找到" & strAppStart & "\" & "ZLUA.ini，无法读取配置文件", vbInformation + vbOKOnly, "提示"
            GoTo Errhand
        End If
        Set rsIni = ReadINIToRec(strAppStart & "\" & "ZLUA.ini")
        rsIni.Filter = ""
        rsIni.Filter = "项目='PortURL'"
        strPortURL = rsIni("内容").Value
        rsIni.Filter = ""
        rsIni.Filter = "项目='UserURL'"
        mstrUserURL = rsIni("内容").Value
        rsIni.Filter = "项目='AppID'"
        mstrAppID = rsIni("内容").Value
        
        strArtifact助诊符 = Split(gstrCommand, "&")(0)
        
        If Trim(strPortURL) = "" Then
            MsgBox "请配置单点登录服务地址", vbInformation + vbOKOnly, "提示"
        ElseIf (Trim(mstrUserURL) = "") Then
            MsgBox "请配置账户服务地址", vbInformation + vbOKOnly, "提示"
        Else
            '采用httprequest方式-----------------
            mstr断言 = SoapEnvelope("GetSAMLResponseByArtifact", strArtifact助诊符, "", "", "")
            Call PostData(strPortURL, "GetSAMLResponseByArtifact", mstr断言, 5)
            strSoapPost = mstrPostData
            strSoapPost = Replace(strSoapPost, "&gt;", ">")
            strSoapPost = Replace(strSoapPost, "&lt;", "<")
            
            '-------------
            '解析XML文本内容并判断是否返回正确验证结果
            If strSoapPost <> "" Then
                Set objDoc = CreateObject("MSXML2.DOMDocument")
                Call objDoc.loadXML(strSoapPost)
                Set ResponseXML = objDoc.documentElement
                Set ResponseNode = ResponseXML.selectSingleNode(".//samlp:StatusCode")
                strStatus = ResponseNode.Attributes(0).Text
                If strStatus <> "" Then
                    Select Case strStatus
                    Case "urn:oasis:names:tc:SAML:2.0:status:Success"
                        '令牌请求成功
                        '获取登录信息:用户名/密码/服务器
                        Set ResponseNode = ResponseXML.selectSingleNode(".//saml:AttributeValue")
                        If ResponseNode Is Nothing Then
                            strStatus = ""
                        Else
                            strStatus = ResponseNode.Text
                        End If
                        
                        '获取ZLUA账户名
                        Set ResponseNode = ResponseXML.selectSingleNode(".//saml:NameID")
                        mstrZLUAUser = ResponseNode.Text
                        
                        Set ResponseNode = ResponseXML.selectSingleNode(".//saml:Assertion")
                        mstrSamlAssertion = ResponseNode.xml
                        '如果信息为空，则显示登录信息框，并调用接口上传信息以便下次成功获取
                        mblnZLUA = True
                        If Trim(strStatus) = "" Then
                            mblnUAAddUser = True
                            '--测试添加ZLUA用户账户
                        Else
                            If InStr(strStatus, "/") > 0 And InStr(strStatus, "@") > 0 And InStr(strStatus, "/") < InStr(strStatus, "@") Then
                               Me.txtUser.Text = Mid(strStatus, 1, InStr(strStatus, "/") - 1)
                               Me.txtPassWord.Text = Mid(strStatus, InStr(strStatus, "/") + 1, InStr(strStatus, "@") - InStr(strStatus, "/") - 1)
                               Me.cboServer.Text = Mid(strStatus, InStr(strStatus, "@") + 1)
                               txtServer.Text = Mid(strStatus, InStr(strStatus, "@") + 1)
                            End If
                            If Trim(txtUser.Text) <> "" And Trim(txtPassWord.Text) <> "" Then cmdOK_Click
                        End If
                    Case Else
                        '令牌请求失败，重新获取三级错误信息
                        Set ResponseNode = ResponseXML.selectSingleNode(".//samlp:StatusMessage")
                        strStatus = ResponseNode.Text
                        strErr = "错误信息：" & strStatus
                        GoTo Errhand
                    End Select
                End If
            End If
            
        End If
    End If

    '单点登录
    ReDim ArrCommandPortal(0)
    If InStr(strCmd, ",") > 0 Then
        If objSoap Is Nothing Then
            Set objSoap = CreateObject("MSSOAP.SoapClient30")
        End If
        
        If Err.Number <> 0 Then
            Screen.MousePointer = 0
            Err.Clear
            MsgBox "无法创建SOAP对象！", vbOKOnly + vbInformation, "提示"
            Set objSoap = Nothing
            GoTo Errhand
        End If
        If Not gobjFile.FileExists(strAppStart & "\" & "Portal.ini") Then
            MsgBox "未找到 " & strAppStart & "\" & "Portal.ini 路径", vbInformation + vbOKOnly, "提示"
            GoTo Errhand
        End If
        Set rsIni = ReadINIToRec(strAppStart & "\" & "Portal.ini")
        rsIni.Filter = ""
        rsIni.Filter = "项目='IP'"
        strIp = rsIni("内容").Value
        rsIni.Filter = ""
        rsIni.Filter = "项目='List'"
        strList = rsIni("内容").Value
        '以前丢失，10.35.10新增
        ArrCommandPortal = Split(strCmd, ",")
    End If
    
    ArrCommand = Split(strCmd, " ")
    
    If UBound(ArrCommandPortal) > 0 Then
        Call objSoap.MSSoapInit("http://" & strIp & "/" & strList & "?wsdl")
        strResult = objSoap.getZLSSORet(ArrCommandPortal(0), ArrCommandPortal(1))
        If strResult <> "" And InStr(strResult, "/") > 0 And InStr(strResult, "@") > 0 And InStr(strResult, "/") < InStr(strResult, "@") Then
           Me.txtUser.Text = Mid(strResult, 1, InStr(strResult, "/") - 1)
           Me.txtPassWord.Text = Mid(strResult, InStr(strResult, "/") + 1, InStr(strResult, "@") - InStr(strResult, "/") - 1)
           Me.cboServer.Text = Mid(strResult, InStr(strResult, "@") + 1)
           txtServer.Text = Mid(strResult, InStr(strResult, "@") + 1)
        End If
        mbln转换 = True
        If Trim(txtUser.Text) <> "" And Trim(txtPassWord.Text) <> "" Then cmdOK_Click
    ElseIf InStr(ArrCommand(0), "=") > 0 And InStr(ArrCommand(0), "&") = 0 Then
        '第三方部件调用导航台登录的格式
        For i = LBound(ArrCommand) To UBound(ArrCommand)
            If UCase(ArrCommand(i)) Like "USER=*" Then
                Me.txtUser.Text = Mid(ArrCommand(i), Len("USER=*"))
            ElseIf UCase(ArrCommand(i)) Like "PASS=*" Then
                Me.txtPassWord.Text = Mid(ArrCommand(i), Len("PASS=*"))
            ElseIf UCase(ArrCommand(i)) Like "SERVER=*" Then
                Me.cboServer.Text = Mid(ArrCommand(i), Len("SERVER=*"))
                txtServer.Text = Mid(ArrCommand(i), Len("SERVER=*"))
            ElseIf UCase(ArrCommand(i)) Like "ONLYONE=*" Then
                If Split(ArrCommand(i), "=")(1) = "1" Then
                    If AppPrevInstance = True Then
                        MsgBox "不能重复运行这个程序！"
                        gblnExitApp = True
                        Exit Function
                    End If
                End If
            ElseIf UCase(ArrCommand(i)) Like "HELPERMAIN=*" Then
                glngHelperMainType = Val(Mid(ArrCommand(i), Len("HELPERMAIN=*")))
            ElseIf UCase(ArrCommand(i)) Like "ISDBPASS=*" Then
                glngDBPass = Val(Mid(ArrCommand(i), Len("ISDBPASS=*")))
            ElseIf UCase(ArrCommand(i)) Like "PARALLELID=*" Then
                glngParallelID = Val(Mid(ArrCommand(i), Len("PARALLELID=*")))
            End If
        Next
        If glngDBPass > 0 Then
            mbln转换 = glngDBPass = 2
        End If
        If Trim(txtUser.Text) <> "" And Trim(txtPassWord.Text) <> "" Then Call cmdOK_Click
    End If
    Docmd = mblnAccess
    Set objSoap = Nothing
    Exit Function
Errhand:
    If strErr <> "" Then
        MsgBox strErr, vbInformation + vbOKOnly, "提示"
        strErr = ""
    Else
        If Err.Number <> 0 Then
            MsgBox Err.Description, vbInformation + vbOKOnly, "提示"
        End If
    End If
    Set objSoap = Nothing
    Err.Clear
End Function

Private Sub txtUser_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = "'" Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtUser_LostFocus()
    Call UpdateUser
End Sub

Private Sub txtUser_Validate(Cancel As Boolean)
    Call UpdateUser
End Sub

Private Sub UpdateUser()
On Error GoTo errH
    If IsNumeric(txtUser.Text) Then
        txtUser.Text = "U" & txtUser.Text
    End If
    Exit Sub
errH:
    MsgBox Err.Description, vbCritical, gstrSysName
    Err.Clear
End Sub

Private Function CheckInput(ByRef strUserName As String, ByRef strPassword As String, ByRef strServerName As String) As String
'功能:检查用户，密码，服务器的输入值
    '分离字符串
    Dim intPos As Integer
    
    '------检验用户是否oracle合法用户----------------
    strUserName = Trim(txtUser.Text)
    strPassword = Trim(txtPassWord.Text)
    strServerName = Trim(txtServer.Text)
    
    '有效字符串效验
    If Len(Trim(txtUser.Text)) = 0 Then
        CheckInput = "请输入用户名"
        txtUser.SetFocus
        Exit Function
    End If
    
    If Len(strUserName) <> 1 Then
        If Mid(strUserName, 1, 1) = "/" Or Mid(strUserName, 1, 1) = "@" Or Mid(strUserName, Len(strUserName) - 1, 1) = "/" Or Mid(strUserName, Len(strUserName) - 1, 1) = "@" Then
            txtUser.SetFocus
            CheckInput = "用户名错误"
            SetConState
            Exit Function
        End If
    End If
    
    If Trim(strPassword) <> "" And Len(strPassword) <> 1 Then
        If Mid(strPassword, Len(strPassword) - 1, 1) = "/" Or Mid(strPassword, Len(strPassword) - 1, 1) = "@" Or Mid(strPassword, 1, 1) = "/" Or Mid(strPassword, 1, 1) = "@" Then
            If txtPassWord.Enabled Then txtPassWord.SetFocus
            CheckInput = "密码错误"
            Exit Function
        End If
    End If
    If Trim(strServerName) <> "" Then
        If Mid(strServerName, Len(strServerName) - 1, 1) = "/" Or Mid(strServerName, Len(strServerName) - 1, 1) = "@" Or Mid(strServerName, 1, 1) = "/" Or Mid(strServerName, 1, 1) = "@" Then
            CheckInput = "主机连接串错误"
            cboServer.SetFocus
            Exit Function
        End If
    End If
    
    intPos = InStr(strUserName, "@")
    If intPos > 0 Then
        strServerName = Mid(strUserName, intPos + 1)
        strUserName = Mid(strUserName, 1, intPos - 1)
    End If
    
    intPos = InStr(strUserName, "/")
    If intPos > 0 Then
        strPassword = Mid(strUserName, intPos + 1)
        strUserName = Mid(strUserName, 1, intPos - 1)
    End If
    
    intPos = InStr(strPassword, "@")
    If intPos > 0 Then
        strServerName = Mid(strPassword, intPos + 1)
        strPassword = Mid(strPassword, 1, intPos - 1)
    End If
    If Len(Trim(strPassword)) = 0 Then
        CheckInput = "请输入密码"
    End If
End Function

Private Function ExecuteCommand(ByVal strCommand As String) As Boolean
'功能：执行指定命令
    Dim lngShell As Long
    
    On Error Resume Next
    lngShell = Shell(strCommand, vbNormalFocus)
    
    If Err <> 0 Then
        Exit Function
    End If
    
    ExecuteCommand = True
End Function

Private Sub InitFaceType()
    picModify.Enabled = gintCallType = 0
    picModify.Visible = gintCallType = 0
    picSet.Enabled = gintCallType = 1
    picSet.Visible = gintCallType = 1
End Sub

Private Function CheckPwdExpiry() As Boolean
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim dtExpiryDate As Date
    Dim dtNow As Date
    Dim intDiff As Integer
    
    On Error GoTo errH
    strSQL = "Select EXPIRY_DATE,sysdate curDate From User_Users Where UserName=[1]"
    Set rsData = OpenSQLRecord(strSQL, "检查密码期效", gstrDBUser)
    
    If rsData.BOF = False Then
        If IsNull(rsData("EXPIRY_DATE").Value) = True Then
            CheckPwdExpiry = False
            Exit Function
        End If
        dtExpiryDate = Format(rsData!EXPIRY_DATE, "YYYY-MM-DD HH:MM:SS")
        '判断过期日期与当前日期相差天数
        dtNow = Format(rsData!curDate, "YYYY-MM-DD HH:MM:SS")
       
        intDiff = DateDiff("d", dtNow, dtExpiryDate)
        
        If intDiff > 7 Then
            CheckPwdExpiry = False
            Exit Function
        End If
        
        If intDiff > 3 And intDiff <= 7 Then
            '提示修改密码
            If MsgBox("密码有效期还有" & intDiff & "天,是否立即修改密码?", vbQuestion + vbYesNo, "密码过期提醒") = vbYes Then
                CheckPwdExpiry = True
            Else
                CheckPwdExpiry = False
                Exit Function
            End If
        ElseIf intDiff <= 3 Then
            CheckPwdExpiry = True
            MsgBox "密码有效期还有" & intDiff & "天，请您立即修改密码。", vbInformation
        Else
            CheckPwdExpiry = False
            Exit Function
        End If
    End If
    Exit Function
errH:
    MsgBox Err.Description, vbCritical, Me.Caption
End Function

Private Function ConvertStr(ByVal strSource As String) As String
    '功能:去掉字符串的空格\换行符,并转换为大写
    
    strSource = UCase(strSource)
    strSource = Replace(strSource, " ", "")
    strSource = Replace(strSource, vbNewLine, "")
    strSource = Replace(strSource, vbCr, "")
    strSource = Replace(strSource, vbLf, "")
    strSource = Replace(strSource, vbTab, "")
    strSource = Replace(strSource, vbBack, "")
    ConvertStr = strSource
End Function

