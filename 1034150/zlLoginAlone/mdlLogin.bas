Attribute VB_Name = "mdlLogin"
Option Explicit
'进程获取
Public Type MODULEENTRY32
    dwSize                                      As Long
    th32ModuleID                                As Long
    th32ProcessID                               As Long
    GlblcntUsage                                As Long
    ProccntUsage                                As Long
    modBaseAddr                                 As Byte
    modBaseSize                                 As Long
    hModule                                     As Long
    szModule                                    As String * 256
    szExePath                                   As String * 1024
End Type

Public Type PROCESSENTRY32
      lSize                                     As Long
      lUsage                                    As Long
      lProcessId                                As Long
      lDefaultHeapId                            As Long
      lModuleId                                 As Long
      lThreads                                  As Long
      lParentProcessId                          As Long
      lPriClassBase                             As Long
      lFlags                                    As Long
      sExeFile                                  As String * 1024
End Type
Public Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Public Const TH32CS_SNAPPROCESS                As Long = &H2
Public Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Public Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Public Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Public Declare Function Module32First Lib "kernel32" (ByVal hSnapshot As Long, lppe As MODULEENTRY32) As Long
Public Declare Function Module32Next Lib "kernel32" (ByVal hSnapshot As Long, lppe As MODULEENTRY32) As Long
Public Const TH32CS_SNAPMODULE                 As Long = &H8
Public Const SYNCHRONIZE                       As Long = &H100000
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'是否是64位进程（Is64bit）
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function IsWow64Process Lib "kernel32" (ByVal hProc As Long, bWow64Process As Long) As Long
'返回系统中可用的输入法个数及各输入法所在Layout,包括英文输入法。
Private Declare Function GetKeyboardLayoutList Lib "user32" (ByVal nBuff As Long, lpList As Long) As Long
'获取某个输入法的名称
Private Declare Function ImmGetDescription Lib "imm32.dll" Alias "ImmGetDescriptionA" (ByVal hkl As Long, ByVal lpsz As String, ByVal uBufLen As Long) As Long
'判断某个输入法是否中文输入法
Private Declare Function ImmIsIME Lib "imm32.dll" (ByVal hkl As Long) As Long
'切换到指定的输入法。
Private Declare Function ActivateKeyboardLayout Lib "user32" (ByVal hkl As Long, ByVal flags As Long) As Long
'电脑名称(ComputerName)
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
'暂停(Wait)
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Const GWL_EXSTYLE = (-20)
Public Const WinStyle = &H40000 'Forces a top-level window onto the taskbar when the window is visible.强制一个可见的顶级视窗到工具栏上
Public Const SWP_NOSIZE = &H1
Public Const SWP_SHOWWINDOW = &H40
Public Const HWND_TOPMOST = -1
'临时IP获取
Private Const MAX_IP = 5                                                    'To make a buffer... i dont think you have more than 5 ip on your pc..
Private Type IPINFO
    dwAddr As Long                                                          ' IP address
    dwIndex As Long                                                         ' interface index
    dwMask As Long                                                          ' subnet mask
    dwBCastAddr As Long                                                     ' broadcast address
    dwReasmSize  As Long                                                    ' assembly size
    unused1 As Integer                                                      ' not currently used
    unused2 As Integer                                                      '; not currently used
End Type
Private Type MIB_IPADDRTABLE
    dEntrys As Long                                                         'number of entries in the table
    mIPInfo(MAX_IP) As IPINFO                                               'array of IP address entries
End Type
Private Declare Function GetIpAddrTable Lib "IPHlpApi" (pIPAdrTable As Byte, pdwSize As Long, ByVal Sort As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
'SM4加密
'/**
' * \brief          SM4-ECB block encryption/decryption
' * \param mode     SM4_ENCRYPT or SM4_DECRYPT
' * \param length   length of the input data
' * \param input    input block
' * \param output   output block
' */
Private Declare Function sm4_crypt_ecb Lib "zlSm4.dll" (ByVal Mode As Long, ByVal Length As Long, key As Byte, in_put As Byte, out_put As Byte) As Long
'SM4分组密码加密
'/**
' * \brief          SM4-CBC buffer encryption/decryption
' * \param mode     SM4_ENCRYPT or SM4_DECRYPT
' * \param length   length of the input data
' * \param iv       initialization vector (updated after use)
' * \param input    buffer holding the input data
' * \param output   buffer holding the output data
' */
Private Declare Function sm4_crypt_cbc Lib "zlSm4.dll" (ByVal Mode As Long, ByVal Length As Long, iv As Byte, key As Byte, in_put As Byte, out_put As Byte) As Long
'获取字符串的哈希编码
'/**
' * \brief          Output = SM3( input buffer )
' *
' * \param input    buffer holding the  data
' * \param ilen     length of the input data
' * \param output   SM3 checksum result
' */
Private Declare Sub sm3_hash Lib "zlSm4.dll" Alias "sm3" (in_put As Byte, ByVal Length As Long, out_put As Byte)
'获取文件的sm哈希编码
'/**
' * \brief          Output = SM3( file contents )
' *
' * \param path     input file name
' * \param output   SM3 checksum result
' *
' * \return         0 if successful, 1 if fopen failed,
' *                 or 2 if fread failed
' */
Private Declare Function sm3_file_hash Lib "zlSm4.dll" Alias "sm3_file" (in_path As Byte, out_put As Byte) As Long
'HMAC是密钥相关的哈希运算消息认证码，HMAC运算利用哈希算法，以一个密钥和一个消息为输入，生成一个消息摘要作为输出。
'/**
' * \brief          Output = HMAC-SM3( hmac key, input buffer )
' *
' * \param key      HMAC secret key
' * \param keylen   length of the HMAC key
' * \param input    buffer holding the  data
' * \param ilen     length of the input data
' * \param output   HMAC-SM3 result
' */
Private Declare Sub sm3_hmac_hash Lib "zlSm4.dll" Alias "sm3_hmac" (key As Byte, ByVal keylen As Long, in_put As Byte, ByVal InputLen As Long, out_put As Byte)
'获取ZLSM4的修改版本
'1:只支持sm4_crypt_ecb,sm4_crypt_cbc
'2:增加支持sm3，sm3_file，sm3_hmac，sm_version
'/**
' * \brief          Output = zlSM4.DLL Version
' */
Private Declare Function get_sm_version Lib "zlSm4.dll" Alias "sm_version" () As Long

Private Enum CrypeMode
    CM_Encrypt = 1   '加密
    CM_Decrypt = 0   '解密
End Enum
Private M_SM4_VERSION As Long
Public Const SM4_CRYPT_RANDOMIZE_KEY    As Long = 999  'sm4加密算法密钥生成器的随机种子
Public Const SM4_CRYPT_RANDOMIZE_IV     As Long = 666   'sm4加密算法初始向量生成器的随机种子
Public Const G_PASSWORD_KEY             As String = "3357F1F2CA0341A5B75DBA7F35666715"

'注册表关键字根类型
Private Enum REGRoot
    HKEY_CLASSES_ROOT = &H80000000 '记录Windows操作系统中所有数据文件的格式和关联信息，主要记录不同文件的文件名后缀和与之对应的应用程序。其下子键可分为两类，一类是已经注册的各类文件的扩展名，这类子键前面都有一个“。”；另一类是各类文件类型有关信息。
    HKEY_CURRENT_USER = &H80000001 '此根键包含了当前登录用户的用户配置文件信息。这些信息保证不同的用户登录计算机时，使用自己的个性化设置，例如自己定义的墙纸、自己的收件箱、自己的安全访问权限等。
    HKEY_LOCaL_MaCHINE = &H80000002 '此根键包含了当前计算机的配置数据，包括所安装的硬件以及软件的设置。这些信息是为所有的用户登录系统服务的。它是整个注册表中最庞大也是最重要的根键！
    HKEY_USERS = &H80000003 '此根键包括默认用户的信息（Default子键）和所有以前登录用户的信息。
    HKEY_PERFORMANCE_DATA = &H80000004 '在Windows NT/2000/XP注册表中虽然没有HKEY_DYN_DATA键，但是它却隐藏了一个名为“HKEY_ PERFOR MANCE_DATA”键。所有系统中的动态信息都是存放在此子键中。系统自带的注册表编辑器无法看到此键
    HKEY_CURRENT_CONFIG = &H80000005  '此根键实际上是HKEY_LOCAL_MACHINE中的一部分，其中存放的是计算机当前设置，如显示器、打印机等外设的设置信息等。它的子键与HKEY_LOCAL_ MACHINE\ Config\0001分支下的数据完全一样。
    HKEY_DYN_DATA = &H80000006 '此根键中保存每次系统启动时，创建的系统配置和当前性能信息。这个根键只存在于Windows 98中。
End Enum

'注册表数据类型
Private Enum REGValueType
    REG_NONE = 0                       ' No value type
    REG_SZ = 1 'Unicode空终结字符串
    REG_EXPAND_SZ = 2 'Unicode空终结字符串
    REG_BINARY = 3 '二进制数值
    REG_DWORD = 4 '32-bit 数字
    REG_DWORD_BIG_ENDIAN = 5
    REG_LINK = 6
    REG_MULTI_SZ = 7 ' 二进制数值串
End Enum
'打开错误
Private Enum REGErr
    ERROR_SUCCESS = &H0
    ERROR_FILE_NOT_FOUND = &H2 'The system cannot find the file specified
    ERROR_BADDB = 1009&
    ERROR_BADKEY = 1010&
    ERROR_CANTOPEN = 1011&
    ERROR_CANTREAD = 1012&
    ERROR_CANTWRITE = 1013&
    ERROR_OUTOFMEMORY = 14&
    ERROR_INVALID_PARAMETER = 87&
    ERROR_ACCESS_DENIED = 5&
    ERROR_NO_MORE_ITEMS = 259&
    ERROR_MORE_DATA = 234&
End Enum
'注册表访问权
Private Enum REGRights
    KEY_QUERY_VaLUE = &H1
    KEY_SET_VaLUE = &H2
    KEY_CREaTE_Sub_KEY = &H4
    KEY_ENUMERaTE_Sub_KEYS = &H8
    KEY_NOTIFY = &H10
    KEY_CREaTE_LINK = &H20
    KEY_aLL_aCCESS = &H3F
    KEY_READ = &H20019
End Enum
' 扩充环境字符串。具体操作过程与命令行处理的所为差不多。也就是说，将由百分号封闭起来的环境变量名转换成那个变量的内容。比如，“%path%”会扩充成完整路径。
Private Declare Function ExpandEnvironmentStrings Lib "kernel32" Alias "ExpandEnvironmentStringsA" (ByVal lpSrc As String, ByVal lpDst As String, ByVal nSize As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal uloptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Private Declare Function RegQueryValueEx_ValueType Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegQueryValueEx_Long Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
Private Declare Function RegQueryValueEx_String Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegQueryValueEx_BINARY Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Private Declare Function RegSetValueEx_String Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal lpcbData As Long) As Long
Private Declare Function RegSetValueEx_Long Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long
Private Declare Function RegSetValueEx_BINARY Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Byte, ByVal cbData As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long

Public gdtStart As Long
Public gstrSysName As String
Public gobjRegister As Object               '注册授权部件zlRegister
Public gcnOracle As ADODB.Connection     '公共数据库连接
Public gstrCommand As String '命令行

Public gobjFile As New FileSystemObject
Public gclsLogin As clsLogin '登录对象
Public gintCallType As Integer '0-不展示修改密码与服务器配置,1-显示修改密码,2-现实服务器配置
Public gblnExitApp  As Boolean '是否因为重复运行，需要退出整个程序

'clsLogin属性缓存
Public gobjEmr             As Object   'EMR新版电子病历

Public gstrInputPwd        As String   'InputPwd属性
Public gstrServerName      As String   'ServerName属性
Public gblnTransPwd        As Boolean  'blnTransPwd属性

Public gstrInputUser        As String  '输入的用户名，例：zyk，未转换大小写
Public gstrDBUser           As String  '登录用户名，例：ZYK，大写
Public gstrUserID           As String  '登录用户ID，例：133
Public gstrUserName         As String  '登录人员姓名，例：张永康
Public gstrDeptName         As String  '登录用户的缺省部门名称
Public gstrDeptNameTerminal As String  '用户登录机器所属的部门名称
Public gstrDeptID           As String  '登录用户缺省部门ID
Public gstrIP               As String  '登录客户端IP地址
Public gstrSessionID        As String  '当前会话ID

Public gblnSysOwner        As Boolean  '是否系统所有者
Public gstrConnString      As String   '连接字符串
Public gstrSystems         As String   '多帐套选择的系统
Public gblnCancel          As Boolean  '是否取消退出
Public gstrMenuGroup       As String   '菜单组名称

Public gstrStation         As String   '用户登录工作站名称
Public gstrNodeNo          As String   '站点编号
Public gstrNodeName        As String   '站点名称
Public gblnEMRProxy         As Boolean
Public gstrEMRPwd           As String
Public gstrEMRUser          As String
Public gstrUsageOccasion    As String      '使用场合，主要用于体检，LIS单独登录。外部设置该变量
Public glngHelperMainType   As Long         '升级助手任务类型
Public glngDBPass           As Long         '0-自动判断，1-数据库密码，2-非数据库密码
Public glngParallelID       As Long         '客户端功能验证并行处理ID
Public gblnTimer            As Boolean  '是否定时器触发的客户端更新检查
Public glngInstanceCount    As Long     '实例计数
Public gobjComlib           As Object
Public gcolTableField       As Collection

Public Sub CollectTableField(ByVal strInfo As String)
'功能：收集表字段是否存在，并存入变量中
'参数：
'  strInfo：指定要收集的表名、字段名。  格式：表名1.字段名1[,表名2.字段名2 ...]
    
    Dim arrItems() As String, strTable As String, strField As String, strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo hErr
    
    arrItems = Split(strInfo, ",")
    For i = LBound(arrItems) To UBound(arrItems)
        If arrItems(i) Like "*.*" Then
            strTable = UCase(Trim(Split(arrItems(i), ".")(0)))
            strField = UCase(Trim(Split(arrItems(i), ".")(1)))
            strSQL = _
                "Select Table_Name, Column_Name " & vbNewLine & _
                "From All_Tab_Columns " & vbNewLine & _
                "Where Owner = 'ZLTOOLS' And Table_Name = [1] And Column_Name = [2] "
            Set rsTemp = OpenSQLRecord(strSQL, "收集字段是否存在", strTable, strField)
            If rsTemp.EOF = False Then
                If gcolTableField Is Nothing Then
                    Set gcolTableField = New Collection
                End If
                
                On Error Resume Next
                gcolTableField.Add "1", rsTemp!Table_Name & rsTemp!Column_Name
                On Error GoTo hErr
            End If
            rsTemp.Close
        End If
    Next
    
    Exit Sub
    
hErr:
    MsgBox Err.Number & "： " & Err.Description, vbInformation, App.Title
End Sub

Public Function GetExistsField(ByVal strTable As String, ByVal strField As String) As Boolean
'功能：检查字段是否存在（兼容性）
'返回：True存在；False不存在

    If gcolTableField Is Nothing Then Exit Function
    
    strTable = Trim(UCase(strTable))
    strField = Trim(UCase(strField))
    
    GetExistsField = False
    On Error Resume Next
    GetExistsField = Val(gcolTableField(strTable & "_" & strField)) = 1
    On Error GoTo 0
End Function

Public Sub SetAppBusyState()
'当其他进程对象未创建完成时，替换在执行主进程功能时弹出的“部件被挂起”对话框
    On Error Resume Next
    App.OleServerBusyMsgTitle = App.ProductName
    App.OleRequestPendingMsgTitle = App.ProductName
    
    App.OleServerBusyMsgText = "相关组件正在创建，请耐心等待。"
    App.OleRequestPendingMsgText = "相关组件正创建，请耐心等待。"
    
    App.OleServerBusyTimeout = 3000
    App.OleRequestPendingTimeout = 10000
    Err.Clear
End Sub

Public Function ShowSplash(Optional ByVal bytType As Byte, Optional ByVal blnRefresh As Boolean) As Boolean
'bytType:0-新窗体；1-老窗体
    Dim strUnitName As String, intCount As Integer
    Dim objPic As IPictureDisp
    '由注册表中获取用户注册相关信息,如果用户单位名称不为空,则显示闪现窗体
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "提示", "")
    strUnitName = GetSetting("ZLSOFT", "注册信息", "单位名称", "")
    
    If bytType = 0 Then
        With frmUserLogin
            If strUnitName <> "" And strUnitName <> "-" Then
                    '有两处需要处理
                    '此时就开始创建clsComLib类实例
                    Call ApplyOEM_Picture(.ImgIndicate, "Picture")
                    Call ApplyOEM_Picture(.imgPic, "PictureB")
                    If gobjFile.FileExists(gstrSetupPath & "\附加文件\logo_login.jpg") Then
                        Set objPic = LoadPicture(gstrSetupPath & "\附加文件\logo_login.jpg")
                        .picHos.Visible = True
                        .picHos.Height = IIf(objPic.Height < 2385, objPic.Height, 2385) '159像素
                        .picHos.Width = IIf(objPic.Width < 4500, objPic.Width, 4730) '322像素
                        .picHos.PaintPicture objPic, 0, 0, .picHos.Width, .picHos.Height
                    Else
                        .picHos.Visible = False
                    End If
                    .LblProductName = GetSetting("ZLSOFT", "注册信息", "产品全称", "")
                    If Len(.LblProductName) > 10 Then
                        .LblProductName.FontSize = 15.75 '三号
                    Else
                        .LblProductName.FontSize = 21.75 '二号
                    End If
                    .lbltag = GetSetting("ZLSOFT", "注册信息", "产品系列", "")
            Else
                .picHos.Visible = False
                .LblProductName.Visible = False
                .lbltag.Visible = False
            End If
        End With
    Else
        If blnRefresh Then
            With frmSplash
                .lblGrant = Replace(strUnitName, ";", vbCrLf)
                .lbl技术支持商.Caption = GetSetting("ZLSOFT", "注册信息", "技术支持商", "")
                
                .LblProductName = GetSetting("ZLSOFT", "注册信息", "产品全称", "")
                .lbltag = GetSetting("ZLSOFT", "注册信息", "产品系列", "")
                strUnitName = GetSetting("ZLSOFT", "注册信息", "开发商", "")
                .lbl开发商.Caption = ""
                For intCount = 0 To UBound(Split(strUnitName, ";"))
                    .lbl开发商.Caption = .lbl开发商.Caption & Split(strUnitName, ";")(intCount) & vbCrLf
                Next
                Call ApplyOEM_Picture(.ImgIndicate, "Picture")
                If gobjFile.FileExists(gstrSetupPath & "\附加文件\logo_login.jpg") Then
                    Set objPic = LoadPicture(gstrSetupPath & "\附加文件\logo_login.jpg")
                    .picHos.Visible = True
                    .picHos.Height = IIf(objPic.Height < 2745, objPic.Height, 2745) '183像素
                    .picHos.Width = IIf(objPic.Width < 4845, objPic.Width, 4845) '323像素
                    .picHos.PaintPicture objPic, 0, 0, .picHos.Width, .picHos.Height
                Else
                    .picHos.Visible = False
                End If
                If InStr(gstrCommand, "=") <= 0 Then .Show
                ShowSplash = True
            End With
        Else
            If strUnitName <> "" And strUnitName <> "-" Then
                gdtStart = Timer
                With frmSplash
                    '有两处需要处理
                    '此时就开始创建clsComLib类实例
                    Call ApplyOEM_Picture(.ImgIndicate, "Picture")
                    Call ApplyOEM_Picture(.imgPic, "PictureB")
                    If gobjFile.FileExists(gstrSetupPath & "\附加文件\logo_login.jpg") Then
                        Set objPic = LoadPicture(gstrSetupPath & "\附加文件\logo_login.jpg")
                        .picHos.Visible = True
                        .picHos.Height = IIf(objPic.Height < 2745, objPic.Height, 2745) '183像素
                        .picHos.Width = IIf(objPic.Width < 4845, objPic.Width, 4845) '323像素
                        .picHos.PaintPicture objPic, 0, 0, .picHos.Width, .picHos.Height
                    Else
                        .picHos.Visible = False
                    End If
                    If InStr(gstrCommand, "=") <= 0 Then .Show
                    
                    .lblGrant = Replace(strUnitName, ";", vbCrLf)
                    strUnitName = GetSetting("ZLSOFT", "注册信息", "开发商", "")
                    If Trim(strUnitName) = "" Then
                        .Label3.Visible = False
                        .lbl开发商.Visible = False
                    Else
                        .Label3.Visible = True
                        .lbl开发商.Visible = True
                        .lbl开发商.Caption = ""
                        For intCount = 0 To UBound(Split(strUnitName, ";"))
                            .lbl开发商.Caption = .lbl开发商.Caption & Split(strUnitName, ";")(intCount) & vbCrLf
                        Next
                    End If
                    .LblProductName = GetSetting("ZLSOFT", "注册信息", "产品全称", "")
                    If Len(.LblProductName) > 10 Then
                        .LblProductName.FontSize = 15.75 '三号
                    Else
                        .LblProductName.FontSize = 21.75 '二号
                    End If
                    .lbl技术支持商 = GetSetting("ZLSOFT", "注册信息", "技术支持商", "")
                    .lbltag = GetSetting("ZLSOFT", "注册信息", "产品系列", "")
                    
                    If Trim$(.lbl技术支持商.Caption) = "" Then
                        .Label1.Visible = False
                        .lbl技术支持商.Visible = False
                    Else
                        .Label1.Visible = True
                        .lbl技术支持商.Visible = True
                    End If
                End With
                Do
                    If (Timer - gdtStart) > 1 Then Exit Do
                    DoEvents
                Loop
                
                ShowSplash = True
            End If
        End If
    End If
End Function

Public Function SaveRegInfo() As Boolean
    Dim strTag As String, strTitle As String
    
    Select Case gobjRegister.zlRegInfo("授权性质")
        Case "1"
            '正式
            SaveSetting "ZLSOFT", "注册信息", "Kind", ""
        Case "2"
            '试用
            SaveSetting "ZLSOFT", "注册信息", "Kind", "试用"
        Case "3"
            '测试
            SaveSetting "ZLSOFT", "注册信息", "Kind", "测试"
        Case Else
            '不对
            MsgBox "授权性质不正确，程序被迫退出！", vbInformation, gstrSysName
            Exit Function
    End Select
    
    gstrSysName = gobjRegister.zlRegInfo("产品简名") & "软件"
    SaveSetting "ZLSOFT", "注册信息", "提示", gstrSysName
    SaveSetting "ZLSOFT", "注册信息", UCase("gstrSysName"), gstrSysName
    strTag = ""
    strTitle = gobjRegister.zlRegInfo("产品标题")
    If strTitle <> "" Then
        If InStr(strTitle, "-") > 0 Then
            If Split(strTitle, "-")(1) = "Ultimate" Then
                strTag = "旗舰版"
            ElseIf Split(strTitle, "-")(1) = "Professional" Then
                strTag = "专业版"
            End If
        End If
    End If
    strTitle = Split(strTitle, "-")(0)
    '将用户注册相关信息写入注册表,供下次启动时显示
    SaveSetting "ZLSOFT", "注册信息", "单位名称", gobjRegister.zlRegInfo("单位名称", , -1)
    SaveSetting "ZLSOFT", "注册信息", "产品全称", strTitle
    SaveSetting "ZLSOFT", "注册信息", "产品名称", gobjRegister.zlRegInfo("产品简名")
    SaveSetting "ZLSOFT", "注册信息", "技术支持商", gobjRegister.zlRegInfo("技术支持商", , -1)
    SaveSetting "ZLSOFT", "注册信息", "开发商", gobjRegister.zlRegInfo("产品开发商", , -1)
    SaveSetting "ZLSOFT", "注册信息", "WEB支持商简名", gobjRegister.zlRegInfo("支持商简名")
    SaveSetting "ZLSOFT", "注册信息", "WEB支持EMAIL", gobjRegister.zlRegInfo("支持商MAIL")
    SaveSetting "ZLSOFT", "注册信息", "WEB支持URL", gobjRegister.zlRegInfo("支持商URL")
    SaveSetting "ZLSOFT", "注册信息", "产品系列", strTag
    SaveRegInfo = True
End Function

Public Function TestComponent() As Boolean
    '如果没有任何部件可使用，则返回假
    TestComponent = False
    
    Dim strObjs As String, strSQL As String
    Dim resComponent As New ADODB.Recordset
    
    On Error GoTo errH
    '部件检测可能回出现错误，导致出现异常进程停滞
    If glngHelperMainType <> 0 Then TestComponent = True: Exit Function
    '--由注册表获取授权部件--
    strObjs = GetSetting("ZLSOFT", "注册信息", "本机部件", "")

    If strObjs <> "" Then
        If InStr(strObjs, "'ZL9REPORT'") = 0 Then
            If CreateComponent("ZL9REPORT.ClsREPORT") Then
                strObjs = strObjs & ",'ZL9REPORT'"
                SaveSetting "ZLSOFT", "注册信息", "本机部件", strObjs
            End If
        End If
        TestComponent = True
        Exit Function
    End If
    '--分析授权安装部件--union已去重
    strSQL = "Select 部件 From (" & _
                "Select Upper(g.部件) As 部件" & vbNewLine & _
                "From zlPrograms G, (Select Distinct 系统, 序号 From zlRegFunc) R" & vbNewLine & _
                "Where g.序号 = r.序号 And Trunc(g.系统 / 100) = r.系统" & vbNewLine & _
                " Union " & _
                " Select Upper(部件) as 部件 From zlPrograms Where 序号 Between 10000 And 19999)"
    Set resComponent = OpenSQLRecord(strSQL, "")
    With resComponent
        Do While Not .EOF
            If CreateComponent(!部件 & ".Cls" & Mid(!部件, 4)) Then
                strObjs = strObjs & IIf(strObjs = "", "", ",") & "'" & !部件 & "'"
            End If
            .MoveNext
        Loop
    End With
    If strObjs = "" Then Exit Function
    TestComponent = True
    SaveSetting "ZLSOFT", "注册信息", "本机部件", strObjs
    Exit Function
errH:
    If Not gobjComlib Is Nothing Then
        If gobjComlib.ErrCenter() = 1 Then
            Resume
        End If
    Else
        MsgBox "检测本机安装部件出错：" & Err.Description, vbInformation, gstrSysName
    End If
End Function

Private Function CreateComponent(StrComponent) As Boolean
    Dim objComponent        As Object
On Error GoTo errH
    Set objComponent = CreateObject(StrComponent)
    CreateComponent = True
    Exit Function
errH:
    Err.Clear
    CreateComponent = False
    Exit Function
End Function

Public Function ValEx(ByVal varInput As Variant) As Variant
'功能：由于Val只能以数字开头识别，ValEx以第一个数字进行识别
    Dim lngPos As Long
    If Val(varInput) = 0 Then
        varInput = varInput & ""
        If Trim(varInput) = "" Then ValEx = 0: Exit Function
        For lngPos = 1 To Len(varInput)
            If IsNumeric(Mid(varInput, lngPos, 1)) Then Exit For
        Next
        If lngPos = Len(varInput) + 1 Then
            ValEx = 0
        Else
            ValEx = Val(Mid(varInput, lngPos))
        End If
    Else
        ValEx = Val(varInput)
    End If
End Function

Public Function CreateRegister() As Boolean
    '创建注册部件(用于登录时获取连接对象)
    Dim strObject As String
    
    '虽然140上取消了Alone部件,但因为要兼容120及以上的低版本，所以此处保留Alone相关逻辑分支
    On Error Resume Next
    If UCase(App.EXEName) = "ZLLOGINALONE" Then
        strObject = "zlRegisterAlone"
    Else
        strObject = "zlRegister"
    End If
    Set gobjRegister = CreateObject(strObject & ".clsRegister")
    If gobjRegister Is Nothing Then
        Err.Clear
        MsgBox "创建" & strObject & ")部件对象失败,请检查文件是否存在并且正确注册。", vbExclamation, gstrSysName
        Exit Function
    End If
    CreateRegister = Not gobjRegister Is Nothing
End Function

Public Function CheckPWDComplex(ByRef cnInput As ADODB.Connection, ByVal strChcekPWD As String, Optional ByRef strToolTip As String) As String
'功能：检查密码复杂度
'参数：cnInput=传入的连接
'          strChcekPWD=等待检查的密码
'          strToolTip=鼠标提示生成
'返回：返回检查错误或检查警告
    Dim strSQL As String, rsData As New ADODB.Recordset
    Dim blnHaveNum As Boolean, blnAlpha As Boolean, blnChar As Boolean
    Dim blnPwdLen As Boolean, intPwdMin As Integer, intPwdMax As Integer
    Dim blnComplex As Boolean, strOterChrs As String
    Dim lngLen As Long, i As Integer, intChr As Integer
    
    On Error GoTo errH
    strToolTip = ""
    strSQL = "Select 参数号,Nvl(参数值,缺省值) 参数值 From zlOptions Where 参数号 in (20,21,22,23)"
    rsData.Open strSQL, cnInput
    blnPwdLen = False: intPwdMin = 0: intPwdMax = 0
    blnComplex = False: strOterChrs = ""
    Do While Not rsData.EOF
        Select Case rsData!参数号
            Case 20 '是否控制密码长度
                blnPwdLen = Val(rsData!参数值 & "") = 1
            Case 21 '密码长度下限
                intPwdMin = Val(rsData!参数值 & "")
            Case 22 '密码长度上限
                intPwdMax = Val(rsData!参数值 & "")
            Case 23 '是否控制密码复杂度
                blnComplex = Val(rsData!参数值 & "") = 1
        End Select
        rsData.MoveNext
    Loop
    '生成悬浮提示
    If blnPwdLen Then
        If intPwdMin = intPwdMax Then
            strToolTip = "密码必须为" & intPwdMax & " 位字符。"
        Else
            strToolTip = "密码必须为" & intPwdMin & "至" & intPwdMax & " 位字符。"
        End If
     End If
     If blnComplex Then
        If strToolTip <> "" Then
            strToolTip = strToolTip & vbNewLine & "至少包含一个数字、一个字母与一个特殊字符组成。"
        Else
            strToolTip = "至少由一个数字、一个字母与一个特殊字符组成。"
        End If
     End If
    '长度检查
    lngLen = ActualLen(strChcekPWD)
    If lngLen <> Len(strChcekPWD) Then
        CheckPWDComplex = "新密码包含双字节字符，请检查！"
        Exit Function
    End If
    If blnPwdLen Then
        If Not (lngLen >= intPwdMin And lngLen <= intPwdMax) Then
            If intPwdMin = intPwdMax Then
                CheckPWDComplex = "密码必须为" & intPwdMax & " 位字符！"
                Exit Function
            Else
                CheckPWDComplex = "密码必须为" & intPwdMin & "至" & intPwdMax & " 位字符！"
                Exit Function
            End If
        End If
    End If
    For i = 1 To Len(strChcekPWD)
        intChr = Asc(UCase(Mid(strChcekPWD, i, 1)))
        If intChr >= 32 And intChr < 127 Then
            'Dim blnHaveNum As Boolean, blnAlpha As Boolean, blnChar As Boolean
            Select Case intChr
                Case 48 To 57 '数字
                    blnHaveNum = True
                Case 65 To 90 '字母
                    blnAlpha = True
                Case 32, 34, 47, 64  '空格,双引号,/,@
                    strOterChrs = strOterChrs & Chr(intChr)
                Case Is < 48, 58 To 64, 91 To 96, Is > 122
                    blnChar = True
            End Select
        Else
            strOterChrs = strOterChrs & Chr(intChr)
        End If
    Next
    If strOterChrs <> "" Then
        CheckPWDComplex = "密码不容许有以下字符：" & strOterChrs
        Exit Function
    ElseIf Not (blnHaveNum And blnAlpha And blnChar) And blnComplex Then
        CheckPWDComplex = "密码至少由一个数字、一个字母与一个特殊字符组成。"
        Exit Function
    End If
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
    CheckPWDComplex = Err.Description
End Function

Public Function CheckSysState() As Boolean
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim blnHaveTools As Boolean, blnDBA As Boolean
    
    On Error Resume Next
    strSQL = "SELECT 1 FROM ZLTOOLS.ZLSYSTEMS WHERE 所有者=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, "读取所有者", gstrDBUser)
    
    If Err.Number <> 0 Then
        blnHaveTools = False
        gclsLogin.IsSysOwner = False
        Err.Clear
    Else
        blnHaveTools = True
        gclsLogin.IsSysOwner = rsTmp.EOF
    End If

    strSQL = "SELECT 1 FROM SESSION_ROLES WHERE ROLE='DBA'"
    Set rsTmp = OpenSQLRecord(strSQL, "判断DBA")
    blnDBA = Not rsTmp.EOF

    If Not (blnDBA) And Not (blnHaveTools) Then
        CheckSysState = False
        MsgBox "尚创建服务器管理工具，请先进行创建！", vbExclamation, gstrSysName
        Exit Function
    End If
    
    If Not (blnDBA) And Not (gclsLogin.IsSysOwner) Then
        CheckSysState = False
        MsgBox "不是数据库DBA或应用系统的所有者，不能使用本工具。", vbExclamation, gstrSysName
        Exit Function
    End If
    If Not blnHaveTools Then
        CheckSysState = False
        MsgBox "尚创建服务器管理工具，请先进行创建！", vbExclamation, gstrSysName
        Exit Function
    End If
    CheckSysState = True
End Function

Public Function GetMenuGroup(ByVal strCommand As String) As String
    Dim ArrCommand As Variant
    Dim i As Long
    '--分析权限菜单--
    
    GetMenuGroup = "缺省"
    
    ArrCommand = Split(strCommand, " ")
    If UBound(ArrCommand) = 0 Then
        '仅仅包含菜单组别（如果含有/，表示是用户加密码的格式，如：zlhis/his）
        If InStr(1, ArrCommand(0), "/") = 0 And InStr(ArrCommand(0), ",") = 0 Then
            GetMenuGroup = ArrCommand(0)
        End If
    Else
        '用户名、密码及菜单组别
        If UBound(ArrCommand) = 2 Then
            If InStr(ArrCommand(0), "=") <= 0 Then GetMenuGroup = ArrCommand(2)
            
        '例：C:\APPSOFT\ZLHIS+.exe USER=用户名 PASS=密码 SERVER=实例名 PROGRAM=模块号 MENUGROUP=缺省
        Else
            For i = 0 To UBound(ArrCommand)
                If Split(ArrCommand(i), "=")(0) = "MENUGROUP" Then
                    GetMenuGroup = Split(ArrCommand(i), "=")(1)
                    Exit For
                End If
            Next
        End If
    End If
End Function

Public Function OpenSQLRecord(ByVal strSQL As String, ByVal strTitle As String, ParamArray arrInput() As Variant) As ADODB.Recordset
    Dim arrPars() As Variant
    arrPars = arrInput
    If gblnTimer And Not gobjComlib Is Nothing Then
        Set OpenSQLRecord = gobjComlib.zldatabase.OpenSQLRecordByArray(strSQL, strTitle, arrPars)
    Else
        Set OpenSQLRecord = OpenSQLRecordByArray(strSQL, strTitle, arrPars)
    End If
End Function

Private Function OpenSQLRecordByArray(ByVal strSQL As String, ByVal strTitle As String, arrInput() As Variant) As ADODB.Recordset
'功能：通过Command对象打开带参数SQL的记录集
'参数：strSQL=条件中包含参数的SQL语句,参数形式为"[x]"
'             x>=1为自定义参数号,"[]"之间不能有空格
'             同一个参数可多处使用,程序自动换为ADO支持的"?"号形式
'             实际使用的参数号可不连续,但传入的参数值必须连续(如SQL组合时不一定要用到的参数)
'      arrInput=不定个数的参数值,按参数号顺序依次传入,必须是明确类型
'               因为使用绑定变量,对带"'"的字符参数,不需要使用"''"形式。
'      strTitle=用于SQLTest识别的调用窗体/模块标题
'      cnOracle=当不使用公共连接时传入
'返回：记录集，CursorLocation=adUseClient,LockType=adLockReadOnly,CursorType=adOpenStatic
'举例：
'SQL语句为="Select 姓名 From 病人信息 Where (病人ID=[3] Or 门诊号=[3] Or 姓名 Like [4]) And 性别=[5] And 登记时间 Between [1] And [2] And 险类 IN([6],[7])"
'调用方式为：Set rsPati=OpenSQLRecord(strSQL, Me.Caption, CDate(Format(rsMove!转出日期,"yyyy-MM-dd")),dtp时间.Value, lng病人ID, "张%", "男", 20, 21)
    Dim cmdData As New ADODB.Command
    Dim strPar As String, arrPar As Variant
    Dim lngLeft As Long, lngRight As Long
    Dim strSeq As String, intMax As Integer, i As Integer
    Dim strLog As String, varValue As Variant
    Dim strSQLTmp As String, arrstr As Variant
    Dim strTmp As String, strSQLtmp1 As String
    
    '检查如果使用了动态内存表，并且没有使用/*+ XXX*/等提示字时自动加上
    strSQLTmp = Trim(UCase(strSQL))
    If Mid(Trim(Mid(strSQLTmp, 7)), 1, 2) <> "/*" And Mid(strSQLTmp, 1, 6) = "SELECT" Then
        arrstr = Split("F_STR2LIST,F_NUM2LIST,F_NUM2LIST2,F_STR2LIST2", ",")
        For i = 0 To UBound(arrstr)
            strSQLtmp1 = strSQLTmp
            Do While InStr(strSQLtmp1, arrstr(i)) > 0
                '判断前面是否用了IN 用了则不加Rule
                '先找到最近一个SELECT
                strTmp = Mid(strSQLtmp1, 1, InStr(strSQLtmp1, arrstr(i)) - 1)
                strTmp = Replace(FromatSQL(Mid(strTmp, 1, InStrRev(strTmp, "SELECT") - 1)), " ", "")
                If Len(strTmp) > 1 Then strTmp = Mid(strTmp, Len(strTmp) - 2)  '取后面3个字符
                
                If strTmp = "IN(" Then '属于in(select这种情况，则继续循环，看是否存在没有使用这种写法的其他动态内存函数
                   strSQLtmp1 = Mid(strSQLtmp1, InStr(strSQLtmp1, arrstr(i)) + Len(arrstr(i)))
                Else
                    Exit For
                End If
            Loop
        Next
        If i <= UBound(arrstr) Then
            strSQL = "Select /*+ RULE*/" & Mid(Trim(strSQL), 7)
        End If
    End If
    
    '分析自定的[x]参数
    lngLeft = InStr(1, strSQL, "[")
    Do While lngLeft > 0
        lngRight = InStr(lngLeft + 1, strSQL, "]")
        If lngRight = 0 Then Exit Do
        '可能是正常的"[编码]名称"
        strSeq = Mid(strSQL, lngLeft + 1, lngRight - lngLeft - 1)
        If IsNumeric(strSeq) Then
            i = CInt(strSeq)
            strPar = strPar & "," & i
            If i > intMax Then intMax = i
        End If
        
        lngLeft = InStr(lngRight + 1, strSQL, "[")
    Loop
    
    If UBound(arrInput) + 1 < intMax Then
        Err.Raise 9527, strTitle, "SQL语句绑定变量不全，调用来源：" & strTitle
    End If

    '替换为"?"参数
    strLog = strSQL
    For i = 1 To intMax
        strSQL = Replace(strSQL, "[" & i & "]", "?")
        
        '产生用于SQL跟踪的语句
        varValue = arrInput(i - 1)
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency", "Decimal" '数字
            strLog = Replace(strLog, "[" & i & "]", varValue)
        Case "String" '字符
            strLog = Replace(strLog, "[" & i & "]", "'" & Replace(varValue, "'", "''") & "'")
        Case "Date" '日期
            strLog = Replace(strLog, "[" & i & "]", "To_Date('" & Format(varValue, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')")
        End Select
    Next
    
    '创建新的参数
    lngLeft = 0: lngRight = 0
    arrPar = Split(Mid(strPar, 2), ",")
    For i = 0 To UBound(arrPar)
        varValue = arrInput((arrPar(i) - 1))
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency", "Decimal" '数字
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adVarNumeric, adParamInput, 30, varValue)
        Case "String" '字符
            intMax = LenB(StrConv(varValue, vbFromUnicode))
            If intMax <= 2000 Then
                intMax = IIf(intMax <= 200, 200, 2000)
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adVarChar, adParamInput, intMax, varValue)
            Else
                If intMax < 4000 Then intMax = 4000
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adLongVarChar, adParamInput, intMax, varValue)
            End If
        Case "Date" '日期
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adDBTimeStamp, adParamInput, , varValue)
        Case "Variant()" '数组
            '这种方式可用于一些IN子句或Union语句
            '表示同一个参数的多个值,参数号不可与其它数组的参数号交叉,且要保证数组的值个数够用
            If arrPar(i) <> lngRight Then lngLeft = 0
            lngRight = arrPar(i)
            Select Case TypeName(varValue(lngLeft))
            Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '数字
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adVarNumeric, adParamInput, 30, varValue(lngLeft))
                strLog = Replace(strLog, "[" & lngRight & "]", varValue(lngLeft), 1, 1)
            Case "String" '字符
                intMax = LenB(StrConv(varValue(lngLeft), vbFromUnicode))
                If intMax <= 2000 Then
                    intMax = IIf(intMax <= 200, 200, 2000)
                    cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adVarChar, adParamInput, intMax, varValue(lngLeft))
                Else
                    If intMax < 4000 Then intMax = 4000
                    cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adLongVarChar, adParamInput, intMax, varValue(lngLeft))
                End If
                
                strLog = Replace(strLog, "[" & lngRight & "]", "'" & Replace(varValue(lngLeft), "'", "''") & "'", 1, 1)
            Case "Date" '日期
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adDBTimeStamp, adParamInput, , varValue(lngLeft))
                strLog = Replace(strLog, "[" & lngRight & "]", "To_Date('" & Format(varValue(lngLeft), "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')", 1, 1)
            End Select
            lngLeft = lngLeft + 1 '该参数在数组中用到第几个值了
        End Select
    Next
'    If gblnSys = True Then
'        Set cmdData.ActiveConnection = gcnSysConn
'    Else
    Set cmdData.ActiveConnection = gcnOracle '这句比较慢(这句执行1000次约0.5x秒)
'    End If
    cmdData.CommandText = strSQL
    
    Set OpenSQLRecordByArray = cmdData.Execute
    Set OpenSQLRecordByArray.ActiveConnection = Nothing
End Function

Public Sub ExecuteProcedure(strSQL As String, ByVal strFormCaption As String)
'功能：执行过程语句,并自动对过程参数进行绑定变量处理
'参数：strSQL=过程语句,可能带参数,形如"过程名(参数1,参数2,...)"。
'      cnOracle=当不使用公共连接时传入
'说明：以下几种情况过程参数不使用绑定变量,仍用老的调用方法：
'  1.参数部份是表达式,这时程序无法处理绑定变量类型和值,如"过程名(参数1,100.12*0.15,...)"
'  2.中间没有传入明确的可选参数,这时程序无法处理绑定变量类型和值,如"过程名(参数1, , ,参数3,...)"
'  3.因为该过程是自动处理,不是一定使用绑定变量,对带"'"的字符参数,仍要使用"''"形式。
    If gblnTimer And Not gobjComlib Is Nothing Then
        Call gobjComlib.zldatabase.ExecuteProcedure(strSQL, strFormCaption)
    Else
        Dim cmdData As New ADODB.Command
        Dim strProc As String, strPar As String
        Dim blnStr As Boolean, intBra As Integer
        Dim strTemp As String, i As Long
        Dim intMax As Integer, datCur As Date
        
        If Right(Trim(strSQL), 1) = ")" Then
            '执行的过程名
            strTemp = Trim(strSQL)
            strProc = Trim(Left(strTemp, InStr(strTemp, "(") - 1))
            
            '执行过程参数
            datCur = CDate(0)
            strTemp = Mid(strTemp, InStr(strTemp, "(") + 1)
            strTemp = Trim(Left(strTemp, Len(strTemp) - 1)) & ","
            For i = 1 To Len(strTemp)
                '是否在字符串内，以及表达式的括号内
                If Mid(strTemp, i, 1) = "'" Then blnStr = Not blnStr
                If Not blnStr And Mid(strTemp, i, 1) = "(" Then intBra = intBra + 1
                If Not blnStr And Mid(strTemp, i, 1) = ")" Then intBra = intBra - 1
                
                If Mid(strTemp, i, 1) = "," And Not blnStr And intBra = 0 Then
                    strPar = Trim(strPar)
                    With cmdData
                        If IsNumeric(strPar) Then '数字
                            .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adVarNumeric, adParamInput, 30, strPar)
                        ElseIf Left(strPar, 1) = "'" And Right(strPar, 1) = "'" Then '字符串
                            strPar = Mid(strPar, 2, Len(strPar) - 2)
                            
                            'Oracle连接符运算:'ABCD'||CHR(13)||'XXXX'||CHR(39)||'1234'
                            If InStr(Replace(strPar, " ", ""), "'||") > 0 Then GoTo NoneVarLine
                            
                            '双"''"的绑定变量处理
                            If InStr(strPar, "''") > 0 Then strPar = Replace(strPar, "''", "'")
                            
                            '电子病历处理LOB时，如果用绑定变量转换为RAW时超过2000个字符要用adLongVarChar
                            intMax = LenB(StrConv(strPar, vbFromUnicode))
                            If intMax <= 2000 Then
                                intMax = IIf(intMax <= 200, 200, 2000)
                                .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adVarChar, adParamInput, intMax, strPar)
                            Else
                                If intMax < 4000 Then intMax = 4000
                                .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adLongVarChar, adParamInput, intMax, strPar)
                            End If
                        ElseIf UCase(strPar) Like "TO_DATE('*','*')" Then '日期
                            strPar = Split(strPar, "(")(1)
                            strPar = Trim(Split(strPar, ",")(0))
                            strPar = Mid(strPar, 2, Len(strPar) - 2)
                            If strPar = "" Then
                                'NULL值当成数字处理可兼容其他类型
                                .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adVarNumeric, adParamInput, , Null)
                            Else
                                If Not IsDate(strPar) Then GoTo NoneVarLine
                                .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adDBTimeStamp, adParamInput, , CDate(strPar))
                            End If
                        ElseIf UCase(strPar) = "SYSDATE" Then '日期
                            If datCur = CDate(0) Then datCur = Currentdate
                            .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adDBTimeStamp, adParamInput, , datCur)
                        ElseIf UCase(strPar) = "NULL" Then 'NULL值当成字符处理可兼容其他类型
                            .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adVarChar, adParamInput, 200, Null)
                        ElseIf strPar = "" Then '可选参数当成NULL处理可能改变了缺省值:因此可选参数不能写在中间
                            GoTo NoneVarLine
                        Else '可能是其他复杂的表达式，无法处理
                            GoTo NoneVarLine
                        End If
                    End With
                    
                    strPar = ""
                Else
                    strPar = strPar & Mid(strTemp, i, 1)
                End If
            Next
            
            '程序员调用过程时书写错误
            If blnStr Or intBra <> 0 Then
                Err.Raise -2147483645, , "调用 Oracle 过程""" & strProc & """时，引号或括号书写不匹配。原始语句如下：" & vbCrLf & vbCrLf & strSQL
                Exit Sub
            End If
            
            '补充?号
            strTemp = ""
            For i = 1 To cmdData.Parameters.Count
                strTemp = strTemp & ",?"
            Next
            strProc = "Call " & strProc & "(" & Mid(strTemp, 2) & ")"
            Set cmdData.ActiveConnection = gcnOracle '这句比较慢
            cmdData.CommandType = adCmdText
            cmdData.CommandText = strProc
            
            Call cmdData.Execute
        Else
            GoTo NoneVarLine
        End If
        Exit Sub
NoneVarLine:
        '说明：为了兼容新连接方式
        '1.新连接用adCmdStoredProc方式在8i下面有问题
        '2.新连接如果不使用{},则即使过程没有参数也要加()
        strSQL = "Call " & strSQL
        If InStr(strSQL, "(") = 0 Then strSQL = strSQL & "()"
        gcnOracle.Execute strSQL, , adCmdText
End If
End Sub

Public Function IP(Optional ByVal strErr As String) As String
    '******************************************************************************************************************
    '功能:通过oracle获取的计算机的IP地址
    '入参:strDefaultIp_Address-缺省IP地址
    '出参:
    '返回:返回IP地址
    '******************************************************************************************************************
    Dim rsTmp As ADODB.Recordset
    Dim strIp_Address As String
    Dim strSQL As String
        
    On Error GoTo Errhand
    
    strSQL = "Select Sys_Context('USERENV', 'IP_ADDRESS') as Ip_Address From Dual"
    Set rsTmp = OpenSQLRecord(strSQL, "获取IP地址")
    If rsTmp.EOF = False Then
        strIp_Address = NVL(rsTmp!Ip_Address)
    End If
    If strIp_Address = "" Then strIp_Address = OSIP(strErr)
    If Replace(strIp_Address, " ", "") = "0.0.0.0" Then strIp_Address = ""
    IP = strIp_Address
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
Errhand:
    strErr = strErr & IIf(strErr = "", "", "|") & Err.Description
    Err.Clear
End Function

Public Function Currentdate() As Date
    '-------------------------------------------------------------
    '功能：提取服务器上当前日期
    '参数：
    '返回：由于Oracle日期格式的问题，所以
    '-------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    Err = 0
    On Error GoTo errH
    With rsTemp
        .CursorLocation = adUseClient
    End With
    Set rsTemp = OpenSQLRecord("SELECT SYSDATE FROM DUAL", "获取服务器时间")
    Currentdate = rsTemp.Fields(0).Value
    rsTemp.Close
    Exit Function
errH:
    Currentdate = 0
    Err = 0
End Function

Public Sub StartInstall()
    Dim strAppPath      As String
    Dim lngErr          As Long
    On Error Resume Next
    strAppPath = gobjFile.GetParentFolderName(App.Path)
    
    If gobjFile.FileExists(strAppPath & "\ZLExFile\ZLExInstall.exe") Then
        strAppPath = strAppPath & "\ZLExFile\ZLExInstall.exe"
    ElseIf gobjFile.FileExists("C:\APPSOFT\ZLExFile\ZLExInstall.exe") Then
        strAppPath = "C:\APPSOFT\ZLExFile\ZLExInstall.exe"
    Else
        Exit Sub
    End If
    '若产品名称不对，则不启动命令行
    If UCase(GetFileDesInfo(strAppPath, "ProductName")) <> "ZLSOFT EXTENSION INSTALL" Then
        Exit Sub
    End If
    lngErr = Shell(strAppPath & " ORAOLEDB -REGSVR -S", vbHide)
End Sub


'功能：获取当前进程的路径
Public Function GetCurExePath() As String
    Dim uProcess        As PROCESSENTRY32, uMdlInfor    As MODULEENTRY32
    Dim lngMdlProcess   As Long, strExeName             As String, lngSnapShot  As Long, strModelPath     As String, strModelName As String
    Dim lngProceess     As Long
    
    On Error GoTo errH
    '创建进程快照
    lngProceess = GetCurrentProcessId
    lngSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    If lngSnapShot > 0 Then
        uProcess.lSize = Len(uProcess)
        If Process32First(lngSnapShot, uProcess) Then
            Do
                If uProcess.lProcessId = lngProceess Then
                    '获得进程的标识符
                    strExeName = UCase(Left(Trim(uProcess.sExeFile), InStr(1, Trim(uProcess.sExeFile), vbNullChar) - 1))
                    lngMdlProcess = CreateToolhelp32Snapshot(TH32CS_SNAPMODULE, uProcess.lProcessId)
                    If lngMdlProcess > 0 Then
                        uMdlInfor.dwSize = Len(uMdlInfor)
                        If Module32First(lngMdlProcess, uMdlInfor) Then
                            Do
                                strModelPath = UCase(Left(Trim(uMdlInfor.szExePath), InStr(1, Trim(uMdlInfor.szExePath), vbNullChar) - 1))
                                strModelName = UCase(Left(Trim(uMdlInfor.szModule), InStr(1, Trim(uMdlInfor.szModule), vbNullChar) - 1))
                                If strModelName = strExeName Then
                                    GetCurExePath = strModelPath
                                    Exit Do
                                End If
                            Loop Until (Module32Next(lngMdlProcess, uMdlInfor) < 1)
                        End If
                        CloseHandle (lngMdlProcess)
                    End If
                End If
            Loop Until (Process32Next(lngSnapShot, uProcess) < 1)
        End If
        CloseHandle (lngSnapShot)
    End If
    Exit Function
errH:
    Err.Clear
End Function

'功能：模仿VB APP.PrevInstance,该属性在DLL中会判断失效，一直为False
'说明：1、当进程路径不同时，两个进程的APP.PrevInstance无关联，尽管EXE文件相同。
'       2、该函数和APP.PrevInstance有一定区别：1）APP.PrevInstance是固定的，进程打开时就固定，尽管关闭其他的相同进程，仍旧不会发生变化。
'                                              2)该函数是动态查询，当进程清单中没有当前EXE路径文件的进程，就是FALSE,否则就是TRUE,和进程的运行有关。
Public Function AppPrevInstance() As Boolean
    Dim uProcess        As PROCESSENTRY32, uMdlInfor    As MODULEENTRY32
    Dim lngMdlProcess   As Long, strExeName             As String, lngSnapShot  As Long, strModelPath     As String, strModelName As String
    Dim lngProceess     As Long
    Dim strCurAppPath   As String
    Dim blnFind         As Boolean
    
    On Error GoTo errH
    '创建进程快照
    strCurAppPath = GetCurExePath()
    lngProceess = GetCurrentProcessId
    blnFind = False
    lngSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    If lngSnapShot > 0 Then
        uProcess.lSize = Len(uProcess)
        If Process32First(lngSnapShot, uProcess) Then
            Do
                If uProcess.lProcessId <> lngProceess Then
                    '获得进程的标识符
                    strExeName = UCase(Left(Trim(uProcess.sExeFile), InStr(1, Trim(uProcess.sExeFile), vbNullChar) - 1))
                    lngMdlProcess = CreateToolhelp32Snapshot(TH32CS_SNAPMODULE, uProcess.lProcessId)
                    If lngMdlProcess > 0 Then
                        uMdlInfor.dwSize = Len(uMdlInfor)
                        If Module32First(lngMdlProcess, uMdlInfor) Then
                            Do
                                strModelPath = UCase(Left(Trim(uMdlInfor.szExePath), InStr(1, Trim(uMdlInfor.szExePath), vbNullChar) - 1))
                                strModelName = UCase(Left(Trim(uMdlInfor.szModule), InStr(1, Trim(uMdlInfor.szModule), vbNullChar) - 1))
                                If strModelName = strExeName Then
                                    If strModelPath = strCurAppPath Then
                                        blnFind = True
                                    End If
                                    Exit Do
                                End If
                            Loop Until (Module32Next(lngMdlProcess, uMdlInfor) < 1)
                        End If
                        CloseHandle (lngMdlProcess)
                    End If
                End If
            Loop Until (Process32Next(lngSnapShot, uProcess) < 1 Or blnFind)
        End If
        CloseHandle (lngSnapShot)
    End If
    AppPrevInstance = blnFind
    Exit Function
errH:
    Err.Clear
End Function

Public Function ComputerName() As String
    '******************************************************************************************************************
    '功能：获取电脑名称
    '参数：
    '说明：
    '******************************************************************************************************************
    Dim strComputer As String * 256
    Call GetComputerName(strComputer, 255)
    ComputerName = strComputer
    ComputerName = Trim(Replace(ComputerName, Chr(0), ""))
End Function

Public Function NVL(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    NVL = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Function IsDesinMode() As Boolean
'功能： 确定当前模式为设计模式
     Err = 0: On Error Resume Next
     Debug.Print 1 / 0
     If Err <> 0 Then
        IsDesinMode = True
     Else
        IsDesinMode = False
     End If
     Err.Clear: Err = 0
 End Function


Public Function OSIP(Optional ByRef strErr As String) As String
    '功能：通过API获取临时IP
    
    Dim ret As Long, Tel As Long
    Dim bBytes() As Byte
    Dim TempList() As String
    Dim TempIP As String
    Dim Tempi As Long
    Dim Listing As MIB_IPADDRTABLE
    Dim L3 As String
    Dim strTmpErr As String, strALLErr As String
    
    strErr = ""
    On Error GoTo Errhand
    GetIpAddrTable ByVal 0&, ret, True
    If ret <= 0 Then Exit Function
    ReDim bBytes(0 To ret - 1) As Byte
    ReDim TempList(0 To ret - 1) As String
    'retrieve the data
    GetIpAddrTable bBytes(0), ret, False
    'Get the first 4 bytes to get the entry's.. ip installed
    CopyMemory Listing.dEntrys, bBytes(0), 4
    For Tel = 0 To Listing.dEntrys - 1
        'Copy whole structure to Listing..
        CopyMemory Listing.mIPInfo(Tel), bBytes(4 + (Tel * Len(Listing.mIPInfo(0)))), Len(Listing.mIPInfo(Tel))
        TempList(Tel) = ConvertAddressToString(Listing.mIPInfo(Tel).dwAddr, strTmpErr)
        If strTmpErr <> "" Then strALLErr = strALLErr & IIf(strALLErr = "", "", "|") & strTmpErr
    Next Tel
    'Sort Out The IP For WAN
        TempIP = TempList(0)
        For Tempi = 0 To Listing.dEntrys - 1
            L3 = Left(TempList(Tempi), 3)
            If L3 <> "169" And L3 <> "127" And L3 <> "192" Then
                TempIP = TempList(Tempi)
            End If
        Next Tempi
        OSIP = TempIP 'Return The TempIP
    Exit Function
    strErr = strALLErr
    '------------------------------------------------------------------------------------------------------------------
Errhand:
    strErr = strALLErr & IIf(strALLErr = "", "", "|") & Err.Description
    Err.Clear
End Function

Private Function ConvertAddressToString(longAddr As Long, Optional ByRef strErr As String) As String
    Dim myByte(3) As Byte
    Dim Cnt As Long
    
    strErr = ""
    On Error GoTo errH
    CopyMemory myByte(0), longAddr, 4
    For Cnt = 0 To 3
        ConvertAddressToString = ConvertAddressToString + CStr(myByte(Cnt)) + "."
    Next Cnt
    ConvertAddressToString = Left$(ConvertAddressToString, Len(ConvertAddressToString) - 1)
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errH:
    strErr = Err.Description
    Err.Clear
End Function

Private Function TruncZeroInside(ByVal strInput As String) As String
'功能：去掉字符串中\0以后的字符,仅用作该工程,可以单独是用clsstring
    Dim lngPos As Long
    
    lngPos = InStr(strInput, Chr(0))
    If lngPos > 0 Then
        TruncZeroInside = Mid(strInput, 1, lngPos - 1)
    Else
        TruncZeroInside = strInput
    End If
End Function
'======================================================================================================================
'方法           Sm4EncryptEcb           SM4加密
'返回值         String                  加密后的值,格式：ZLSV+版本号+:+加密后的字符串
'入参列表:
'参数名         类型                    说明
'strInput       String                  要加密的字符串
'strKey         String(Optional)        加密密钥（32位的16进制字符串，可以通过HexStringToByte返回）
'======================================================================================================================
Public Function Sm4EncryptEcb(ByVal strInput As String, Optional ByVal strKey As String) As String
    Dim arrKey()    As Byte
    Dim arrInput()  As Byte
    Dim arrOutPut() As Byte
    
    If M_SM4_VERSION = 0 Then
        M_SM4_VERSION = sm_version
    End If
    If strInput = "" Then
        Sm4EncryptEcb = ""
    Else
        arrKey = GetKey(strKey, 2)
        arrInput = BytePadding(strInput, M_SM4_VERSION)
        ReDim arrOutPut(UBound(arrInput))
        Call sm4_crypt_ecb(CM_Encrypt, UBound(arrInput) + 1, arrKey(0), arrInput(0), arrOutPut(0))
        Sm4EncryptEcb = "ZLSV" & M_SM4_VERSION & ":" & ByteToHexString(arrOutPut())
    End If
End Function

'======================================================================================================================
'方法           Sm4DecryptEcb           SM4解密
'返回值         String                  解密后的值
'入参列表:
'参数名         类型                    说明
'strInput       String                  要解密的字符串（该字符串是Sm4EncryptEcb生成的结果）
'strKey         String(Optional)        加密密钥也就是解密密钥（32位的16进制字符串，可以通过HexStringToByte返回）
'======================================================================================================================
Public Function Sm4DecryptEcb(ByVal strInput As String, Optional ByVal strKey As String) As String
    Dim arrKey()        As Byte
    Dim arrInput()      As Byte
    Dim arrOutPut()     As Byte
    Dim lngVersion      As Long

    If M_SM4_VERSION = 0 Then
        M_SM4_VERSION = sm_version
    End If
    If strInput Like "ZLSV*:*" Then
        lngVersion = Val(Mid(strInput, 5, InStr(strInput, ":") - 5))
        strInput = Mid(strInput, InStr(strInput, ":") + 1)
        '当前客户端的ZLSM4不支持该版本的加密字符串解密，仍旧解密，因为一般来说都能解密出相同的字符串
'        If lngVersion > M_SM4_VERSION Then
'            Exit Function
'        End If
    Else
        Exit Function
    End If
    
    arrKey = GetKey(strKey, 2)
    arrInput = HexStringToByte(strInput)
    ReDim arrOutPut(UBound(arrInput))
    
    Call sm4_crypt_ecb(CM_Decrypt, UBound(arrInput) + 1, arrKey(0), arrInput(0), arrOutPut(0))
    If lngVersion = 1 Then
        Sm4DecryptEcb = Trim(StrConv(arrOutPut(), vbUnicode))
    Else
        Sm4DecryptEcb = TruncZeroInside(StrConv(arrOutPut(), vbUnicode))
    End If
End Function
'======================================================================================================================
'方法           Sm4EncryptCbc           SM4分组加密
'返回值         String                  加密后的值
'入参列表:
'参数名         类型                    说明
'strInput       String                  要加密的字符串
'strKey         String(Optional)        加密密钥（32位的16进制字符串，可以通过HexStringToByte返回）
'strIv          String(Optional)        分组加密密钥（32位的16进制字符串，可以通过HexStringToByte返回）
'======================================================================================================================
Public Function Sm4EncryptCbc(ByVal strInput As String, Optional ByVal strKey As String, Optional ByVal strIv As String) As String
    Dim arrKey()        As Byte
    Dim arrInput()      As Byte
    Dim arrOutPut()     As Byte
    Dim arrIv()         As Byte
    
    If M_SM4_VERSION = 0 Then
        M_SM4_VERSION = sm_version
    End If
    If strInput = "" Then
        Sm4EncryptCbc = ""
    Else
        arrKey = GetKey(strKey, 2)
        arrIv = GetKey(strIv, 1)
        
        arrInput = BytePadding(strInput, M_SM4_VERSION)
        ReDim arrOutPut(UBound(arrInput))
        
        Call sm4_crypt_cbc(CM_Encrypt, UBound(arrInput) + 1, arrIv(0), arrKey(0), arrInput(0), arrOutPut(0))
        Sm4EncryptCbc = "ZLSV" & M_SM4_VERSION & ":" & ByteToHexString(arrOutPut)
    End If
End Function

'======================================================================================================================
'方法           Sm4EncryptCbc           SM4分组加密对应的解密过程
'返回值         String                  解密后的值
'入参列表:
'参数名         类型                    说明
'strInput       String                  已经加密的字符串
'strKey         String(Optional)        解密密钥也就是加密密钥（32位的16进制字符串，可以通过HexStringToByte返回）
'strIv          String(Optional)        分组解密密钥也就是分组加密密钥（32位的16进制字符串，可以通过HexStringToByte返回）
'======================================================================================================================
Public Function Sm4DecryptCbc(ByVal strInput As String, Optional ByVal strKey As String, Optional ByVal strIv As String) As String
    Dim arrKey() As Byte
    Dim arrInput() As Byte
    Dim arrOutPut() As Byte
    Dim arrIv() As Byte
    Dim lngVersion      As Long

    If M_SM4_VERSION = 0 Then
        M_SM4_VERSION = sm_version
    End If
    If strInput Like "ZLSV*:*" Then
        lngVersion = Val(Mid(strInput, 5, InStr(strInput, ":") - 5))
        strInput = Mid(strInput, InStr(strInput, ":") + 1)
        '当前客户端的ZLSM4不支持该版本的加密字符串解密，仍旧解密，因为一般来说都能解密出相同的字符串
'        If lngVersion > M_SM4_VERSION Then
'            Exit Function
'        End If
    Else
        Exit Function
    End If
    
    arrKey = GetKey(strKey, 2)
    arrIv = GetKey(strIv, 1)
    
    arrInput = HexStringToByte(strInput)
    ReDim arrOutPut(UBound(arrInput))

    Call sm4_crypt_cbc(CM_Decrypt, UBound(arrInput) + 1, arrIv(0), arrKey(0), arrInput(0), arrOutPut(0))
    
    If lngVersion = 1 Then
        Sm4DecryptCbc = Trim(StrConv(arrOutPut(), vbUnicode))
    Else
        Sm4DecryptCbc = TruncZeroInside(StrConv(arrOutPut(), vbUnicode))
    End If
End Function

'======================================================================================================================
'方法           Sm3                     计算字符串的哈希值（用来检测字符串的变动）
'返回值         String(32)              字符串的哈希值
'入参列表:
'参数名         类型                    说明
'strInput       String                  字符串内容
'======================================================================================================================
Public Function Sm3(ByRef strInput As String) As String
    Dim arrInput()  As Byte
    Dim lngLength   As Long
    Dim arrOut(31)  As Byte

    '先将字符串由 Unicode 转成系统的缺省码页
    arrInput = StrConv(strInput, vbFromUnicode)
    lngLength = UBound(arrInput) + 1
    
    Call sm3_hash(arrInput(0), lngLength, arrOut(0))
    '将返回值转换为16进制字符串
    Sm3 = ByteToHexString(arrOut)
End Function
'======================================================================================================================
'方法           Sm3_File                计算文件的哈希值（用来检测 文件内容的变动）
'返回值         String(32)              文件的哈希值
'入参列表:
'参数名         类型                    说明
'strFile        String                  文件路径
'======================================================================================================================
Public Function Sm3_File(ByRef strFile As String) As String
    Dim arrInput()  As Byte
    Dim lngLength   As Long
    Dim arrOut(31)  As Byte
    Dim lngReturn As Long

    '先将字符串由 Unicode 转成系统的缺省码页
    arrInput = StrConv(strFile, vbFromUnicode)
    '由于API没有传递长度，因此增加字符串Chr(0)
    lngLength = UBound(arrInput) + 1
    ReDim Preserve arrInput(lngLength)
    '计算结果
    lngReturn = sm3_file_hash(arrInput(0), arrOut(0))
    '判断是否成功处理
    If lngReturn = 0 Then
        '将返回值转换为16进制字符串
        Sm3_File = ByteToHexString(arrOut)
    ElseIf lngReturn = 1 Then
        Sm3_File = "ERROR:文件打开失败"
    ElseIf lngReturn = 2 Then
        Sm3_File = "ERROR:文件读取失败"
    End If
End Function
'======================================================================================================================
'方法           sm3_hmac                给定义一个密钥对传入的消息产生消息摘要
'返回值         String(32)              密钥加密消息后生成的消息摘要
'入参列表:
'参数名         类型                    说明
'strKey         String                  密钥
'strMsg         String                  消息内容
'======================================================================================================================
Public Function sm3_hmac(ByRef strKey As String, ByVal strMsg As String) As String
    Dim arrInput()  As Byte
    Dim lngLength   As Long
    Dim arrOut(31)  As Byte
    Dim arrKey()    As Byte
    Dim lngKeyLen   As Long
    
    '先将字符串由 Unicode 转成系统的缺省码页
    arrInput = StrConv(strMsg, vbFromUnicode)
    lngLength = UBound(arrInput) + 1
    '先将字符串由 Unicode 转成系统的缺省码页
    arrKey = StrConv(strKey, vbFromUnicode)
    lngKeyLen = UBound(arrKey) + 1
    Call sm3_hmac_hash(arrKey(0), lngKeyLen, arrInput(0), lngLength, arrOut(0))
    '将返回值转换为16进制字符串
    sm3_hmac = ByteToHexString(arrOut)
End Function
'======================================================================================================================
'方法           sm_version              获取ZLSM4的版本号
'返回值         Long                    ZLSM4的版本号
'入参列表:
'======================================================================================================================
Public Function sm_version() As Long
    Dim lngVersion As Long
    On Error Resume Next
    lngVersion = get_sm_version
    If Err.Number <> 0 Then
        Err.Clear
        sm_version = 1
    Else
        sm_version = lngVersion
    End If
End Function
'======================================================================================================================
'方法           ByteToHexString         将字节组转换为16进制字符串
'返回值         String                  字节组转换的16进制字符串
'入参列表:
'参数名         类型                    说明
'bytInpu        Byte(）                 字节数组
'======================================================================================================================
Public Function ByteToHexString(bytInpu() As Byte) As String
    Dim i           As Long
    Dim strReturn   As String
    
    For i = LBound(bytInpu) To UBound(bytInpu)
        If Len("" & Hex(bytInpu(i))) = 1 Then
            strReturn = strReturn & "0" & Hex(bytInpu(i))
        Else
            strReturn = strReturn & Hex(bytInpu(i))
        End If
    Next
    
    ByteToHexString = strReturn
End Function
'======================================================================================================================
'方法           ByteToHexString         将16进制字符串转换为字节组
'返回值         Byte()                  16进制字符串转换的字节组
'入参列表:
'参数名         类型                    说明
'bstrInput      String                  16进制字符串
'lngRetBytLen   Long(Optional)          指定返回的字节组的长度,0-按原始长度返回，<>0返回指定的长度，不足补齐（补0），多了截取
'======================================================================================================================
Public Function HexStringToByte(ByVal strInput As String, Optional ByVal lngRetBytLen As Long) As Byte()
    Dim arrReturn() As Byte
    Dim i           As Long
    Dim lngLen      As Long
    
    lngLen = Len(strInput)
    If lngRetBytLen <> 0 Then
        lngLen = lngLen \ 2
        If lngLen > lngRetBytLen Then
            lngLen = lngRetBytLen
        End If
        ReDim arrReturn(lngRetBytLen - 1)
    Else
        lngLen = lngLen \ 2
        ReDim arrReturn(lngLen - 1)
    End If
    
    For i = 0 To lngLen - 1
        arrReturn(i) = Val("&H" & Mid(strInput, 2 * i + 1, 2))
    Next
    
    HexStringToByte = arrReturn()
End Function

'======================================================================================================================
'方法           BytePadding             将指定字符串按照16字节补齐，
'返回值         Byte()                  补齐后的字符串字节组
'入参列表:
'参数名         类型                    说明
'strInput       String                  字符串
'lngVersion     Long(Optional,2)        字符串补齐的版本（ZLSM4.DLL的版本，以及加密算法前缀中的版本），1-空格补齐，>1:Chr(0)补齐
'lngPaddingNum  Long(Optional,16)        补齐的字节数，缺省按照16进制补齐
'======================================================================================================================
Public Function BytePadding(ByVal strInput As String, Optional ByVal lngVersion As Long = 2, Optional ByVal lngPaddingNum As Long = 16) As Byte()
    Dim arrReturn()     As Byte
    Dim lngLenBef       As Long
    Dim i               As Long
    Dim lngLenAft       As Long
    
    '先将字符串由 Unicode 转成系统的缺省码页
    arrReturn = StrConv(strInput, vbFromUnicode)
    lngLenBef = UBound(arrReturn) + 1
    '判断得到的数组的长度，若不是16的整数倍，则补空格或:Chr(0)
    lngLenAft = -Int(-lngLenBef / lngPaddingNum) * lngPaddingNum
    If lngLenBef <> lngLenAft Then
        ReDim Preserve arrReturn(lngLenAft - 1)
        For i = lngLenBef To lngLenAft - 1
            If lngVersion = 1 Then
                arrReturn(i) = 32
            Else
                arrReturn(i) = 0
            End If
        Next
    End If
    BytePadding = arrReturn()
End Function

Private Function GetKey(ByVal strKey As String, ByVal intType As Integer) As Byte()
    Dim arrReturn() As Byte
    Dim i           As Long
    If strKey <> "" Then
        arrReturn = HexStringToByte(strKey, 16)
    Else
        ReDim arrReturn(15)
        If intType = 0 Then
            For i = 0 To 15
                arrReturn(i) = i * 15
            Next
        ElseIf intType = 1 Then
            Rnd (-1)
            Randomize (SM4_CRYPT_RANDOMIZE_IV)
            For i = 0 To 15
                arrReturn(i) = Int(Rnd() * 256)
            Next
        ElseIf intType = 2 Then
            Rnd (-1)
            Randomize (SM4_CRYPT_RANDOMIZE_KEY)
            For i = 0 To 15
                arrReturn(i) = Int(Rnd() * 256)
            Next
        End If
    End If
    GetKey = arrReturn
End Function

Public Sub ApplyOEM_Picture(objPicture As Object, ByVal str属性 As String, Optional ByVal strProductName As String)
'针对各种图标应用OEM策略
    Dim strOEM As String
    Dim blnCorp As Boolean
    On Error Resume Next
    
    If strProductName = "" Then
        strProductName = GetSetting("ZLSOFT", "注册信息", "产品名称", "")
    End If

    If strProductName <> "中联" And strProductName <> "-" Then
        '处理状态栏图标的OEM策略
        If Right(str属性, 1) = "B" Then
            '表示产品图片
            blnCorp = False
            str属性 = Mid(str属性, 1, Len(str属性) - 1)
        Else
            '表示公司徽标
            blnCorp = True
        End If
        
        strOEM = mGetOEM(strProductName, blnCorp)
        If str属性 = "Picture" Then
            Set objPicture.Picture = LoadCustomPicture(strOEM)
        ElseIf str属性 = "Icon" Then
            Set objPicture.Icon = LoadCustomPicture(strOEM)
        End If
        
        If Err <> 0 Then
            Err.Clear
        End If
    
    End If
End Sub

Private Function mGetOEM(ByVal strAsk As String, Optional ByVal blnCorp As Boolean = True) As String
    '-------------------------------------------------------------
    '功能：返回每个字线的ASCII码
    '参数：
    '返回：
    '-------------------------------------------------------------
    Dim intBit As Integer
    Dim strCode As String
    
    'OEM图片有两种类型 ，一是指公司徽标，另一个是产品标识
    strCode = IIf(blnCorp = True, "OEM_", "PIC_")
    For intBit = 1 To Len(strAsk)
        '取每个字的ASCII码
        strCode = strCode & Hex(Asc(Mid(strAsk, intBit, 1)))
    Next
    mGetOEM = strCode
End Function

Public Function ActualLen(ByVal strAsk As String) As Long
    '--------------------------------------------------------------
    '功能：求取指定字符串的实际长度，用于判断实际包含双字节字符串的
    '       实际数据存储长度
    '参数：
    '       strAsk
    '返回：
    '-------------------------------------------------------------
    ActualLen = LenB(StrConv(strAsk, vbFromUnicode))
End Function

Public Function FromatSQL(ByVal strText As String, Optional ByVal blnCrlf As Boolean) As String
'功能：去掉TAB字符，两边空格，回车，最后只由单空格分隔。
'参数：strText=处理字符
'         blnCrlf=是否去掉换行符
    Dim i As Long
    
    If blnCrlf Then
        strText = Replace(strText, vbCrLf, " ")
        strText = Replace(strText, vbCr, " ")
        strText = Replace(strText, vbLf, " ")
    End If
    strText = Trim(Replace(strText, vbTab, " "))
    
    i = 5
    Do While i > 1
        strText = Replace(strText, String(i, " "), " ")
        If InStr(strText, String(i, " ")) = 0 Then i = i - 1
    Loop
    FromatSQL = strText
End Function

Public Function LoadCustomPicture(strID As String) As StdPicture
'功能:将资源文件中的指定资源生成磁盘文件
'参数:ID=资源号,strExt=要生成文件的扩展名(如BMP)
'返回:生成文件名
    Dim arrData() As Byte
    Dim intFile As Integer
    Dim strFile As String * 255, strR As String
    
    arrData = LoadResData(strID, "CUSTOM")
    intFile = FreeFile
    
    GetTempPath 255, strFile
    strR = Trim(Left(strFile, InStr(strFile, Chr(0)) - 1)) & CLng(Timer * 100) & ".pic"

    Open strR For Binary As intFile
    Put intFile, , arrData()
    Close intFile
    Set LoadCustomPicture = VB.LoadPicture(strR)
    Kill strR
End Function

Public Function OpenIme(Optional blnOpen As Boolean = False, Optional strImeName As String) As Boolean
'功能:打开中文输入法，或关闭输入法
'参数：strImeName-打开指定的输入法，没有指定时打开系统选项设置的缺省输入法
    Dim arrIme(99) As Long, lngCount As Long, strName As String * 255
    Dim strIme As String
    
    If strImeName = "不自动开启" Then OpenIme = True: Exit Function
    '用户没进行设置，就不处理
    If blnOpen Then
        If strImeName <> "" Then
            strIme = Trim(strImeName)
        End If
        If strIme = "" Then Exit Function                '要求打开输入法，但是又没有设置
    End If
    
    lngCount = GetKeyboardLayoutList(UBound(arrIme) + 1, arrIme(0))

    Do
        lngCount = lngCount - 1
        If ImmIsIME(arrIme(lngCount)) = 1 Then
            If blnOpen = True Then
                '需要打开输入法。接着判断是否指定输入法
                ImmGetDescription arrIme(lngCount), strName, Len(strName)
                If InStr(1, Mid(strName, 1, InStr(1, strName, Chr(0)) - 1), strIme) > 0 Then
                    If ActivateKeyboardLayout(arrIme(lngCount), 0) <> 0 Then
                        OpenIme = True
                        Exit Function
                    End If
                End If
            End If
        ElseIf blnOpen = False Then
            '不是中文输入法，正好是应了关闭输入法的请求
            If ActivateKeyboardLayout(arrIme(lngCount), 0) <> 0 Then OpenIme = True: Exit Function
        End If
    Loop Until lngCount = 0
    
    If blnOpen = False Then
        '由于windows Vista系统的英文输入法用ImmIsIME测试出是1的输入法,因此,需要单独处理.
        '刘兴宏:2008/09/03
        If ActivateKeyboardLayout(arrIme(0), 0) <> 0 Then OpenIme = True: Exit Function
    End If
End Function

Public Function GetAllSubKey(ByVal strKey As String) As Variant
'功能:获取某项的所有子项
'返回：=子项数组
    Dim lnghKey As Long, lngRet As Long, strName As String, lngIdx As Long
    Dim hRootKey As Long, strKeyName As String
    Dim strSubKey As Variant
    strSubKey = Array()
    lngIdx = 0: strName = String(256, Chr(0))
     If Not GetKeyValueInfo(strKey, "", hRootKey, strKeyName) Then Exit Function
    lngRet = RegOpenKey(hRootKey, strKeyName, lnghKey)
    If lngRet = 0 Then
        Do
            lngRet = RegEnumKey(lnghKey, lngIdx, strName, Len(strName))
            If lngRet = 0 Then
                ReDim Preserve strSubKey(UBound(strSubKey) + 1)
                strSubKey(UBound(strSubKey)) = Left(strName, InStr(strName, Chr(0)) - 1)
                lngIdx = lngIdx + 1
            End If
        Loop Until lngRet <> 0
    End If
    RegCloseKey lnghKey
    GetAllSubKey = strSubKey
End Function

Private Function GetKeyValueInfo(ByVal strKey As String, Optional ByVal strValueName As String, Optional ByRef hRootKey As REGRoot, Optional ByRef strSubKey As String, Optional ByRef lngType As Long) As Boolean
'功能：根据键位获取根键值与子健,以及值类型
'参数：strKey=注册表键位，如“HKEY_CURRENT_USER\Printers\DevModePerUser"
'          strValueName=变量名
'出参：
'          hRootKey=根键
'          strSubKey=子健
'          lngType=键类型
'返回：是否获取成功
    Dim strRoot As String, lngPos As String, hKey As Long
    Dim lngReturn As Long, strName As String * 255
    
    On Error GoTo errH
    hRootKey = 0: strSubKey = "": lngType = 0
    lngPos = InStr(strKey, "\")
    If lngPos = 0 Then Exit Function
    strRoot = Mid(strKey, 1, lngPos - 1)
    strSubKey = Mid(strKey, lngPos + 1)
    
    hRootKey = Decode(UCase(strRoot), "HKEY_CLASSES_ROOT", HKEY_CLASSES_ROOT, _
                                                                         "HKEY_CURRENT_USER", HKEY_CURRENT_USER, _
                                                                         "HKEY_LOCAL_MACHINE", HKEY_LOCaL_MaCHINE, _
                                                                         "HKEY_USERS", HKEY_USERS, _
                                                                         "HKEY_PERFORMANCE_DATA", HKEY_PERFORMANCE_DATA, _
                                                                         "HKEY_CURRENT_CONFIG", HKEY_CURRENT_CONFIG, _
                                                                         "HKEY_DYN_DATA", HKEY_DYN_DATA, 0)
    If hRootKey = 0 Then Exit Function
    If lngType <> -1 Then
        '使用查询方式打开，进行键名类型查询
        lngReturn = RegOpenKeyEx(hRootKey, strSubKey, 0, KEY_QUERY_VaLUE, hKey)
        If lngReturn <> ERROR_SUCCESS Then
            Exit Function
        End If
        If strValueName <> "" Then
            lngReturn = RegQueryValueEx_ValueType(hKey, strValueName, ByVal 0&, lngType, ByVal strName, Len(strName))
            'SetRegKey这种情况返回的类型为很大的数，数值不固定,因此设置为0，根据传入数据类型判断
            If lngReturn = ERROR_BADKEY Then
                If lngType < REG_NONE Or lngType > REG_MULTI_SZ Then lngType = REG_NONE
            End If
            '可能字段超长，长度不够，所以出错不退出
            'If lngReturn <> ERROR_SUCCESS Then: RegCloseKey (hKey): Exit Function
        End If
        RegCloseKey (hKey)
    End If
    GetKeyValueInfo = True
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
    Err.Clear
End Function

Public Function GetRegValue(ByVal strKey As String, ByVal strValueName As String, ByRef varValue As Variant, Optional blnOneString As Boolean = False) As Boolean
'功能：获取注册表中指定位置的值
'参数：strKey=注册表键位，如“HKEY_CURRENT_USER\Printers\DevModePerUser"
'          strValueName=变量名
'          strValue=变量值
'          strValueType=变量类型，默认为字符串
'           blnOneString = 对REG_EXPAND_SZ、REG_MULTI_SZ,REG_BINARY有效。-  True 则函数返回单一字符串，且不经任何处理，只去掉字符串尾！
'返回：是否读取成功
'说明：当前只对REG_SZ, REG_EXPAND_SZ, REG_MULTI_SZ，REG_DWORD，REG_BINARY实现了读取。没有查询到可以自动查找键名
    Dim hRootKey As REGRoot, strSubKey As String
    Dim lngReturn As Long
    Dim lngKey As Long, ruType As REGValueType
    Dim lngLength As Long, strBufVar() As String, lngBuf As Long, bytBuf() As Byte, strBuf As String
    Dim i As Long, strReturn As String, strTmp As String
    '不是有效的注册表键位,获取键名类型
    If Not GetKeyValueInfo(strKey, strValueName, hRootKey, strSubKey, ruType) Then Exit Function
    '打开变量
    lngReturn = RegOpenKeyEx(hRootKey, strSubKey, 0, KEY_QUERY_VaLUE, lngKey)
    If lngReturn <> ERROR_SUCCESS Then
        Exit Function
    End If
    On Error GoTo errH
    Select Case ruType
        Case REG_SZ, REG_EXPAND_SZ, REG_MULTI_SZ '字符串类型读取
'            lngReturn = RegQueryValueEx(lngKey, strValueName, 0, ruType, 0, lngLength)
'            If lngReturn <> ERROR_SUCCESS Then Err.Clear '可能出错，因此这样处理
            lngLength = 1024: strBuf = Space(lngLength)
            lngReturn = RegQueryValueEx_String(lngKey, strValueName, 0, ruType, strBuf, lngLength)
            If lngReturn <> ERROR_SUCCESS Then: RegCloseKey (lngKey): Exit Function
            Select Case ruType
                Case REG_SZ
                    varValue = TruncZero(strBuf)
                Case REG_EXPAND_SZ ' 扩充环境字符串，查询环境变量和返回定义值
                    If Not blnOneString Then
                        varValue = TruncZero(ExpandEnvStr(TruncZero(strBuf)))
                    Else
                        varValue = TruncZero(strBuf)
                    End If
                Case REG_MULTI_SZ ' 多行字符串
                    If Not blnOneString Then
                        If Len(strBuf) <> 0 Then ' 读到的是非空字符串，可以分割。
                            strBufVar = Split(Left$(strBuf, Len(strBuf) - 1), Chr$(0))
                        Else ' 若是空字符串，要定义S(0) ，否则出错！
                            ReDim strBufVar(0) As String
                        End If
                        ' 函数返回值，返回一个字符串数组？！
                        varValue = strBufVar()
                    Else
                        varValue = TruncZero(strBuf)
                    End If
            End Select
        Case REG_DWORD
            lngReturn = RegQueryValueEx_Long(lngKey, strValueName, ByVal 0&, ruType, lngBuf, Len(lngBuf))
            If lngReturn <> ERROR_SUCCESS Then: RegCloseKey (lngKey): varValue = 0: Exit Function
            varValue = lngBuf
        Case REG_BINARY
            lngReturn = RegQueryValueEx_BINARY(lngKey, strValueName, 0, ruType, ByVal 0, lngLength)
            If lngReturn <> ERROR_SUCCESS And lngReturn <> ERROR_MORE_DATA Then
                RegCloseKey lngKey: Exit Function
                If blnOneString Then
                    varValue = "00"
                Else
                    ReDim bytBuf(0)
                    varValue = bytBuf()
                End If
            End If
            ReDim bytBuf(lngLength - 1)
            lngReturn = RegQueryValueEx_BINARY(lngKey, strValueName, 0, ruType, bytBuf(0), lngLength)
            If lngReturn <> ERROR_SUCCESS And lngReturn <> ERROR_MORE_DATA Then
                RegCloseKey lngKey: Exit Function
                If blnOneString Then
                    varValue = "00"
                Else
                    ReDim bytBuf(0)
                    varValue = bytBuf()
                End If
            End If
            If lngLength <> UBound(bytBuf) + 1 Then
               ReDim Preserve bytBuf(0 To lngLength - 1) As Byte
            End If
            ' 返回字符串，注意：要将字节数组进行转化！
            If blnOneString Then
                '循环数据，把字节转换为16进制字符串
                For i = LBound(bytBuf) To UBound(bytBuf)
                   strTmp = CStr(Hex(bytBuf(i)))
                   If (Len(strTmp) = 1) Then strTmp = "0" & strTmp
                   strReturn = strReturn & " " & strTmp
                Next i
                varValue = Trim$(strReturn)
            Else
                varValue = bytBuf()
            End If
    End Select
    RegCloseKey lngKey
    GetRegValue = True
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
End Function

Public Function ExpandEnvStr(ByVal strInput As String) As String
'功能：将字符串中的环境变量替换为常规值
'         strInput=包含环境变量的字符串
'返回：用实际的值替换字符串中的环境变量后的字符串
    '// 如： %PATH% 则返回 "c:\;c:\windows;"
    Dim lngLen As Long, strBuf As String, strOld As String
    strOld = strInput & "  " ' 不知为什么要加两个字符，否则返回值会少最后两个字符！
    strBuf = "" '// 不支持Windows 95
    '// get the length
    lngLen = ExpandEnvironmentStrings(strOld, strBuf, lngLen)
    '// 展开字符串
    strBuf = String$(lngLen - 1, Chr$(0))
    lngLen = ExpandEnvironmentStrings(strOld, strBuf, LenB(strBuf))
    '// 返回环境变量
    ExpandEnvStr = TruncZero(strBuf)
End Function

Public Sub OSWait(ByVal lngMilliseconds As Long)
'功能：执行挂起一段时间
'lngMilliseconds=毫秒数，1000毫秒=1秒
    Call Sleep(lngMilliseconds)
End Sub

Public Function Is64bit() As Boolean
    '******************************************************************************************************************
    '功能：是否是64位系统
    '返回：
    '******************************************************************************************************************
    Dim handle As Long
    Dim lngFunc As Long
        
    lngFunc = 0
    handle = GetProcAddress(GetModuleHandle("kernel32"), "IsWow64Process")
    If handle > 0 Then
        IsWow64Process GetCurrentProcess(), lngFunc
    End If
    Is64bit = lngFunc <> 0
End Function

Public Function zlGetComLib() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取公共部件相关对象
    '返回:获取成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-05-15 15:34:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If gobjComlib Is Nothing Then
        On Error Resume Next
        Set gobjComlib = GetObject("", "zl9Comlib.clsComlib")
        If gobjComlib Is Nothing Then
            Err.Clear
            Set gobjComlib = CreateObject("zl9Comlib.clsComlib")
            Err.Clear
        End If
    End If
    zlGetComLib = Not gobjComlib Is Nothing
End Function

Public Function Decode(ParamArray arrPar() As Variant) As Variant
'功能：模拟Oracle的Decode函数
    Dim varValue As Variant, i As Integer
    
    i = 1
    varValue = arrPar(0)
    Do While i <= UBound(arrPar)
        If i = UBound(arrPar) Then
            Decode = arrPar(i): Exit Function
        ElseIf varValue = arrPar(i) Then
            Decode = arrPar(i + 1): Exit Function
        Else
            i = i + 2
        End If
    Loop
End Function

Public Function TruncZero(ByVal strInput As String) As String
'功能：去掉字符串中\0以后的字符
    Dim lngPos As Long
    
    lngPos = InStr(strInput, Chr(0))
    If lngPos > 0 Then
        TruncZero = Mid(strInput, 1, lngPos - 1)
    Else
        TruncZero = strInput
    End If
End Function

Public Sub SetWindowsInTaskBar(ByVal lnghwnd As Long, ByVal blnShow As Boolean)
'功能：设置窗体是否在任务条上显示
    Dim LngStyle As Long
    
    LngStyle = GetWindowLong(lnghwnd, GWL_EXSTYLE)
    If blnShow Then
        LngStyle = LngStyle Or &H40000
    Else
        LngStyle = LngStyle And Not &H40000
    End If
    Call SetWindowLong(lnghwnd, GWL_EXSTYLE, LngStyle)
End Sub

Public Function GetGeneralAccountKey(ByRef strKey As String) As String
    Dim arrTmp()    As Byte
    Dim i           As Long
    arrTmp = HexStringToByte(strKey, 16)
    For i = LBound(arrTmp) To UBound(arrTmp)
        If i Mod 2 = 0 Then
            arrTmp(i) = 255 - arrTmp(i)
        ElseIf i Mod 3 = 0 Then
            arrTmp(i) = (arrTmp(i) + i) Mod 256
        End If
    Next
    GetGeneralAccountKey = ByteToHexString(arrTmp)
End Function


Public Function GetZLOptions(ByVal strParNO As String) As ADODB.Recordset
'参数：strParNO-参数号，如果有多个，则以逗号隔开
    Dim strSQL As String
    
    On Error GoTo errH
    If InStr(strParNO, ",") > 0 Then
        strSQL = "Select 参数号,参数名,Nvl(参数值,缺省值) as 参数值  From zlOptions Where 参数号 in(Select Column_value From Table(f_num2list([1])))"
        Set GetZLOptions = OpenSQLRecord(strSQL, "读取zlOptions参数", strParNO)
    Else
        strSQL = "Select 参数号,参数名,Nvl(参数值,缺省值) as 参数值 From zlOptions Where 参数号=[1]"
        Set GetZLOptions = OpenSQLRecord(strSQL, "读取zlOptions参数", Val(strParNO))
    End If
    
    Exit Function
errH:
    MsgBox Err.Description & vbCrLf & strSQL, vbExclamation, "获取管理工具参数"
    Set GetZLOptions = New ADODB.Recordset
End Function
