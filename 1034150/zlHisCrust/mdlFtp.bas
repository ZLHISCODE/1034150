Attribute VB_Name = "mdlFTP"
Option Explicit
'**************************
'功能:FTP的处理方式
'编写修改:祝庆
'**************************

'************************************************************************************
'打开连接Internet的会话
'************************************************************************************
Public Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" _
   (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, _
   ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
'************************************************************************************
'说明：
'************************************************************************************
'    sAgent--要调用Internet对话的应用程序名
'    lAccessType--请求的网络访问的类型，包括：
'************************************************************************************
'        常量                                                          值         说明
'        INTERNET_OPEN_TYPE_PRECONFIG        0          预配置（缺省）
'        INTERNET_OPEN_TYPE_DIRECT               1          直接连接到Internet
'        INTERNET_OPEN_TYPE_PROXY                3          通过代理服务器连接
'************************************************************************************
'    备注：如果lAccessType设置为INTERNET_OPEN_TYPE_PRECONFIG，连接时就要基于
'    HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings
'    注册表路径下的注册表数值ProxyEnable、ProxyServer和 ProxyOverride
'************************************************************************************
'    sProxyName--指定代理服务器的名字，访问类型设置为INTERNET_OPEN_TYPE_PROXY才有效
'    sProxyBypass--指定代理服务器的名字或地址，有设置此项时lpszProxyName指定的将失效
'    lFlags--会话的选项，可包括下列值：
'************************************************************************************
'        常量                                                         值          说明
'        INTERNET_FLAG_DONT_CACHE                           不对数据进行本地缓冲或通过网关服务器缓冲
'        INTERNET_FLAG_ASYNC                                      使用异步连接
'        INTERNET_FLAG_OFFLINE                                   只通过永久缓冲进行下载操作
'************************************************************************************
'函数返回值：如果函数调用失败，lngINet 为0。
'************************************************************************************

'************************************************************************************
'建立Internet连接，打开FTP会话
'************************************************************************************
Public Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" _
    (ByVal hInternetSession As Long, ByVal sServerName As String, _
    ByVal nServerPort As Integer, ByVal sUsername As String, _
    ByVal sPassword As String, ByVal lService As Long, _
    ByVal lFlags As Long, ByVal lContext As Long) As Long
'************************************************************************************
'说明：
'************************************************************************************
'    hInternetSession--函数InternetOpen返回的Internet会话句柄
'    sServerName--要连接的服务器的名称或IP
'    nServerPort--要连接的Internet端口
'    sUsername--登录的用户帐号
'    sPassword--登录的口令
'    lService--要连接的服务器类型（这里是连接FTP服务器，连接的类型为常数INTERNET_SERVICE_FTP）
'    lFlags--如果传递x8000000，连接将使用被动FTP语义，传递0使用非被动语义
'    lContext--当使用回调函数时使用该参数，不使用回调服务传递0
'************************************************************************************
'函数返回值：如果函数调用失败，lngINetConn 为0
'************************************************************************************

'************************************************************************************
'从FTP服务器上下载一个文件
'************************************************************************************
Public Declare Function FtpGetFile Lib "wininet.dll" Alias "FtpGetFileA" _
    (ByVal hFtpSession As Long, ByVal lpszRemoteFile As String, _
    ByVal lpszNewFile As String, ByVal fFailIfExists As Boolean, _
    ByVal dwFlagsAndAttributes As Long, ByVal dwFlags As Long, _
    ByVal dwContext As Long) As Boolean
'************************************************************************************
'说明：
'************************************************************************************
'    hFtpSession--函数InternetConnect返回的Internet连接句柄
'    lpszRemoteFile--想要获得的FTP服务器上的文件名
'    lpszNewFile--要保存在本地机器中的文件名
'    fFailIfExists--0（替换本地文件）或1 （如果本地文件已经存在则调用失败）。
'    dwFlagsAndAttributes--用来指定本地文件的文件属性，传递0忽略
'    dwFlags--文件的传输方式可能包括下列值：
'************************************************************************************
'        常量                                                         值          说明
'        FTP_TRANSFER_TYPE_ASCII                   1           用ASCII 传输文件（A类传输方法）
'        FTP_TRANSFER_TYPE_BINARY                 2           用二进制传输文件（B类传输方法）
'************************************************************************************
'    dwContext--要取回的文件的描述表标识符
'************************************************************************************
'函数返回值：如果函数调用失败，blnRC 为FALSE
'************************************************************************************

'Public Declare Function FtpPutFile Lib "wininet.dll" Alias "FtpPutFileA" _
'(ByVal hConnect As Long, ByVal lpszLocalFile As String, _
'ByVal lpszNewRemoteFile As String, ByVal dwFlags As Long, _
'ByVal dwContext As Long) As Boolean


'************************************************************************************
'关闭Internet连接
'************************************************************************************
Public Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer
'************************************************************************************
'说明：
'************************************************************************************
'hInet--要关闭的会话（InternetOpen）或连接（InternetConnect）句柄
'************************************************************************************
'函数返回值：
'************************************************************************************

'************************************************************************************
'常量定义
Public Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Public Const INTERNET_OPEN_TYPE_DIRECT = 1
Public Const INTERNET_OPEN_TYPE_PROXY = 3
Public Const INTERNET_SERVICE_FTP = 1
Public Const FTP_TRANSFER_TYPE_BINARY = &H2
Public Const FTP_TRANSFER_TYPE_ASCII = &H1
'************************************************************************************

'Public gstrFtpServer As String    'FTP服务器
'Public gstrFtpUser As String      'FTP用户名
'Public gstrFtpPassword As String  'FTP密码
'Public gstrFtpPort                'FTP端口

Public glngINet As Long
Public glngINetConn As Long

Public Function FtpDownFile(ByVal srcPath As String, ByVal descPath As String) As Boolean
        FtpDownFile = FtpGetFile(glngINetConn, srcPath, descPath, False, 0, FTP_TRANSFER_TYPE_BINARY, 0)
End Function

Public Function FtpupFile(ByVal srcPath As String, ByVal descPath As String) As Boolean
    ' FtpupFile = FtpPutFile(glngINetConn, srcPath, descPath, FTP_TRANSFER_TYPE_BINARY, 0)
    Dim API As New APILoad
    Dim strSystemDirectory As String '系统目录
    
    strSystemDirectory = GetWinSystemPath
    FtpupFile = API.ExecuteAPI(strSystemDirectory & "\wininet.dll", "FtpPutFile " & glngINetConn & "," & srcPath & "," & descPath & ",FTP_TRANSFER_TYPE_BINARY,0")

End Function


