Attribute VB_Name = "mdlFTP"
Option Explicit
'**************************
'����:FTP�Ĵ���ʽ
'��д�޸�:ף��
'**************************

'************************************************************************************
'������Internet�ĻỰ
'************************************************************************************
Public Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" _
   (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, _
   ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
'************************************************************************************
'˵����
'************************************************************************************
'    sAgent--Ҫ����Internet�Ի���Ӧ�ó�����
'    lAccessType--�����������ʵ����ͣ�������
'************************************************************************************
'        ����                                                          ֵ         ˵��
'        INTERNET_OPEN_TYPE_PRECONFIG        0          Ԥ���ã�ȱʡ��
'        INTERNET_OPEN_TYPE_DIRECT               1          ֱ�����ӵ�Internet
'        INTERNET_OPEN_TYPE_PROXY                3          ͨ���������������
'************************************************************************************
'    ��ע�����lAccessType����ΪINTERNET_OPEN_TYPE_PRECONFIG������ʱ��Ҫ����
'    HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings
'    ע���·���µ�ע�����ֵProxyEnable��ProxyServer�� ProxyOverride
'************************************************************************************
'    sProxyName--ָ����������������֣�������������ΪINTERNET_OPEN_TYPE_PROXY����Ч
'    sProxyBypass--ָ����������������ֻ��ַ�������ô���ʱlpszProxyNameָ���Ľ�ʧЧ
'    lFlags--�Ự��ѡ��ɰ�������ֵ��
'************************************************************************************
'        ����                                                         ֵ          ˵��
'        INTERNET_FLAG_DONT_CACHE                           �������ݽ��б��ػ����ͨ�����ط���������
'        INTERNET_FLAG_ASYNC                                      ʹ���첽����
'        INTERNET_FLAG_OFFLINE                                   ֻͨ�����û���������ز���
'************************************************************************************
'��������ֵ�������������ʧ�ܣ�lngINet Ϊ0��
'************************************************************************************

'************************************************************************************
'����Internet���ӣ���FTP�Ự
'************************************************************************************
Public Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" _
    (ByVal hInternetSession As Long, ByVal sServerName As String, _
    ByVal nServerPort As Integer, ByVal sUsername As String, _
    ByVal sPassword As String, ByVal lService As Long, _
    ByVal lFlags As Long, ByVal lContext As Long) As Long
'************************************************************************************
'˵����
'************************************************************************************
'    hInternetSession--����InternetOpen���ص�Internet�Ự���
'    sServerName--Ҫ���ӵķ����������ƻ�IP
'    nServerPort--Ҫ���ӵ�Internet�˿�
'    sUsername--��¼���û��ʺ�
'    sPassword--��¼�Ŀ���
'    lService--Ҫ���ӵķ��������ͣ�����������FTP�����������ӵ�����Ϊ����INTERNET_SERVICE_FTP��
'    lFlags--�������x8000000�����ӽ�ʹ�ñ���FTP���壬����0ʹ�÷Ǳ�������
'    lContext--��ʹ�ûص�����ʱʹ�øò�������ʹ�ûص����񴫵�0
'************************************************************************************
'��������ֵ�������������ʧ�ܣ�lngINetConn Ϊ0
'************************************************************************************

'************************************************************************************
'��FTP������������һ���ļ�
'************************************************************************************
Public Declare Function FtpGetFile Lib "wininet.dll" Alias "FtpGetFileA" _
    (ByVal hFtpSession As Long, ByVal lpszRemoteFile As String, _
    ByVal lpszNewFile As String, ByVal fFailIfExists As Boolean, _
    ByVal dwFlagsAndAttributes As Long, ByVal dwFlags As Long, _
    ByVal dwContext As Long) As Boolean
'************************************************************************************
'˵����
'************************************************************************************
'    hFtpSession--����InternetConnect���ص�Internet���Ӿ��
'    lpszRemoteFile--��Ҫ��õ�FTP�������ϵ��ļ���
'    lpszNewFile--Ҫ�����ڱ��ػ����е��ļ���
'    fFailIfExists--0���滻�����ļ�����1 ����������ļ��Ѿ����������ʧ�ܣ���
'    dwFlagsAndAttributes--����ָ�������ļ����ļ����ԣ�����0����
'    dwFlags--�ļ��Ĵ��䷽ʽ���ܰ�������ֵ��
'************************************************************************************
'        ����                                                         ֵ          ˵��
'        FTP_TRANSFER_TYPE_ASCII                   1           ��ASCII �����ļ���A�ഫ�䷽����
'        FTP_TRANSFER_TYPE_BINARY                 2           �ö����ƴ����ļ���B�ഫ�䷽����
'************************************************************************************
'    dwContext--Ҫȡ�ص��ļ����������ʶ��
'************************************************************************************
'��������ֵ�������������ʧ�ܣ�blnRC ΪFALSE
'************************************************************************************

'Public Declare Function FtpPutFile Lib "wininet.dll" Alias "FtpPutFileA" _
'(ByVal hConnect As Long, ByVal lpszLocalFile As String, _
'ByVal lpszNewRemoteFile As String, ByVal dwFlags As Long, _
'ByVal dwContext As Long) As Boolean


'************************************************************************************
'�ر�Internet����
'************************************************************************************
Public Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer
'************************************************************************************
'˵����
'************************************************************************************
'hInet--Ҫ�رյĻỰ��InternetOpen�������ӣ�InternetConnect�����
'************************************************************************************
'��������ֵ��
'************************************************************************************

'************************************************************************************
'��������
Public Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Public Const INTERNET_OPEN_TYPE_DIRECT = 1
Public Const INTERNET_OPEN_TYPE_PROXY = 3
Public Const INTERNET_SERVICE_FTP = 1
Public Const FTP_TRANSFER_TYPE_BINARY = &H2
Public Const FTP_TRANSFER_TYPE_ASCII = &H1
'************************************************************************************

'Public gstrFtpServer As String    'FTP������
'Public gstrFtpUser As String      'FTP�û���
'Public gstrFtpPassword As String  'FTP����
'Public gstrFtpPort                'FTP�˿�

Public glngINet As Long
Public glngINetConn As Long

Public Function FtpDownFile(ByVal srcPath As String, ByVal descPath As String) As Boolean
        FtpDownFile = FtpGetFile(glngINetConn, srcPath, descPath, False, 0, FTP_TRANSFER_TYPE_BINARY, 0)
End Function

Public Function FtpupFile(ByVal srcPath As String, ByVal descPath As String) As Boolean
    ' FtpupFile = FtpPutFile(glngINetConn, srcPath, descPath, FTP_TRANSFER_TYPE_BINARY, 0)
    Dim API As New APILoad
    Dim strSystemDirectory As String 'ϵͳĿ¼
    
    strSystemDirectory = GetWinSystemPath
    FtpupFile = API.ExecuteAPI(strSystemDirectory & "\wininet.dll", "FtpPutFile " & glngINetConn & "," & srcPath & "," & descPath & ",FTP_TRANSFER_TYPE_BINARY,0")

End Function


