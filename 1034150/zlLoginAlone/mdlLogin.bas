Attribute VB_Name = "mdlLogin"
Option Explicit
'���̻�ȡ
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
'�Ƿ���64λ���̣�Is64bit��
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function IsWow64Process Lib "kernel32" (ByVal hProc As Long, bWow64Process As Long) As Long
'����ϵͳ�п��õ����뷨�����������뷨����Layout,����Ӣ�����뷨��
Private Declare Function GetKeyboardLayoutList Lib "user32" (ByVal nBuff As Long, lpList As Long) As Long
'��ȡĳ�����뷨������
Private Declare Function ImmGetDescription Lib "imm32.dll" Alias "ImmGetDescriptionA" (ByVal hkl As Long, ByVal lpsz As String, ByVal uBufLen As Long) As Long
'�ж�ĳ�����뷨�Ƿ��������뷨
Private Declare Function ImmIsIME Lib "imm32.dll" (ByVal hkl As Long) As Long
'�л���ָ�������뷨��
Private Declare Function ActivateKeyboardLayout Lib "user32" (ByVal hkl As Long, ByVal flags As Long) As Long
'��������(ComputerName)
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
'��ͣ(Wait)
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Const GWL_EXSTYLE = (-20)
Public Const WinStyle = &H40000 'Forces a top-level window onto the taskbar when the window is visible.ǿ��һ���ɼ��Ķ����Ӵ�����������
Public Const SWP_NOSIZE = &H1
Public Const SWP_SHOWWINDOW = &H40
Public Const HWND_TOPMOST = -1
'��ʱIP��ȡ
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
'SM4����
'/**
' * \brief          SM4-ECB block encryption/decryption
' * \param mode     SM4_ENCRYPT or SM4_DECRYPT
' * \param length   length of the input data
' * \param input    input block
' * \param output   output block
' */
Private Declare Function sm4_crypt_ecb Lib "zlSm4.dll" (ByVal Mode As Long, ByVal Length As Long, key As Byte, in_put As Byte, out_put As Byte) As Long
'SM4�����������
'/**
' * \brief          SM4-CBC buffer encryption/decryption
' * \param mode     SM4_ENCRYPT or SM4_DECRYPT
' * \param length   length of the input data
' * \param iv       initialization vector (updated after use)
' * \param input    buffer holding the input data
' * \param output   buffer holding the output data
' */
Private Declare Function sm4_crypt_cbc Lib "zlSm4.dll" (ByVal Mode As Long, ByVal Length As Long, iv As Byte, key As Byte, in_put As Byte, out_put As Byte) As Long
'��ȡ�ַ����Ĺ�ϣ����
'/**
' * \brief          Output = SM3( input buffer )
' *
' * \param input    buffer holding the  data
' * \param ilen     length of the input data
' * \param output   SM3 checksum result
' */
Private Declare Sub sm3_hash Lib "zlSm4.dll" Alias "sm3" (in_put As Byte, ByVal Length As Long, out_put As Byte)
'��ȡ�ļ���sm��ϣ����
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
'HMAC����Կ��صĹ�ϣ������Ϣ��֤�룬HMAC�������ù�ϣ�㷨����һ����Կ��һ����ϢΪ���룬����һ����ϢժҪ��Ϊ�����
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
'��ȡZLSM4���޸İ汾
'1:ֻ֧��sm4_crypt_ecb,sm4_crypt_cbc
'2:����֧��sm3��sm3_file��sm3_hmac��sm_version
'/**
' * \brief          Output = zlSM4.DLL Version
' */
Private Declare Function get_sm_version Lib "zlSm4.dll" Alias "sm_version" () As Long

Private Enum CrypeMode
    CM_Encrypt = 1   '����
    CM_Decrypt = 0   '����
End Enum
Private M_SM4_VERSION As Long
Public Const SM4_CRYPT_RANDOMIZE_KEY    As Long = 999  'sm4�����㷨��Կ���������������
Public Const SM4_CRYPT_RANDOMIZE_IV     As Long = 666   'sm4�����㷨��ʼ�������������������
Public Const G_PASSWORD_KEY             As String = "3357F1F2CA0341A5B75DBA7F35666715"

'ע���ؼ��ָ�����
Private Enum REGRoot
    HKEY_CLASSES_ROOT = &H80000000 '��¼Windows����ϵͳ�����������ļ��ĸ�ʽ�͹�����Ϣ����Ҫ��¼��ͬ�ļ����ļ�����׺����֮��Ӧ��Ӧ�ó��������Ӽ��ɷ�Ϊ���࣬һ�����Ѿ�ע��ĸ����ļ�����չ���������Ӽ�ǰ�涼��һ������������һ���Ǹ����ļ������й���Ϣ��
    HKEY_CURRENT_USER = &H80000001 '�˸��������˵�ǰ��¼�û����û������ļ���Ϣ����Щ��Ϣ��֤��ͬ���û���¼�����ʱ��ʹ���Լ��ĸ��Ի����ã������Լ������ǽֽ���Լ����ռ��䡢�Լ��İ�ȫ����Ȩ�޵ȡ�
    HKEY_LOCaL_MaCHINE = &H80000002 '�˸��������˵�ǰ��������������ݣ���������װ��Ӳ���Լ���������á���Щ��Ϣ��Ϊ���е��û���¼ϵͳ����ġ���������ע��������Ӵ�Ҳ������Ҫ�ĸ�����
    HKEY_USERS = &H80000003 '�˸�������Ĭ���û�����Ϣ��Default�Ӽ�����������ǰ��¼�û�����Ϣ��
    HKEY_PERFORMANCE_DATA = &H80000004 '��Windows NT/2000/XPע�������Ȼû��HKEY_DYN_DATA����������ȴ������һ����Ϊ��HKEY_ PERFOR MANCE_DATA����������ϵͳ�еĶ�̬��Ϣ���Ǵ���ڴ��Ӽ��С�ϵͳ�Դ���ע���༭���޷������˼�
    HKEY_CURRENT_CONFIG = &H80000005  '�˸���ʵ������HKEY_LOCAL_MACHINE�е�һ���֣����д�ŵ��Ǽ������ǰ���ã�����ʾ������ӡ���������������Ϣ�ȡ������Ӽ���HKEY_LOCAL_ MACHINE\ Config\0001��֧�µ�������ȫһ����
    HKEY_DYN_DATA = &H80000006 '�˸����б���ÿ��ϵͳ����ʱ��������ϵͳ���ú͵�ǰ������Ϣ���������ֻ������Windows 98�С�
End Enum

'ע�����������
Private Enum REGValueType
    REG_NONE = 0                       ' No value type
    REG_SZ = 1 'Unicode���ս��ַ���
    REG_EXPAND_SZ = 2 'Unicode���ս��ַ���
    REG_BINARY = 3 '��������ֵ
    REG_DWORD = 4 '32-bit ����
    REG_DWORD_BIG_ENDIAN = 5
    REG_LINK = 6
    REG_MULTI_SZ = 7 ' ��������ֵ��
End Enum
'�򿪴���
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
'ע������Ȩ
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
' ���价���ַ�����������������������д������Ϊ��ࡣҲ����˵�����ɰٷֺŷ�������Ļ���������ת�����Ǹ����������ݡ����磬��%path%�������������·����
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
Public gobjRegister As Object               'ע����Ȩ����zlRegister
Public gcnOracle As ADODB.Connection     '�������ݿ�����
Public gstrCommand As String '������

Public gobjFile As New FileSystemObject
Public gclsLogin As clsLogin '��¼����
Public gintCallType As Integer '0-��չʾ�޸����������������,1-��ʾ�޸�����,2-��ʵ����������
Public gblnExitApp  As Boolean '�Ƿ���Ϊ�ظ����У���Ҫ�˳���������

'clsLogin���Ի���
Public gobjEmr             As Object   'EMR�°���Ӳ���

Public gstrInputPwd        As String   'InputPwd����
Public gstrServerName      As String   'ServerName����
Public gblnTransPwd        As Boolean  'blnTransPwd����

Public gstrInputUser        As String  '������û���������zyk��δת����Сд
Public gstrDBUser           As String  '��¼�û���������ZYK����д
Public gstrUserID           As String  '��¼�û�ID������133
Public gstrUserName         As String  '��¼��Ա����������������
Public gstrDeptName         As String  '��¼�û���ȱʡ��������
Public gstrDeptNameTerminal As String  '�û���¼���������Ĳ�������
Public gstrDeptID           As String  '��¼�û�ȱʡ����ID
Public gstrIP               As String  '��¼�ͻ���IP��ַ
Public gstrSessionID        As String  '��ǰ�ỰID

Public gblnSysOwner        As Boolean  '�Ƿ�ϵͳ������
Public gstrConnString      As String   '�����ַ���
Public gstrSystems         As String   '������ѡ���ϵͳ
Public gblnCancel          As Boolean  '�Ƿ�ȡ���˳�
Public gstrMenuGroup       As String   '�˵�������

Public gstrStation         As String   '�û���¼����վ����
Public gstrNodeNo          As String   'վ����
Public gstrNodeName        As String   'վ������
Public gblnEMRProxy         As Boolean
Public gstrEMRPwd           As String
Public gstrEMRUser          As String
Public gstrUsageOccasion    As String      'ʹ�ó��ϣ���Ҫ������죬LIS������¼���ⲿ���øñ���
Public glngHelperMainType   As Long         '����������������
Public glngDBPass           As Long         '0-�Զ��жϣ�1-���ݿ����룬2-�����ݿ�����
Public glngParallelID       As Long         '�ͻ��˹�����֤���д���ID
Public gblnTimer            As Boolean  '�Ƿ�ʱ�������Ŀͻ��˸��¼��
Public glngInstanceCount    As Long     'ʵ������
Public gobjComlib           As Object
Public gcolTableField       As Collection

Public Sub CollectTableField(ByVal strInfo As String)
'���ܣ��ռ����ֶ��Ƿ���ڣ������������
'������
'  strInfo��ָ��Ҫ�ռ��ı������ֶ�����  ��ʽ������1.�ֶ���1[,����2.�ֶ���2 ...]
    
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
            Set rsTemp = OpenSQLRecord(strSQL, "�ռ��ֶ��Ƿ����", strTable, strField)
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
    MsgBox Err.Number & "�� " & Err.Description, vbInformation, App.Title
End Sub

Public Function GetExistsField(ByVal strTable As String, ByVal strField As String) As Boolean
'���ܣ�����ֶ��Ƿ���ڣ������ԣ�
'���أ�True���ڣ�False������

    If gcolTableField Is Nothing Then Exit Function
    
    strTable = Trim(UCase(strTable))
    strField = Trim(UCase(strField))
    
    GetExistsField = False
    On Error Resume Next
    GetExistsField = Val(gcolTableField(strTable & "_" & strField)) = 1
    On Error GoTo 0
End Function

Public Sub SetAppBusyState()
'���������̶���δ�������ʱ���滻��ִ�������̹���ʱ�����ġ����������𡱶Ի���
    On Error Resume Next
    App.OleServerBusyMsgTitle = App.ProductName
    App.OleRequestPendingMsgTitle = App.ProductName
    
    App.OleServerBusyMsgText = "���������ڴ����������ĵȴ���"
    App.OleRequestPendingMsgText = "�������������������ĵȴ���"
    
    App.OleServerBusyTimeout = 3000
    App.OleRequestPendingTimeout = 10000
    Err.Clear
End Sub

Public Function ShowSplash(Optional ByVal bytType As Byte, Optional ByVal blnRefresh As Boolean) As Boolean
'bytType:0-�´��壻1-�ϴ���
    Dim strUnitName As String, intCount As Integer
    Dim objPic As IPictureDisp
    '��ע����л�ȡ�û�ע�������Ϣ,����û���λ���Ʋ�Ϊ��,����ʾ���ִ���
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "��ʾ", "")
    strUnitName = GetSetting("ZLSOFT", "ע����Ϣ", "��λ����", "")
    
    If bytType = 0 Then
        With frmUserLogin
            If strUnitName <> "" And strUnitName <> "-" Then
                    '��������Ҫ����
                    '��ʱ�Ϳ�ʼ����clsComLib��ʵ��
                    Call ApplyOEM_Picture(.ImgIndicate, "Picture")
                    Call ApplyOEM_Picture(.imgPic, "PictureB")
                    If gobjFile.FileExists(gstrSetupPath & "\�����ļ�\logo_login.jpg") Then
                        Set objPic = LoadPicture(gstrSetupPath & "\�����ļ�\logo_login.jpg")
                        .picHos.Visible = True
                        .picHos.Height = IIf(objPic.Height < 2385, objPic.Height, 2385) '159����
                        .picHos.Width = IIf(objPic.Width < 4500, objPic.Width, 4730) '322����
                        .picHos.PaintPicture objPic, 0, 0, .picHos.Width, .picHos.Height
                    Else
                        .picHos.Visible = False
                    End If
                    .LblProductName = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒȫ��", "")
                    If Len(.LblProductName) > 10 Then
                        .LblProductName.FontSize = 15.75 '����
                    Else
                        .LblProductName.FontSize = 21.75 '����
                    End If
                    .lbltag = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒϵ��", "")
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
                .lbl����֧����.Caption = GetSetting("ZLSOFT", "ע����Ϣ", "����֧����", "")
                
                .LblProductName = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒȫ��", "")
                .lbltag = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒϵ��", "")
                strUnitName = GetSetting("ZLSOFT", "ע����Ϣ", "������", "")
                .lbl������.Caption = ""
                For intCount = 0 To UBound(Split(strUnitName, ";"))
                    .lbl������.Caption = .lbl������.Caption & Split(strUnitName, ";")(intCount) & vbCrLf
                Next
                Call ApplyOEM_Picture(.ImgIndicate, "Picture")
                If gobjFile.FileExists(gstrSetupPath & "\�����ļ�\logo_login.jpg") Then
                    Set objPic = LoadPicture(gstrSetupPath & "\�����ļ�\logo_login.jpg")
                    .picHos.Visible = True
                    .picHos.Height = IIf(objPic.Height < 2745, objPic.Height, 2745) '183����
                    .picHos.Width = IIf(objPic.Width < 4845, objPic.Width, 4845) '323����
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
                    '��������Ҫ����
                    '��ʱ�Ϳ�ʼ����clsComLib��ʵ��
                    Call ApplyOEM_Picture(.ImgIndicate, "Picture")
                    Call ApplyOEM_Picture(.imgPic, "PictureB")
                    If gobjFile.FileExists(gstrSetupPath & "\�����ļ�\logo_login.jpg") Then
                        Set objPic = LoadPicture(gstrSetupPath & "\�����ļ�\logo_login.jpg")
                        .picHos.Visible = True
                        .picHos.Height = IIf(objPic.Height < 2745, objPic.Height, 2745) '183����
                        .picHos.Width = IIf(objPic.Width < 4845, objPic.Width, 4845) '323����
                        .picHos.PaintPicture objPic, 0, 0, .picHos.Width, .picHos.Height
                    Else
                        .picHos.Visible = False
                    End If
                    If InStr(gstrCommand, "=") <= 0 Then .Show
                    
                    .lblGrant = Replace(strUnitName, ";", vbCrLf)
                    strUnitName = GetSetting("ZLSOFT", "ע����Ϣ", "������", "")
                    If Trim(strUnitName) = "" Then
                        .Label3.Visible = False
                        .lbl������.Visible = False
                    Else
                        .Label3.Visible = True
                        .lbl������.Visible = True
                        .lbl������.Caption = ""
                        For intCount = 0 To UBound(Split(strUnitName, ";"))
                            .lbl������.Caption = .lbl������.Caption & Split(strUnitName, ";")(intCount) & vbCrLf
                        Next
                    End If
                    .LblProductName = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒȫ��", "")
                    If Len(.LblProductName) > 10 Then
                        .LblProductName.FontSize = 15.75 '����
                    Else
                        .LblProductName.FontSize = 21.75 '����
                    End If
                    .lbl����֧���� = GetSetting("ZLSOFT", "ע����Ϣ", "����֧����", "")
                    .lbltag = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒϵ��", "")
                    
                    If Trim$(.lbl����֧����.Caption) = "" Then
                        .Label1.Visible = False
                        .lbl����֧����.Visible = False
                    Else
                        .Label1.Visible = True
                        .lbl����֧����.Visible = True
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
    
    Select Case gobjRegister.zlRegInfo("��Ȩ����")
        Case "1"
            '��ʽ
            SaveSetting "ZLSOFT", "ע����Ϣ", "Kind", ""
        Case "2"
            '����
            SaveSetting "ZLSOFT", "ע����Ϣ", "Kind", "����"
        Case "3"
            '����
            SaveSetting "ZLSOFT", "ע����Ϣ", "Kind", "����"
        Case Else
            '����
            MsgBox "��Ȩ���ʲ���ȷ���������˳���", vbInformation, gstrSysName
            Exit Function
    End Select
    
    gstrSysName = gobjRegister.zlRegInfo("��Ʒ����") & "���"
    SaveSetting "ZLSOFT", "ע����Ϣ", "��ʾ", gstrSysName
    SaveSetting "ZLSOFT", "ע����Ϣ", UCase("gstrSysName"), gstrSysName
    strTag = ""
    strTitle = gobjRegister.zlRegInfo("��Ʒ����")
    If strTitle <> "" Then
        If InStr(strTitle, "-") > 0 Then
            If Split(strTitle, "-")(1) = "Ultimate" Then
                strTag = "�콢��"
            ElseIf Split(strTitle, "-")(1) = "Professional" Then
                strTag = "רҵ��"
            End If
        End If
    End If
    strTitle = Split(strTitle, "-")(0)
    '���û�ע�������Ϣд��ע���,���´�����ʱ��ʾ
    SaveSetting "ZLSOFT", "ע����Ϣ", "��λ����", gobjRegister.zlRegInfo("��λ����", , -1)
    SaveSetting "ZLSOFT", "ע����Ϣ", "��Ʒȫ��", strTitle
    SaveSetting "ZLSOFT", "ע����Ϣ", "��Ʒ����", gobjRegister.zlRegInfo("��Ʒ����")
    SaveSetting "ZLSOFT", "ע����Ϣ", "����֧����", gobjRegister.zlRegInfo("����֧����", , -1)
    SaveSetting "ZLSOFT", "ע����Ϣ", "������", gobjRegister.zlRegInfo("��Ʒ������", , -1)
    SaveSetting "ZLSOFT", "ע����Ϣ", "WEB֧���̼���", gobjRegister.zlRegInfo("֧���̼���")
    SaveSetting "ZLSOFT", "ע����Ϣ", "WEB֧��EMAIL", gobjRegister.zlRegInfo("֧����MAIL")
    SaveSetting "ZLSOFT", "ע����Ϣ", "WEB֧��URL", gobjRegister.zlRegInfo("֧����URL")
    SaveSetting "ZLSOFT", "ע����Ϣ", "��Ʒϵ��", strTag
    SaveRegInfo = True
End Function

Public Function TestComponent() As Boolean
    '���û���κβ�����ʹ�ã��򷵻ؼ�
    TestComponent = False
    
    Dim strObjs As String, strSQL As String
    Dim resComponent As New ADODB.Recordset
    
    On Error GoTo errH
    '���������ܻس��ִ��󣬵��³����쳣����ͣ��
    If glngHelperMainType <> 0 Then TestComponent = True: Exit Function
    '--��ע����ȡ��Ȩ����--
    strObjs = GetSetting("ZLSOFT", "ע����Ϣ", "��������", "")

    If strObjs <> "" Then
        If InStr(strObjs, "'ZL9REPORT'") = 0 Then
            If CreateComponent("ZL9REPORT.ClsREPORT") Then
                strObjs = strObjs & ",'ZL9REPORT'"
                SaveSetting "ZLSOFT", "ע����Ϣ", "��������", strObjs
            End If
        End If
        TestComponent = True
        Exit Function
    End If
    '--������Ȩ��װ����--union��ȥ��
    strSQL = "Select ���� From (" & _
                "Select Upper(g.����) As ����" & vbNewLine & _
                "From zlPrograms G, (Select Distinct ϵͳ, ��� From zlRegFunc) R" & vbNewLine & _
                "Where g.��� = r.��� And Trunc(g.ϵͳ / 100) = r.ϵͳ" & vbNewLine & _
                " Union " & _
                " Select Upper(����) as ���� From zlPrograms Where ��� Between 10000 And 19999)"
    Set resComponent = OpenSQLRecord(strSQL, "")
    With resComponent
        Do While Not .EOF
            If CreateComponent(!���� & ".Cls" & Mid(!����, 4)) Then
                strObjs = strObjs & IIf(strObjs = "", "", ",") & "'" & !���� & "'"
            End If
            .MoveNext
        Loop
    End With
    If strObjs = "" Then Exit Function
    TestComponent = True
    SaveSetting "ZLSOFT", "ע����Ϣ", "��������", strObjs
    Exit Function
errH:
    If Not gobjComlib Is Nothing Then
        If gobjComlib.ErrCenter() = 1 Then
            Resume
        End If
    Else
        MsgBox "��Ȿ����װ��������" & Err.Description, vbInformation, gstrSysName
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
'���ܣ�����Valֻ�������ֿ�ͷʶ��ValEx�Ե�һ�����ֽ���ʶ��
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
    '����ע�Ჿ��(���ڵ�¼ʱ��ȡ���Ӷ���)
    Dim strObject As String
    
    '��Ȼ140��ȡ����Alone����,����ΪҪ����120�����ϵĵͰ汾�����Դ˴�����Alone����߼���֧
    On Error Resume Next
    If UCase(App.EXEName) = "ZLLOGINALONE" Then
        strObject = "zlRegisterAlone"
    Else
        strObject = "zlRegister"
    End If
    Set gobjRegister = CreateObject(strObject & ".clsRegister")
    If gobjRegister Is Nothing Then
        Err.Clear
        MsgBox "����" & strObject & ")��������ʧ��,�����ļ��Ƿ���ڲ�����ȷע�ᡣ", vbExclamation, gstrSysName
        Exit Function
    End If
    CreateRegister = Not gobjRegister Is Nothing
End Function

Public Function CheckPWDComplex(ByRef cnInput As ADODB.Connection, ByVal strChcekPWD As String, Optional ByRef strToolTip As String) As String
'���ܣ�������븴�Ӷ�
'������cnInput=���������
'          strChcekPWD=�ȴ���������
'          strToolTip=�����ʾ����
'���أ����ؼ�������龯��
    Dim strSQL As String, rsData As New ADODB.Recordset
    Dim blnHaveNum As Boolean, blnAlpha As Boolean, blnChar As Boolean
    Dim blnPwdLen As Boolean, intPwdMin As Integer, intPwdMax As Integer
    Dim blnComplex As Boolean, strOterChrs As String
    Dim lngLen As Long, i As Integer, intChr As Integer
    
    On Error GoTo errH
    strToolTip = ""
    strSQL = "Select ������,Nvl(����ֵ,ȱʡֵ) ����ֵ From zlOptions Where ������ in (20,21,22,23)"
    rsData.Open strSQL, cnInput
    blnPwdLen = False: intPwdMin = 0: intPwdMax = 0
    blnComplex = False: strOterChrs = ""
    Do While Not rsData.EOF
        Select Case rsData!������
            Case 20 '�Ƿ�������볤��
                blnPwdLen = Val(rsData!����ֵ & "") = 1
            Case 21 '���볤������
                intPwdMin = Val(rsData!����ֵ & "")
            Case 22 '���볤������
                intPwdMax = Val(rsData!����ֵ & "")
            Case 23 '�Ƿ�������븴�Ӷ�
                blnComplex = Val(rsData!����ֵ & "") = 1
        End Select
        rsData.MoveNext
    Loop
    '����������ʾ
    If blnPwdLen Then
        If intPwdMin = intPwdMax Then
            strToolTip = "�������Ϊ" & intPwdMax & " λ�ַ���"
        Else
            strToolTip = "�������Ϊ" & intPwdMin & "��" & intPwdMax & " λ�ַ���"
        End If
     End If
     If blnComplex Then
        If strToolTip <> "" Then
            strToolTip = strToolTip & vbNewLine & "���ٰ���һ�����֡�һ����ĸ��һ�������ַ���ɡ�"
        Else
            strToolTip = "������һ�����֡�һ����ĸ��һ�������ַ���ɡ�"
        End If
     End If
    '���ȼ��
    lngLen = ActualLen(strChcekPWD)
    If lngLen <> Len(strChcekPWD) Then
        CheckPWDComplex = "���������˫�ֽ��ַ������飡"
        Exit Function
    End If
    If blnPwdLen Then
        If Not (lngLen >= intPwdMin And lngLen <= intPwdMax) Then
            If intPwdMin = intPwdMax Then
                CheckPWDComplex = "�������Ϊ" & intPwdMax & " λ�ַ���"
                Exit Function
            Else
                CheckPWDComplex = "�������Ϊ" & intPwdMin & "��" & intPwdMax & " λ�ַ���"
                Exit Function
            End If
        End If
    End If
    For i = 1 To Len(strChcekPWD)
        intChr = Asc(UCase(Mid(strChcekPWD, i, 1)))
        If intChr >= 32 And intChr < 127 Then
            'Dim blnHaveNum As Boolean, blnAlpha As Boolean, blnChar As Boolean
            Select Case intChr
                Case 48 To 57 '����
                    blnHaveNum = True
                Case 65 To 90 '��ĸ
                    blnAlpha = True
                Case 32, 34, 47, 64  '�ո�,˫����,/,@
                    strOterChrs = strOterChrs & Chr(intChr)
                Case Is < 48, 58 To 64, 91 To 96, Is > 122
                    blnChar = True
            End Select
        Else
            strOterChrs = strOterChrs & Chr(intChr)
        End If
    Next
    If strOterChrs <> "" Then
        CheckPWDComplex = "���벻�����������ַ���" & strOterChrs
        Exit Function
    ElseIf Not (blnHaveNum And blnAlpha And blnChar) And blnComplex Then
        CheckPWDComplex = "����������һ�����֡�һ����ĸ��һ�������ַ���ɡ�"
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
    strSQL = "SELECT 1 FROM ZLTOOLS.ZLSYSTEMS WHERE ������=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, "��ȡ������", gstrDBUser)
    
    If Err.Number <> 0 Then
        blnHaveTools = False
        gclsLogin.IsSysOwner = False
        Err.Clear
    Else
        blnHaveTools = True
        gclsLogin.IsSysOwner = rsTmp.EOF
    End If

    strSQL = "SELECT 1 FROM SESSION_ROLES WHERE ROLE='DBA'"
    Set rsTmp = OpenSQLRecord(strSQL, "�ж�DBA")
    blnDBA = Not rsTmp.EOF

    If Not (blnDBA) And Not (blnHaveTools) Then
        CheckSysState = False
        MsgBox "�д��������������ߣ����Ƚ��д�����", vbExclamation, gstrSysName
        Exit Function
    End If
    
    If Not (blnDBA) And Not (gclsLogin.IsSysOwner) Then
        CheckSysState = False
        MsgBox "�������ݿ�DBA��Ӧ��ϵͳ�������ߣ�����ʹ�ñ����ߡ�", vbExclamation, gstrSysName
        Exit Function
    End If
    If Not blnHaveTools Then
        CheckSysState = False
        MsgBox "�д��������������ߣ����Ƚ��д�����", vbExclamation, gstrSysName
        Exit Function
    End If
    CheckSysState = True
End Function

Public Function GetMenuGroup(ByVal strCommand As String) As String
    Dim ArrCommand As Variant
    Dim i As Long
    '--����Ȩ�޲˵�--
    
    GetMenuGroup = "ȱʡ"
    
    ArrCommand = Split(strCommand, " ")
    If UBound(ArrCommand) = 0 Then
        '���������˵�����������/����ʾ���û�������ĸ�ʽ���磺zlhis/his��
        If InStr(1, ArrCommand(0), "/") = 0 And InStr(ArrCommand(0), ",") = 0 Then
            GetMenuGroup = ArrCommand(0)
        End If
    Else
        '�û��������뼰�˵����
        If UBound(ArrCommand) = 2 Then
            If InStr(ArrCommand(0), "=") <= 0 Then GetMenuGroup = ArrCommand(2)
            
        '����C:\APPSOFT\ZLHIS+.exe USER=�û��� PASS=���� SERVER=ʵ���� PROGRAM=ģ��� MENUGROUP=ȱʡ
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
'���ܣ�ͨ��Command����򿪴�����SQL�ļ�¼��
'������strSQL=�����а���������SQL���,������ʽΪ"[x]"
'             x>=1Ϊ�Զ��������,"[]"֮�䲻���пո�
'             ͬһ�������ɶദʹ��,�����Զ���ΪADO֧�ֵ�"?"����ʽ
'             ʵ��ʹ�õĲ����ſɲ�����,������Ĳ���ֵ��������(��SQL���ʱ��һ��Ҫ�õ��Ĳ���)
'      arrInput=���������Ĳ���ֵ,��������˳�����δ���,��������ȷ����
'               ��Ϊʹ�ð󶨱���,�Դ�"'"���ַ�����,����Ҫʹ��"''"��ʽ��
'      strTitle=����SQLTestʶ��ĵ��ô���/ģ�����
'      cnOracle=����ʹ�ù�������ʱ����
'���أ���¼����CursorLocation=adUseClient,LockType=adLockReadOnly,CursorType=adOpenStatic
'������
'SQL���Ϊ="Select ���� From ������Ϣ Where (����ID=[3] Or �����=[3] Or ���� Like [4]) And �Ա�=[5] And �Ǽ�ʱ�� Between [1] And [2] And ���� IN([6],[7])"
'���÷�ʽΪ��Set rsPati=OpenSQLRecord(strSQL, Me.Caption, CDate(Format(rsMove!ת������,"yyyy-MM-dd")),dtpʱ��.Value, lng����ID, "��%", "��", 20, 21)
    Dim cmdData As New ADODB.Command
    Dim strPar As String, arrPar As Variant
    Dim lngLeft As Long, lngRight As Long
    Dim strSeq As String, intMax As Integer, i As Integer
    Dim strLog As String, varValue As Variant
    Dim strSQLTmp As String, arrstr As Variant
    Dim strTmp As String, strSQLtmp1 As String
    
    '������ʹ���˶�̬�ڴ������û��ʹ��/*+ XXX*/����ʾ��ʱ�Զ�����
    strSQLTmp = Trim(UCase(strSQL))
    If Mid(Trim(Mid(strSQLTmp, 7)), 1, 2) <> "/*" And Mid(strSQLTmp, 1, 6) = "SELECT" Then
        arrstr = Split("F_STR2LIST,F_NUM2LIST,F_NUM2LIST2,F_STR2LIST2", ",")
        For i = 0 To UBound(arrstr)
            strSQLtmp1 = strSQLTmp
            Do While InStr(strSQLtmp1, arrstr(i)) > 0
                '�ж�ǰ���Ƿ�����IN �����򲻼�Rule
                '���ҵ����һ��SELECT
                strTmp = Mid(strSQLtmp1, 1, InStr(strSQLtmp1, arrstr(i)) - 1)
                strTmp = Replace(FromatSQL(Mid(strTmp, 1, InStrRev(strTmp, "SELECT") - 1)), " ", "")
                If Len(strTmp) > 1 Then strTmp = Mid(strTmp, Len(strTmp) - 2)  'ȡ����3���ַ�
                
                If strTmp = "IN(" Then '����in(select��������������ѭ�������Ƿ����û��ʹ������д����������̬�ڴ溯��
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
    
    '�����Զ���[x]����
    lngLeft = InStr(1, strSQL, "[")
    Do While lngLeft > 0
        lngRight = InStr(lngLeft + 1, strSQL, "]")
        If lngRight = 0 Then Exit Do
        '������������"[����]����"
        strSeq = Mid(strSQL, lngLeft + 1, lngRight - lngLeft - 1)
        If IsNumeric(strSeq) Then
            i = CInt(strSeq)
            strPar = strPar & "," & i
            If i > intMax Then intMax = i
        End If
        
        lngLeft = InStr(lngRight + 1, strSQL, "[")
    Loop
    
    If UBound(arrInput) + 1 < intMax Then
        Err.Raise 9527, strTitle, "SQL���󶨱�����ȫ��������Դ��" & strTitle
    End If

    '�滻Ϊ"?"����
    strLog = strSQL
    For i = 1 To intMax
        strSQL = Replace(strSQL, "[" & i & "]", "?")
        
        '��������SQL���ٵ����
        varValue = arrInput(i - 1)
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency", "Decimal" '����
            strLog = Replace(strLog, "[" & i & "]", varValue)
        Case "String" '�ַ�
            strLog = Replace(strLog, "[" & i & "]", "'" & Replace(varValue, "'", "''") & "'")
        Case "Date" '����
            strLog = Replace(strLog, "[" & i & "]", "To_Date('" & Format(varValue, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')")
        End Select
    Next
    
    '�����µĲ���
    lngLeft = 0: lngRight = 0
    arrPar = Split(Mid(strPar, 2), ",")
    For i = 0 To UBound(arrPar)
        varValue = arrInput((arrPar(i) - 1))
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency", "Decimal" '����
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adVarNumeric, adParamInput, 30, varValue)
        Case "String" '�ַ�
            intMax = LenB(StrConv(varValue, vbFromUnicode))
            If intMax <= 2000 Then
                intMax = IIf(intMax <= 200, 200, 2000)
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adVarChar, adParamInput, intMax, varValue)
            Else
                If intMax < 4000 Then intMax = 4000
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adLongVarChar, adParamInput, intMax, varValue)
            End If
        Case "Date" '����
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adDBTimeStamp, adParamInput, , varValue)
        Case "Variant()" '����
            '���ַ�ʽ������һЩIN�Ӿ��Union���
            '��ʾͬһ�������Ķ��ֵ,�����Ų�������������Ĳ����Ž���,��Ҫ��֤�����ֵ��������
            If arrPar(i) <> lngRight Then lngLeft = 0
            lngRight = arrPar(i)
            Select Case TypeName(varValue(lngLeft))
            Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '����
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adVarNumeric, adParamInput, 30, varValue(lngLeft))
                strLog = Replace(strLog, "[" & lngRight & "]", varValue(lngLeft), 1, 1)
            Case "String" '�ַ�
                intMax = LenB(StrConv(varValue(lngLeft), vbFromUnicode))
                If intMax <= 2000 Then
                    intMax = IIf(intMax <= 200, 200, 2000)
                    cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adVarChar, adParamInput, intMax, varValue(lngLeft))
                Else
                    If intMax < 4000 Then intMax = 4000
                    cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adLongVarChar, adParamInput, intMax, varValue(lngLeft))
                End If
                
                strLog = Replace(strLog, "[" & lngRight & "]", "'" & Replace(varValue(lngLeft), "'", "''") & "'", 1, 1)
            Case "Date" '����
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adDBTimeStamp, adParamInput, , varValue(lngLeft))
                strLog = Replace(strLog, "[" & lngRight & "]", "To_Date('" & Format(varValue(lngLeft), "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')", 1, 1)
            End Select
            lngLeft = lngLeft + 1 '�ò������������õ��ڼ���ֵ��
        End Select
    Next
'    If gblnSys = True Then
'        Set cmdData.ActiveConnection = gcnSysConn
'    Else
    Set cmdData.ActiveConnection = gcnOracle '���Ƚ���(���ִ��1000��Լ0.5x��)
'    End If
    cmdData.CommandText = strSQL
    
    Set OpenSQLRecordByArray = cmdData.Execute
    Set OpenSQLRecordByArray.ActiveConnection = Nothing
End Function

Public Sub ExecuteProcedure(strSQL As String, ByVal strFormCaption As String)
'���ܣ�ִ�й������,���Զ��Թ��̲������а󶨱�������
'������strSQL=�������,���ܴ�����,����"������(����1,����2,...)"��
'      cnOracle=����ʹ�ù�������ʱ����
'˵�������¼���������̲�����ʹ�ð󶨱���,�����ϵĵ��÷�����
'  1.���������Ǳ��ʽ,��ʱ�����޷�����󶨱������ͺ�ֵ,��"������(����1,100.12*0.15,...)"
'  2.�м�û�д�����ȷ�Ŀ�ѡ����,��ʱ�����޷�����󶨱������ͺ�ֵ,��"������(����1, , ,����3,...)"
'  3.��Ϊ�ù������Զ�����,����һ��ʹ�ð󶨱���,�Դ�"'"���ַ�����,��Ҫʹ��"''"��ʽ��
    If gblnTimer And Not gobjComlib Is Nothing Then
        Call gobjComlib.zldatabase.ExecuteProcedure(strSQL, strFormCaption)
    Else
        Dim cmdData As New ADODB.Command
        Dim strProc As String, strPar As String
        Dim blnStr As Boolean, intBra As Integer
        Dim strTemp As String, i As Long
        Dim intMax As Integer, datCur As Date
        
        If Right(Trim(strSQL), 1) = ")" Then
            'ִ�еĹ�����
            strTemp = Trim(strSQL)
            strProc = Trim(Left(strTemp, InStr(strTemp, "(") - 1))
            
            'ִ�й��̲���
            datCur = CDate(0)
            strTemp = Mid(strTemp, InStr(strTemp, "(") + 1)
            strTemp = Trim(Left(strTemp, Len(strTemp) - 1)) & ","
            For i = 1 To Len(strTemp)
                '�Ƿ����ַ����ڣ��Լ����ʽ��������
                If Mid(strTemp, i, 1) = "'" Then blnStr = Not blnStr
                If Not blnStr And Mid(strTemp, i, 1) = "(" Then intBra = intBra + 1
                If Not blnStr And Mid(strTemp, i, 1) = ")" Then intBra = intBra - 1
                
                If Mid(strTemp, i, 1) = "," And Not blnStr And intBra = 0 Then
                    strPar = Trim(strPar)
                    With cmdData
                        If IsNumeric(strPar) Then '����
                            .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adVarNumeric, adParamInput, 30, strPar)
                        ElseIf Left(strPar, 1) = "'" And Right(strPar, 1) = "'" Then '�ַ���
                            strPar = Mid(strPar, 2, Len(strPar) - 2)
                            
                            'Oracle���ӷ�����:'ABCD'||CHR(13)||'XXXX'||CHR(39)||'1234'
                            If InStr(Replace(strPar, " ", ""), "'||") > 0 Then GoTo NoneVarLine
                            
                            '˫"''"�İ󶨱�������
                            If InStr(strPar, "''") > 0 Then strPar = Replace(strPar, "''", "'")
                            
                            '���Ӳ�������LOBʱ������ð󶨱���ת��ΪRAWʱ����2000���ַ�Ҫ��adLongVarChar
                            intMax = LenB(StrConv(strPar, vbFromUnicode))
                            If intMax <= 2000 Then
                                intMax = IIf(intMax <= 200, 200, 2000)
                                .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adVarChar, adParamInput, intMax, strPar)
                            Else
                                If intMax < 4000 Then intMax = 4000
                                .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adLongVarChar, adParamInput, intMax, strPar)
                            End If
                        ElseIf UCase(strPar) Like "TO_DATE('*','*')" Then '����
                            strPar = Split(strPar, "(")(1)
                            strPar = Trim(Split(strPar, ",")(0))
                            strPar = Mid(strPar, 2, Len(strPar) - 2)
                            If strPar = "" Then
                                'NULLֵ�������ִ���ɼ�����������
                                .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adVarNumeric, adParamInput, , Null)
                            Else
                                If Not IsDate(strPar) Then GoTo NoneVarLine
                                .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adDBTimeStamp, adParamInput, , CDate(strPar))
                            End If
                        ElseIf UCase(strPar) = "SYSDATE" Then '����
                            If datCur = CDate(0) Then datCur = Currentdate
                            .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adDBTimeStamp, adParamInput, , datCur)
                        ElseIf UCase(strPar) = "NULL" Then 'NULLֵ�����ַ�����ɼ�����������
                            .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adVarChar, adParamInput, 200, Null)
                        ElseIf strPar = "" Then '��ѡ��������NULL������ܸı���ȱʡֵ:��˿�ѡ��������д���м�
                            GoTo NoneVarLine
                        Else '�������������ӵı��ʽ���޷�����
                            GoTo NoneVarLine
                        End If
                    End With
                    
                    strPar = ""
                Else
                    strPar = strPar & Mid(strTemp, i, 1)
                End If
            Next
            
            '����Ա���ù���ʱ��д����
            If blnStr Or intBra <> 0 Then
                Err.Raise -2147483645, , "���� Oracle ����""" & strProc & """ʱ�����Ż�������д��ƥ�䡣ԭʼ������£�" & vbCrLf & vbCrLf & strSQL
                Exit Sub
            End If
            
            '����?��
            strTemp = ""
            For i = 1 To cmdData.Parameters.Count
                strTemp = strTemp & ",?"
            Next
            strProc = "Call " & strProc & "(" & Mid(strTemp, 2) & ")"
            Set cmdData.ActiveConnection = gcnOracle '���Ƚ���
            cmdData.CommandType = adCmdText
            cmdData.CommandText = strProc
            
            Call cmdData.Execute
        Else
            GoTo NoneVarLine
        End If
        Exit Sub
NoneVarLine:
        '˵����Ϊ�˼��������ӷ�ʽ
        '1.��������adCmdStoredProc��ʽ��8i����������
        '2.�����������ʹ��{},��ʹ����û�в���ҲҪ��()
        strSQL = "Call " & strSQL
        If InStr(strSQL, "(") = 0 Then strSQL = strSQL & "()"
        gcnOracle.Execute strSQL, , adCmdText
End If
End Sub

Public Function IP(Optional ByVal strErr As String) As String
    '******************************************************************************************************************
    '����:ͨ��oracle��ȡ�ļ������IP��ַ
    '���:strDefaultIp_Address-ȱʡIP��ַ
    '����:
    '����:����IP��ַ
    '******************************************************************************************************************
    Dim rsTmp As ADODB.Recordset
    Dim strIp_Address As String
    Dim strSQL As String
        
    On Error GoTo Errhand
    
    strSQL = "Select Sys_Context('USERENV', 'IP_ADDRESS') as Ip_Address From Dual"
    Set rsTmp = OpenSQLRecord(strSQL, "��ȡIP��ַ")
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
    '���ܣ���ȡ�������ϵ�ǰ����
    '������
    '���أ�����Oracle���ڸ�ʽ�����⣬����
    '-------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    Err = 0
    On Error GoTo errH
    With rsTemp
        .CursorLocation = adUseClient
    End With
    Set rsTemp = OpenSQLRecord("SELECT SYSDATE FROM DUAL", "��ȡ������ʱ��")
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
    '����Ʒ���Ʋ��ԣ�������������
    If UCase(GetFileDesInfo(strAppPath, "ProductName")) <> "ZLSOFT EXTENSION INSTALL" Then
        Exit Sub
    End If
    lngErr = Shell(strAppPath & " ORAOLEDB -REGSVR -S", vbHide)
End Sub


'���ܣ���ȡ��ǰ���̵�·��
Public Function GetCurExePath() As String
    Dim uProcess        As PROCESSENTRY32, uMdlInfor    As MODULEENTRY32
    Dim lngMdlProcess   As Long, strExeName             As String, lngSnapShot  As Long, strModelPath     As String, strModelName As String
    Dim lngProceess     As Long
    
    On Error GoTo errH
    '�������̿���
    lngProceess = GetCurrentProcessId
    lngSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    If lngSnapShot > 0 Then
        uProcess.lSize = Len(uProcess)
        If Process32First(lngSnapShot, uProcess) Then
            Do
                If uProcess.lProcessId = lngProceess Then
                    '��ý��̵ı�ʶ��
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

'���ܣ�ģ��VB APP.PrevInstance,��������DLL�л��ж�ʧЧ��һֱΪFalse
'˵����1��������·����ͬʱ���������̵�APP.PrevInstance�޹���������EXE�ļ���ͬ��
'       2���ú�����APP.PrevInstance��һ������1��APP.PrevInstance�ǹ̶��ģ����̴�ʱ�͹̶������ܹر���������ͬ���̣��Ծɲ��ᷢ���仯��
'                                              2)�ú����Ƕ�̬��ѯ���������嵥��û�е�ǰEXE·���ļ��Ľ��̣�����FALSE,�������TRUE,�ͽ��̵������йء�
Public Function AppPrevInstance() As Boolean
    Dim uProcess        As PROCESSENTRY32, uMdlInfor    As MODULEENTRY32
    Dim lngMdlProcess   As Long, strExeName             As String, lngSnapShot  As Long, strModelPath     As String, strModelName As String
    Dim lngProceess     As Long
    Dim strCurAppPath   As String
    Dim blnFind         As Boolean
    
    On Error GoTo errH
    '�������̿���
    strCurAppPath = GetCurExePath()
    lngProceess = GetCurrentProcessId
    blnFind = False
    lngSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    If lngSnapShot > 0 Then
        uProcess.lSize = Len(uProcess)
        If Process32First(lngSnapShot, uProcess) Then
            Do
                If uProcess.lProcessId <> lngProceess Then
                    '��ý��̵ı�ʶ��
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
    '���ܣ���ȡ��������
    '������
    '˵����
    '******************************************************************************************************************
    Dim strComputer As String * 256
    Call GetComputerName(strComputer, 255)
    ComputerName = strComputer
    ComputerName = Trim(Replace(ComputerName, Chr(0), ""))
End Function

Public Function NVL(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'���ܣ��൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
    NVL = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Function IsDesinMode() As Boolean
'���ܣ� ȷ����ǰģʽΪ���ģʽ
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
    '���ܣ�ͨ��API��ȡ��ʱIP
    
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
'���ܣ�ȥ���ַ�����\0�Ժ���ַ�,�������ù���,���Ե�������clsstring
    Dim lngPos As Long
    
    lngPos = InStr(strInput, Chr(0))
    If lngPos > 0 Then
        TruncZeroInside = Mid(strInput, 1, lngPos - 1)
    Else
        TruncZeroInside = strInput
    End If
End Function
'======================================================================================================================
'����           Sm4EncryptEcb           SM4����
'����ֵ         String                  ���ܺ��ֵ,��ʽ��ZLSV+�汾��+:+���ܺ���ַ���
'����б�:
'������         ����                    ˵��
'strInput       String                  Ҫ���ܵ��ַ���
'strKey         String(Optional)        ������Կ��32λ��16�����ַ���������ͨ��HexStringToByte���أ�
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
'����           Sm4DecryptEcb           SM4����
'����ֵ         String                  ���ܺ��ֵ
'����б�:
'������         ����                    ˵��
'strInput       String                  Ҫ���ܵ��ַ��������ַ�����Sm4EncryptEcb���ɵĽ����
'strKey         String(Optional)        ������ԿҲ���ǽ�����Կ��32λ��16�����ַ���������ͨ��HexStringToByte���أ�
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
        '��ǰ�ͻ��˵�ZLSM4��֧�ָð汾�ļ����ַ������ܣ��Ծɽ��ܣ���Ϊһ����˵���ܽ��ܳ���ͬ���ַ���
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
'����           Sm4EncryptCbc           SM4�������
'����ֵ         String                  ���ܺ��ֵ
'����б�:
'������         ����                    ˵��
'strInput       String                  Ҫ���ܵ��ַ���
'strKey         String(Optional)        ������Կ��32λ��16�����ַ���������ͨ��HexStringToByte���أ�
'strIv          String(Optional)        ���������Կ��32λ��16�����ַ���������ͨ��HexStringToByte���أ�
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
'����           Sm4EncryptCbc           SM4������ܶ�Ӧ�Ľ��ܹ���
'����ֵ         String                  ���ܺ��ֵ
'����б�:
'������         ����                    ˵��
'strInput       String                  �Ѿ����ܵ��ַ���
'strKey         String(Optional)        ������ԿҲ���Ǽ�����Կ��32λ��16�����ַ���������ͨ��HexStringToByte���أ�
'strIv          String(Optional)        ���������ԿҲ���Ƿ��������Կ��32λ��16�����ַ���������ͨ��HexStringToByte���أ�
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
        '��ǰ�ͻ��˵�ZLSM4��֧�ָð汾�ļ����ַ������ܣ��Ծɽ��ܣ���Ϊһ����˵���ܽ��ܳ���ͬ���ַ���
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
'����           Sm3                     �����ַ����Ĺ�ϣֵ����������ַ����ı䶯��
'����ֵ         String(32)              �ַ����Ĺ�ϣֵ
'����б�:
'������         ����                    ˵��
'strInput       String                  �ַ�������
'======================================================================================================================
Public Function Sm3(ByRef strInput As String) As String
    Dim arrInput()  As Byte
    Dim lngLength   As Long
    Dim arrOut(31)  As Byte

    '�Ƚ��ַ����� Unicode ת��ϵͳ��ȱʡ��ҳ
    arrInput = StrConv(strInput, vbFromUnicode)
    lngLength = UBound(arrInput) + 1
    
    Call sm3_hash(arrInput(0), lngLength, arrOut(0))
    '������ֵת��Ϊ16�����ַ���
    Sm3 = ByteToHexString(arrOut)
End Function
'======================================================================================================================
'����           Sm3_File                �����ļ��Ĺ�ϣֵ��������� �ļ����ݵı䶯��
'����ֵ         String(32)              �ļ��Ĺ�ϣֵ
'����б�:
'������         ����                    ˵��
'strFile        String                  �ļ�·��
'======================================================================================================================
Public Function Sm3_File(ByRef strFile As String) As String
    Dim arrInput()  As Byte
    Dim lngLength   As Long
    Dim arrOut(31)  As Byte
    Dim lngReturn As Long

    '�Ƚ��ַ����� Unicode ת��ϵͳ��ȱʡ��ҳ
    arrInput = StrConv(strFile, vbFromUnicode)
    '����APIû�д��ݳ��ȣ���������ַ���Chr(0)
    lngLength = UBound(arrInput) + 1
    ReDim Preserve arrInput(lngLength)
    '������
    lngReturn = sm3_file_hash(arrInput(0), arrOut(0))
    '�ж��Ƿ�ɹ�����
    If lngReturn = 0 Then
        '������ֵת��Ϊ16�����ַ���
        Sm3_File = ByteToHexString(arrOut)
    ElseIf lngReturn = 1 Then
        Sm3_File = "ERROR:�ļ���ʧ��"
    ElseIf lngReturn = 2 Then
        Sm3_File = "ERROR:�ļ���ȡʧ��"
    End If
End Function
'======================================================================================================================
'����           sm3_hmac                ������һ����Կ�Դ������Ϣ������ϢժҪ
'����ֵ         String(32)              ��Կ������Ϣ�����ɵ���ϢժҪ
'����б�:
'������         ����                    ˵��
'strKey         String                  ��Կ
'strMsg         String                  ��Ϣ����
'======================================================================================================================
Public Function sm3_hmac(ByRef strKey As String, ByVal strMsg As String) As String
    Dim arrInput()  As Byte
    Dim lngLength   As Long
    Dim arrOut(31)  As Byte
    Dim arrKey()    As Byte
    Dim lngKeyLen   As Long
    
    '�Ƚ��ַ����� Unicode ת��ϵͳ��ȱʡ��ҳ
    arrInput = StrConv(strMsg, vbFromUnicode)
    lngLength = UBound(arrInput) + 1
    '�Ƚ��ַ����� Unicode ת��ϵͳ��ȱʡ��ҳ
    arrKey = StrConv(strKey, vbFromUnicode)
    lngKeyLen = UBound(arrKey) + 1
    Call sm3_hmac_hash(arrKey(0), lngKeyLen, arrInput(0), lngLength, arrOut(0))
    '������ֵת��Ϊ16�����ַ���
    sm3_hmac = ByteToHexString(arrOut)
End Function
'======================================================================================================================
'����           sm_version              ��ȡZLSM4�İ汾��
'����ֵ         Long                    ZLSM4�İ汾��
'����б�:
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
'����           ByteToHexString         ���ֽ���ת��Ϊ16�����ַ���
'����ֵ         String                  �ֽ���ת����16�����ַ���
'����б�:
'������         ����                    ˵��
'bytInpu        Byte(��                 �ֽ�����
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
'����           ByteToHexString         ��16�����ַ���ת��Ϊ�ֽ���
'����ֵ         Byte()                  16�����ַ���ת�����ֽ���
'����б�:
'������         ����                    ˵��
'bstrInput      String                  16�����ַ���
'lngRetBytLen   Long(Optional)          ָ�����ص��ֽ���ĳ���,0-��ԭʼ���ȷ��أ�<>0����ָ���ĳ��ȣ����㲹�루��0�������˽�ȡ
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
'����           BytePadding             ��ָ���ַ�������16�ֽڲ��룬
'����ֵ         Byte()                  �������ַ����ֽ���
'����б�:
'������         ����                    ˵��
'strInput       String                  �ַ���
'lngVersion     Long(Optional,2)        �ַ�������İ汾��ZLSM4.DLL�İ汾���Լ������㷨ǰ׺�еİ汾����1-�ո��룬>1:Chr(0)����
'lngPaddingNum  Long(Optional,16)        ������ֽ�����ȱʡ����16���Ʋ���
'======================================================================================================================
Public Function BytePadding(ByVal strInput As String, Optional ByVal lngVersion As Long = 2, Optional ByVal lngPaddingNum As Long = 16) As Byte()
    Dim arrReturn()     As Byte
    Dim lngLenBef       As Long
    Dim i               As Long
    Dim lngLenAft       As Long
    
    '�Ƚ��ַ����� Unicode ת��ϵͳ��ȱʡ��ҳ
    arrReturn = StrConv(strInput, vbFromUnicode)
    lngLenBef = UBound(arrReturn) + 1
    '�жϵõ�������ĳ��ȣ�������16�����������򲹿ո��:Chr(0)
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

Public Sub ApplyOEM_Picture(objPicture As Object, ByVal str���� As String, Optional ByVal strProductName As String)
'��Ը���ͼ��Ӧ��OEM����
    Dim strOEM As String
    Dim blnCorp As Boolean
    On Error Resume Next
    
    If strProductName = "" Then
        strProductName = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒ����", "")
    End If

    If strProductName <> "����" And strProductName <> "-" Then
        '����״̬��ͼ���OEM����
        If Right(str����, 1) = "B" Then
            '��ʾ��ƷͼƬ
            blnCorp = False
            str���� = Mid(str����, 1, Len(str����) - 1)
        Else
            '��ʾ��˾�ձ�
            blnCorp = True
        End If
        
        strOEM = mGetOEM(strProductName, blnCorp)
        If str���� = "Picture" Then
            Set objPicture.Picture = LoadCustomPicture(strOEM)
        ElseIf str���� = "Icon" Then
            Set objPicture.Icon = LoadCustomPicture(strOEM)
        End If
        
        If Err <> 0 Then
            Err.Clear
        End If
    
    End If
End Sub

Private Function mGetOEM(ByVal strAsk As String, Optional ByVal blnCorp As Boolean = True) As String
    '-------------------------------------------------------------
    '���ܣ�����ÿ�����ߵ�ASCII��
    '������
    '���أ�
    '-------------------------------------------------------------
    Dim intBit As Integer
    Dim strCode As String
    
    'OEMͼƬ���������� ��һ��ָ��˾�ձ꣬��һ���ǲ�Ʒ��ʶ
    strCode = IIf(blnCorp = True, "OEM_", "PIC_")
    For intBit = 1 To Len(strAsk)
        'ȡÿ���ֵ�ASCII��
        strCode = strCode & Hex(Asc(Mid(strAsk, intBit, 1)))
    Next
    mGetOEM = strCode
End Function

Public Function ActualLen(ByVal strAsk As String) As Long
    '--------------------------------------------------------------
    '���ܣ���ȡָ���ַ�����ʵ�ʳ��ȣ������ж�ʵ�ʰ���˫�ֽ��ַ�����
    '       ʵ�����ݴ洢����
    '������
    '       strAsk
    '���أ�
    '-------------------------------------------------------------
    ActualLen = LenB(StrConv(strAsk, vbFromUnicode))
End Function

Public Function FromatSQL(ByVal strText As String, Optional ByVal blnCrlf As Boolean) As String
'���ܣ�ȥ��TAB�ַ������߿ո񣬻س������ֻ�ɵ��ո�ָ���
'������strText=�����ַ�
'         blnCrlf=�Ƿ�ȥ�����з�
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
'����:����Դ�ļ��е�ָ����Դ���ɴ����ļ�
'����:ID=��Դ��,strExt=Ҫ�����ļ�����չ��(��BMP)
'����:�����ļ���
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
'����:���������뷨����ر����뷨
'������strImeName-��ָ�������뷨��û��ָ��ʱ��ϵͳѡ�����õ�ȱʡ���뷨
    Dim arrIme(99) As Long, lngCount As Long, strName As String * 255
    Dim strIme As String
    
    If strImeName = "���Զ�����" Then OpenIme = True: Exit Function
    '�û�û�������ã��Ͳ�����
    If blnOpen Then
        If strImeName <> "" Then
            strIme = Trim(strImeName)
        End If
        If strIme = "" Then Exit Function                'Ҫ������뷨��������û������
    End If
    
    lngCount = GetKeyboardLayoutList(UBound(arrIme) + 1, arrIme(0))

    Do
        lngCount = lngCount - 1
        If ImmIsIME(arrIme(lngCount)) = 1 Then
            If blnOpen = True Then
                '��Ҫ�����뷨�������ж��Ƿ�ָ�����뷨
                ImmGetDescription arrIme(lngCount), strName, Len(strName)
                If InStr(1, Mid(strName, 1, InStr(1, strName, Chr(0)) - 1), strIme) > 0 Then
                    If ActivateKeyboardLayout(arrIme(lngCount), 0) <> 0 Then
                        OpenIme = True
                        Exit Function
                    End If
                End If
            End If
        ElseIf blnOpen = False Then
            '�����������뷨��������Ӧ�˹ر����뷨������
            If ActivateKeyboardLayout(arrIme(lngCount), 0) <> 0 Then OpenIme = True: Exit Function
        End If
    Loop Until lngCount = 0
    
    If blnOpen = False Then
        '����windows Vistaϵͳ��Ӣ�����뷨��ImmIsIME���Գ���1�����뷨,���,��Ҫ��������.
        '���˺�:2008/09/03
        If ActivateKeyboardLayout(arrIme(0), 0) <> 0 Then OpenIme = True: Exit Function
    End If
End Function

Public Function GetAllSubKey(ByVal strKey As String) As Variant
'����:��ȡĳ�����������
'���أ�=��������
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
'���ܣ����ݼ�λ��ȡ����ֵ���ӽ�,�Լ�ֵ����
'������strKey=ע����λ���硰HKEY_CURRENT_USER\Printers\DevModePerUser"
'          strValueName=������
'���Σ�
'          hRootKey=����
'          strSubKey=�ӽ�
'          lngType=������
'���أ��Ƿ��ȡ�ɹ�
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
        'ʹ�ò�ѯ��ʽ�򿪣����м������Ͳ�ѯ
        lngReturn = RegOpenKeyEx(hRootKey, strSubKey, 0, KEY_QUERY_VaLUE, hKey)
        If lngReturn <> ERROR_SUCCESS Then
            Exit Function
        End If
        If strValueName <> "" Then
            lngReturn = RegQueryValueEx_ValueType(hKey, strValueName, ByVal 0&, lngType, ByVal strName, Len(strName))
            'SetRegKey����������ص�����Ϊ�ܴ��������ֵ���̶�,�������Ϊ0�����ݴ������������ж�
            If lngReturn = ERROR_BADKEY Then
                If lngType < REG_NONE Or lngType > REG_MULTI_SZ Then lngType = REG_NONE
            End If
            '�����ֶγ��������Ȳ��������Գ����˳�
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
'���ܣ���ȡע�����ָ��λ�õ�ֵ
'������strKey=ע����λ���硰HKEY_CURRENT_USER\Printers\DevModePerUser"
'          strValueName=������
'          strValue=����ֵ
'          strValueType=�������ͣ�Ĭ��Ϊ�ַ���
'           blnOneString = ��REG_EXPAND_SZ��REG_MULTI_SZ,REG_BINARY��Ч��-  True �������ص�һ�ַ������Ҳ����κδ���ֻȥ���ַ���β��
'���أ��Ƿ��ȡ�ɹ�
'˵������ǰֻ��REG_SZ, REG_EXPAND_SZ, REG_MULTI_SZ��REG_DWORD��REG_BINARYʵ���˶�ȡ��û�в�ѯ�������Զ����Ҽ���
    Dim hRootKey As REGRoot, strSubKey As String
    Dim lngReturn As Long
    Dim lngKey As Long, ruType As REGValueType
    Dim lngLength As Long, strBufVar() As String, lngBuf As Long, bytBuf() As Byte, strBuf As String
    Dim i As Long, strReturn As String, strTmp As String
    '������Ч��ע����λ,��ȡ��������
    If Not GetKeyValueInfo(strKey, strValueName, hRootKey, strSubKey, ruType) Then Exit Function
    '�򿪱���
    lngReturn = RegOpenKeyEx(hRootKey, strSubKey, 0, KEY_QUERY_VaLUE, lngKey)
    If lngReturn <> ERROR_SUCCESS Then
        Exit Function
    End If
    On Error GoTo errH
    Select Case ruType
        Case REG_SZ, REG_EXPAND_SZ, REG_MULTI_SZ '�ַ������Ͷ�ȡ
'            lngReturn = RegQueryValueEx(lngKey, strValueName, 0, ruType, 0, lngLength)
'            If lngReturn <> ERROR_SUCCESS Then Err.Clear '���ܳ��������������
            lngLength = 1024: strBuf = Space(lngLength)
            lngReturn = RegQueryValueEx_String(lngKey, strValueName, 0, ruType, strBuf, lngLength)
            If lngReturn <> ERROR_SUCCESS Then: RegCloseKey (lngKey): Exit Function
            Select Case ruType
                Case REG_SZ
                    varValue = TruncZero(strBuf)
                Case REG_EXPAND_SZ ' ���价���ַ�������ѯ���������ͷ��ض���ֵ
                    If Not blnOneString Then
                        varValue = TruncZero(ExpandEnvStr(TruncZero(strBuf)))
                    Else
                        varValue = TruncZero(strBuf)
                    End If
                Case REG_MULTI_SZ ' �����ַ���
                    If Not blnOneString Then
                        If Len(strBuf) <> 0 Then ' �������Ƿǿ��ַ��������Էָ
                            strBufVar = Split(Left$(strBuf, Len(strBuf) - 1), Chr$(0))
                        Else ' ���ǿ��ַ�����Ҫ����S(0) ���������
                            ReDim strBufVar(0) As String
                        End If
                        ' ��������ֵ������һ���ַ������飿��
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
            ' �����ַ�����ע�⣺Ҫ���ֽ��������ת����
            If blnOneString Then
                'ѭ�����ݣ����ֽ�ת��Ϊ16�����ַ���
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
'���ܣ����ַ����еĻ��������滻Ϊ����ֵ
'         strInput=���������������ַ���
'���أ���ʵ�ʵ�ֵ�滻�ַ����еĻ�����������ַ���
    '// �磺 %PATH% �򷵻� "c:\;c:\windows;"
    Dim lngLen As Long, strBuf As String, strOld As String
    strOld = strInput & "  " ' ��֪ΪʲôҪ�������ַ������򷵻�ֵ������������ַ���
    strBuf = "" '// ��֧��Windows 95
    '// get the length
    lngLen = ExpandEnvironmentStrings(strOld, strBuf, lngLen)
    '// չ���ַ���
    strBuf = String$(lngLen - 1, Chr$(0))
    lngLen = ExpandEnvironmentStrings(strOld, strBuf, LenB(strBuf))
    '// ���ػ�������
    ExpandEnvStr = TruncZero(strBuf)
End Function

Public Sub OSWait(ByVal lngMilliseconds As Long)
'���ܣ�ִ�й���һ��ʱ��
'lngMilliseconds=��������1000����=1��
    Call Sleep(lngMilliseconds)
End Sub

Public Function Is64bit() As Boolean
    '******************************************************************************************************************
    '���ܣ��Ƿ���64λϵͳ
    '���أ�
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
    '����:��ȡ����������ض���
    '����:��ȡ�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-05-15 15:34:05
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
'���ܣ�ģ��Oracle��Decode����
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
'���ܣ�ȥ���ַ�����\0�Ժ���ַ�
    Dim lngPos As Long
    
    lngPos = InStr(strInput, Chr(0))
    If lngPos > 0 Then
        TruncZero = Mid(strInput, 1, lngPos - 1)
    Else
        TruncZero = strInput
    End If
End Function

Public Sub SetWindowsInTaskBar(ByVal lnghwnd As Long, ByVal blnShow As Boolean)
'���ܣ����ô����Ƿ�������������ʾ
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
'������strParNO-�����ţ�����ж�������Զ��Ÿ���
    Dim strSQL As String
    
    On Error GoTo errH
    If InStr(strParNO, ",") > 0 Then
        strSQL = "Select ������,������,Nvl(����ֵ,ȱʡֵ) as ����ֵ  From zlOptions Where ������ in(Select Column_value From Table(f_num2list([1])))"
        Set GetZLOptions = OpenSQLRecord(strSQL, "��ȡzlOptions����", strParNO)
    Else
        strSQL = "Select ������,������,Nvl(����ֵ,ȱʡֵ) as ����ֵ From zlOptions Where ������=[1]"
        Set GetZLOptions = OpenSQLRecord(strSQL, "��ȡzlOptions����", Val(strParNO))
    End If
    
    Exit Function
errH:
    MsgBox Err.Description & vbCrLf & strSQL, vbExclamation, "��ȡ�����߲���"
    Set GetZLOptions = New ADODB.Recordset
End Function
