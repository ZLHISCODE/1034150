VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls7zZip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mclsZip     As Object       'ZL7z��ѹ����
Private mclsUnZip   As Object       'ZL7z�Ľ�ѹ��
Private mblnZL7z    As Boolean      '�Ƿ�������ʹ��Zl7z
Private mstr7zExe   As String       '7z.exe��·��

Public Function Init7zZip(Optional ByVal str7zPath As String) As Boolean
'���ܣ�����7z��ع��ܵĳ�ʼ��
    On Error Resume Next
    Set mclsZip = CreateObject("zl7z.Zip")
    If err.Number <> 0 Then
       err.Clear
    End If
    Set mclsUnZip = CreateObject("zl7z.UnZip")
    If err.Number <> 0 Then
       err.Clear
    End If
    mblnZL7z = Not (mclsZip Is Nothing Or mclsUnZip Is Nothing)
    If mblnZL7z Then
        mblnZL7z = Val(GetSetting("ZLSOFT", "����ģ��\�Զ�����", "7ZEXE����", "0")) = 0
    End If
    If str7zPath = "" Or Not gobjFSO.FileExists(mstr7zExe) Then
        mstr7zExe = Get7zExePath
    Else
        mstr7zExe = str7zPath
    End If
    Init7zZip = mstr7zExe <> "" Or mblnZL7z
End Function

Public Function ZipFile(ByVal strFile As String, ByVal strZipName As String, Optional ByVal intRate As Integer = 5, Optional ByRef strErr As String) As Boolean
'���ܣ�����ѹ��
'������
'       strZipName=������ļ�·��������
'       strFile=ԭʼ�ļ�
'       intRate=ѹ���ȼ���ѹ��ʹ�á�
    If mblnZL7z Then
        Call mclsZip.ZipSingle(strFile, gobjFSO.GetFileName(strFile), strZipName, intRate)
        ZipFile = gobjFSO.FileExists(strZipName)
    Else
        strErr = ""
        ZipFile = DeCompression(strZipName, strFile, intRate, True, strErr)
        ZipFile = ZipFile And strErr = "" And gobjFSO.FileExists(strZipName)
    End If
End Function

Public Function UnZipFile(ByVal strZipName As String, ByVal strFile As String, Optional ByVal blnDeleteZipFile As Boolean = True, Optional ByRef strErr As String, Optional blnFolder As Boolean) As Boolean
'���ܣ�����ѹ��
'������
'       strZipName=ѹ���ļ���
'       strFile=��ѹ�������
    If mblnZL7z Then
        mclsUnZip.Open strZipName
        mclsUnZip.unzip gobjFSO.GetParentFolderName(strFile)
        UnZipFile = (gobjFSO.FileExists(strFile) Or gobjFSO.FolderExists(strFile))
    Else
        strErr = ""
        UnZipFile = DeCompression(strFile, strZipName, , , strErr, blnFolder)
        If blnFolder Then
            UnZipFile = UnZipFile And strErr = "" And gobjFSO.FolderExists(strFile)
        Else
            UnZipFile = UnZipFile And strErr = "" And gobjFSO.FileExists(strFile)
        End If
    End If
    If blnDeleteZipFile Then
        On Error Resume Next
        If FileSystem.GetAttr(strZipName) <> vbNormal Then
             Call FileSystem.SetAttr(strZipName, vbNormal)
        End If
        'ɾ��ԭʼ�ļ�
        Call gobjFSO.DeleteFile(strZipName, True)
        If err.Number <> 0 Then err.Clear
    End If
End Function

'===========================================================================
'                2��˽�з���
'===========================================================================
Private Function DeCompression(ByVal strDesFile As String, ByVal strSourceFile As String, Optional ByVal intRate As Integer, Optional ByVal blnCompression As Boolean, Optional ByRef strErr As String, Optional blnFolder As Boolean) As Boolean
'���ܣ�����ѹ����ѹ(��ǰ��֧�ֵ��ļ���
'������
'       strDesFile=������ļ�·��������
'       strSourceFile=ԭʼ�ļ�
'       intRate=ѹ���ȼ���ѹ��ʹ�á�
'                   ѹ���ȼ� ѹ���㷨 �ֵ��С �����ֽ� ƥ���� ������ ����
'                   0           Copy    ��ѹ��
'                   1           LZMA    64KB     32       HC4   BCJ   ���ѹ��
'                   3           LZMA    1MB      32       HC4   BCJ   ����ѹ��
'                   5           LZMA    16MB     32       BT4   BCJ   ����ѹ��(Ĭ�ϵȼ���
'                   7           LZMA    32MB     64       BT4   BCJ   ���ѹ��
'                   9           LZMA    64MB     64       BT4   BCJ2  ����ѹ��
'       blnCompression=True-ѹ����False-��ѹ
'���أ��Ƿ�ɹ�
'˵������ѹ���ļ�������,��ɾ��ѹ��ԭʼ�ļ�
    Dim strCommand As String, strReturn As String
    '��ȡ����7Z�ļ�·������ֱ���˳�
    If mstr7zExe = "" Then
        strErr = "7Z.EXE��ѹ���򲻴���"
        Exit Function
    End If
    If Not gobjFSO.FileExists(strSourceFile) Then
        strErr = "Դ�ļ�" & strSourceFile & "������"
        Exit Function
    End If
    If gobjFSO.FileExists(strDesFile) Then
        On Error Resume Next
        'ɾ�����ڵ�Ŀ���ļ�
        If FileSystem.GetAttr(strDesFile) <> vbNormal Then
             Call FileSystem.SetAttr(strDesFile, vbNormal)
        End If
        Call gobjFSO.DeleteFile(strDesFile, True)
        If err.Number <> 0 Then err.Clear
    End If
    On Error GoTo ErrH
    If blnCompression Then
        '-m �̶������ַ� x=���õȼ� mt������رն��߳�ѹ��ģʽ
        strCommand = """" & mstr7zExe & """  a -y """ & strDesFile & """ """ & strSourceFile & """ -mx=" & intRate & " -mmt"
    Else
        '-o �̶������ַ�
        If blnFolder Then
            strCommand = """" & mstr7zExe & """  e -y """ & strSourceFile & """ -o""" & strDesFile & """"
        Else
            strCommand = """" & mstr7zExe & """  x -y """ & strSourceFile & """ -o""" & gobjFSO.GetParentFolderName(strDesFile) & """"
        End If
    End If
    strReturn = RunCommand(strCommand, strErr, , 5000)
    If strErr = "" And strReturn <> "" Then strErr = strReturn
    If blnFolder Then
        DeCompression = gobjFSO.FolderExists(strDesFile)
    Else
        DeCompression = gobjFSO.FileExists(strDesFile)
    End If
    Exit Function
ErrH:
    If strErr = "" Then strErr = err.Description
    If 0 = 1 Then
        Resume
    End If
End Function

Private Function Get7zExePath() As String
    Dim strSystem As String
    
    strSystem = gobjFSO.GetSpecialFolder(SystemFolder)
    If Is64bit Then '64ϵͳ��32λ����Ӧ�÷���C:\windows\SysWOW64
        strSystem = gobjFSO.GetParentFolderName(strSystem) & "\SysWOW64"
    End If
    If gobjFSO.FileExists(strSystem & "\7z.exe") Then
        Get7zExePath = strSystem & "\7z.exe"
    End If
End Function