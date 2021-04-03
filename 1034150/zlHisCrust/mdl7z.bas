Attribute VB_Name = "mdl7z"
Option Explicit

'**************************
'����:ѹ��/��ѹ���ļ�
'��д����:ף��
'**************************

Public Const PROAPPCTION = "7z.exe" 'ִ�г���
Public Const COMPRESSIONRATE = 5 '��׼ѹ��

'''ѹ���ȼ� ѹ���㷨 �ֵ��С �����ֽ� ƥ���� ������ ����
'''0 Copy ��ѹ��
'''1 LZMA 64 KB 32 HC4 BCJ ���ѹ��
'''3 LZMA 1 MB 32 HC4 BCJ ����ѹ��
'''5 LZMA 16 MB 32 BT4 BCJ ����ѹ��
'''7 LZMA 32 MB 64 BT4 BCJ ���ѹ��
'''9 LZMA 64 MB 64 BT4 BCJ2 ����ѹ��

Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

'==============================================================================
'=���ܣ� ����ѹ����
'=������ strDescPath ѹ���ļ�����λ��
'        strFiles    ѡ��Ĵ�ѹ���ļ�
'        strRate     ѹ���ȼ�
'==============================================================================
Public Function CompressionCmd(ByVal strDescPath As String, ByVal strFiles As String, ByVal strRate As String) As String
    On Error GoTo errH
    Dim strShellCmd As String '��ִ�д�
    Dim str7zPath As String '7z.exe λ��
    Dim strM      As String
    Dim strParameter As String '�����ܴ�
    Dim strFile() As String '�ļ�����
    Dim i As Integer
    '��ϲ���
    strM = "-m"  '�̶������ַ�
    strParameter = strM & "x=" & strRate & " " '���õȼ�
    strParameter = strParameter & strM & "mt" & " " '������رն��߳�ѹ��ģʽ
    
    '����﷨
'    str7zPath = App.Path & "\" & PROAPPCTION  '�Ȳ��ұ����Ƿ���7zg.exe
'    If Dir(str7zPath, 63) = "" Then
    str7zPath = GetWinSystemPath & "\" & PROAPPCTION  '��ϵͳ���Ƿ���7zg.exe
    If Dir(str7zPath, 63) = "" Then
        CompressionCmd = ""
        Exit Function
    End If
'    End If
    strShellCmd = """" & str7zPath & """" & " "
    strShellCmd = strShellCmd & "a -y "
    strShellCmd = strShellCmd & """" & strDescPath & """" & " "
    
    '''''�ֽ��ļ�
    strFiles = Trim(strFiles)
    strFile = Split(strFiles, " ")
    strFiles = ""
    For i = 0 To UBound(strFile)
        strFiles = strFiles & """" & strFile(i) & """" & " "
    Next
    
    strShellCmd = strShellCmd & Trim(strFiles) & " "
    strShellCmd = strShellCmd & strParameter
    
    CompressionCmd = strShellCmd
    Exit Function
errH:
    If Err Then
       CompressionCmd = ""
    End If
End Function


'==============================================================================
'=���ܣ� ���ɽ�ѹ����
'=������ strSavePath ��ѹ�ļ�����λ��
'        strFile     ѡ��Ĵ���ѹ�ļ�
'==============================================================================
Public Function DeCompressionCmd(ByVal strSavePath As String, ByVal strFile As String) As String
    On Error GoTo errH
    Dim strShellCmd As String '��ִ�д�
    Dim str7zPath As String '7z.exe λ��
    Dim strM      As String
    Dim i As Integer
    '��ϲ���
    strM = "-o"  '�̶������ַ�

    '����﷨
'    str7zPath = App.Path & "\" & PROAPPCTION  '�Ȳ��ұ����Ƿ���7zg.exe
'    If Dir(str7zPath, 63) = "" Then
    str7zPath = GetWinSystemPath & "\" & PROAPPCTION   '��ϵͳ���Ƿ���7zg.exe
    If Dir(str7zPath, 63) = "" Then
        DeCompressionCmd = ""
        Exit Function
    End If
'    End If
    strShellCmd = """" & str7zPath & """" & " "
    strShellCmd = strShellCmd & "e -y "
    strShellCmd = strShellCmd & """" & strFile & """" & " "
    
    
    strShellCmd = strShellCmd & strM
    strShellCmd = strShellCmd & """" & strSavePath & """"
    
    DeCompressionCmd = strShellCmd
    Exit Function
errH:
    If Err Then
        DeCompressionCmd = ""
    End If
End Function

Private Function GetWinSystemPath() As String
    
    Dim Buffer As String
    Dim strSystem As String
    Dim rtn As Long
    Const MAX_PATH = 260
    
    Buffer = Space(MAX_PATH)
    rtn = GetSystemDirectory(Buffer, Len(Buffer))
    strSystem = Left(Buffer, rtn)
    
    GetWinSystemPath = strSystem
End Function

