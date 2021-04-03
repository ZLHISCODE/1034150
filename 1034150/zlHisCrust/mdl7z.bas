Attribute VB_Name = "mdl7z"
Option Explicit

'**************************
'功能:压缩/解压缩文件
'编写整理:祝庆
'**************************

Public Const PROAPPCTION = "7z.exe" '执行程序
Public Const COMPRESSIONRATE = 5 '标准压缩

'''压缩等级 压缩算法 字典大小 快速字节 匹配器 过滤器 描述
'''0 Copy 无压缩
'''1 LZMA 64 KB 32 HC4 BCJ 最快压缩
'''3 LZMA 1 MB 32 HC4 BCJ 快速压缩
'''5 LZMA 16 MB 32 BT4 BCJ 正常压缩
'''7 LZMA 32 MB 64 BT4 BCJ 最大压缩
'''9 LZMA 64 MB 64 BT4 BCJ2 极限压缩

Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

'==============================================================================
'=功能： 生成压缩串
'=参数： strDescPath 压缩文件保存位置
'        strFiles    选择的待压缩文件
'        strRate     压缩等级
'==============================================================================
Public Function CompressionCmd(ByVal strDescPath As String, ByVal strFiles As String, ByVal strRate As String) As String
    On Error GoTo errH
    Dim strShellCmd As String '总执行串
    Dim str7zPath As String '7z.exe 位置
    Dim strM      As String
    Dim strParameter As String '参数总串
    Dim strFile() As String '文件数组
    Dim i As Integer
    '组合参数
    strM = "-m"  '固定传输字符
    strParameter = strM & "x=" & strRate & " " '设置等级
    strParameter = strParameter & strM & "mt" & " " '开启或关闭多线程压缩模式
    
    '组合语法
'    str7zPath = App.Path & "\" & PROAPPCTION  '先查找本地是否有7zg.exe
'    If Dir(str7zPath, 63) = "" Then
    str7zPath = GetWinSystemPath & "\" & PROAPPCTION  '查系统下是否有7zg.exe
    If Dir(str7zPath, 63) = "" Then
        CompressionCmd = ""
        Exit Function
    End If
'    End If
    strShellCmd = """" & str7zPath & """" & " "
    strShellCmd = strShellCmd & "a -y "
    strShellCmd = strShellCmd & """" & strDescPath & """" & " "
    
    '''''分解文件
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
'=功能： 生成解压缩串
'=参数： strSavePath 解压文件保存位置
'        strFile     选择的待解压文件
'==============================================================================
Public Function DeCompressionCmd(ByVal strSavePath As String, ByVal strFile As String) As String
    On Error GoTo errH
    Dim strShellCmd As String '总执行串
    Dim str7zPath As String '7z.exe 位置
    Dim strM      As String
    Dim i As Integer
    '组合参数
    strM = "-o"  '固定传输字符

    '组合语法
'    str7zPath = App.Path & "\" & PROAPPCTION  '先查找本地是否有7zg.exe
'    If Dir(str7zPath, 63) = "" Then
    str7zPath = GetWinSystemPath & "\" & PROAPPCTION   '查系统下是否有7zg.exe
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

