VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCommFun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'---------------------------------------------------------------------
'模块名称：clsCommFun
'包含对象：实现算法和特殊功能的函数
'命名规则：操作类型+操作对象
'说明    ：各段的含义与上面类似。因为这类函数功能千差万别，不一定好按上面的规则命名，特殊情况也灵活命名。
'例      ：SetRegValue，CheckStrValidity，ToVarchar，NVL
'--------------------------------------------------------------------------
Private Const REG_BINARY = 3
Private Const REG_DWORD = 4
Private Const REG_DWORD_BIG_ENDIAN = 5
Private Const REG_DWORD_LITTLE_ENDIAN = 4
Private Const REG_EXPAND_SZ = 2
Private Const REG_MULTI_SZ = 7
Private Const REG_SZ = 1
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long

'切换到指定的输入法。
Private Declare Function ActivateKeyboardLayout Lib "user32" (ByVal hkl As Long, ByVal flags As Long) As Long
'返回系统中可用的输入法个数及各输入法所在Layout,包括英文输入法。
Private Declare Function GetKeyboardLayoutList Lib "user32" (ByVal nBuff As Long, lpList As Long) As Long
'获取某个输入法的名称
Private Declare Function ImmGetDescription Lib "imm32.dll" Alias "ImmGetDescriptionA" (ByVal hkl As Long, ByVal lpsz As String, ByVal uBufLen As Long) As Long
'判断某个输入法是否中文输入法
Private Declare Function ImmIsIME Lib "imm32.dll" (ByVal hkl As Long) As Long

Private Type BrowseInfo
   hwndOwner      As Long
   pIDLRoot       As Long
   pszDisplayName As Long
   lpszTitle      As Long
   ulFlags        As Long
   lpfnCallback   As Long
   lParam         As Long
   iImage         As Long
End Type
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
'Window版本函数

Private Type OSVERSIONINFO 'for GetVersionEx API call
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Private Const KEYEVENTF_EXTENDEDKEY = &H1
Private Const KEYEVENTF_KEYUP = &H2
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

'---------------------------------------------------
Private mfrmTipInfo As frmTipInfo

Public Sub PressKey(bytKey As Byte)
'功能：向键盘发送一个键,类似SendKey
'参数：bytKey=VirtualKey Codes，1-254，可以用vbKeyTab,vbKeyReturn,vbKeyF4
    Call keybd_event(bytKey, 0, KEYEVENTF_EXTENDEDKEY, 0)
    Call keybd_event(bytKey, 0, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0)
End Sub

Public Sub PressKeyEx(ByVal KeyCode As Integer, Optional ByVal Shift As Integer)
'功能：向键盘发送一个键,类似SendKey
'参数：KeyCode=VirtualKey Codes，1-254，可以用vbKeyTab,vbKeyReturn,vbKeyF4,vbKeyA等
'      Shift=vbKeyControl,vbKeyShift,vbKeyMenu(ALT)，如果是组合，用Or运算
    
    '按下组合键
    If (Shift And vbKeyControl) = vbKeyControl Then
        Call keybd_event(vbKeyControl, 0, KEYEVENTF_EXTENDEDKEY, 0)
    End If
    If (Shift And vbKeyMenu) = vbKeyMenu Then
        Call keybd_event(vbKeyMenu, 0, KEYEVENTF_EXTENDEDKEY, 0)
    End If
    If (Shift And vbKeyShift) = vbKeyShift Then
        Call keybd_event(vbKeyShift, 0, KEYEVENTF_EXTENDEDKEY, 0)
    End If
    
    '按下普通键
    Call keybd_event(KeyCode, 0, KEYEVENTF_EXTENDEDKEY, 0)
    
    '松开普通键
    Call keybd_event(KeyCode, 0, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0)
    
    '松开组合键
    If (Shift And vbKeyShift) = vbKeyShift Then
        Call keybd_event(vbKeyShift, 0, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0)
    End If
    If (Shift And vbKeyMenu) = vbKeyMenu Then
        Call keybd_event(vbKeyMenu, 0, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0)
    End If
    If (Shift And vbKeyControl) = vbKeyControl Then
        Call keybd_event(vbKeyControl, 0, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0)
    End If
End Sub

Public Function GetRegValue(ByVal hKey As Long, ByVal strSubKey As String, ByVal strValueName As String, strValue As String) As Boolean
'功能：获取注册表中指定位置的值
    Dim strData As String
    Dim lngLength As Long, lngReturn As Long
    Dim lngKey As Long, lngType As Long
    
    
    lngReturn = RegOpenKeyEx(hKey, strSubKey, 0, KEY_QUERY_VALUE, lngKey)
    If lngReturn <> 0 Then
        Exit Function
    End If
    
    strData = Space(255)
    lngLength = 255
    lngType = REG_SZ
    lngReturn = RegQueryValueEx(lngKey, strValueName, 0, lngType, strData, lngLength)
    If lngReturn <> 0 Then
        RegCloseKey lngKey
        Exit Function
    End If
    
    strValue = Mid(strData, 1, lngLength)
    
    If InStr(strValue, Chr(0)) > 0 Then
        strValue = Mid(strValue, 1, InStr(strValue, Chr(0)) - 1)
    End If
    RegCloseKey lngKey
    GetRegValue = True
End Function

Public Function SetRegValue(ByVal hKey As Long, ByVal strSubKey As String, ByVal strValueName As String, ByVal strValue As String) As Boolean
'功能：获取注册表中指定位置的值
    Dim lngLength As Long, lngReturn As Long
    Dim lngKey As Long, lngType As Long
    
    
    lngReturn = RegOpenKeyEx(hKey, strSubKey, 0, KEY_SET_VALUE, lngKey)
    If lngReturn <> 0 Then
        Exit Function
    End If
    
    lngLength = LenB(StrConv(strValue, vbFromUnicode))
    lngType = REG_SZ
    lngReturn = RegSetValueEx(lngKey, strValueName, 0, lngType, strValue, lngLength)
    If lngReturn <> 0 Then
        RegCloseKey lngKey
        Exit Function
    End If
    
    RegCloseKey lngKey
    SetRegValue = True
End Function

Public Function DeleteRegValue(ByVal hKey As Long, ByVal strSubKey As String, ByVal strValueName As String) As Boolean
'功能：获取注册表中指定位置的值
    Dim lngLength As Long, lngReturn As Long
    Dim lngKey As Long, lngType As Long
    
    
    lngReturn = RegOpenKeyEx(hKey, strSubKey, 0, KEY_SET_VALUE, lngKey)
    If lngReturn <> 0 Then
        Exit Function
    End If
    
    lngReturn = RegDeleteValue(lngKey, strValueName)
    If lngReturn <> 0 Then
        RegCloseKey lngKey
        Exit Function
    End If
    
    RegCloseKey lngKey
    DeleteRegValue = True
End Function

Public Sub ShowFlash(Optional strNote As String, Optional frmParent As Object)
    '------------------------------------------------
    '功能： 显示等待的动态窗体
    '参数：
    '   strNote:提示信息
    '   frmParent：关于窗体的父窗体
    '返回：
    '------------------------------------------------
    Dim strStyle As String   '导航台风格
    
    With frmFlash
        If strNote <> "" Then .lbl提示.Caption = strNote
        Err = 0
        On Error Resume Next
        .avi.Open gstrAviPath & "\" & "Findfile.avi"
        If Err <> 0 Then
            .lblFile.Visible = True
        End If
        .Refresh
        
        
        If frmParent Is Nothing Then
            .Show
        Else
            '10674问题 在MDI风格下, 这种显示方式会造成系统跳到另一个程序,影响用户操作.
             strStyle = UCase(gobjComLib.zlDatabase.GetPara("导航台"))
             If strStyle = "ZLMDI" Then
                .Show
             Else
                .Show , frmParent
             End If
        End If
        .Refresh
        If Not .lblFile.Visible Then .avi.Play
    End With
End Sub

Public Sub StopFlash()
    '------------------------------------------------
    '功能： 停止并关闭等待的动态窗体
    '参数：
    '返回：
    '------------------------------------------------
    On Error Resume Next
    frmFlash.avi.Stop
    Unload frmFlash
End Sub

Public Function SpellCode(ByVal strAsk As String) As String
'-------------------------------------------------------------
'功能：返回指定字符串的简码
'说明：根据指定字符串生成简码，可以生成三种类型的简码
'        0、拼音，取每字的首字母构成简码
'        1、五笔，取每字的首字母构成简码
'        2、五笔，按五笔规则构成简码
'        3、拼音，完整的拼音码
'      在传入的参数中未发现※符号，就按用户在系统选项中设置的方式生成简码；
'        否则就按在※符号后的数字指定的方式强制生成简码，如※1表示按五笔首字母生成
'-------------------------------------------------------------
    Dim lngSplit As Long, lng方式 As Long
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    '取完整拼音要用的变量
    Dim i As Integer            '临时变量
    Dim j As Integer            '临时变量
    Dim PYStr() As String      '存储拼音对照表
    Dim tmpStr As String       '临时字串
    '取五笔规则要用的变量
    Dim var编码表 As Variant, str汉字 As String, intBit As Integer, strBit As String, str编码 As String
    
    If Len(Trim(strAsk)) < 1 Then Exit Function                 '为空时退出
    
    lngSplit = InStr(strAsk, "※")
    If lngSplit = 0 Then
        '按用户在系统选项中设置的方式生成简码
        lng方式 = Val(gobjComLib.zlDatabase.GetPara("简码方式"))
    Else
        '指定的方式强制生成简码
        lng方式 = Val(Mid(strAsk, lngSplit + 1))
        strAsk = Mid(strAsk, 1, lngSplit - 1)
    End If
    
    Select Case lng方式
    
    Case 0, 1
        '拼音，五笔简码
        If lng方式 = 0 Then
            Set rsTmp = gobjComLib.zlDatabase.OpenCursor("clsCommFun", "zlTools", "b_ComFunc.Get_Spell_Code", strAsk, 0)
        Else
            Set rsTmp = gobjComLib.zlDatabase.OpenCursor("clsCommFun", "zlTools", "b_ComFunc.Get_Spell_Code", strAsk, 1)
        End If
        
        On Error GoTo errHand
        SpellCode = IIf(IsNull(rsTmp.Fields(0).value), "", rsTmp.Fields(0).value)

        
    Case 2
        '按五笔规则
        strAsk = StrConv(Trim(strAsk), vbNarrow + vbUpperCase)
        
        For intBit = 1 To Len(strAsk)
            strBit = Mid(strAsk, intBit, 1)
            If LenB(StrConv(strBit, vbFromUnicode)) = 2 Then
                '汉字
                str汉字 = str汉字 & strBit
            End If
        Next
        
        var编码表 = Split(LoadResString(102) & LoadResString(103), ",")
        For j = 1 To Len(str汉字)
            For i = 0 To UBound(var编码表)
               
                If InStr(var编码表(i), Mid(str汉字, j, 1)) > 0 Then
                    tmpStr = Mid(var编码表(i), InStr(var编码表(i), "|") + 1)
                    Exit For
                End If
            Next
            If Trim(tmpStr) <> "" Then
                If str编码 = "" Then
                    str编码 = Trim(tmpStr)
                Else
                    str编码 = str编码 & "|" & Trim(tmpStr)
                End If
            End If
        Next
        
        If str编码 <> "" Then
            var编码表 = Split(str编码, "|")
            Select Case Len(str汉字)
            Case 0
            Case 1
               SpellCode = var编码表(0)
            Case 2
               SpellCode = Mid(var编码表(0), 1, 2) & Mid(var编码表(1), 1, 2)
            Case 3
               SpellCode = Mid(var编码表(0), 1, 1) & Mid(var编码表(1), 1, 1) & Mid(var编码表(2), 1, 2)
            Case Else
               SpellCode = Mid(var编码表(0), 1, 1) & Mid(var编码表(1), 1, 1) & Mid(var编码表(2), 1, 1) & Mid(var编码表((UBound(var编码表))), 1, 1)
            End Select
        End If
     Case Else
        '------------------------------------------------
        '功能：取得汉字的全拼
        '编制人：曾超 2005-8-31
        '------------------------------------------------
            
        
        PYStr = Split(LoadResString(101), ",")
        For j = 1 To Len(strAsk)
            For i = 0 To UBound(PYStr)
                If InStr(PYStr(i), Mid(strAsk, j, 1)) > 0 Then
                    tmpStr = Mid(PYStr(i), 1, InStr(PYStr(i), "=>") - 1)
                    Exit For
                End If
            Next
            If Trim(tmpStr) <> "" Then
                If SpellCode = "" Then
                    SpellCode = Trim(tmpStr)
                Else
                    SpellCode = SpellCode & " " & Trim(tmpStr)
                End If
            End If
        Next

    End Select
    
    Exit Function
errHand:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
    SpellCode = "-"
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


Public Function IsNumOrChar(ByVal strAsk As String) As Boolean
    '-------------------------------------------------------------
    '功能：判断指定字符串是否全部由数字和英文字母构成，用于允许数字
    '       和字母但不允许特殊字符的情况下的检测，isnumberic只能判断数字。
    '参数：（SSC编制）
    '       strAsk
    '返回：
    '-------------------------------------------------------------
    Dim i As Integer, j As Integer
    
    If Len(Trim(strAsk)) > 0 Then
        For i = 1 To Len(Trim(strAsk))
            j = Asc(Mid(Trim(strAsk), i, 1))
            If Not ((j > 47 And j < 58) Or (j > 64 And j < 91) Or (j > 96 And j < 123)) Then
                IsNumOrChar = False
                Exit Function
            End If
        Next
    End If
    IsNumOrChar = True

End Function

Public Function IsCharAlpha(ByVal strAsk As String) As Boolean
    '-------------------------------------------------------------
    '功能：判断指定字符串是否全部由英文字母构成    '
    '参数：
    '       strAsk
    '返回：
    '-------------------------------------------------------------
    Dim i As Integer, j As Integer
    
    If Len(Trim(strAsk)) > 0 Then
        For i = 1 To Len(Trim(strAsk))
            j = Asc(Mid(Trim(strAsk), i, 1))
            If Not ((j > 64 And j < 91) Or (j > 96 And j < 123)) Then
                IsCharAlpha = False
                Exit Function
            End If
        Next
    End If
    IsCharAlpha = True
End Function

Public Function IsCharChinese(ByVal strAsk As String) As Boolean
    '-------------------------------------------------------------
    '功能：判断指定字符串是否含有汉字
    '参数：
    '       strAsk
    '返回：
    '-------------------------------------------------------------
    Dim i As Integer, j As Integer
    
    If Len(Trim(strAsk)) > 0 Then
        For i = 1 To Len(Trim(strAsk))
            j = Asc(Mid(Trim(strAsk), i, 1))
            If j < 0 Then
                IsCharChinese = True
                Exit Function
            End If
        Next
    End If
    IsCharChinese = False
End Function

Public Function UppeMoney(curMoney) As String
    '-------------------------------------------------------------
    '功能：将指定的金额数值转换为大写金额显示
    '参数：
    '       curMoney:需要转换的金额数值
    '返回：
    '-------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    Set rsTmp = gobjComLib.zlDatabase.OpenCursor("clsCommFun", "zltools", "b_ComFunc.get_UppMoney", curMoney)
    
    UppeMoney = rsTmp.Fields(0).value
    Exit Function
errH:
    If gobjComLib.ErrCenter() = 1 Then Resume
End Function

Public Function ZyMod(x1, x2) As Long
'-------------------------------------------------------------
'功能：取余函数，由于vb的Mod函数仅仅对亿位以下起作用，超过时出错
'参数：
'       X1-除数；X2-被除数
'返回：
'-------------------------------------------------------------
    ZyMod = x1 - Int(x1 / x2) * x2
End Function

Public Function SelOper(ByVal frmParent As Object) As String
'-------------------------------------------------------------
'功能：显示人员选择器，按部门进行选择
'-------------------------------------------------------------
    Dim frmTemp As New frmManSelect
    frmTemp.Show vbModal, frmParent
    SelOper = frmTemp.mstrReturn
    
    Set frmTemp = Nothing
End Function

Public Function OpenIme(Optional blnOpen As Boolean = False) As Boolean
    
    '功能:打开中文输入法，或关闭输入法
    Dim arrIme(99) As Long, lngCount As Long, strName As String * 255
    Dim strIme As String, blnNotCloseIme As Boolean
     
    '用户没进行设置，就不处理
    strIme = gobjComLib.zlDatabase.GetPara("输入法")
    If strIme = "" And blnOpen = True Then Exit Function                 '要求打开输入法，但是又没有设置
    
    lngCount = GetKeyboardLayoutList(UBound(arrIme) + 1, arrIme(0))
    blnNotCloseIme = True
    Do
        lngCount = lngCount - 1
        If ImmIsIME(arrIme(lngCount)) = 1 Then
            If blnOpen = True Then
                '需要打开输入法。接着判断是否批定输入法
                ImmGetDescription arrIme(lngCount), strName, Len(strName)
                If InStr(1, Mid(strName, 1, InStr(1, strName, Chr(0)) - 1), strIme) > 0 And strIme <> "" Then
                    If ActivateKeyboardLayout(arrIme(lngCount), 0) <> 0 Then OpenIme = True: Exit Function
                End If
            End If
        ElseIf blnOpen = False Then
            '不是输入法，正好是应了关闭输入法的请求
            If ActivateKeyboardLayout(arrIme(lngCount), 0) <> 0 Then OpenIme = True: Exit Function
        End If
    Loop Until lngCount = 0
    
    If blnNotCloseIme And blnOpen = False Then
        '由于windows Vista系统的英文输入法用ImmIsIME测试出是true的输入法,因此,需要单独处理.
        '刘兴宏:2008/09/03
        If ActivateKeyboardLayout(arrIme(0), 0) <> 0 Then OpenIme = True: Exit Function
    End If
    
End Function

Public Function GetLike(ByVal strTable As String, ByVal strField As String, ByVal strInput As String) As String
'功能：产生指定字段的Like条件子句
'参数：strTable     表名，与SQL语句相关，可以为别名
'      strField     字段名
'      strInput     输入的字符串
    Dim strMatch  As String
    
    strMatch = IIf(Val(gobjComLib.zlDatabase.GetPara("输入匹配")) = 0, "%", "")
    strInput = UCase(Replace(Trim(strInput), "'", "''"))
    
    GetLike = "Upper(" & IIf(strTable = "", "", strTable & ".") & strField & _
                   ") LIKE '" & strMatch & strInput & "%'"
End Function

Public Function ToVarchar(ByVal varText As Variant, ByVal lngLength As Long, Optional strDefault As String = "") As String
'功能：将文本按Varchar2的长度计算方法进行截断
    Dim strText As String
    
    strText = IIf(IsNull(varText), strDefault, varText)
    ToVarchar = StrConv(LeftB(StrConv(strText, vbFromUnicode), lngLength), vbUnicode)
    '去掉可能出现的半个字符
    ToVarchar = Replace(ToVarchar, Chr(0), "")
End Function

Public Function AddDate(ByVal strOrin As String, Optional ByVal bln时 As Boolean = False) As String
'功能：为不全的日期信息补充完整
'参数：strOrin  用户输入的原始值
'      bln时    是否增加小时部分
    Dim strTemp As String
    Dim intPos As Integer
    
    strTemp = Trim(strOrin)
    
    If strTemp = "" Then
        AddDate = ""
        Exit Function
    End If
    
    intPos = InStr(strTemp, "-")
    If intPos = 0 Then
        intPos = InStr(strTemp, ".")
        If intPos <> 0 Then
            '使用 . 隔
            strTemp = Replace(strTemp, ".", "-")
        End If
    End If
    
    If intPos = 0 Then
        '没有"-",手工加上
        intPos = Len(strTemp)
        If intPos <= 8 Then
            If intPos = 8 Then
                strTemp = Mid(strTemp, 1, 4) & "-" & Mid(strTemp, 5, 2) & "-" & Mid(strTemp, 7, 2)
            ElseIf intPos > 4 Then
                strTemp = Left(strTemp, intPos - 4) & "-" & Mid(Right(strTemp, 4), 1, 2) & "-" & Right(strTemp, 2)
            ElseIf intPos > 2 Then
                strTemp = Format(Date, "yyyy") & "-" & Left(strTemp, intPos - 2) & "-" & Right(strTemp, 2)
            Else
                strTemp = Format(Date, "yyyy") & "-" & Format(Date, "MM") & "-" & strTemp
            End If
        End If
    Else
        If bln时 = False Then
            If IsDate(strTemp) Then
                strTemp = Format(CDate(strTemp), "yyyy-MM-dd")
            End If
        Else
            '处理小时
            If InStr(strTemp, " ") > 0 Then
                '输入了小时
                If IsDate(strTemp & ":00") Then
                    strTemp = Format(CDate(strTemp & ":00"), "yyyy-MM-dd HH:mm")
                End If
            Else
                If IsDate(strTemp) Then
                    strTemp = Format(CDate(strTemp), "yyyy-MM-dd HH:mm")
                End If
            End If
        End If
    End If
    
    AddDate = strTemp
End Function

Public Function NVL(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    NVL = IIf(IsNull(varValue), DefaultValue, varValue)
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

Public Function OpenDir(hwndOwner As Long, Optional strTitle As String) As String
'功能：选择Windows的文件夹
    Const BIF_RETURNONLYFSDIRS = 1
    Const BIF_DONTGOBELOWDOMAIN = 2

    Dim lpIDList As Long
    Dim sBuffer As String
    Dim tBrowseInfo As BrowseInfo
    With tBrowseInfo
       .hwndOwner = hwndOwner
       .lpszTitle = lstrcat(strTitle, "")
       .ulFlags = BIF_RETURNONLYFSDIRS ' + BIF_DONTGOBELOWDOMAIN
    End With
    lpIDList = SHBrowseForFolder(tBrowseInfo)
    If (lpIDList) Then
       sBuffer = Space(MAX_PATH)
       SHGetPathFromIDList lpIDList, sBuffer
       sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
       OpenDir = sBuffer
    End If
End Function

Public Function InputIsCard(ByRef txtInput As Object, ByVal KeyAscii As Integer, ByVal lngSys As Long) As Boolean
'功能：判断指定文本框中当前输入是否在刷卡(是否达到卡号长度，在调用程序中判断),并根据系统参数处理是否密文显示
'参数：KeyAscii=在KeyPress事件中调用的参数
    Static sngInputBegin As Single
    Dim sngNow As Single, blnCard As Boolean, strText As String
        
    '刷卡时含有特殊符号的要取消
    If InStr(":：;；?？", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Function
                    
    '处理当前键入后显示的内容(还未显示出来)
    strText = txtInput.Text
    If txtInput.SelLength = Len(txtInput.Text) Then strText = ""
    If KeyAscii = 8 Then
        If strText <> "" Then strText = Mid(strText, 1, Len(strText) - 1)
    Else
        strText = strText & Chr(KeyAscii)
    End If
        
    '判断是否在刷卡
    If IsNumeric(strText) And IsNumeric(Left(strText, 1)) Then  '姓名输入框如果输的是全数字，认为是刷卡
        blnCard = True
    ElseIf KeyAscii > 32 Then
        sngNow = Timer
        If txtInput.Text = "" Or strText = "" Then
            sngInputBegin = sngNow
        Else
            If Format((sngNow - sngInputBegin) / Len(strText), "0.000") < 0.04 Then blnCard = True   '用一台笔记本测试，一般在0.014左右
        End If
    End If
    
    '刷卡时卡号是否密文显示
    If blnCard Then
        txtInput.PasswordChar = IIf(gobjComLib.zlDatabase.GetPara(12, lngSys) = "0", "", "*")
    Else
        txtInput.PasswordChar = ""
    End If
    
    InputIsCard = blnCard
End Function


Public Function IncStr(ByVal strVal As String) As String
'功能：对一个字符串自动加1。
'说明：每一位进位时,如果是数字,则按十进制处理,否则按26进制处理
'参数：strVal=要加1的字符串
    Dim i As Long, strTmp As String, bytUp As Byte, bytAdd As Byte
    
    For i = Len(strVal) To 1 Step -1
        If i = Len(strVal) Then
            bytAdd = 1
        Else
            bytAdd = 0
        End If
        If IsNumeric(Mid(strVal, i, 1)) Then
            If CByte(Mid(strVal, i, 1)) + bytAdd + bytUp < 10 Then
                strVal = Left(strVal, i - 1) & CByte(Mid(strVal, i, 1)) + bytAdd + bytUp & Mid(strVal, i + 1)
                bytUp = 0
            Else
                strVal = Left(strVal, i - 1) & "0" & Mid(strVal, i + 1)
                bytUp = 1
            End If
        Else
            If Asc(Mid(strVal, i, 1)) + bytAdd + bytUp <= Asc("Z") Then
                strVal = Left(strVal, i - 1) & Chr(Asc(Mid(strVal, i, 1)) + bytAdd + bytUp) & Mid(strVal, i + 1)
                bytUp = 0
            Else
                strVal = Left(strVal, i - 1) & "0" & Mid(strVal, i + 1)
                bytUp = 1
            End If
        End If
        If bytUp = 0 Then Exit For
    Next
    IncStr = strVal
End Function

Public Function GetNeedName(strList As String) As String
'功能：从编码名称组合串中分离出名称
'参数：strList=编码名称组合串,如"012-内科","(012)内科","[012]内科"
    If InStr(strList, Chr(13)) > 0 Then
        GetNeedName = LTrim(Mid(strList, InStr(strList, Chr(13)) + 1))
    ElseIf InStr(strList, "]") > 0 And InStr(strList, "-") = 0 Then
        GetNeedName = LTrim(Mid(strList, InStr(strList, "]") + 1))
    ElseIf InStr(strList, ")") > 0 And InStr(strList, "-") = 0 Then
        GetNeedName = LTrim(Mid(strList, InStr(strList, ")") + 1))
    Else
        GetNeedName = LTrim(Mid(strList, InStr(strList, "-") + 1))
    End If
End Function

Private Function PreFixNO(Optional curDate As Date = #1/1/1900#) As String
'功能：根据ZLHIS的单据号前缀规则,返回大写的单据号年前缀
'参数：curDate=单据号前缀基准时间,不传时基准时间为服务器时间
'应用：该函数主要是配合GetFullNO函数应用

    If curDate = #1/1/1900# Then
        PreFixNO = CStr(CInt(Format(gobjComLib.zlDatabase.CurrentDate, "YYYY")) - 1990)
    Else
        PreFixNO = CStr(CInt(Format(curDate, "YYYY")) - 1990)
    End If
    PreFixNO = IIf(CInt(PreFixNO) < 10, PreFixNO, Chr(55 + CInt(PreFixNO)))
End Function

Public Function GetFullNO(strNO As String) As String
'功能：由用户输入的部份单号，返回当年的完整单据号。
'参数：strNO=由用户输入的单据号部分

    If Len(strNO) >= 8 Then GetFullNO = Right(strNO, 8): Exit Function
    GetFullNO = PreFixNO & Format(strNO, "0000000")
End Function

Public Function GetIDCardDate(strCardID As String) As String
'功能：根据身份证号返回出生日期
'参数：ID=身份证号,应该为15位或18位
'返回：格式"yyyy-MM-dd"
    Dim strTmp As String
    
    If Len(strCardID) = 15 Then
        strTmp = Mid(strCardID, 7, 6)
        If Len(strTmp) = 6 And IsNumeric(strTmp) Then
            strTmp = "19" & Left(strTmp, 2) & "-" & Mid(strTmp, 3, 2) & "-" & Right(strTmp, 2)
        End If
    ElseIf Len(strCardID) = 18 Then
        strTmp = Mid(strCardID, 7, 8)
        If Len(strTmp) = 8 And IsNumeric(strTmp) Then
            strTmp = Left(strTmp, 4) & "-" & Mid(strTmp, 5, 2) & "-" & Right(strTmp, 2)
        End If
    End If
    If IsDate(strTmp) Then GetIDCardDate = strTmp
End Function

Private Function GetWinPlatform() As Long
    Dim osvi As OSVERSIONINFO
    Dim strCSDVersion As String

    osvi.dwOSVersionInfoSize = Len(osvi)

    If GetVersionEx(osvi) = 0 Then
        Exit Function
    End If

    GetWinPlatform = osvi.dwPlatformId
End Function

Public Function IsWindowsNT() As Boolean
'功能：判断是否WindowNT操作系统
    Const dwMaskNT = &H2&

    IsWindowsNT = (GetWinPlatform() And dwMaskNT)
End Function

Public Function IsWindows9X() As Boolean
'功能：判断是否Window95操作系统
    Const dwMask95 = &H1&

    IsWindows9X = (GetWinPlatform() And dwMask95)
End Function

Public Sub SetWindowsInTaskBar(ByVal lngHwnd As Long, ByVal blnShow As Boolean)
'功能：设置窗体是否在任务条上显示
    Dim lngStyle As Long
    
    lngStyle = GetWindowLong(lngHwnd, GWL_EXSTYLE)
    If blnShow Then
        lngStyle = lngStyle Or &H40000
    Else
        lngStyle = lngStyle And Not &H40000
    End If
    Call SetWindowLong(lngHwnd, GWL_EXSTYLE, lngStyle)
End Sub

Public Function StrIsValid(ByVal strInput As String, Optional ByVal intMax As Integer = 0, Optional ByVal hwnd As Long = 0, Optional str项目 As String) As Boolean
'检查字符串是否含有非法字符；如果提供长度，对长度的合法性也作检测。
    If str项目 = "" Then str项目 = "所输入内容"
    
    If InStr(strInput, "'") > 0 Or InStr(strInput, "|") > 0 Then
        MsgBox str项目 & "含有非法字符。", vbExclamation, gstrSysName
        If hwnd <> 0 Then SetFocusHwnd hwnd              '设置焦点
        Exit Function
    End If
    If intMax > 0 Then
        If LenB(StrConv(strInput, vbFromUnicode)) > intMax Then
            MsgBox str项目 & "不能超过" & Int(intMax / 2) & "个汉字" & "或" & intMax & "个字母。", vbExclamation, gstrSysName
            If hwnd <> 0 Then SetFocusHwnd hwnd              '设置焦点
            Exit Function
        End If
    End If
    
    StrIsValid = True
End Function

Public Function ShowMsgBox(ByVal strCaption As String, ByVal strInfo As String, ByVal strCmds As String, _
    frmParent As Object, Optional vStyle As VbMsgBoxStyle = vbQuestion) As String
'参数：strCaption=消息窗体标题
'      strInfo=具体提示内容,可用"^"表示换行,">"表示缩进。
'      strCmds=按钮描述,如"重试(&R),!忽略(&A),?取消(&C)"
'              至少要有两个按钮,"!"表示缺省按钮,"?"表示取消按钮
'              每个按钮文字最多支持4个汉字
'      vStyle=vbInformation,vbQuestion,vbExclamation,vbCritical
'返回：按钮文字,如"按钮2"(不包含()和&),如果按关闭或取消则返回""
    ShowMsgBox = frmMsgBox.ShowMsgBox(strCaption, strInfo, strCmds, frmParent, vStyle)
End Function

Public Function VerifyPassWord(frmParent As Object, ByVal strPass As String, Optional ByVal strName As String, Optional ByVal strSex As String, Optional ByVal strOld As String) As Boolean
'功能：对密码进行验证
'参数：frmParent=显示的父窗体
'      strPass=正确的密码
'      strName,strSex,strOld=可选参数，病人姓名、性别、年龄，当不传入时不显示这个区域。
'返回：True=密码验证通过,False=取消输入，或连续3次输入错误的密码
    VerifyPassWord = frmVerifyPassword.ShowMe(frmParent, strPass, strName, strSex, strOld)
End Function

Public Sub DeviceSetup(frmParent As Object, ByVal lngSys As Long, ByVal lngModule As Long)
'功能:根据不同的系统和模块,调用不同的设备配置
    
    Call frmDeviceSetup.ShowMe(frmParent, lngSys, lngModule)
End Sub

Public Sub ShowTipInfo(ByVal lngHwnd As Long, ByVal strInfo As String, Optional blnMultiRow As Boolean)
'功能：显示或者隐藏提示
'参数：lngHwnd=提示所针对的控件句柄,当传入为0时隐藏提示
'      strInfo=提示信息,当传入为空时隐藏提示
'      blnMultiRow=以一定的间距分行显示多行信息，每行按vbcrlf分隔

    Call mfrmTipInfo.ShowTipInfo(lngHwnd, strInfo, blnMultiRow)
End Sub

Private Sub Class_Initialize()
    If mfrmTipInfo Is Nothing Then
        Set mfrmTipInfo = New frmTipInfo
    End If
End Sub

Private Sub Class_Terminate()
    If Not mfrmTipInfo Is Nothing Then
        Unload mfrmTipInfo
        Set mfrmTipInfo = Nothing
    End If
End Sub
