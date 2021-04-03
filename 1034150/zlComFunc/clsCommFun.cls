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
'ģ�����ƣ�clsCommFun
'��������ʵ���㷨�����⹦�ܵĺ���
'�������򣺲�������+��������
'˵��    �����εĺ������������ơ���Ϊ���ຯ������ǧ����𣬲�һ���ð�����Ĺ����������������Ҳ���������
'��      ��SetRegValue��CheckStrValidity��ToVarchar��NVL
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

'�л���ָ�������뷨��
Private Declare Function ActivateKeyboardLayout Lib "user32" (ByVal hkl As Long, ByVal flags As Long) As Long
'����ϵͳ�п��õ����뷨�����������뷨����Layout,����Ӣ�����뷨��
Private Declare Function GetKeyboardLayoutList Lib "user32" (ByVal nBuff As Long, lpList As Long) As Long
'��ȡĳ�����뷨������
Private Declare Function ImmGetDescription Lib "imm32.dll" Alias "ImmGetDescriptionA" (ByVal hkl As Long, ByVal lpsz As String, ByVal uBufLen As Long) As Long
'�ж�ĳ�����뷨�Ƿ��������뷨
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
'Window�汾����

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
'���ܣ�����̷���һ����,����SendKey
'������bytKey=VirtualKey Codes��1-254��������vbKeyTab,vbKeyReturn,vbKeyF4
    Call keybd_event(bytKey, 0, KEYEVENTF_EXTENDEDKEY, 0)
    Call keybd_event(bytKey, 0, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0)
End Sub

Public Sub PressKeyEx(ByVal KeyCode As Integer, Optional ByVal Shift As Integer)
'���ܣ�����̷���һ����,����SendKey
'������KeyCode=VirtualKey Codes��1-254��������vbKeyTab,vbKeyReturn,vbKeyF4,vbKeyA��
'      Shift=vbKeyControl,vbKeyShift,vbKeyMenu(ALT)���������ϣ���Or����
    
    '������ϼ�
    If (Shift And vbKeyControl) = vbKeyControl Then
        Call keybd_event(vbKeyControl, 0, KEYEVENTF_EXTENDEDKEY, 0)
    End If
    If (Shift And vbKeyMenu) = vbKeyMenu Then
        Call keybd_event(vbKeyMenu, 0, KEYEVENTF_EXTENDEDKEY, 0)
    End If
    If (Shift And vbKeyShift) = vbKeyShift Then
        Call keybd_event(vbKeyShift, 0, KEYEVENTF_EXTENDEDKEY, 0)
    End If
    
    '������ͨ��
    Call keybd_event(KeyCode, 0, KEYEVENTF_EXTENDEDKEY, 0)
    
    '�ɿ���ͨ��
    Call keybd_event(KeyCode, 0, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0)
    
    '�ɿ���ϼ�
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
'���ܣ���ȡע�����ָ��λ�õ�ֵ
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
'���ܣ���ȡע�����ָ��λ�õ�ֵ
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
'���ܣ���ȡע�����ָ��λ�õ�ֵ
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
    '���ܣ� ��ʾ�ȴ��Ķ�̬����
    '������
    '   strNote:��ʾ��Ϣ
    '   frmParent�����ڴ���ĸ�����
    '���أ�
    '------------------------------------------------
    Dim strStyle As String   '����̨���
    
    With frmFlash
        If strNote <> "" Then .lbl��ʾ.Caption = strNote
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
            '10674���� ��MDI�����, ������ʾ��ʽ�����ϵͳ������һ������,Ӱ���û�����.
             strStyle = UCase(gobjComLib.zlDatabase.GetPara("����̨"))
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
    '���ܣ� ֹͣ���رյȴ��Ķ�̬����
    '������
    '���أ�
    '------------------------------------------------
    On Error Resume Next
    frmFlash.avi.Stop
    Unload frmFlash
End Sub

Public Function SpellCode(ByVal strAsk As String) As String
'-------------------------------------------------------------
'���ܣ�����ָ���ַ����ļ���
'˵��������ָ���ַ������ɼ��룬���������������͵ļ���
'        0��ƴ����ȡÿ�ֵ�����ĸ���ɼ���
'        1����ʣ�ȡÿ�ֵ�����ĸ���ɼ���
'        2����ʣ�����ʹ��򹹳ɼ���
'        3��ƴ����������ƴ����
'      �ڴ���Ĳ�����δ���֡����ţ��Ͱ��û���ϵͳѡ�������õķ�ʽ���ɼ��룻
'        ����Ͱ��ڡ����ź������ָ���ķ�ʽǿ�����ɼ��룬���1��ʾ���������ĸ����
'-------------------------------------------------------------
    Dim lngSplit As Long, lng��ʽ As Long
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    'ȡ����ƴ��Ҫ�õı���
    Dim i As Integer            '��ʱ����
    Dim j As Integer            '��ʱ����
    Dim PYStr() As String      '�洢ƴ�����ձ�
    Dim tmpStr As String       '��ʱ�ִ�
    'ȡ��ʹ���Ҫ�õı���
    Dim var����� As Variant, str���� As String, intBit As Integer, strBit As String, str���� As String
    
    If Len(Trim(strAsk)) < 1 Then Exit Function                 'Ϊ��ʱ�˳�
    
    lngSplit = InStr(strAsk, "��")
    If lngSplit = 0 Then
        '���û���ϵͳѡ�������õķ�ʽ���ɼ���
        lng��ʽ = Val(gobjComLib.zlDatabase.GetPara("���뷽ʽ"))
    Else
        'ָ���ķ�ʽǿ�����ɼ���
        lng��ʽ = Val(Mid(strAsk, lngSplit + 1))
        strAsk = Mid(strAsk, 1, lngSplit - 1)
    End If
    
    Select Case lng��ʽ
    
    Case 0, 1
        'ƴ������ʼ���
        If lng��ʽ = 0 Then
            Set rsTmp = gobjComLib.zlDatabase.OpenCursor("clsCommFun", "zlTools", "b_ComFunc.Get_Spell_Code", strAsk, 0)
        Else
            Set rsTmp = gobjComLib.zlDatabase.OpenCursor("clsCommFun", "zlTools", "b_ComFunc.Get_Spell_Code", strAsk, 1)
        End If
        
        On Error GoTo errHand
        SpellCode = IIf(IsNull(rsTmp.Fields(0).value), "", rsTmp.Fields(0).value)

        
    Case 2
        '����ʹ���
        strAsk = StrConv(Trim(strAsk), vbNarrow + vbUpperCase)
        
        For intBit = 1 To Len(strAsk)
            strBit = Mid(strAsk, intBit, 1)
            If LenB(StrConv(strBit, vbFromUnicode)) = 2 Then
                '����
                str���� = str���� & strBit
            End If
        Next
        
        var����� = Split(LoadResString(102) & LoadResString(103), ",")
        For j = 1 To Len(str����)
            For i = 0 To UBound(var�����)
               
                If InStr(var�����(i), Mid(str����, j, 1)) > 0 Then
                    tmpStr = Mid(var�����(i), InStr(var�����(i), "|") + 1)
                    Exit For
                End If
            Next
            If Trim(tmpStr) <> "" Then
                If str���� = "" Then
                    str���� = Trim(tmpStr)
                Else
                    str���� = str���� & "|" & Trim(tmpStr)
                End If
            End If
        Next
        
        If str���� <> "" Then
            var����� = Split(str����, "|")
            Select Case Len(str����)
            Case 0
            Case 1
               SpellCode = var�����(0)
            Case 2
               SpellCode = Mid(var�����(0), 1, 2) & Mid(var�����(1), 1, 2)
            Case 3
               SpellCode = Mid(var�����(0), 1, 1) & Mid(var�����(1), 1, 1) & Mid(var�����(2), 1, 2)
            Case Else
               SpellCode = Mid(var�����(0), 1, 1) & Mid(var�����(1), 1, 1) & Mid(var�����(2), 1, 1) & Mid(var�����((UBound(var�����))), 1, 1)
            End Select
        End If
     Case Else
        '------------------------------------------------
        '���ܣ�ȡ�ú��ֵ�ȫƴ
        '�����ˣ����� 2005-8-31
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
    '���ܣ���ȡָ���ַ�����ʵ�ʳ��ȣ������ж�ʵ�ʰ���˫�ֽ��ַ�����
    '       ʵ�����ݴ洢����
    '������
    '       strAsk
    '���أ�
    '-------------------------------------------------------------
    ActualLen = LenB(StrConv(strAsk, vbFromUnicode))
End Function


Public Function IsNumOrChar(ByVal strAsk As String) As Boolean
    '-------------------------------------------------------------
    '���ܣ��ж�ָ���ַ����Ƿ�ȫ�������ֺ�Ӣ����ĸ���ɣ�������������
    '       ����ĸ�������������ַ�������µļ�⣬isnumbericֻ���ж����֡�
    '��������SSC���ƣ�
    '       strAsk
    '���أ�
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
    '���ܣ��ж�ָ���ַ����Ƿ�ȫ����Ӣ����ĸ����    '
    '������
    '       strAsk
    '���أ�
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
    '���ܣ��ж�ָ���ַ����Ƿ��к���
    '������
    '       strAsk
    '���أ�
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
    '���ܣ���ָ���Ľ����ֵת��Ϊ��д�����ʾ
    '������
    '       curMoney:��Ҫת���Ľ����ֵ
    '���أ�
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
'���ܣ�ȡ�ຯ��������vb��Mod������������λ���������ã�����ʱ����
'������
'       X1-������X2-������
'���أ�
'-------------------------------------------------------------
    ZyMod = x1 - Int(x1 / x2) * x2
End Function

Public Function SelOper(ByVal frmParent As Object) As String
'-------------------------------------------------------------
'���ܣ���ʾ��Աѡ�����������Ž���ѡ��
'-------------------------------------------------------------
    Dim frmTemp As New frmManSelect
    frmTemp.Show vbModal, frmParent
    SelOper = frmTemp.mstrReturn
    
    Set frmTemp = Nothing
End Function

Public Function OpenIme(Optional blnOpen As Boolean = False) As Boolean
    
    '����:���������뷨����ر����뷨
    Dim arrIme(99) As Long, lngCount As Long, strName As String * 255
    Dim strIme As String, blnNotCloseIme As Boolean
     
    '�û�û�������ã��Ͳ�����
    strIme = gobjComLib.zlDatabase.GetPara("���뷨")
    If strIme = "" And blnOpen = True Then Exit Function                 'Ҫ������뷨��������û������
    
    lngCount = GetKeyboardLayoutList(UBound(arrIme) + 1, arrIme(0))
    blnNotCloseIme = True
    Do
        lngCount = lngCount - 1
        If ImmIsIME(arrIme(lngCount)) = 1 Then
            If blnOpen = True Then
                '��Ҫ�����뷨�������ж��Ƿ��������뷨
                ImmGetDescription arrIme(lngCount), strName, Len(strName)
                If InStr(1, Mid(strName, 1, InStr(1, strName, Chr(0)) - 1), strIme) > 0 And strIme <> "" Then
                    If ActivateKeyboardLayout(arrIme(lngCount), 0) <> 0 Then OpenIme = True: Exit Function
                End If
            End If
        ElseIf blnOpen = False Then
            '�������뷨��������Ӧ�˹ر����뷨������
            If ActivateKeyboardLayout(arrIme(lngCount), 0) <> 0 Then OpenIme = True: Exit Function
        End If
    Loop Until lngCount = 0
    
    If blnNotCloseIme And blnOpen = False Then
        '����windows Vistaϵͳ��Ӣ�����뷨��ImmIsIME���Գ���true�����뷨,���,��Ҫ��������.
        '���˺�:2008/09/03
        If ActivateKeyboardLayout(arrIme(0), 0) <> 0 Then OpenIme = True: Exit Function
    End If
    
End Function

Public Function GetLike(ByVal strTable As String, ByVal strField As String, ByVal strInput As String) As String
'���ܣ�����ָ���ֶε�Like�����Ӿ�
'������strTable     ��������SQL�����أ�����Ϊ����
'      strField     �ֶ���
'      strInput     ������ַ���
    Dim strMatch  As String
    
    strMatch = IIf(Val(gobjComLib.zlDatabase.GetPara("����ƥ��")) = 0, "%", "")
    strInput = UCase(Replace(Trim(strInput), "'", "''"))
    
    GetLike = "Upper(" & IIf(strTable = "", "", strTable & ".") & strField & _
                   ") LIKE '" & strMatch & strInput & "%'"
End Function

Public Function ToVarchar(ByVal varText As Variant, ByVal lngLength As Long, Optional strDefault As String = "") As String
'���ܣ����ı���Varchar2�ĳ��ȼ��㷽�����нض�
    Dim strText As String
    
    strText = IIf(IsNull(varText), strDefault, varText)
    ToVarchar = StrConv(LeftB(StrConv(strText, vbFromUnicode), lngLength), vbUnicode)
    'ȥ�����ܳ��ֵİ���ַ�
    ToVarchar = Replace(ToVarchar, Chr(0), "")
End Function

Public Function AddDate(ByVal strOrin As String, Optional ByVal blnʱ As Boolean = False) As String
'���ܣ�Ϊ��ȫ��������Ϣ��������
'������strOrin  �û������ԭʼֵ
'      blnʱ    �Ƿ�����Сʱ����
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
            'ʹ�� . ��
            strTemp = Replace(strTemp, ".", "-")
        End If
    End If
    
    If intPos = 0 Then
        'û��"-",�ֹ�����
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
        If blnʱ = False Then
            If IsDate(strTemp) Then
                strTemp = Format(CDate(strTemp), "yyyy-MM-dd")
            End If
        Else
            '����Сʱ
            If InStr(strTemp, " ") > 0 Then
                '������Сʱ
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
'���ܣ��൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
    NVL = IIf(IsNull(varValue), DefaultValue, varValue)
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

Public Function OpenDir(hwndOwner As Long, Optional strTitle As String) As String
'���ܣ�ѡ��Windows���ļ���
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
'���ܣ��ж�ָ���ı����е�ǰ�����Ƿ���ˢ��(�Ƿ�ﵽ���ų��ȣ��ڵ��ó������ж�),������ϵͳ���������Ƿ�������ʾ
'������KeyAscii=��KeyPress�¼��е��õĲ���
    Static sngInputBegin As Single
    Dim sngNow As Single, blnCard As Boolean, strText As String
        
    'ˢ��ʱ����������ŵ�Ҫȡ��
    If InStr(":��;��?��", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Function
                    
    '������ǰ�������ʾ������(��δ��ʾ����)
    strText = txtInput.Text
    If txtInput.SelLength = Len(txtInput.Text) Then strText = ""
    If KeyAscii = 8 Then
        If strText <> "" Then strText = Mid(strText, 1, Len(strText) - 1)
    Else
        strText = strText & Chr(KeyAscii)
    End If
        
    '�ж��Ƿ���ˢ��
    If IsNumeric(strText) And IsNumeric(Left(strText, 1)) Then  '�����������������ȫ���֣���Ϊ��ˢ��
        blnCard = True
    ElseIf KeyAscii > 32 Then
        sngNow = Timer
        If txtInput.Text = "" Or strText = "" Then
            sngInputBegin = sngNow
        Else
            If Format((sngNow - sngInputBegin) / Len(strText), "0.000") < 0.04 Then blnCard = True   '��һ̨�ʼǱ����ԣ�һ����0.014����
        End If
    End If
    
    'ˢ��ʱ�����Ƿ�������ʾ
    If blnCard Then
        txtInput.PasswordChar = IIf(gobjComLib.zlDatabase.GetPara(12, lngSys) = "0", "", "*")
    Else
        txtInput.PasswordChar = ""
    End If
    
    InputIsCard = blnCard
End Function


Public Function IncStr(ByVal strVal As String) As String
'���ܣ���һ���ַ����Զ���1��
'˵����ÿһλ��λʱ,���������,��ʮ���ƴ���,����26���ƴ���
'������strVal=Ҫ��1���ַ���
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
'���ܣ��ӱ���������ϴ��з��������
'������strList=����������ϴ�,��"012-�ڿ�","(012)�ڿ�","[012]�ڿ�"
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
'���ܣ�����ZLHIS�ĵ��ݺ�ǰ׺����,���ش�д�ĵ��ݺ���ǰ׺
'������curDate=���ݺ�ǰ׺��׼ʱ��,����ʱ��׼ʱ��Ϊ������ʱ��
'Ӧ�ã��ú�����Ҫ�����GetFullNO����Ӧ��

    If curDate = #1/1/1900# Then
        PreFixNO = CStr(CInt(Format(gobjComLib.zlDatabase.CurrentDate, "YYYY")) - 1990)
    Else
        PreFixNO = CStr(CInt(Format(curDate, "YYYY")) - 1990)
    End If
    PreFixNO = IIf(CInt(PreFixNO) < 10, PreFixNO, Chr(55 + CInt(PreFixNO)))
End Function

Public Function GetFullNO(strNO As String) As String
'���ܣ����û�����Ĳ��ݵ��ţ����ص�����������ݺš�
'������strNO=���û�����ĵ��ݺŲ���

    If Len(strNO) >= 8 Then GetFullNO = Right(strNO, 8): Exit Function
    GetFullNO = PreFixNO & Format(strNO, "0000000")
End Function

Public Function GetIDCardDate(strCardID As String) As String
'���ܣ���������֤�ŷ��س�������
'������ID=����֤��,Ӧ��Ϊ15λ��18λ
'���أ���ʽ"yyyy-MM-dd"
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
'���ܣ��ж��Ƿ�WindowNT����ϵͳ
    Const dwMaskNT = &H2&

    IsWindowsNT = (GetWinPlatform() And dwMaskNT)
End Function

Public Function IsWindows9X() As Boolean
'���ܣ��ж��Ƿ�Window95����ϵͳ
    Const dwMask95 = &H1&

    IsWindows9X = (GetWinPlatform() And dwMask95)
End Function

Public Sub SetWindowsInTaskBar(ByVal lngHwnd As Long, ByVal blnShow As Boolean)
'���ܣ����ô����Ƿ�������������ʾ
    Dim lngStyle As Long
    
    lngStyle = GetWindowLong(lngHwnd, GWL_EXSTYLE)
    If blnShow Then
        lngStyle = lngStyle Or &H40000
    Else
        lngStyle = lngStyle And Not &H40000
    End If
    Call SetWindowLong(lngHwnd, GWL_EXSTYLE, lngStyle)
End Sub

Public Function StrIsValid(ByVal strInput As String, Optional ByVal intMax As Integer = 0, Optional ByVal hwnd As Long = 0, Optional str��Ŀ As String) As Boolean
'����ַ����Ƿ��зǷ��ַ�������ṩ���ȣ��Գ��ȵĺϷ���Ҳ����⡣
    If str��Ŀ = "" Then str��Ŀ = "����������"
    
    If InStr(strInput, "'") > 0 Or InStr(strInput, "|") > 0 Then
        MsgBox str��Ŀ & "���зǷ��ַ���", vbExclamation, gstrSysName
        If hwnd <> 0 Then SetFocusHwnd hwnd              '���ý���
        Exit Function
    End If
    If intMax > 0 Then
        If LenB(StrConv(strInput, vbFromUnicode)) > intMax Then
            MsgBox str��Ŀ & "���ܳ���" & Int(intMax / 2) & "������" & "��" & intMax & "����ĸ��", vbExclamation, gstrSysName
            If hwnd <> 0 Then SetFocusHwnd hwnd              '���ý���
            Exit Function
        End If
    End If
    
    StrIsValid = True
End Function

Public Function ShowMsgBox(ByVal strCaption As String, ByVal strInfo As String, ByVal strCmds As String, _
    frmParent As Object, Optional vStyle As VbMsgBoxStyle = vbQuestion) As String
'������strCaption=��Ϣ�������
'      strInfo=������ʾ����,����"^"��ʾ����,">"��ʾ������
'      strCmds=��ť����,��"����(&R),!����(&A),?ȡ��(&C)"
'              ����Ҫ��������ť,"!"��ʾȱʡ��ť,"?"��ʾȡ����ť
'              ÿ����ť�������֧��4������
'      vStyle=vbInformation,vbQuestion,vbExclamation,vbCritical
'���أ���ť����,��"��ť2"(������()��&),������رջ�ȡ���򷵻�""
    ShowMsgBox = frmMsgBox.ShowMsgBox(strCaption, strInfo, strCmds, frmParent, vStyle)
End Function

Public Function VerifyPassWord(frmParent As Object, ByVal strPass As String, Optional ByVal strName As String, Optional ByVal strSex As String, Optional ByVal strOld As String) As Boolean
'���ܣ������������֤
'������frmParent=��ʾ�ĸ�����
'      strPass=��ȷ������
'      strName,strSex,strOld=��ѡ�����������������Ա����䣬��������ʱ����ʾ�������
'���أ�True=������֤ͨ��,False=ȡ�����룬������3��������������
    VerifyPassWord = frmVerifyPassword.ShowMe(frmParent, strPass, strName, strSex, strOld)
End Function

Public Sub DeviceSetup(frmParent As Object, ByVal lngSys As Long, ByVal lngModule As Long)
'����:���ݲ�ͬ��ϵͳ��ģ��,���ò�ͬ���豸����
    
    Call frmDeviceSetup.ShowMe(frmParent, lngSys, lngModule)
End Sub

Public Sub ShowTipInfo(ByVal lngHwnd As Long, ByVal strInfo As String, Optional blnMultiRow As Boolean)
'���ܣ���ʾ����������ʾ
'������lngHwnd=��ʾ����ԵĿؼ����,������Ϊ0ʱ������ʾ
'      strInfo=��ʾ��Ϣ,������Ϊ��ʱ������ʾ
'      blnMultiRow=��һ���ļ�������ʾ������Ϣ��ÿ�а�vbcrlf�ָ�

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