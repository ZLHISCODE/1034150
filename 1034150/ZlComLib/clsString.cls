VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'##############################################################################
'#���ܣ��ַ�����صĴ������жϵȹ�������
'�������򣺺�����������Ϊͨ���׶���Ӣ�ļ��ɣ����ü�Str�ؼ���
'�ϼ��ࣺclsComLib
'##############################################################################

Public Function CheckCharScope(ByVal strText As String, ByVal strScopeChars As String) As Boolean
'���ܣ�����ַ����Ƿ�ֻ����ָ�����ַ�
    Dim i As Integer
    
    For i = 1 To Len(strText)
        If InStr(strScopeChars, Mid(strText, i, 1)) = 0 Then Exit Function
    Next
    CheckCharScope = True
End Function

Public Function ListMinus(ByVal strLeft As String, ByVal strExists As String, Optional ByVal strSplit As String = ",") As String
'���ܣ�����"������Ŀ��"����"������Ŀ��"�в��ظ�����Ŀ��
    Dim arrIDs As Variant, strResult As String, i As Long
    
    arrIDs = Split(strLeft, strSplit)
    For i = 0 To UBound(arrIDs)
        If InStr(strSplit & strExists & strSplit, strSplit & arrIDs(i) & strSplit) = 0 _
            And InStr(strResult & strSplit, strSplit & arrIDs(i) & strSplit) = 0 Then
            strResult = strResult & strSplit & arrIDs(i)
        End If
    Next
    
    ListMinus = Mid(strResult, Len(strSplit) + 1)
End Function

Public Function TrimEx(ByVal strTrim As String, Optional ByVal strTrmChar As String = " ") As String
'���ܣ�ȥ��strTrim���ߵ�strTrmChar,��������Trim
'         ����strTrmChar���ߴ��ո�ʱ���൱Trim
    Dim i As Integer, intB As Integer, intE As Integer
    
    If strTrim = "" Or strTrmChar = "" Then TrimEx = strTrim: Exit Function
    If strTrmChar = " " Then TrimEx = Trim(strTrim): Exit Function
    
    intB = 1
    For i = 1 To Len(strTrim)
        If Mid(strTrim, i, 1) <> strTrmChar Then intB = i: Exit For
    Next
    intE = Len(strTrim)
    For i = Len(strTrim) To 1 Step -1
        If Mid(strTrim, i, 1) <> strTrmChar Then intE = i: Exit For
    Next
    TrimEx = Mid(strTrim, intB, intE - intB + 1)
End Function

Public Function Encode(ByVal strPutString As String) As String
'����:�ַ�������
'���:strPutString-��Ҫ���ܵĴ�
'����:���ܴ�
    Encode = gobjComLib.zlCommFun.zlStringEncode(strPutString)
End Function

Public Function NeedCode(strList As String, Optional ByVal strSplit As String) As String
'���ܣ��ӱ���������ϴ��з��������
'������strList=����������ϴ�,��"012-�ڿ�","(012)�ڿ�","[012]�ڿ�"
'          strSplit=ָ���ı������Ʒָ����û��ָ������Ĭ�����ȼ����н���,������ֻ�������»������������м�ָ��
'˵��:1-strList��()��[]�ָ����������ʱ��������[����]��(����)��ͷ,�������Ϊ���ֻ���ĸ
'     2-�ָ��������ȼ����س���(Chr(13)��>����(Chr(10))> - > [] > ()
    Dim intType As Integer
    
    If strList = "" Then Exit Function
    intType = gobjComLib.Decode(strSplit, "", 0, Chr(13), 1, Chr(10), 2, "-", 3, "[]", 4, "()", 5, 6)
    If intType = 0 Or intType = 1 Then
        '�����ж��Իس����ָ�
        If InStr(strList, Chr(13)) > 0 Then
            NeedCode = LTrim(Mid(strList, 1, InStr(strList, Chr(13)) - 1))
            Exit Function
        End If
    End If
    
    If intType = 0 Or intType = 2 Then
        '�Ի��з��ָ�
        If InStr(strList, Chr(10)) > 0 Then
            NeedCode = Trim(Mid(strList, InStr(strList, Chr(10)) - 1))
            Exit Function
        End If
    End If
    
    If intType = 0 Or intType = 4 Then
        '��[]�ָ�
        If InStr(strList, "]") > 0 And InStr(strList, "-") = 0 And Left(LTrim(strList), 1) = "[" Then
            If IsNumOrChar(Mid(strList, 2, InStr(strList, "]") - 2)) Then
                NeedCode = Trim(Mid(strList, 2, InStr(strList, "]") - 2))
                Exit Function
            End If
        End If
    End If
    
    If intType = 0 Or intType = 5 Then
        '��()�ָ�
        If InStr(strList, ")") > 0 And InStr(strList, "-") = 0 And Left(LTrim(strList), 1) = "(" Then
            If IsNumOrChar(Mid(strList, 2, InStr(strList, ")") - 2)) Then
                NeedCode = Trim(Mid(strList, 2, InStr(strList, ")") - 2))
                Exit Function
            End If
        End If
    End If
    If intType = 0 Or intType = 3 Then
        '��-�ָ�
        If InStr(strList, "-") > 0 Then
            NeedCode = Trim(Mid(strList, 1, InStr(strList, "-") - 1))
        End If
    Else
        If InStr(strList, strSplit) > 0 And strSplit <> "" Then
            NeedCode = Trim(Mid(strList, 1, InStr(strList, strSplit) - 1))
        End If
    End If
End Function

Public Function NeedName(strList As String, Optional ByVal strSplit As String) As String
'���ܣ��ӱ���������ϴ��з��������
'������strList=����������ϴ�,��"012-�ڿ�","(012)�ڿ�","[012]�ڿ�"
'          strSplit=ָ���ı������Ʒָ����û��ָ������Ĭ�����ȼ����н���,������ֻ�������»������������м�ָ��
'˵��:1-strList��()��[]�ָ����������ʱ��������[����]��(����)��ͷ,�������Ϊ���ֻ���ĸ
'     2-�ָ��������ȼ����س���(Chr(13)��>����(Chr(10))> - > [] > ()
    NeedName = gobjComLib.zlCommFun.GetNeedName(strList, strSplit)
End Function

Public Function To_Date(ByVal strDate As String, Optional ByVal strType As String = "YMDHMS") As String
'���ܣ���ȡORACLE Date���ʹ�
'������strDate=ʱ���ַ���
'         strType=��ʽ�ַ������ͣ�ymd-�����գ�yyyy-mm-dd)��ymdhm-��yyyy-mm-dd hh:mm),ymdhms-��yyyy-mm-dd hh:mm:ss)
'���أ�ORACLE Date���ʹ�
    If Not IsDate(strDate) Then To_Date = "Null": Exit Function
    Select Case UCase(strType)
        Case "YMD"
           To_Date = "To_Date('" & Format(strDate, "yyyy-MM-dd") & "','YYYY-MM-DD')"
        Case "YMDHM"
           To_Date = "To_Date('" & Format(strDate, "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI')"
        Case "YMDHMS"
           To_Date = "To_Date('" & Format(strDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
        Case Else
           To_Date = "Null"
    End Select
End Function

Public Function SubB(ByVal strInfor As String, ByVal lngStart As Long, ByVal lngLen As Long) As String
'����:��ȡָ���ִ���ֵ,�ִ��п��԰�������
 '���:strInfor-ԭ��
 '         lngStart-ֱʼλ��
'         lngLen-����
'����:�Ӵ�
    Dim strTmp As String, i As Long
    Err = 0: On Error GoTo errH:
    SubB = StrConv(MidB(StrConv(strInfor, vbFromUnicode), lngStart, lngLen), vbUnicode)
    SubB = Replace(SubB, Chr(0), "")
    Exit Function
errH:
    Err.Clear
    SubB = ""
End Function

Public Function SetBit(ByVal StrBit As String, ByVal intBit As Integer, Optional ByVal intVal As Integer = -1) As String
'���ܣ���ָ��λ�ַ���strBit�еĵ�intBitλ����Ϊ0��1
'������intVal=����ֵ,0��1,������ʾ��ת
    If intVal = -1 Then intVal = IIf(Val(Mid(StrBit, intBit, 1)) = 0, 1, 0)
    SetBit = Left(StrBit, intBit - 1) & intVal & Mid(StrBit, intBit + 1)
End Function

Public Function PrefixNO(Optional curDate As Date = #1/1/1900#) As String
'���ܣ�����ZLHIS�ĵ��ݺ�ǰ׺����,���ش�д�ĵ��ݺ���ǰ׺
'������curDate=���ݺ�ǰ׺��׼ʱ��,����ʱ��׼ʱ��Ϊ������ʱ��
'Ӧ�ã��ú�����Ҫ�����GetFullNO����Ӧ��

    If curDate = #1/1/1900# Then
        PrefixNO = CStr(CInt(Format(gobjComLib.zlDatabase.Currentdate, "YYYY")) - 1990)
    Else
        PrefixNO = CStr(CInt(Format(curDate, "YYYY")) - 1990)
    End If
    PrefixNO = IIf(CInt(PrefixNO) < 10, PrefixNO, Chr(55 + CInt(PrefixNO)))
End Function

Public Function CPAD(ByVal strText As String, ByVal intCount As Integer, Optional ByVal StrPAD As String = " ", Optional ByVal blnAutoSub As Boolean) As String
'���ܣ�ʹ�ַ������У�����������ַ���
'������
'       strText:����ַ���
'       intCount:����ĳ���
'       StrPAD:�����ַ�
'       blnAutoSub:�ַ����������Զ���ȡ
'���أ���������ַ���
'˵����һ�����ֵ��������ַ����ȴ���
    Dim lngTmp As Long, lngFill As Long
    If StrPAD = "" Then
        StrPAD = " "
    Else
        StrPAD = Mid(StrPAD, 1, 1)
    End If
    
    lngFill = ActualLen(StrPAD)
    lngTmp = ActualLen(strText)
    If lngTmp <= intCount - lngFill * 2 Then
        CPAD = String((intCount - lngTmp) \ (lngFill * 2), StrPAD) & strText & String((intCount - lngTmp) \ (lngFill * 2), StrPAD)
    ElseIf lngTmp > intCount And blnAutoSub Then
        CPAD = SubB(strText, 1, intCount)
    Else
        CPAD = strText
    End If
End Function

Public Function RPAD(ByVal strText As String, ByVal intCount As Integer, Optional ByVal StrPAD As String = " ", Optional ByVal blnAutoSub As Boolean) As String
'���ܣ���ͬOracle��RPAD����
'����:��ָ���������ƿո�
 '������
 '       strText:����ַ���
 '       intCount:����ĳ���
 '       StrPAD:�����ַ�
 '       blnAutoSub:�ַ����������Զ���ȡ
'����:�����ִ�
   
    Dim lngTmp As Long, lngFill As Long
    If StrPAD = "" Then
        StrPAD = " "
    Else
        StrPAD = Mid(StrPAD, 1, 1)
    End If
    
    lngFill = ActualLen(StrPAD)
    lngTmp = ActualLen(strText)
    If lngTmp <= intCount - lngFill Then
        RPAD = strText & String((intCount - lngTmp) \ lngFill, StrPAD)
    ElseIf lngTmp > intCount And blnAutoSub Then
        RPAD = SubB(strText, 1, intCount)
    Else
        RPAD = strText
    End If
End Function

Public Function LPAD(ByVal strText As String, ByVal intCount As Integer, Optional ByVal StrPAD As String = " ", Optional ByVal blnAutoSub As Boolean) As String
'���ܣ���ͬOracle��LPAD����
 '����:��ָ���������ƿո�
 '������
 '  strText:����ַ���
 '  intCount:����ĳ���
 '  StrPAD:�����ַ�
 '  blnAutoSub:�ַ����������Զ���ȡ
 '����:�����ִ�
 
    Dim lngTmp As Long, lngFill As Long
    If StrPAD = "" Then
        StrPAD = " "
    Else
        StrPAD = Mid(StrPAD, 1, 1)
    End If
    lngFill = ActualLen(StrPAD)
    lngTmp = ActualLen(strText)
    If lngTmp <= intCount - lngFill Then
        LPAD = String((intCount - lngTmp) \ lngFill, StrPAD) & strText
    ElseIf lngTmp > intCount And blnAutoSub Then
        LPAD = SubB(strText, 1, intCount)
    Else
        LPAD = strText
    End If
End Function

Public Function IsNumOrChar(ByVal strAsk As String) As Boolean
'���ܣ��ж�ָ���ַ����Ƿ�ȫ�������ֺ�Ӣ����ĸ���ɣ�������������
'       ����ĸ�������������ַ�������µļ�⣬isnumbericֻ���ж����֡�
'��������SSC���ƣ�strAsk
'���أ�
    IsNumOrChar = gobjComLib.zlCommFun.IsNumOrChar(strAsk)
End Function

Public Function IsHavePrivs(ByVal strPrivs As String, ByVal strMyPriv As String) As Boolean
'����:���ָ����Ȩ���Ƿ����
'����:strPrivs-Ȩ�޴�
'     strMyPriv-����Ȩ��
'����,����Ȩ��,����true,���򷵻�False
    IsHavePrivs = InStr(";" & strPrivs & ";", ";" & strMyPriv & ";") > 0
End Function

Public Function IsCharChinese(ByVal strAsk As String) As Boolean
'���ܣ��ж�ָ���ַ����Ƿ��к���
'������strAsk
'���أ�
    IsCharChinese = gobjComLib.zlCommFun.IsCharChinese(strAsk)
End Function

Public Function IsCharAlpha(ByVal strAsk As String) As Boolean
'���ܣ��ж�ָ���ַ����Ƿ�ȫ����Ӣ����ĸ����    '
'������ strAsk
'���أ�
    IsCharAlpha = gobjComLib.zlCommFun.IsCharAlpha(strAsk)
End Function

Public Function Increase(ByVal strVal As String, Optional ByVal blnDown As Boolean, Optional ByRef strErr As String) As String
'���ܣ���һ���ַ����Զ���1��
'˵����ÿһλ��λʱ,���������,��ʮ���ƴ���,����26���ƴ���
'������strVal=Ҫ��1���ַ���
'         blnDown=True ,��һ��False,��һ
    Increase = gobjComLib.zlCommFun.IncStr(strVal, blnDown, strErr)
End Function

Public Function FullDate(ByVal strText As String, Optional blnTime As Boolean = True, Optional ByVal strMintime As String, Optional strMaxtTime As String) As String
'���ܣ�������������ڼ�,�������������ڴ�(yyyy-MM-dd[ HH:mm])
'������blnTime=�Ƿ���ʱ�䲿��
'������strMintime=����ʱ�������
'          strOutTime=����ʱ�������
    Dim curDate As Date, strTmp As String
    
    If strText = "" Then Exit Function
    curDate = gobjComLib.zlDatabase.Currentdate
    strTmp = strText
    
    If InStr(strTmp, "-") > 0 Or InStr(strTmp, "/") Or InStr(strTmp, ":") > 0 Then
        '���봮�а������ڷָ���
        If IsDate(strTmp) Then
            strTmp = Format(strTmp, "yyyy-MM-dd HH:mm")
            If Right(strTmp, 5) = "00:00" And InStr(strText, ":") = 0 Then
                'ֻ���������ڲ���
                strTmp = Mid(strTmp, 1, 11) & Format(curDate, "HH:mm")
            ElseIf Left(strTmp, 10) = "1899-12-30" Then
                'ֻ������ʱ�䲿��
                strTmp = Format(curDate, "yyyy-MM-dd") & Right(strTmp, 6)
            End If
        Else
            '����Ƿ�����,����ԭ����
            strTmp = strText
        End If
    Else
        '���������ڷָ���
        If Len(strTmp) <= 2 Then
            '��������dd
            strTmp = Format(strTmp, "00")
            strTmp = Format(curDate, "yyyy-MM") & "-" & strTmp & " " & Format(curDate, "HH:mm")
        ElseIf Len(strTmp) <= 4 Then
            '��������MMdd
            strTmp = Format(strTmp, "0000")
            strTmp = Format(curDate, "yyyy") & "-" & Left(strTmp, 2) & "-" & Right(strTmp, 2) & " " & Format(curDate, "HH:mm")
        ElseIf Len(strTmp) <= 6 Then
            '��������yyMMdd
            strTmp = Format(strTmp, "000000")
            strTmp = Format(Left(strTmp, 2) & "-" & Mid(strTmp, 3, 2) & "-" & Right(strTmp, 2), "yyyy-MM-dd") & " " & Format(curDate, "HH:mm")
        ElseIf Len(strTmp) <= 8 Then
            '��������MMddHHmm
            strTmp = Format(strTmp, "00000000")
            strTmp = Format(curDate, "yyyy") & "-" & Left(strTmp, 2) & "-" & Mid(strTmp, 3, 2) & " " & Mid(strTmp, 5, 2) & ":" & Right(strTmp, 2)
            If Not IsDate(strTmp) Then
                '��������yyyyMMdd
                strTmp = Format(strText, "00000000")
                strTmp = Left(strTmp, 4) & "-" & Mid(strTmp, 5, 2) & "-" & Right(strTmp, 2) & " " & Format(curDate, "HH:mm")
            End If
        Else
            '��������yyyyMMddHHmm
            strTmp = Format(strTmp, "000000000000")
            strTmp = Left(strTmp, 4) & "-" & Mid(strTmp, 5, 2) & "-" & Mid(strTmp, 7, 2) & " " & Mid(strTmp, 9, 2) & ":" & Right(strTmp, 2)
        End If
    End If
    
    If IsDate(strTmp) Then
        If strMintime <> "" Then
            If Format(strTmp, "yyyy-MM-dd HH:mm") < Format(strMintime, "yyyy-MM-dd HH:mm") Then
                strTmp = strMintime
            End If
        End If
        If strMaxtTime <> "" Then
            If Format(strTmp, "yyyy-MM-dd HH:mm") > Format(strMaxtTime, "yyyy-MM-dd HH:mm") Then
                strTmp = strMaxtTime
            End If
        End If
        If Not blnTime Then
            strTmp = Format(strTmp, "yyyy-MM-dd")
        End If
    End If
    FullDate = strTmp
End Function

Public Function ActualLen(ByVal strAsk As String) As Long
'���ܣ���ȡָ���ַ�����ʵ�ʳ��ȣ������ж�ʵ�ʰ���˫�ֽ��ַ�����
'       ʵ�����ݴ洢����
    ActualLen = gobjComLib.zlCommFun.ActualLen(strAsk)
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

Public Function TruncZero(ByVal strInput As String) As String
'���ܣ�ȥ���ַ�����\0�Ժ���ַ�
    TruncZero = gobjComLib.zlCommFun.TruncZero(strInput)
End Function

Public Function ToVarchar(ByVal varText As Variant, ByVal lngLength As Long, Optional strDefault As String = "") As String
'���ܣ����ı���Varchar2�ĳ��ȼ��㷽�����нض�
    ToVarchar = gobjComLib.zlCommFun.ToVarchar(varText, lngLength, strDefault)
End Function

Public Function FormatString(ByVal strFormat As String, ParamArray arrInput() As Variant) As String
'���ܣ��ò������������strFormat�е�[x]ռλ����X>=1
'������
'  strFormat������ʽ��[1..n]Ϊ�����Źؼ��֣����ӣ�"����ֵΪ��[1]"
'  arrInput������ʽ�Ĳ�������ӦstrFormat�еĲ����Źؼ���
'���أ���ʽ������ַ���

    Dim strPar As String, arrPar As Variant, arrTmp As Variant
    Dim lngLeft As Long, lngRight As Long
    Dim strSeq As String, intPar As Integer, i As Integer
    Dim strReturn As String, strTmp As String
    
    FormatString = strFormat
    If Len(strFormat) > 60000 Then Exit Function
    If UBound(arrInput) < 0 Then Exit Function
    On Error GoTo errH
    '�����Զ���[x]����
    lngLeft = InStr(1, strFormat, "[")
    Do While lngLeft > 0
        lngRight = InStr(lngLeft + 1, strFormat, "]")
        If lngRight = 0 Then Exit Do
        strSeq = Mid(strFormat, lngLeft + 1, lngRight - lngLeft - 1)
        If IsNumeric(strSeq) Then
            i = CInt(strSeq)
            If i > 0 And i - 1 <= UBound(arrInput) Then '��Ч����
                strPar = strPar & ";" & i & "," & lngLeft & "," & lngRight
                If i > intPar Then intPar = i
            End If
        End If
        lngLeft = InStr(lngRight + 1, strFormat, "[")
    Loop
    '��װ�ַ���
    lngLeft = 0: lngRight = 0
    arrPar = Split(Mid(strPar, 2), ";")
    strReturn = ""
    For i = LBound(arrPar) To UBound(arrPar)
        arrTmp = Split(arrPar(i), ","): intPar = Val(arrTmp(0)): lngLeft = Val(arrTmp(1))
        strReturn = strReturn & Mid(strFormat, lngRight + 1, lngLeft - lngRight - 1) & arrInput(intPar - 1)
        lngRight = Val(arrTmp(2))
        If i = UBound(arrPar) Then strReturn = strReturn & Mid(strFormat, lngRight + 1)
    Next
    FormatString = strReturn
    Exit Function
errH:
    Err.Clear
End Function

Public Function ZVal(ByVal varValue As Variant, Optional ByVal blnForceNum As Boolean, Optional ByVal varDefault As Variant = 0) As String
'���ܣ���0��ת��Ϊ"NULL"��,������SQL���ʱ��
'������blnForceNum=��ΪNullʱ���Ƿ�ǿ�Ʊ�ʾΪ������
    Dim varTmp As Variant
    varTmp = IIf(Val(varValue & "") = 0, varDefault, varValue) & ""
    ZVal = IIf(Val(varTmp) = 0, IIf(blnForceNum, "-NULL", "NULL"), Val(varTmp))
End Function

Public Function RoundEx(ByVal dblNumber As Double, ByVal intBit As Integer) As Double
'���ܣ��������뷽ʽ��ʽ������
'������intBit=���С��λ��
'����ţ�94552
'˵����VB�Դ���Round�����м����뷨,��ʵ�ʲ�һ�¡���Round(57.575,2)=57.58,Round(57.565,2)=57.56
    If intBit >= 0 Then
        RoundEx = Val(Format(dblNumber, "0" & IIf(intBit = 0, "", ".") & String(intBit, "0")))
    Else
        RoundEx = dblNumber
    End If
End Function

Public Function NVL(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'clsCommFun���ڸú���
'���ܣ��൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
    NVL = gobjComLib.NVL(varValue, DefaultValue)
End Function

Public Function VerCompare(ByVal strVerCur As String, Optional ByVal strVerCom As String) As Integer
'���ܣ��Ƚ������汾��,�ȵ�ǰ�汾��С������1����ȷ���0���ȵ�ǰ�汾�Ŵ󷵻�-1
'������strVerCur=��ǰ�汾��
'         strVerCom=�Աȵİ汾��
' ���أ��ԱȰ汾�űȵ�ǰ�汾��С������1����ȷ���0���ȵ�ǰ�汾�Ŵ󷵻�-1
    If VerFull(strVerCur) < VerFull(strVerCom) Then
        VerCompare = -1
    ElseIf VerFull(strVerCur) > VerFull(strVerCom) Then
        VerCompare = 1
    Else
        VerCompare = 0
    End If
End Function

Public Function VerFull(ByVal strVer As String, Optional ByVal blnMax As Boolean) As String
'���ܣ�����VB���֧�ֵİ汾����ʽ:9999.9999.9999,��С�汾��0000.0000.0000
'������strVer=��ǰ�汾��
'           blnMax=True,����Ϊ�գ��򷵻����֧�ְ汾��False=����Ϊ�գ��򷵻���С֧�ְ汾
    Dim arrVer As Variant
    If Not IsVersion(strVer) Then
        VerFull = IIf(blnMax, "9999.9999.9999", "0000.0000.0000")
        Exit Function
    End If
    arrVer = Split(strVer, ".")
    VerFull = Format(arrVer(0), "0000") & "." & Format(arrVer(1), "0000") & "." & Format(arrVer(2), "0000")
End Function

Public Function IsVersion(ByVal strVer As String) As Boolean
'���ܣ��ж��ַ����Ƿ��ǰ汾��
    Dim arrVer As Variant
    Dim i As Integer
    If Not strVer Like "*.*.*" Then Exit Function
    arrVer = Split(strVer, ".")
    If UBound(arrVer) <> 2 Then Exit Function
    
    For i = LBound(arrVer) To UBound(arrVer)
        If Not IsNumeric(arrVer(i)) Then Exit Function
        If Val(arrVer(i)) < 0 Or Val(arrVer(i)) > 9999 Then Exit Function
        If Val(arrVer(i)) & "" <> arrVer(i) Then Exit Function
    Next
    
    IsVersion = True
End Function

Public Function PinYinCode(ByVal strAsk As String) As String
'���ܣ�����ָ���ַ�����ƴ������
'������strAsk  ���������ַ���
    Dim aryStard As Variant
    Dim intBit As Integer, iCount As Integer
    Dim strCode As String, StrBit As String

'    aryStard = Split("��;��;��;��;��;��;��;��;;��;��;��;��;��;ž;��;��;��;��;��;;��;��;Ѿ;��", ";")
    aryStard = Split("��;��;��;��;��;�;��;��;;��;��;��;��;��;ž;��;��;��;��;��;;;��;Ѿ;��", ";")
    strAsk = StrConv(Trim(strAsk), vbNarrow + vbUpperCase)         '��ȫ��ת��Ϊ��ǣ�Сдת��Ϊ��д
    
    strCode = ""
    For intBit = 1 To Len(strAsk)
        StrBit = Mid(strAsk, intBit, 1)
        If InStr(1, "��������������������¦���ſ������Ϧϫ�������������������������������", StrBit) > 0 Then
            '�����ֵĴ���
            strCode = strCode & Switch(StrBit = "��", "1", StrBit = "��", "2", StrBit = "��", "3", StrBit = "��", "4", StrBit = "��", "5" _
                            , StrBit = "��", "6", StrBit = "��", "7", StrBit = "��", "8", StrBit = "��", "9" _
                            , StrBit = "��", "A", StrBit = "��", "B", StrBit = "��", "G" _
                            , StrBit = "��", "N", StrBit = "ſ", "P", StrBit = "��", "S", StrBit = "��", "W" _
                            , StrBit = "��", "W", StrBit = "Ϧ", "X", StrBit = "ϫ", "X", StrBit = "��", "S" _
                            , StrBit = "��", "X", StrBit = "��", "P", StrBit = "��", "C", StrBit = "�", "X" _
                            , StrBit = "�", "C", StrBit = "��", "D", StrBit = "��", "C", StrBit = "�", "Q" _
                            , StrBit = "��", "T", StrBit = "��", "N", StrBit = "��", "H", StrBit = "��", "D" _
                            , StrBit = "��", "P", StrBit = "��", "Q", StrBit = "��", "Q", StrBit = "��", "T")
        ElseIf Asc(StrBit) < 0 Then
            For iCount = 0 To UBound(aryStard)
                If Len(aryStard(iCount)) <> 0 Then
                    If StrComp(StrBit, aryStard(iCount), vbTextCompare) = -1 Then
                        strCode = strCode & Chr(65 + iCount)
                        Exit For
                    ElseIf iCount = UBound(aryStard) Then
                        strCode = strCode & "Z"
                    End If
                End If
            Next
        Else
            If StrBit >= "A" And StrBit <= "Z" Then
                strCode = strCode & StrBit
            End If
        End If
        '���ű��޸���������¼��100���ֽڳ��ȣ�����¼���������������һ�£��˾䲻�ڿ���
'        If Len(strCode) >= 10 Then Exit For
    Next
    PinYinCode = strCode

End Function

Public Function WBCode(ByVal strAsk As String, ByVal lng��ʽ As Long) As String
'���ܣ�����ָ���ַ���������ͼ���
'������strAsk  ���������ַ���
'      lng��ʽ 1-ȡ����ĸ��2-����ʹ���
    Static blnNotFound As Boolean
    Dim lngFile As Long, strFile As String, strReturn As String
    Dim str����� As String, str���� As String, blnǰ��ĸ As Boolean, str���� As String
    Dim intBit As Integer, StrBit As String
    
    If blnNotFound = True Then
        'wbx.txt�ļ�δ�ҵ������ܽ��б����ѯ
        Exit Function
    End If
    
    '���ļ�
    strFile = gstrAviPath
    If Right(strFile, 1) <> "\" Then strFile = strFile & "\"
'    strFile = "C:\AppSoft\"
    strFile = strFile & "wbx.txt"
    
    On Error Resume Next
    lngFile = FreeFile
    Open strFile For Input Access Read As lngFile
    If Err <> 0 Then
        blnNotFound = True
        MsgBox "δ����" & strFile & "�ļ���", vbInformation, gstrSysName
        Exit Function
    End If
    
    '�ҵ�ÿһ���ֶ�Ӧ����
    Do Until EOF(lngFile)
        Line Input #lngFile, strReturn
        If InStr(strAsk, Left(strReturn, 1)) > 0 Then
            '������жϷ����ڲ�����Ҫ��Ϊ�˼ӿ��ٶȣ���Ϊֻ���־�����һ���ж�
            If InStr(strReturn, " ") > 0 Then
                str����� = str����� & strReturn & "|"
            End If
        End If
    Loop
    Close #lngFile
    str����� = UCase(str�����)
    
    '�õ��ַ������к���
    strAsk = StrConv(Trim(strAsk), vbNarrow + vbUpperCase)         '��ȫ��ת��Ϊ��ǣ����ַ�������ת��Сд
    If lng��ʽ = 1 Then
        '������ĸ
        For intBit = 1 To Len(strAsk)
            StrBit = Mid(strAsk, intBit, 1)
            If LenB(StrConv(StrBit, vbFromUnicode)) = 2 Then
                '����
                str���� = str���� & mGet����By����(str�����, StrBit, 1)
                blnǰ��ĸ = False
            ElseIf InStr(" ,.;:", StrBit) > 0 Then
                '�ո�
                blnǰ��ĸ = False
            Else
                If blnǰ��ĸ = False And StrBit >= "A" And StrBit <= "Z" Then
                    'ֻȡһ���ַ���������ĸ
                    str���� = str���� & StrBit
                End If
                blnǰ��ĸ = True
            End If
        Next
    Else
        '����ʹ���
        For intBit = 1 To Len(strAsk)
            StrBit = Mid(strAsk, intBit, 1)
            If LenB(StrConv(StrBit, vbFromUnicode)) = 2 Then
                '����
                str���� = str���� & StrBit
            End If
        Next
        
        Select Case Len(str����)
            Case 0
            Case 1
               str���� = mGet����By����(str�����, str����, 4)
            Case 2
               str���� = mGet����By����(str�����, Mid(str����, 1, 1), 2) & mGet����By����(str�����, Mid(str����, 2, 1), 2)
            Case 3
               str���� = mGet����By����(str�����, Mid(str����, 1, 1), 1) & mGet����By����(str�����, Mid(str����, 2, 1), 1) & mGet����By����(str�����, Mid(str����, 3, 1), 2)
            Case Else
               str���� = mGet����By����(str�����, Mid(str����, 1, 1), 1) & mGet����By����(str�����, Mid(str����, 2, 1), 1) & _
                         mGet����By����(str�����, Mid(str����, 3, 1), 1) & mGet����By����(str�����, Right(str����, 1), 1)
        End Select
    End If
    
    WBCode = str����
    If Err.Number <> 0 Then Err.Clear
End Function

Private Function mGet����By����(ByVal str����� As String, ByVal str���� As String, ByVal lngLen As Long) As String
'���ܣ����ݺ��ֵõ������
    Dim lngStart As Long, lngEnd As Long
    Dim str���� As String
    
    lngStart = InStr(str�����, str����)
    If lngStart = 0 Then
        'δ�ڱ�����ҵ����ֱ���
        mGet����By���� = "Z"
        Exit Function
    End If
    
    lngEnd = InStr(lngStart, str�����, "|")
    str���� = Mid(str�����, lngStart, lngEnd - lngStart)
    mGet����By���� = Mid(Split(str����, " ")(1), 1, lngLen)
End Function

Public Function GetCodeByVB(ByVal strAsk As String) As String
'-------------------------------------------------------------
'���ܣ�����ָ���ַ����ļ���
'˵��������ָ���ַ������ɼ��룬���������������͵ļ���
'        0��ƴ����ȡÿ�ֵ�����ĸ���ɼ���
'        1����ʣ�ȡÿ�ֵ�����ĸ���ɼ���
'        2����ʣ�����ʹ��򹹳ɼ���
'      �ڴ���Ĳ�����δ���֡����ţ��Ͱ��û���ϵͳѡ�������õķ�ʽ���ɼ��룻
'        ����Ͱ��ڡ����ź������ָ���ķ�ʽǿ�����ɼ��룬���1��ʾ���������ĸ����
'-------------------------------------------------------------
    GetCodeByVB = gobjComLib.zlCommFun.SpellCode(strAsk)
End Function

Public Function GetCodeByORCL(strInput As String, Optional ByVal blnWB As Boolean, Optional lngLen As Long = 10) As String
'���ܣ������ַ����ļ���
'��Σ�strInput-�����ַ�����bytIsWB-�Ƿ����(����Ϊƴ��)
    '          lngLen=<1��>40���ص��ַ�������Ϊ10�����򷵻�ָ�������ַ���
'���Σ���ȷ�����ַ��������󷵻�"-"
     GetCodeByORCL = gobjComLib.zlCommFun.zlGetSymbol(strInput, IIf(blnWB, 1, 0), lngLen)
End Function

Public Function FullPinYin(strText As String, Optional intCapital As Integer = 0, Optional blnUseSpliter As Boolean = True) As String
'���ܣ�����GBK������ȡ���ֵ�ȫƴ��
'������ strText         Ҫȡ��ȫƴ���ַ���
'       intCapital      ��Сд��ǣ�0-��д��1-Сд��2-����ĸ��д
'       blnUseSpliter   ƴ��֮���Ƿ񷵻ؿո�True-���ո�False-�����ո�
'���أ�             ����ȡ�õ�ȫƴ���û����Ϊ��
    FullPinYin = gobjComLib.zlCommFun.mGetFullPY(strText, intCapital, blnUseSpliter)
End Function

Public Function ChineseMoney(curMoney) As String
'���ܣ���ָ���Ľ����ֵת��Ϊ��д�����ʾ
'������
'       curMoney:��Ҫת���Ľ����ֵ
'���أ�
    ChineseMoney = gobjComLib.zlCommFun.UppeMoney(curMoney)
End Function

Public Function FormatEx(ByVal vNumber As Variant, ByVal intBit As Integer, Optional blnShowZero As Boolean = True, Optional ByVal blnAddZero As Boolean) As String
'���ܣ��������뷽ʽ��ʽ����ʾ����,��֤С������󲻳���0,С����ǰҪ��0
'������vNumber=Single,Double,Currency���͵�����,
'          intBit=���С��λ��
'         blnShowZero=vNumberΪ0ʱ�Ƿ���ʾ0ֵ
'         blnAddZero=С��λ�����Ƿ���
'���أ���ʽ������ַ���
    Dim strNumber As String
            
    If TypeName(vNumber) = "String" Then
        If vNumber = "" Then Exit Function
        If Not IsNumeric(vNumber) Then Exit Function
        vNumber = Val(vNumber)
    End If
    
    If Not blnAddZero Then 'С��λ�����ʾ�㡣��1.0100 ��Ϊ1.01
        If vNumber = 0 Then
            strNumber = IIf(blnShowZero, 0, "")
        ElseIf Int(vNumber) = vNumber Then
            strNumber = vNumber
        Else
            strNumber = Format(vNumber, "0." & String(intBit, "0"))
            If Left(strNumber, 1) = "." Then strNumber = "0" & strNumber
            If InStr(strNumber, ".") > 0 Then
                Do While Right(strNumber, 1) = "0"
                    strNumber = Left(strNumber, Len(strNumber) - 1)
                Loop
                If Right(strNumber, 1) = "." Then strNumber = Left(strNumber, Len(strNumber) - 1)
            End If
        End If
    Else 'С��λ�����㲹��.��3λС����1.1��Ϊ1.100
        strNumber = Format(vNumber, "#0." & String(intBit, "0"))
    End If
    FormatEx = strNumber
End Function

Public Function ExpressValue(strExpress As String, Optional ByRef strErr As String) As Variant
    '����               �������ʽ
    '����               strExpress = �������ʽ
    '����               ������
    On Error GoTo errH
    Dim sc As Object
    Set sc = CreateObject("ScriptControl")
    sc.Language = "VBScript"
    ExpressValue = sc.Eval(Trim(strExpress))
    Exit Function
errH:
    strErr = "��������(ExpressValue),������Ϣ:" & Err.Number & " " & Err.Description
    Err.Clear
End Function

Public Function SQLAdjust(Str As String) As String
'���ܣ�������"'"���ŵ��ַ�������ΪOracle����ʶ����ַ�����
'˵�����Զ�(����)�����߼�"'"�綨����

    Dim i As Long, strTmp As String
    
    If InStr(1, Str, "'") = 0 Then SQLAdjust = "'" & Str & "'": Exit Function
    
    For i = 1 To Len(Str)
        If Mid(Str, i, 1) = "'" Then
            If i = 1 Then
                strTmp = "CHR(39)||'"
            ElseIf i = Len(Str) Then
                strTmp = strTmp & "'||CHR(39)"
            Else
                strTmp = strTmp & "'||CHR(39)||'"
            End If
        Else
            If i = 1 Then
                strTmp = "'" & Mid(Str, i, 1)
            ElseIf i = Len(Str) Then
                strTmp = strTmp & Mid(Str, i, 1) & "'"
            Else
                strTmp = strTmp & Mid(Str, i, 1)
            End If
        End If
    Next
    SQLAdjust = strTmp
End Function