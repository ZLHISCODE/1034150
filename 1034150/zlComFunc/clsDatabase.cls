VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'---------------------------------------------------------------------
'模块名称：clsDatabase
'包含对象：操作数据库的函数
'命名规则：操作类型 + 操作对象
'说明    ：各段的含义与上面类似；操作对象一般为名词，也可配合修饰词使用。
'例      ：OpenRecordset，GetMaxCode
'--------------------------------------------------------------------------
Private mcolPatiColor As Collection '记录各种类型的病人颜色

Public Function GetPatiColor(ByVal strPatiType As String, Optional ByVal blnLoad As Boolean) As Long
'功能：根据指定的病人类型，返回病人颜色
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    If strPatiType = "" Then GetPatiColor = 0: Exit Function
    
    If mcolPatiColor Is Nothing Or blnLoad Then
        Set mcolPatiColor = New Collection
            
        On Error GoTo errH
        strSQL = "Select 名称,颜色 From 病人类型"
        Set rsTmp = New ADODB.Recordset
        Call OpenRecordset(rsTmp, strSQL, "GetPatiColor")
        Do While Not rsTmp.EOF
            mcolPatiColor.Add CLng(IIf(IsNull(rsTmp!颜色), 0, rsTmp!颜色)), "_" & rsTmp!名称
            rsTmp.MoveNext
        Loop
    End If
    
    On Error Resume Next
    GetPatiColor = mcolPatiColor("_" & strPatiType)
    Exit Function
errH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Function

Public Function GetMax(ByVal strTable As String, ByVal strField As String, ByVal intLength As Integer, Optional ByVal strWhere As String) As String
'功能：读取指定表的本级编码的最大值
'参数：strTable  表名;
'      strField  字段名;
'      intLength 字段长度
'返回：成功返回 下级最大编码; 否者返回 0
    Dim rsTemp As New ADODB.Recordset
    Dim varTemp As Variant, strSQL As String
    Dim lngLengh As Long
    
    On Error GoTo errHand
    With rsTemp
        strSQL = "SELECT MAX(LPAD(" & strField & "," & intLength & ",' ')) as ""最大值"",max(length(" & _
             strField & ")) as ""最长值"" FROM " & strTable & strWhere
        rsTemp.Open strSQL, gcnOracle, adOpenStatic, adLockReadOnly
        
        If rsTemp.EOF Then
            GetMax = Format(1, String(intLength, "0"))
            Exit Function
        End If
        varTemp = IIf(IsNull(.Fields("最大值").value), "0", .Fields("最大值").value)
        lngLengh = IIf(IsNull(.Fields("最长值").value), intLength, .Fields("最长值").value)
        If IsNumeric(varTemp) Then
            GetMax = CStr(Val(varTemp) + 1)
            GetMax = Format(GetMax, String(lngLengh, "0"))
        Else
            GetMax = Mid(varTemp, 1, Len(varTemp) - 1) & Chr(Asc(Right(varTemp, 1)) + 1)
            GetMax = Trim(GetMax)
        End If
        .Close
    End With
    Exit Function
    
errHand:
    If gobjComLib.ErrCenter() = 1 Then Resume
End Function

Public Function GetNextId(strTable As String) As Long
    '------------------------------------------------------------------------------------
    '功能：读取指定表名对应的序列(按规范，其序列名称为“表名称_id”)的下一数值
    '参数：
    '   strTable：表名称
    '返回：
    '------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select " & Trim(strTable) & "_ID.Nextval From Dual"
    Call gobjComLib.SQLTest(App.ProductName, "mdlCommon", strSQL)
    rsTmp.Open strSQL, gcnOracle, adOpenKeyset
    Call gobjComLib.SQLTest
    GetNextId = rsTmp.Fields(0).value
    Exit Function
errH:
    If gobjComLib.ErrCenter() = 1 Then Resume
End Function

Public Function GetNextNo(ByVal int序号 As Integer, Optional ByVal lng科室ID As Long, Optional ByVal strTag As String, Optional ByVal intStep As Integer = 1) As Variant
'功能:根据特定规则产生新的号码,本函数规则只适于ZLHIS10，且需要Oracle 8i(8.1.5)以上版本支持
'参数：
'int序号=项目序号:
'  1   病人ID 数字
'  2   住院号 数字
'  3   门诊号 数字
'  10  医嘱发送号 数字,顺序递增编号
'  x   其它单据号 字符,根据编号规则顺序递增编号,不自动补缺
'lng科室ID=按科室号码编号规则的项目需要
'返回：最大号码
'说明：
'  编号规则：0-按年顺序编号,1-按日顺序编号,2-按执行科室分月编号(需要读取科室号码表)
'            对门诊号：0-顺序编号,1-年月日(YYMMDD)+顺序号(0000)
'            对住院号：0-顺序编号,1-年月(YYMM)+顺序号(0000),2-年(YYYY)+顺序号(00000)
'  年度位确定：以1990为基数，随年度增长，按“0～9/A～Z”顺序作为年度编码
'  最大号码-10存入号码控制表,用于并发情况下补缺号(取了号,但未使用)
'  For Update在并发情况下锁定行,不用Wait选项以避免向调用者返回空
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
        
    GetNextNo = Null
    
    On Error GoTo errH
    
    strSQL = "Select NextNO([1],[2],[3],[4]) as NO From Dual"
    Set rsTmp = OpenSQLRecord(strSQL, "GetNextNo", int序号, lng科室ID, strTag, intStep)
    
    If gcnOracle.Errors.count > 0 Then 'Select中函数出错时,在VB中不自动触发错误
        Err.Raise gcnOracle.Errors(0).Number, , gcnOracle.Errors(0).Description
    End If
    
    If Not rsTmp.EOF Then
        If Not IsNull(rsTmp!NO) Then GetNextNo = rsTmp!NO
    End If
    Exit Function
errH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Function

Public Function GetUserInfo() As ADODB.Recordset
'功能：获取当前用户的基本信息
'返回：返回Ado记录集
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = _
        " Select User,A.ID,A.编号,A.简码,A.姓名,B.用户名," & _
        " C.部门ID,D.编码 as 部门码,D.名称 as 部门名" & _
        " From 人员表 A,上机人员表 B,部门人员 C,部门表 D" & _
        " Where A.ID = B.人员ID And A.ID = C.人员ID And C.部门ID=D.ID" & _
        " And C.缺省 = 1 And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null ) And B.用户名 = User"
    Call OpenRecordset(rsTmp, strSQL, "GetUserInfo")
    Set GetUserInfo = rsTmp
    Exit Function
errH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Function

Public Function CurrentDate() As Date
    '-------------------------------------------------------------
    '功能：提取服务器上当前日期
    '参数：
    '返回：服务器上当前日期时间
    '-------------------------------------------------------------
    
    Dim rsTmp  As ADODB.Recordset
    
    Err = 0
    On Error GoTo errHandle
    
    Set rsTmp = OpenCursor("clsDataBase", "ZLTOOLS", "B_Public.Get_Current_Date")
    
    If rsTmp.RecordCount > 0 Then
        CurrentDate = rsTmp.Fields(0)
    Else
        CurrentDate = 0
    End If
    Exit Function
    
errHandle:
    If gobjComLib.ErrCenter() = 1 Then Resume
    CurrentDate = 0
    Err = 0
End Function

Public Function SavePicture(strFile As String, rsTable As ADODB.Recordset, strField As String) As Boolean
    '-------------------------------------------------------------
    '功能：将指定的图形文件保存到指定的记录集字段中
    '参数：
    '       strFile：图形文件名称
    '       rsTable：图形存储记录集
    '       strField：图形字段
    '返回：
    '-------------------------------------------------------------
    Const conChunkSize As Integer = 10240
    Dim lngFileSize As Long, lngCurSize As Long, lngModSize As Long
    Dim intBolcks As Integer, FileNum, j
    Dim aryChunk() As Byte
    
    On Error GoTo errH
    FileNum = FreeFile
    Open strFile For Binary Access Read As FileNum
    lngFileSize = LOF(FileNum)
    
    lngModSize = lngFileSize Mod conChunkSize
    intBolcks = lngFileSize \ conChunkSize - IIf(lngModSize = 0, 1, 0)
    rsTable.Fields(strField).value = Null
    For j = 0 To intBolcks
        If j = lngFileSize \ conChunkSize Then
            lngCurSize = lngModSize
        Else
            lngCurSize = conChunkSize
        End If
        ReDim aryChunk(lngCurSize - 1) As Byte
        Get FileNum, , aryChunk()
        rsTable.Fields(strField).AppendChunk aryChunk()
    Next j
    rsTable.Update
    Close FileNum
    SavePicture = True
    Exit Function

errH:
    rsTable.CancelUpdate
    Close FileNum
    SavePicture = False

End Function

Public Function ReadPicture(rsTable As ADODB.Recordset, strField As String, Optional strFile As String) As String
'-------------------------------------------------------------
'功能：将指定的记录集图形字段复制为图形临时文件
'参数：
'       rsTable   图形存储记录集
'       strField  图形字段
'       strFile   用户定义的文件名（可选项）
'返回：
'-------------------------------------------------------------
    Const conChunkSize As Integer = 10240
    Dim lngFileSize As Long, lngCurSize As Long, lngModSize As Long
    Dim intBolcks As Integer, FileNum, j
    Dim aryChunk() As Byte
    Dim strTempFile As String
    
    On Error GoTo errH
    lngFileSize = rsTable.Fields(strField).ActualSize
    If lngFileSize = 0 Then
        '未读取有效数据
        Exit Function
    End If
    
    FileNum = FreeFile
    If strFile = "" Then
        '当用户并没定义文件名时
        j = 0
        Do While True
            strTempFile = CurDir & "\zlNewPicture" & CStr(j) & ".pic"
            If Len(Dir(strTempFile)) = 0 Then Exit Do
            j = j + 1
        Loop
        strFile = strTempFile
    End If
    Open strFile For Binary As FileNum
    
    lngModSize = lngFileSize Mod conChunkSize
    intBolcks = lngFileSize \ conChunkSize - IIf(lngModSize = 0, 1, 0)
    rsTable.Move 0
    For j = 0 To intBolcks
        If j = lngFileSize \ conChunkSize Then
            lngCurSize = lngModSize
        Else
            lngCurSize = conChunkSize
        End If
        ReDim aryChunk(lngCurSize - 1) As Byte
        aryChunk() = rsTable.Fields(strField).GetChunk(lngCurSize)
        Put FileNum, , aryChunk()
    Next
    Close FileNum
    ReadPicture = strFile
    Exit Function

errH:
    Close FileNum
    Kill strFile
    ReadPicture = ""
End Function

Public Sub OpenRecordset(rsTemp As ADODB.Recordset, strSQL As String, ByVal strFormCaption As String, _
        Optional CursorType As CursorTypeEnum = adOpenStatic, Optional LockType As LockTypeEnum = adLockReadOnly)
'功能：打开记录。同时保存SQL语句
    
    If rsTemp.State = adStateOpen Then rsTemp.Close
    rsTemp.CursorLocation = adUseClient
    Call gobjComLib.SQLTest(App.ProductName, strFormCaption, strSQL)
    rsTemp.Open strSQL, gcnOracle, CursorType, LockType
    rsTemp.ActiveConnection = Nothing
    Call gobjComLib.SQLTest
End Sub

Public Function OpenSQLRecord(ByVal strSQL As String, ByVal strTitle As String, ParamArray arrInput() As Variant) As ADODB.Recordset
'功能：通过Command对象打开带参数SQL的记录集
'参数：strSQL=条件中包含参数的SQL语句,参数形式为"[x]"
'             x>=1为自定义参数号,"[]"之间不能有空格
'             同一个参数可多处使用,程序自动换为ADO支持的"?"号形式
'             实际使用的参数号可不连续,但传入的参数值必须连续(如SQL组合时不一定要用到的参数)
'      arrInput=不定个数的参数值,按参数号顺序依次传入,必须是明确类型
'      strTitle=用于SQLTest识别的调用窗体/模块标题
'返回：记录集，CursorLocation=adUseClient,LockType=adLockReadOnly,CursorType=adOpenStatic
'举例：
'SQL语句为="Select 姓名 From 病人信息 Where (病人ID=[3] Or 门诊号=[3] Or 姓名 Like [4]) And 性别=[5] And 登记时间 Between [1] And [2] And 险类 IN([6],[7])"
'调用方式为：Set rsPati=OpenSQLRecord(strSQL, Me.Caption, CDate(Format(rsMove!转出日期,"yyyy-MM-dd")),dtp时间.Value, lng病人ID, "张%", "男", 20, 21)
    Static cmdData As New ADODB.Command
    Dim strPar As String, arrPar As Variant
    Dim lngLeft As Long, lngRight As Long
    Dim strSeq As String, intMax As Integer, i As Integer
    Dim strLog As String, varValue As Variant
    
    '分析自定的[x]参数
    lngLeft = InStr(1, strSQL, "[")
    Do While lngLeft > 0
        lngRight = InStr(lngLeft + 1, strSQL, "]")
        
        '可能是正常的"[编码]名称"
        strSeq = Mid(strSQL, lngLeft + 1, lngRight - lngLeft - 1)
        If IsNumeric(strSeq) Then
            i = CInt(strSeq)
            strPar = strPar & "," & i
            If i > intMax Then intMax = i
        End If
        
        lngLeft = InStr(lngRight + 1, strSQL, "[")
    Loop

    '替换为"?"参数
    strLog = strSQL
    For i = 1 To intMax
        strSQL = Replace(strSQL, "[" & i & "]", "?")
        
        '产生用于SQL跟踪的语句
        varValue = arrInput(i - 1)
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '数字
            strLog = Replace(strLog, "[" & i & "]", varValue)
        Case "String" '字符
            strLog = Replace(strLog, "[" & i & "]", "'" & Replace(varValue, "'", "''") & "'")
        Case "Date" '日期
            strLog = Replace(strLog, "[" & i & "]", "To_Date('" & Format(varValue, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')")
        End Select
    Next

    '清除原有参数:不然不能重复执行
    cmdData.CommandText = "" '不为空有时清除参数出错
    Do While cmdData.Parameters.count > 0
        cmdData.Parameters.Delete 0
    Loop
    
    '创建新的参数
    lngLeft = 0: lngRight = 0
    arrPar = Split(Mid(strPar, 2), ",")
    For i = 0 To UBound(arrPar)
        varValue = arrInput((arrPar(i) - 1))
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '数字
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

    '执行返回记录集
    If cmdData.ActiveConnection Is Nothing Then
        Set cmdData.ActiveConnection = gcnOracle '这句比较慢
    End If
    cmdData.CommandText = strSQL
    
    Call gobjComLib.SQLTest(App.ProductName, strTitle, strLog)
    Set OpenSQLRecord = cmdData.Execute
    OpenSQLRecord.ActiveConnection = Nothing
    Call gobjComLib.SQLTest
End Function

Public Sub ExecuteProcedure(strSQL As String, ByVal strFormCaption As String)
'功能：执行过程语句,并自动对过程参数进行绑定变量处理
'参数：strSQL=过程语句,可能带参数,形如"过程名(参数1,参数2,...)"。
'说明：以下几种情况过程参数不使用绑定变量,仍用老的调用方法：
'  1.参数部份是表达式,这时程序无法处理绑定变量类型和值,如"过程名(参数1,100.12*0.15,...)"
'  2.中间没有传入明确的可选参数,这时程序无法处理绑定变量类型和值,如"过程名(参数1, , ,参数3,...)"
'  3.因为该过程是自动处理,不是一定使用绑定变量,对带"'"的字符参数,仍要使用"''"形式。
    Dim cmdData As New ADODB.Command
    Dim strProc As String, strPar As String
    Dim blnStr As Boolean, intBra As Integer
    Dim strTemp As String, i As Long
    Dim intMax As Integer, datCur As Date
    
    If Right(Trim(strSQL), 1) = ")" Then
        '清除原有参数:不然不能重复执行
'        cmdData.CommandText = "" '不为空有时清除参数出错
'        Do While cmdData.Parameters.Count > 0
'            cmdData.Parameters.Delete 0
'        Loop
        
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
                        .Parameters.Append .CreateParameter("PAR" & .Parameters.count, adVarNumeric, adParamInput, 30, strPar)
                    ElseIf Left(strPar, 1) = "'" And Right(strPar, 1) = "'" Then '字符串
                        strPar = Mid(strPar, 2, Len(strPar) - 2)
                        
                        'Oracle连接符运算:'ABCD'||CHR(13)||'XXXX'||CHR(39)||'1234'
                        If InStr(Replace(strPar, " ", ""), "'||") > 0 Then GoTo NoneVarLine
                        
                        '双"''"的绑定变量处理
                        If InStr(strPar, "''") > 0 Then strPar = Replace(strPar, "''", "'")
                        
                        '电子病历处理LOB时，如果用绑定变量转换为RAW时第2000个字符不正确
                        intMax = LenB(StrConv(strPar, vbFromUnicode))
                                                
                        If intMax <= 2000 Then
                            intMax = IIf(intMax <= 200, 200, 2000)
                            .Parameters.Append .CreateParameter("PAR" & .Parameters.count, adVarChar, adParamInput, intMax, strPar)
                        Else
                            If intMax < 4000 Then intMax = 4000
                            .Parameters.Append .CreateParameter("PAR" & .Parameters.count, adLongVarChar, adParamInput, intMax, strPar)
                        End If
                        
                    ElseIf UCase(strPar) Like "TO_DATE('*','*')" Then '日期
                        strPar = Split(strPar, "(")(1)
                        strPar = Trim(Split(strPar, ",")(0))
                        strPar = Mid(strPar, 2, Len(strPar) - 2)
                        If strPar = "" Then
                            'NULL值当成数字处理可兼容其他类型
                            .Parameters.Append .CreateParameter("PAR" & .Parameters.count, adVarNumeric, adParamInput, , Null)
                        Else
                            If Not IsDate(strPar) Then GoTo NoneVarLine
                            .Parameters.Append .CreateParameter("PAR" & .Parameters.count, adDBTimeStamp, adParamInput, , CDate(strPar))
                        End If
                    ElseIf UCase(strPar) = "SYSDATE" Then '日期
                        If datCur = CDate(0) Then datCur = CurrentDate
                        .Parameters.Append .CreateParameter("PAR" & .Parameters.count, adDBTimeStamp, adParamInput, , datCur)
                    ElseIf UCase(strPar) = "NULL" Then 'NULL值当成字符处理可兼容其他类型
                        .Parameters.Append .CreateParameter("PAR" & .Parameters.count, adVarChar, adParamInput, 200, Null)
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
        For i = 1 To cmdData.Parameters.count
            strTemp = strTemp & ",?"
        Next
        strProc = "Call " & strProc & "(" & Mid(strTemp, 2) & ")"
        
        '执行过程
        'If cmdData.ActiveConnection Is Nothing Then
            Set cmdData.ActiveConnection = gcnOracle '这句比较慢
            cmdData.CommandType = adCmdText
        'End If
        cmdData.CommandText = strProc
        
        Call gobjComLib.SQLTest(App.ProductName, strFormCaption, strSQL)
        Call cmdData.Execute
        Call gobjComLib.SQLTest
    Else
        GoTo NoneVarLine
    End If
    Exit Sub
NoneVarLine:
    Call gobjComLib.SQLTest(App.ProductName, strFormCaption, strSQL)
    
    '说明：为了兼容新连接方式
    '1.新连接用adCmdStoredProc方式在8i下面有问题
    '2.新连接如果不使用{},则即使过程没有参数也要加()
    strSQL = "Call " & strSQL
    If InStr(strSQL, "(") = 0 Then strSQL = strSQL & "()"
    gcnOracle.Execute strSQL, , adCmdText
    
    Call gobjComLib.SQLTest
End Sub

Public Function ShowSelect(frmParent As Object, ByVal strSQL As String, bytStyle As Byte, _
    Optional ByVal strTitle As String, Optional bln末级 As Boolean, _
    Optional ByVal strSeek As String, Optional ByVal strNote As String, _
    Optional ByVal blnShowSub As Boolean, Optional blnShowRoot As Boolean, _
    Optional ByVal blnNoneWin As Boolean, Optional ByVal X As Long = -1, _
    Optional ByVal Y As Long = -1, Optional ByVal txtH As Long = -1, _
    Optional ByRef Cancel As Boolean, Optional ByVal blnMultiOne As Boolean, _
    Optional ByVal blnSearch As Boolean) As ADODB.Recordset
'功能：多功能选择器
'参数：
'     frmParent=显示的父窗体
'     strSQL=数据来源,不同风格的选择器对SQL中的字段有不同要求
'     bytStyle=选择器风格
'       为0时:列表风格:ID,…
'       为1时:树形风格:ID,上级ID,编码,名称(如果bln末级，则需要末级字段)
'       为2时:双表风格:ID,上级ID,编码,名称,末级…；ListView只显示末级=1的项目
'     strTitle=选择器功能命名,也用于个性化区分
'     bln末级=当树形选择器(bytStyle=1)时,是否只能选择末级为1的项目
'     strSeek=当bytStyle<>2时有效,缺省定位的项目。
'             bytStyle=0时,以ID和上级ID之后的第一个字段为准。
'             bytStyle=1时,可以是编码或名称
'     strNote=选择器的说明文字
'     blnShowSub=当选择一个非根结点时,是否显示所有下级子树中的项目(项目多时较慢)
'     blnShowRoot=当选择根结点时,是否显示所有项目(项目多时较慢)
'     blnNoneWin,X,Y,txtH=处理成非窗体风格,X,Y,txtH表示调用界面输入框的坐标(相对于屏幕)和高度
'     Cancel=返回参数,表示是否取消,主要用于blnNoneWin=True时
'     blnMultiOne=当bytStyle=0时,是否将对多行相同记录当作一行判断
'     blnSearch=是否显示行号,并可以输入行号定位
'返回：取消=Nothing,选择=SQL源的单行记录集
'说明：
'     1.ID和上级ID可以为字符型数据
'     2.末级等字段不要带空值
'应用：可用于各个程序中数据量不是很大的选择器,输入匹配列表等。
    Dim frmNew As New frmPubSel
    
    Set ShowSelect = frmNew.ShowSelect(frmParent, strSQL, bytStyle, strTitle, bln末级, strSeek, strNote, blnShowSub, blnShowRoot, blnNoneWin, X, Y, txtH, Cancel, blnMultiOne, blnSearch, False)
End Function

Public Function ShowSQLSelect(frmParent As Object, ByVal strSQL As String, bytStyle As Byte, _
    ByVal strTitle As String, ByVal bln末级 As Boolean, _
    ByVal strSeek As String, ByVal strNote As String, _
    ByVal blnShowSub As Boolean, ByVal blnShowRoot As Boolean, _
    ByVal blnNoneWin As Boolean, ByVal X As Long, _
    ByVal Y As Long, ByVal txtH As Long, _
    ByRef Cancel As Boolean, ByVal blnMultiOne As Boolean, _
    ByVal blnSearch As Boolean, ParamArray arrInput() As Variant) As ADODB.Recordset
'功能：多功能选择器,使用ADO.Command打开,允许使用[x]参数
'参数：
'     frmParent=显示的父窗体
'     strSQL=数据来源,不同风格的选择器对SQL中的字段有不同要求
'     bytStyle=选择器风格
'       为0时:列表风格:ID,…
'       为1时:树形风格:ID,上级ID,编码,名称(如果bln末级，则需要末级字段)
'       为2时:双表风格:ID,上级ID,编码,名称,末级…；ListView只显示末级=1的项目
'     strTitle=选择器功能命名,也用于个性化区分
'     bln末级=当树形选择器(bytStyle=1)时,是否只能选择末级为1的项目
'     strSeek=当bytStyle<>2时有效,缺省定位的项目。
'             bytStyle=0时,以ID和上级ID之后的第一个字段为准。
'             bytStyle=1时,可以是编码或名称
'     strNote=选择器的说明文字
'     blnShowSub=当选择一个非根结点时,是否显示所有下级子树中的项目(项目多时较慢)
'     blnShowRoot=当选择根结点时,是否显示所有项目(项目多时较慢)
'     blnNoneWin,X,Y,txtH=处理成非窗体风格,X,Y,txtH表示调用界面输入框的坐标(相对于屏幕)和高度
'     Cancel=返回参数,表示是否取消,主要用于blnNoneWin=True时
'     blnMultiOne=当bytStyle=0时,是否将对多行相同记录当作一行判断
'     blnSearch=是否显示行号,并可以输入行号定位
'     arrInput=对应的各个SQL参数值,按顺序传入,必须为明确类型
'返回：取消=Nothing,选择=SQL源的单行记录集
'说明：
'     1.ID和上级ID可以为字符型数据
'     2.末级等字段不要带空值
'应用：可用于各个程序中数据量不是很大的选择器,输入匹配列表等。
    Dim frmNew As New frmPubSel
    Dim arrPar() As Variant
    
    arrPar = arrInput
    Set ShowSQLSelect = frmNew.ShowSelect(frmParent, strSQL, bytStyle, strTitle, bln末级, strSeek, strNote, blnShowSub, blnShowRoot, blnNoneWin, X, Y, txtH, Cancel, blnMultiOne, blnSearch, False, arrPar)
End Function

Public Function ShowSQLMultiSelect(frmParent As Object, ByVal strSQL As String, bytStyle As Byte, _
    ByVal strTitle As String, ByVal bln末级 As Boolean, _
    ByVal strSeek As String, ByVal strNote As String, _
    ByVal blnShowSub As Boolean, ByVal blnShowRoot As Boolean, _
    ByVal blnNoneWin As Boolean, ByVal X As Long, _
    ByVal Y As Long, ByVal txtH As Long, _
    ByRef Cancel As Boolean, ByVal blnMultiOne As Boolean, _
    ByVal blnSearch As Boolean, ParamArray arrInput() As Variant) As ADODB.Recordset
'功能：功能同ShowSQLSelect,除树形风格外(bytStyle=1)，可以多选
'返回：取消=Nothing,选择=SQL源的单行或多行记录集
    Dim frmNew As New frmPubSel
    Dim arrPar() As Variant
    
    arrPar = arrInput
    Set ShowSQLMultiSelect = frmNew.ShowSelect(frmParent, strSQL, bytStyle, strTitle, bln末级, strSeek, strNote, blnShowSub, blnShowRoot, blnNoneWin, X, Y, txtH, Cancel, blnMultiOne, blnSearch, True, arrPar)
End Function

Public Function zlShowListSelect(ByVal frmMain As Object, ByVal lngSys As Long, ByVal lngModule As Long, ByVal objControl As Object, ByVal rsBindings As ADODB.Recordset, _
     Optional ByVal blnShowHead As Boolean = False, _
     Optional ByVal str参数名 As String = "", _
     Optional ByVal strHideCols As String = "", _
     Optional ByRef rsOutSel As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:选择器入口
    '入参:frmMain-调用的主窗口
    '     lngSys-系统号
    '     lngModule-模块号
    '     objControl-控件对象(目前只支:textBox,Combox)
    '     rsBindings-绑定的记录集(不能为空,主要字段,ID,......)
    '     str参数-个性化保存的参数名.
    '     blnShowHead-是否显示现列头
    '
    '出参:rsOutSel-选择后的记录集
    '返回:选中返回True, 否则返回False(可以按Esc进行返回)
    '编制:刘兴洪
    '日期:2009-01-01 15:35:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim frmNew As New frmListSel
   zlShowListSelect = frmNew.ShowSelect(frmMain, lngSys, lngModule, objControl, rsBindings, blnShowHead, str参数名, strHideCols, rsOutSel)
End Function

Public Function ShowILLSelect(frmParent As Object, ByVal str类别 As String, ByVal lng病人科室ID As Long, _
    Optional ByVal str性别 As String, Optional ByVal blnMultiSel As Boolean, Optional ByVal blnICD10 As Boolean = True) As ADODB.Recordset
'功能：ICD10疾病编码/疾病诊断选择器
'参数：str类别=多个疾病编码类别,形如"D,M",为空时表示所有
'         D:ICD-10疾病编码
'         Y:损伤中毒的外部原因
'         M:肿瘤形态学编码
'         S:ICD-9-CM3手术编码
'         B:中医疾病编码
'         Z:中医病征编码
'         1:疾病诊断-西医
'         2:疾病诊断-中医
'      str性别=病人性别,用于处理疾病的性别限制
'      blnMultiSel=是否允许进行复选
'      blnICD10=是否显示ICD-10疾病编码选择器，否则显示疾病诊断选择器
'返回：包含选择疾病项目的记录集，
'      ICD-10疾病编码选择器包含字段：项目ID,编码,序号,附码,名称,说明
'      疾病诊断选择器包含字段：项目ID,编码,名称,说明,编者
'      如果取消选择，返回为Nothing
'注意：调用该函数时需具有以下权限
'      zl_疾病编码科室_Insert,Zl_疾病编码科室_Delete
'      疾病编码目录,疾病编码科室,疾病编码分类,疾病编码类别
'      Zl_疾病诊断科室_Insert,Zl_疾病诊断科室_Delete
'      疾病诊断目录,疾病诊断分类,疾病诊断属类
'      部门表,部门人员,部门性质说明,上机人员表
    Set ShowILLSelect = frmILLSelect.ShowMe(frmParent, str类别, lng病人科室ID, str性别, blnMultiSel, blnICD10)
End Function

Public Function ShowPatiColorTip(frmParent As Object) As Boolean
'功能:在frmParent窗口右下角显示一窗体，内容为各种病人类型的颜色说明
    Call frmPatiType.ShowPatiType(frmParent)
End Function

Public Function CopyNewRec(ByVal rsSource As ADODB.Recordset) As ADODB.Recordset
'编制人:朱玉宝
'编制日期:2000-11-02
'复制记录集
'在程序中，经常会涉及到相互传递记录集，而使用ADO的Clone复制产生的记录集，当其中一个记录集的数据发生变化的时候，所有副本都将发生相同的变化（通常指修改或删除），而我们往往希望这些记录集相互间保持独立
    Dim rsClone As New ADODB.Recordset
    Dim rsTarget As New ADODB.Recordset
    Dim intFields As Integer
    
    Set rsClone = rsSource.Clone
    rsClone.Filter = rsSource.Filter
    Set rsTarget = New ADODB.Recordset
    With rsTarget
        For intFields = 0 To rsClone.Fields.count - 1
            .Fields.Append rsClone.Fields(intFields).Name, IIf(rsClone.Fields(intFields).type = adNumeric, adDouble, rsClone.Fields(intFields).type), rsClone.Fields(intFields).DefinedSize, adFldIsNullable    '0:表示新增
        Next
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
        
        If rsClone.RecordCount <> 0 Then rsClone.MoveFirst
        Do While Not rsClone.EOF
            .AddNew
            For intFields = 0 To rsClone.Fields.count - 1
                .Fields(intFields) = rsClone.Fields(intFields).value
            Next
            .Update
            rsClone.MoveNext
        Loop
    End With
    
    Set CopyNewRec = rsTarget
End Function

Public Function NOMoved(ByVal strTableName As String, ByVal strNO As String, Optional ByVal strIF As String = "记录性质 IN", _
        Optional ByVal strIFValue As String = "", Optional ByVal strFormCaption As String = "ZL9ComLib.NOMoved", Optional ByVal blnOnline As Boolean = False) As Boolean
'功能:根据指定的表名、单据号、记录性质,判断指定的单据号是否已转出到后备数据表中
'参数:NO可以是"F0000007"或"Between 'F0000001' And 'F0000007',如果是非单据号条件,则要求长度不低于8位
'       strIFValue在费用系统中指记录性质,可以用1,2这种形式
'       strIF,strIFValue:除单据号外的附加条件及条件值,默认为费用系统的条件,两个条件必须配对使用
'返回:
'     注意:函数名NOMoved中的NO表示单据号NO,而不是no,不要理解为没有被转出,正确的含义是:指定的单据被转出,则为真,否则为假,出错也返回假.
'张永康 2005-04-14
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    If strTableName = "" Or strNO = "" Then Exit Function
    
    On Error GoTo errH
    
    If strIFValue <> "" Then
        If strIF = "记录性质 IN" Then
            strSQL = " And " & strIF & " (" & strIFValue & ")" '仅费用系统适用
        Else
            strSQL = " And " & strIF & strIFValue
        End If
    End If
    
    If Not blnOnline Then strTableName = "H" & strTableName
    
    strSQL = "Select * From " & strTableName & " Where Rownum<2 And " & _
            IIf(Len(strNO) <= 8, "NO='" & strNO & "' ", "NO " & strNO) & strSQL
    
    Call OpenRecordset(rsTmp, strSQL, strFormCaption)
    
    If Not rsTmp Is Nothing Then
        NOMoved = Not rsTmp.EOF
    End If
    Exit Function
    
errH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Function

Public Function DateMoved(ByVal strDate As String, Optional ByVal int组号 As Integer = 1, Optional ByVal int系统 As Integer = 100, Optional ByVal strFormCaption As String = "ZL9ComLib.DateMoved") As Boolean
'功能:根据指定的日期、组号、系统判断指定日期的数据是否已转出到后备数据表中
'张永康 2005-04-14
    Dim rsTmp As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    Dim prmPara As ADODB.Parameter
    Dim strSQL As String
    
    If strDate = "" Then Exit Function
    On Error GoTo errH
    
    Set rsTmp = OpenCursor("clsDatabase", "zlTools", "b_ComFunc.Get_DataMoved", int组号, int系统, CDate(strDate))
    DateMoved = Not rsTmp.EOF
    Exit Function
    
errH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Function

Public Function ReturnMovedExes(ByVal strOwner As String, ByVal strNO As String, Optional ByVal bytType As Byte = 1, Optional ByVal strFormCaption As String = "ZL9ComLib.ReturnMovedExes") As Boolean
'功能:抽选返回指定费用单据号的数据(Oracle过程:Zl_Retu_Exes)
'参数:NO单据号,bytType表示单据类型,值::1-收费,2-记帐,3-自动记帐,4-挂号,5-就诊卡,6-预交,7-结帐；
'张永康 2005-04-19
    Dim strSQL As String
    
    On Error GoTo errH
    
    'strSQL = "Zl_Retu_Exes('" & strNO & "'," & CStr(bytType) & ")"
    Call RunProcedure("clsDatabase", strOwner, "Zl_Retu_Exes", strNO, CStr(bytType))
    ReturnMovedExes = True
    Exit Function
    
errH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Function

Public Function ReturnMovedClinic(ByVal strOwner As String, ByVal intPatientID As Integer, ByVal strTimes As String, Optional ByVal bytFlag As Byte = 1, Optional ByVal strFormCaption As String = "ZL9ComLib.ReturnMovedClinic") As Boolean
'功能:抽选返回病人某次门诊住院医疗数据(Oracle过程:Zl_Retu_Clinic)
'参数:  Times,挂号单号或住院主页id
'       Flag,门诊或住院标志:0-门诊,1-住院
'张永康 2005-04-19
    Dim strSQL As String
    
    On Error GoTo errH
    'strSQL = "Zl_Retu_Clinic(" & CStr(intPatientID) & ",'" & strTimes & "'," & CStr(bytFlag) & ")"
    Call RunProcedure("clsDatabase", strOwner, "Zl_Retu_Clinic", CStr(intPatientID), strTimes, CStr(bytFlag))
    
    ReturnMovedClinic = True
    Exit Function
    
errH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Function

Public Sub ShowReportMenu(objForm As Object, ByVal lngSys As Long, ByVal lngModul As Long, ByVal strPrivs As String, ParamArray arrHide() As Variant)
'功能：在指定窗体的菜单中显示发布到该窗体模块的自定义报表
'参数：objForm=需要显示报表菜单的模块窗体
'      lngSys,lngModul=系统编号和模块编号
'      strPrivs=指定模块已分析出的当前具有的权限
'      arrHide=用于指定需要隐藏不显示在菜单中的系统固有报表(传入编号, "ZL1_INSIDE_1234_1"," ZL1_INSIDE_1234_2",…)
'              比如少量报表可能是虽然发布到了模块，但并不想自动在菜单中直接显示，而是程序另外处理调用的情况。
'说明:
'1.窗体中必须要有VB菜单或CommandBar控件
'  如果是用VB菜单,则调用函数前需要预先定义菜单项，命名：mnuReport\mnuReportItem(0),标题:报表(&R)\-，mnuReport应为隐藏,下面可预先加入一些需要的菜单项(快捷字母用0-9)
'  如果是用CommandBar控件,则报表根菜单ID=4,具体菜单项ID=401开始,根菜单标题为:报表(&R)，没有时函数自动增加在“查看,帮助”菜单项前面。
'2.报表的"系统,编号"存放于VBMenu.Tag和CommandControl.Parameter属性中。
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strHide As String
    Dim objItem As Object, objControl As Object
    Dim objBar As Object, objPop As Object
    Dim intMenu As Integer, lngMenuID As Long
    Dim i As Long, K As Long
    
    Const xtpControlButton = 1
    Const xtpControlPopup = 2
    Const conMenu_ReportPopup = 4
    Const conMenu_Report_Item = 451 '最小开始ID号
    
    
    '判断窗体菜单类型:intMenu=0-无,1-VBMenu,2=CommandBar
    '---------------------------------------------------
    If TypeName(objForm) = "CommandBars" Then
        Set objBar = objForm: intMenu = 2
    Else
    On Error Resume Next
    If objForm.mnuReport.Name <> "" And objForm.mnuReportItem(0).Name <> "" Then
        If Err.Number = 0 Then intMenu = 1
    End If
    Err.Clear
    If intMenu = 0 Then
        For Each objItem In objForm.Controls
            If TypeName(objItem) = "CommandBars" Then
                If objItem.ActiveMenuBar.Visible Then
                    Set objBar = objItem: intMenu = 2
                End If
                Exit For
            End If
        Next
    End If
    Err.Clear: On Error GoTo 0
    End If
    If intMenu = 0 Then Exit Sub
    
    '读取要显示的报表项
    '---------------------------------------------------
    If UBound(arrHide) <> -1 Then
        For i = 0 To UBound(arrHide)
            strHide = strHide & "," & arrHide(i)
        Next
        If strHide <> "" Then strHide = strHide & ","
    End If
    
    '固定发布报表(不含票据,通过_BILL_判断) + 用户发布报表
    On Error GoTo errH
    Set rsTmp = OpenCursor("clsDatabase", "zltools", "b_ComFunc.Get_Report_Menu", lngSys, lngModul, ";" & strPrivs & ";", strHide)
    On Error GoTo 0
    If rsTmp.EOF Then Exit Sub
    
    '在VBMenu上加载要显示的报表项
    '---------------------------------------------------
    If intMenu = 1 Then
        Set objItem = objForm.mnuReportItem(objForm.mnuReportItem.UBound)
        If Not (objItem.Caption = "-" Or Trim(objItem.Caption) = "" Or Not objItem.Visible) Then
            Load objForm.mnuReportItem(objForm.mnuReportItem.UBound + 1)
            Set objItem = objForm.mnuReportItem(objForm.mnuReportItem.UBound)
        End If
        
        '不是第一个菜单项时，加入分隔
        If objItem.Index > 0 Then
            objItem.Enabled = True
            objItem.Caption = "-"
            objItem.Tag = ""
            objItem.Visible = True
            Set objItem = Nothing
        End If
        
        K = 0
        For i = 1 To rsTmp.RecordCount
            '加入固定与用户报表的分隔
            If K <> rsTmp!标志 And K <> 0 Then
                Load objForm.mnuReportItem(objForm.mnuReportItem.UBound + 1)
                Set objItem = objForm.mnuReportItem(objForm.mnuReportItem.UBound)
                objItem.Enabled = True
                objItem.Caption = "-"
                objItem.Tag = ""
                objItem.Visible = True
                Set objItem = Nothing
            End If
                        
            '加入报表项:用Nothing标识是否需要新加
            If objItem Is Nothing Then
                Load objForm.mnuReportItem(objForm.mnuReportItem.UBound + 1)
            End If
            Set objItem = objForm.mnuReportItem(objForm.mnuReportItem.UBound)
            objItem.Enabled = True
            objItem.Caption = rsTmp!名称
            If i <= 26 Then
                objItem.Caption = objItem.Caption & "(&" & Chr(Asc("A") + i - 1) & ")"
            End If
            objItem.Tag = CStr(gobjComLib.zlCommFun.NVL(rsTmp!系统, 0) & "," & rsTmp!编号)
            objItem.Visible = True
            
            K = rsTmp!标志
            Set objItem = Nothing
            
            rsTmp.MoveNext
        Next
        objForm.mnuReport.Visible = True
    End If
    
    '在CommandBar上加载要显示的报表项
    '---------------------------------------------------
    If intMenu = 2 Then
        '寻找报表菜单基项(菜单标题为报表或菜单ID=4,没有则在查看,帮助前面创建)
        K = 0
        For i = 1 To objBar.ActiveMenuBar.Controls.count
            If objBar.ActiveMenuBar.Controls(i).id = conMenu_ReportPopup _
                Or objBar.ActiveMenuBar.Controls(i).Caption Like "报表*" Then
                Set objPop = objBar.ActiveMenuBar.Controls(i)
                Exit For
            End If
            If objBar.ActiveMenuBar.Controls(i).Caption Like "查看*" Then
                K = objBar.ActiveMenuBar.Controls(i).Index
            End If
            If objBar.ActiveMenuBar.Controls(i).Caption Like "帮助*" And K = 0 Then
                K = objBar.ActiveMenuBar.Controls(i).Index
            End If
        Next
        If objPop Is Nothing Then
            'Before=0时加到最后
            Set objPop = objBar.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ReportPopup, "报表(&R)", K)
            objPop.id = conMenu_ReportPopup
        End If
        '加入报表项
        With objPop.CommandBar.Controls
            lngMenuID = conMenu_ReportPopup * 100# + .count + 1 '确定开始索引
            If lngMenuID < conMenu_Report_Item Then lngMenuID = conMenu_Report_Item
            
            K = IIf(.count = 0, 0, -1) '确定新增的第一项是否带分隔
            
            For i = 1 To rsTmp.RecordCount
                Set objControl = .Add(xtpControlButton, lngMenuID, rsTmp!名称)
                If i <= 26 Then
                    objControl.Caption = objControl.Caption & "(&" & Chr(Asc("A") + i - 1) & ")"
                End If
                objControl.Parameter = CStr(gobjComLib.zlCommFun.NVL(rsTmp!系统, 0) & "," & rsTmp!编号)
                objControl.IconId = conMenu_ReportPopup * 100# + 1 '取第一个菜单项的图标
                If K <> rsTmp!标志 And K <> 0 Then
                    objControl.BeginGroup = True
                End If
                
                K = rsTmp!标志
                lngMenuID = lngMenuID + 1
                
                rsTmp.MoveNext
            Next
        End With
    End If
    Exit Sub
errH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Sub

Public Function Event_10046_Begin() As Boolean
'功能:对当前会话启用10046事件,跟踪SQL及等待事件
'     注意,该事件启用后,对当前用户的操作速度影响明显,建议不要在正式程序中调用
'说明:在Oracle服务器的udump目录下生成.trc文件,可用tkprof工具(Oracle自带)生成分析文件
'   tkprof常用的sort选项,fchdsk可以帮助快速发现全表扫描,fchqry可以帮助快速发现索引使用不当或不良SQL
    Dim strSQL As String
    
    On Error Resume Next    '可能没有alter session权限
    strSQL = "alter session set timed_statistics=true"
    gcnOracle.Execute strSQL
    strSQL = "alter session set events '10046 trace name context forever ,level 12'"
    gcnOracle.Execute strSQL
    
    Event_10046_Begin = (Err.Number = 0)
End Function

Public Function Event_10046_End() As Boolean
'功能:关闭当前会话的10046事件跟踪
'说明:如果忘记关闭,会在本次会话结束时自动关闭
    Dim strSQL As String
    
    On Error Resume Next    '可能没有alter session权限
    strSQL = "alter session set timed_statistics=false"
    gcnOracle.Execute strSQL
    strSQL = "alter session set events '10046 trace name context off'"
    gcnOracle.Execute strSQL
    
    Event_10046_End = (Err.Number = 0)
End Function


Public Function OpenCursor(ByVal strFormCaption As String, _
                           ByVal strOwner As String, _
                           ByVal strPackagesName As String, _
                           ParamArray varParValue() As Variant) As ADODB.Recordset
'-----------------------------------------
'功能：调用存储过程返回记录集
'入参：strPackagesName ，格式为 包.过程名
'-----------------------------------------
    Dim cmdPackage As New ADODB.Command
    Dim parPackage As ADODB.Parameter
    Dim arrPar As Variant, i As Integer
    Dim varValue As Variant, intMax As Integer
    Dim intMaxArr As Integer  '记录参数个数
    Dim varOutPar As Variant
    Dim strNode As String
    '添加所有者
    If strOwner <> "" Then
        strPackagesName = strOwner & "." & strPackagesName
    End If
    '清除原有参数:不然不能重复执行
    cmdPackage.CommandText = "" '不为空有时清除参数出错
    Do While cmdPackage.Parameters.count > 0
        cmdPackage.Parameters.Delete 0
    Loop
    
    '------ IN 参数
    strNode = ""
    For i = 0 To UBound(varParValue)
        varValue = varParValue(i)
        Select Case TypeName(varValue)
            Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '数字
                cmdPackage.Parameters.Append cmdPackage.CreateParameter("P" & i, adVarNumeric, adParamInput, 30, varValue)
            Case "String" '字符
                intMax = LenB(StrConv(varValue, vbFromUnicode))
                If intMax = 0 Or intMax < 10 Then intMax = 10
                cmdPackage.Parameters.Append cmdPackage.CreateParameter("P" & i, adVarChar, adParamInput, intMax, varValue)
            Case "Date" '日期
                cmdPackage.Parameters.Append cmdPackage.CreateParameter("P" & i, adDBTimeStamp, adParamInput, , varValue)
        End Select
        strNode = strNode & CStr(varValue) & ","
    Next

    If cmdPackage.ActiveConnection Is Nothing Then
        Set cmdPackage.ActiveConnection = gcnOracle
    End If
    
    
    cmdPackage.CommandType = adCmdStoredProc
    cmdPackage.CommandText = strPackagesName
    
    Call gobjComLib.SQLTest(App.ProductName, strFormCaption, strPackagesName, strNode)
    cmdPackage.Properties("PLSQLRSet") = True
    Set OpenCursor = cmdPackage.Execute
    OpenCursor.ActiveConnection = Nothing
    Call gobjComLib.SQLTest
    cmdPackage.Properties("PLSQLRSet") = False

End Function

Public Sub RunProcedure(ByVal strFormCaption As String, _
                        ByVal strOwner As String, _
                        ByVal strProcedure As String, _
                        ParamArray varParValue() As Variant)
'-----------------------------------------
'功能：调用Oracle存储过程
'入参：ParamArray 入参数组
'-----------------------------------------
    Dim cmdPackage As New ADODB.Command
    
    Dim varValue As Variant, intMax As Integer, i As Integer
    Dim strSQL As String
    Dim strNode As String
    
    If strOwner <> "" Then
        strProcedure = strOwner & "." & strProcedure
    End If


    '清除原有参数:不然不能重复执行
    cmdPackage.CommandType = adCmdText
    
    cmdPackage.CommandText = "" '不为空有时清除参数出错
    Do While cmdPackage.Parameters.count > 0
        cmdPackage.Parameters.Delete 0
    Loop
    
    '------ IN 参数
    strSQL = "Call " & strProcedure
    For i = 0 To UBound(varParValue)
        varValue = varParValue(i)
        Select Case TypeName(varValue)
            Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '数字
                cmdPackage.Parameters.Append cmdPackage.CreateParameter("P" & i, adVarNumeric, adParamInput, 30, varValue)
            Case "String" '字符
                intMax = LenB(StrConv(varValue, vbFromUnicode)) + 100
                If intMax = 0 Or intMax < 10 Then intMax = 10
                cmdPackage.Parameters.Append cmdPackage.CreateParameter("P" & i, adVarChar, adParamInput, intMax, varValue)
            Case "Date" '日期
                cmdPackage.Parameters.Append cmdPackage.CreateParameter("P" & i, adDBTimeStamp, adParamInput, , varValue)
        End Select
        If i = 0 Then
            strSQL = strSQL & "(?,"
        Else
            strSQL = strSQL & "?,"
        End If
        
        If i = UBound(varParValue) Then
            strSQL = Mid(strSQL, 1, Len(strSQL) - 1) & ")"
        End If
        strNode = strNode & CStr(varValue) & ","
    Next
    
    '设置连接
    If cmdPackage.ActiveConnection Is Nothing Then
        Set cmdPackage.ActiveConnection = gcnOracle
    End If
    
    '执行
    cmdPackage.CommandText = strSQL
    cmdPackage.Properties("NDatatype") = True
    Call gobjComLib.SQLTest(App.ProductName, strFormCaption, strSQL, strNode)
    cmdPackage.Execute
    Call gobjComLib.SQLTest
    cmdPackage.Properties("NDatatype") = False

End Sub

Public Function GetOwner(ByVal lngSys As Long) As String
'-----------------------------------------
'功能：根据系统号取所有者
'入参：lngSys 系统号
'-----------------------------------------
    Dim rsTmp As ADODB.Recordset
    
    If lngSys >= 100 Then
        Set rsTmp = OpenCursor("clsDatabase", "zltools", "b_ComFunc.Get_Owner", lngSys)
        If rsTmp.RecordCount > 0 Then
            GetOwner = IIf(IsNull(rsTmp.Fields(0)), "", rsTmp.Fields(0))
        Else
            GetOwner = ""
        End If
    Else
        GetOwner = ""
    End If
    
End Function

Public Sub ClearParaCache()
'功能：强制清除参数缓存
    Set grsParas = Nothing
    Set grsUserParas = Nothing
End Sub

Public Function SetPara(ByVal varPara As Variant, ByVal strValue As String, Optional ByVal lngSys As Long, _
    Optional ByVal lngModual As Long, Optional ByVal blnSetup As Boolean = True) As Boolean
'功能：设置指定的参数值
'参数：varPara=参数号或参数名，以数字或字符类型传入区分
'      strValue=要设置的参数值
'      lngSys=使用该参数的系统编号，如100
'      lngModual=使用该参数的模块号，如1230
'      blnSetup=调用模块是否有参数设置权限
'返回：设置是否成功
    Dim strSQL As String
    
    '检查参数值，如果没有变化则不处理
    strSQL = GetPara(varPara, lngSys, lngModual)
    If strSQL = strValue Then SetPara = True: Exit Function

    On Error GoTo errH
        
    strSQL = "zl_Parameters_Update('" & varPara & "','" & strValue & "'," & lngSys & "," & lngModual & "," & IIf(blnSetup, 1, 0) & ")"
    Call ExecuteProcedure(strSQL, "SetPara")
    
    '设置成功后清除缓存
    Set grsParas = Nothing
    Set grsUserParas = Nothing
    
    SetPara = True
    Exit Function
errH:
    If gobjComLib.ErrCenter() = 1 Then Resume
End Function

Public Function GetPara(ByVal varPara As Variant, Optional ByVal lngSys As Long, Optional ByVal lngModual As Long, Optional ByVal strDefault As String, _
    Optional ByVal arrControl As Variant, Optional ByVal blnSetup As Boolean, Optional intType As Integer) As String
'功能：读取指定的参数值
'参数：varPara=参数号或参数名，以数字或字符类型传入区分
'      lngSys=使用该参数的系统编号，如100
'      lngModual=使用该参数的模块号，如1230
'      strDefault=当数据库中没有该参数时使用的缺省值(注意不是为空时)
'      blnNotCache=是否不从缓存中读取
'      arrControl=控件数组，如Array(Me.Text1, Me.CheckBox1)，用于函数内部自动处理对应控件的显示颜色，是否禁止设置。
'      blnSetup=调用模块是否有参数设置权限
'      intType=返回参数，返回参数类型
'返回：参数值，字符串形式
    Dim strSQL As String, i As Integer
    Dim blnNew As Boolean, blnEnabled As Boolean
    
    On Error GoTo errH
    
    intType = 0
    
    '第一次加载参数缓存
    If grsParas Is Nothing Then
        blnNew = True
    ElseIf grsParas.State = 0 Then
        blnNew = True
    End If
    If blnNew Then
        strSQL = "Select ID,Nvl(系统,0) as 系统,Nvl(模块,0) as 模块,Nvl(私有,0) as 私有,Nvl(本机,0) as 本机,Nvl(授权,0) as 授权,参数号,参数名," & _
            " Nvl(参数值,缺省值) as 参数值,User as 用户名,SYS_CONTEXT('USERENV','TERMINAL') as 机器名 From zlParameters"
        Set grsParas = New ADODB.Recordset
        Call OpenRecordset(grsParas, strSQL, "GetPara")
        
        strSQL = _
            " Select 参数ID,Nvl(用户名,'NullUser') as 用户名,Nvl(机器名,'NullMachine') as 机器名,参数值 From zlUserParas Where 用户名=User" & _
            " Union" & _
            " Select 参数ID,Nvl(用户名,'NullUser') as 用户名,Nvl(机器名,'NullMachine') as 机器名,参数值 From zlUserParas Where 机器名=SYS_CONTEXT('USERENV','TERMINAL')"
        Set grsUserParas = New ADODB.Recordset
        Call OpenRecordset(grsUserParas, strSQL, "GetPara")
    End If
    
    '使用缓存
    If TypeName(varPara) = "String" Then
        grsParas.Filter = "参数名='" & CStr(varPara) & "' And 模块=" & lngModual & " And 系统=" & lngSys
    Else
        grsParas.Filter = "参数号=" & Val(varPara) & " And 模块=" & lngModual & " And 系统=" & lngSys
    End If
    If Not grsParas.EOF Then
        '获取参数值
        If grsParas!私有 = 1 Or grsParas!本机 = 1 Then
            grsUserParas.Filter = "参数ID=" & grsParas!id & _
                IIf(grsParas!私有 = 1, " And 用户名='" & grsParas!用户名 & "'", " And 用户名='NullUser'") & _
                IIf(grsParas!本机 = 1, " And 机器名='" & grsParas!机器名 & "'", " And 机器名='NullMachine'")
            If Not grsUserParas.EOF Then
                GetPara = gobjComLib.zlCommFun.NVL(grsUserParas!参数值, strDefault)
            Else
                GetPara = gobjComLib.zlCommFun.NVL(grsParas!参数值, strDefault)
            End If
        Else
            GetPara = gobjComLib.zlCommFun.NVL(grsParas!参数值, strDefault)
        End If
        
        '返回参数类型：1.公共全局,2.私有全局,3.公共模块,4.私有模块,5.本机公共模块(不授权控制),6.本机私有模块,15.本机公共模块(要授权控制)
        If grsParas!系统 <> 0 And grsParas!模块 = 0 And grsParas!私有 = 0 And grsParas!本机 = 0 Then
            intType = 1
        ElseIf grsParas!模块 = 0 And grsParas!私有 = 1 And grsParas!本机 = 0 Then
            intType = 2
        ElseIf grsParas!系统 <> 0 And grsParas!模块 <> 0 And grsParas!私有 = 0 And grsParas!本机 = 0 Then
            intType = 3
        ElseIf grsParas!系统 <> 0 And grsParas!模块 <> 0 And grsParas!私有 = 1 And grsParas!本机 = 0 Then
            intType = 4
        ElseIf grsParas!系统 <> 0 And grsParas!模块 <> 0 And grsParas!私有 = 0 And grsParas!本机 = 1 Then
            intType = IIf(grsParas!授权 = 1, 15, 5)
        ElseIf grsParas!系统 <> 0 And grsParas!模块 <> 0 And grsParas!私有 = 1 And grsParas!本机 = 1 Then
            intType = 6
        End If
        
        '处理对应的控件颜色，可控状态
        If IsArray(arrControl) And (intType = 3 Or (intType Mod 10) = 5) Then
            blnEnabled = Not ((intType = 3 Or (intType Mod 10) = 5 And grsParas!授权 = 1) And Not blnSetup)
            For i = 0 To UBound(arrControl)
                Select Case TypeName(arrControl(i))
                Case "Label"
                    arrControl(i).ForeColor = vbBlue
                Case "TextBox", "MaskEdBox", "CheckBox", "OptionButton", "ComboBox", "ListBox", "Frame", "PictureBox", "ListView"
                    arrControl(i).ForeColor = vbBlue
                    If Not blnEnabled Then arrControl(i).Enabled = False
                Case "CommandButton", "DTPicker"
                    If Not blnEnabled Then arrControl(i).Enabled = False
                Case "MSHFlexGrid"
                    arrControl(i).ForeColor = vbBlue
                    arrControl(i).ForeColorFixed = vbBlue
                    If Not blnEnabled Then arrControl(i).Enabled = False
                Case "VSFlexGrid"
                    arrControl(i).ForeColor = vbBlue
                    arrControl(i).ForeColorFixed = vbBlue
                    If Not blnEnabled Then arrControl(i).Editable = 0
                Case Else
                    On Error Resume Next
                    arrControl(i).ForeColor = vbBlue
                    If Not blnEnabled Then arrControl(i).Enabled = False
                    Err.Clear: On Error GoTo errH
                End Select
            Next
        End If
    Else
        GetPara = strDefault
    End If
    
    Exit Function
errH:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function PatiIdentify(frmParent As Object, ByVal lngSys As Long, ByVal lng病人ID As Long, ByVal cur金额 As Currency) As Boolean
'功能：对指定病人进行刷卡身份验证
'参数：lngSys=调用程序的系统编号
'      lng病人ID=要验证身份的病人ID
'      cur金额=本次消费金额
'返回：验证成功能返回True，失败或取消返回False
'注意：1.只适用于标准版
'      2.调用该函数的地方要具有"病人信息,病人余额"的查询权限
    PatiIdentify = frmIdentify.ShowMe(frmParent, lngSys, lng病人ID, cur金额)
End Function

Public Function UserIdentify(frmParent As Object, ByVal strNote As String, ByVal lngSys As Long, ByVal lngProgID As Long, ByVal strFunc As String, Optional cnNew As ADODB.Connection) As String
'参数：strNote=提示信息(简短)
'      lngProgID=程序序号
'      strFunc=授权功能
'      cnNew=要返回的连接,需要返回时,必须传入非Nothing的对象,并且需要由调用程序关闭连接；如果是当前登录用户,返回Nothing
'返回：成功返回人员姓名
    UserIdentify = frmUserIdentify.ShowMe(frmParent, strNote, lngSys, lngProgID, strFunc, cnNew)
End Function

Private Sub Class_Terminate()
    Set mcolPatiColor = Nothing
End Sub
