VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCompatibility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const MSTR_CR As String = " " & vbCrLf                                      'SQL拆行

Private mdicTableFields As Scripting.Dictionary

Public Enum enmSplitType
    sltAndString = 0
    sltCommaSymbol = 1
End Enum

Private Sub Class_Initialize()
    '
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    Set mdicTableFields = Nothing
End Sub

Public Function BuildInsertString(ByVal strTable As String, ByVal strFieldInfo As String _
    , ByRef strSQL As String) As Boolean
'功能：生成Insert语句
'参数：
'  strTable：表名
'  strFieldInfo：字段信息。格式：字段名1=字段值[,字段名n=字段值n]
'返回：True成功；False失败

    Dim arrInfo() As String
    Dim strNewInfo As String, strFieldName As String, strFieldValue As String, strTemp As String
    Dim i As Integer, intPos As Integer

    If Trim$(strFieldInfo) = "" Then Exit Function
    
    Call TablesAdd(strTable, mdicTableFields)
    If Not mdicTableFields.Exists(UCase$(strTable)) Then Exit Function
    
    strNewInfo = ""
    Call BuildExpressionString(strTable, strFieldInfo, strNewInfo, sltCommaSymbol)
    If Trim$(strNewInfo) = "" Then Exit Function
    
    strFieldName = ""
    strFieldValue = ""
    arrInfo = Split(strNewInfo, ", ")
    For i = LBound(arrInfo) To UBound(arrInfo)
        strTemp = arrInfo(i)
        intPos = InStr(strTemp, "=")
        If intPos <= 0 Then
            strSQL = ""
            Exit Function
        Else
            strFieldName = strFieldName & ", " & left$(strTemp, intPos - 1)
            strFieldValue = strFieldValue & ", " & Mid$(strTemp, intPos + 1)
        End If
    Next
    If strFieldName <> "" Then strFieldName = Mid$(strFieldName, 3)
    If strFieldValue <> "" Then strFieldValue = Mid$(strFieldValue, 3)
    
    strSQL = "Insert Into " & strTable & MSTR_CR & _
             "(" & strFieldName & ")" & MSTR_CR & _
             "Values (" & strFieldValue & ")"
    
    BuildInsertString = True
End Function

Public Function BuildUpdateString(ByVal strTable As String, ByVal strSet As String, ByVal strWhere As String _
    , ByRef strSQL As String) As Boolean
'功能：生成Update语句
'参数：
'返回：True成功；False失败
    
    Dim strNewSet As String, strNewWhere As String
    
    If Trim$(strSet) = "" Or Trim$(strWhere) = "" Then Exit Function
    
    Call TablesAdd(strTable, mdicTableFields)
    If Not mdicTableFields.Exists(UCase$(strTable)) Then Exit Function
    
    strNewSet = ""
    Call BuildExpressionString(strTable, strWhere, strNewSet, sltCommaSymbol)
    If strNewSet = "" Then Exit Function
    
    strNewWhere = ""
    Call BuildExpressionString(strTable, strWhere, strNewWhere, sltCommaSymbol)
    If strNewWhere = "" Then Exit Function
    
    strSQL = "Update " & strTable & MSTR_CR & _
             "Set " & strNewSet & MSTR_CR & _
             "Where " & strNewWhere
             
    BuildUpdateString = True
End Function

Public Function BuildSelectString(ByVal strTable As String, ByVal strSelect As String, ByVal strWhere As String _
    , ByRef strSQL As String, ByVal blnNeedNull As Boolean) As Boolean
'功能：生成Select语句
'参数：
'返回：True成功；False失败
    
    Dim arrFields() As String, arrExists() As Variant
    Dim strFieldName As String, strNewWhere As String
    Dim i As Integer, intPos As Integer
    
    Call TablesAdd(strTable, mdicTableFields)
    If Not mdicTableFields.Exists(UCase$(strTable)) Then Exit Function
    
    'Select转数组
    arrFields = Split(strSelect, ",")
    
    '清除Select在数据库中没有字段
    arrExists = Array()
    For i = LBound(arrFields) To UBound(arrFields)
        strFieldName = Trim$(UCase$(arrFields(i)))
        intPos = InStr(strFieldName, " ")
        If intPos > 0 Then
            strFieldName = left$(strFieldName, intPos - 1)
        End If
        '检查字段
        ReDim Preserve arrExists(UBound(arrExists) + 1)
        If InStr(";ROWNUM;ROWID;1;", strFieldName) > 0 Then
            '特殊字段名
            arrExists(UBound(arrExists)) = arrFields(i)
        ElseIf FieldExists(strTable, strFieldName) Then
            '存在的字段名
            arrExists(UBound(arrExists)) = arrFields(i)
        Else
            If blnNeedNull Then
                '不存在的字段名强制以别名形式补充
                arrExists(UBound(arrExists)) = "Null " & arrFields(i)
            Else
                '不存在的字符名不收集
                ReDim Preserve arrExists(UBound(arrExists) - 1)
            End If
        End If
    Next
    
    strSQL = "Select " & Join(arrExists, ", ") & MSTR_CR & _
             "From " & strTable

    If Trim$(strWhere) <> "" Then
        strNewWhere = ""
        Call BuildExpressionString(strTable, strWhere, strNewWhere)
        If strNewWhere <> "" Then
            strSQL = strSQL & MSTR_CR & "Where " & strNewWhere
        End If
    End If
    
    BuildSelectString = True
End Function

Public Function BuildExpressionString(ByVal strTable As String, ByVal strWhere As String, ByRef strNew As String _
    , Optional ByVal enmType As enmSplitType = sltAndString) As Boolean
'功能：生成表达式
'参数：
'  enmType：每个表达式的分隔关键字
'返回：True成功；False失败
'注意：字段必须置于=的左边

    Dim arrExpressionSub() As String, arrNew() As Variant
    Dim strSub As String, strFieldName As String, strSplitString As String, strTmp As String
    Dim lngPos As Long
    Dim i As Integer
    
    strNew = strWhere
    If Trim$(strWhere) = "" Then Exit Function
    
    Call TablesAdd(strTable, mdicTableFields)
    If Not mdicTableFields.Exists(UCase$(strTable)) Then Exit Function
    
    '区分分隔串
    If enmType = sltCommaSymbol Then
        strSplitString = ","
        arrExpressionSub = Split(strWhere, strSplitString)
    Else
        strSplitString = " AND "
        arrExpressionSub = Split(UCase$(strWhere), strSplitString)
    End If
    
    arrNew = Array()
    For i = LBound(arrExpressionSub) To UBound(arrExpressionSub)
        strSub = Trim$(arrExpressionSub(i))
        lngPos = InStr(strSub, "=")
        If lngPos > 0 Then
            strFieldName = left$(strSub, lngPos - 1)
            lngPos = InStr(strFieldName, "(")
            If lngPos > 0 Then
                '存在函数套字段的写法
                strTmp = Mid$(strFieldName, lngPos + 1)
                lngPos = InStr(strTmp, ",")
                If lngPos <= 0 Then lngPos = InStr(strTmp, ")")
                If lngPos <= 0 Then lngPos = Len(strTmp) + 1
                strFieldName = left$(strTmp, lngPos - 1)
            End If
            If FieldExists(strTable, strFieldName) Then
                '子表达式字段存在就收集
                ReDim Preserve arrNew(UBound(arrNew) + 1)
                arrNew(UBound(arrNew)) = arrExpressionSub(i)
            End If
        Else
            '无等号直接收集
            If strSub Like "*(*)*" Or strSub Like "*(*" Then
                ReDim Preserve arrNew(UBound(arrNew) + 1)
                arrNew(UBound(arrNew)) = arrExpressionSub(i)
            Else
                If UBound(arrNew) < 0 Then
                    ReDim Preserve arrNew(UBound(arrNew) + 1)
                End If
                arrNew(UBound(arrNew)) = arrNew(UBound(arrNew)) & arrExpressionSub(i)
            End If
        End If
    Next
    
    If UBound(arrNew) >= 0 Then
        strNew = Join(arrNew, strSplitString)
    Else
        strNew = ""
    End If
    
    BuildExpressionString = True
End Function

Public Function FieldExists(ByVal strTable As String, ByVal strField As String) As Boolean
'功能：判断字段是否存在
'参数：
'返回：True存在；False不存在

    Dim rsTable As ADODB.Recordset
    Dim i As Integer
    
    FieldExists = False
    strField = Replace(Replace(Replace(strField, vbCrLf, ""), vbCr, ""), vbLf, "")
    strField = UCase$(Trim$(strField))
    strTable = UCase$(Trim$(strTable))
    
    Call TablesAdd(strTable, mdicTableFields)
    If Not mdicTableFields.Exists(strTable) Then Exit Function
    
    Set rsTable = mdicTableFields.Item(UCase$(strTable))
    For i = 0 To rsTable.Fields.Count - 1
        If UCase$(Trim$(rsTable.Fields(i).name)) = strField Then
            FieldExists = True
            Exit For
        End If
    Next
End Function
 
Private Sub TablesAdd(ByVal strTableName As String, ByRef objDic As Scripting.Dictionary)
'功能：收集表的信息并且存入Dictionary对象
'参数：

    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    If Trim$(strTableName) = "" Then Exit Sub
    
    strTableName = UCase$(strTableName)
    
    On Error Resume Next
    If Err.Number = 0 Then
        If objDic Is Nothing Then
            Set objDic = New Scripting.Dictionary
        End If
        If Not objDic.Exists(strTableName) Then
            '收集信息
            strSQL = "Select * From zlTools." & strTableName & " Where Rownum < 1"
            Set rsTmp = mdlPublic.OpenSQLRecord(strSQL, "收集" & strTableName & "表的字段信息")
            objDic.Add strTableName, rsTmp
        End If
    End If
End Sub


