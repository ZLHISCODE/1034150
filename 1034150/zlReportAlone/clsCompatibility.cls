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

Private Const MSTR_CR As String = " " & vbCrLf                                      'SQL����

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
'���ܣ�����Insert���
'������
'  strTable������
'  strFieldInfo���ֶ���Ϣ����ʽ���ֶ���1=�ֶ�ֵ[,�ֶ���n=�ֶ�ֵn]
'���أ�True�ɹ���Falseʧ��

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
'���ܣ�����Update���
'������
'���أ�True�ɹ���Falseʧ��
    
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
'���ܣ�����Select���
'������
'���أ�True�ɹ���Falseʧ��
    
    Dim arrFields() As String, arrExists() As Variant
    Dim strFieldName As String, strNewWhere As String
    Dim i As Integer, intPos As Integer
    
    Call TablesAdd(strTable, mdicTableFields)
    If Not mdicTableFields.Exists(UCase$(strTable)) Then Exit Function
    
    'Selectת����
    arrFields = Split(strSelect, ",")
    
    '���Select�����ݿ���û���ֶ�
    arrExists = Array()
    For i = LBound(arrFields) To UBound(arrFields)
        strFieldName = Trim$(UCase$(arrFields(i)))
        intPos = InStr(strFieldName, " ")
        If intPos > 0 Then
            strFieldName = left$(strFieldName, intPos - 1)
        End If
        '����ֶ�
        ReDim Preserve arrExists(UBound(arrExists) + 1)
        If InStr(";ROWNUM;ROWID;1;", strFieldName) > 0 Then
            '�����ֶ���
            arrExists(UBound(arrExists)) = arrFields(i)
        ElseIf FieldExists(strTable, strFieldName) Then
            '���ڵ��ֶ���
            arrExists(UBound(arrExists)) = arrFields(i)
        Else
            If blnNeedNull Then
                '�����ڵ��ֶ���ǿ���Ա�����ʽ����
                arrExists(UBound(arrExists)) = "Null " & arrFields(i)
            Else
                '�����ڵ��ַ������ռ�
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
'���ܣ����ɱ���ʽ
'������
'  enmType��ÿ������ʽ�ķָ��ؼ���
'���أ�True�ɹ���Falseʧ��
'ע�⣺�ֶα�������=�����

    Dim arrExpressionSub() As String, arrNew() As Variant
    Dim strSub As String, strFieldName As String, strSplitString As String, strTmp As String
    Dim lngPos As Long
    Dim i As Integer
    
    strNew = strWhere
    If Trim$(strWhere) = "" Then Exit Function
    
    Call TablesAdd(strTable, mdicTableFields)
    If Not mdicTableFields.Exists(UCase$(strTable)) Then Exit Function
    
    '���ַָ���
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
                '���ں������ֶε�д��
                strTmp = Mid$(strFieldName, lngPos + 1)
                lngPos = InStr(strTmp, ",")
                If lngPos <= 0 Then lngPos = InStr(strTmp, ")")
                If lngPos <= 0 Then lngPos = Len(strTmp) + 1
                strFieldName = left$(strTmp, lngPos - 1)
            End If
            If FieldExists(strTable, strFieldName) Then
                '�ӱ���ʽ�ֶδ��ھ��ռ�
                ReDim Preserve arrNew(UBound(arrNew) + 1)
                arrNew(UBound(arrNew)) = arrExpressionSub(i)
            End If
        Else
            '�޵Ⱥ�ֱ���ռ�
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
'���ܣ��ж��ֶ��Ƿ����
'������
'���أ�True���ڣ�False������

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
'���ܣ��ռ�������Ϣ���Ҵ���Dictionary����
'������

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
            '�ռ���Ϣ
            strSQL = "Select * From zlTools." & strTableName & " Where Rownum < 1"
            Set rsTmp = mdlPublic.OpenSQLRecord(strSQL, "�ռ�" & strTableName & "�����ֶ���Ϣ")
            objDic.Add strTableName, rsTmp
        End If
    End If
End Sub

