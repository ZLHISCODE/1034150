Attribute VB_Name = "mdlPubJson"
Option Explicit
Private mobjServiceCall As Object
'JSON节点类型
Public Enum JSON_TYPE
    Json_Text = 0 '字符
    Json_num = 1 '数值
End Enum


Public Function zlGetNodeValueFromCollect(ByVal cllData As Collection, ByVal strKey As String, ByVal strType As String) As Variant
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定节点的数据集
    '入参:cllData-当前个集合
    '     strKey-Key
    '     strType-"N"-数字;"C"字符
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-14 16:20:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varTemp As Variant
    Err = 0: On Error Resume Next
    varTemp = cllData(strKey)
    If Err <> 0 Then
        Err = 0: On Error GoTo 0
        If strType = "N" Then zlGetNodeValueFromCollect = Empty: Exit Function
        zlGetNodeValueFromCollect = "": Exit Function
    End If
    zlGetNodeValueFromCollect = varTemp
End Function

Public Function zlGetNodeObjectFromCollect(ByVal cllData As Collection, ByVal strKey As String) As Collection
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定节点的对象集
    '入参:cllData-当前个集合
    '     strKey-Key
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-14 16:20:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllTemp As Collection
    Err = 0: On Error Resume Next
    
    Set cllTemp = cllData(strKey)
    If Err <> 0 Then
        Err = 0: On Error GoTo 0
       Set zlGetNodeObjectFromCollect = cllTemp
       Exit Function
    End If
    Set zlGetNodeObjectFromCollect = cllTemp
End Function


Public Function ToJsonStr(ByVal strValue As String) As String
'功能：处理要组合成Json串的字符串中的特殊符号
    If strValue <> "" Then
        strValue = Replace(strValue, "\", "\\")
        strValue = Replace(strValue, """", "\""")
        strValue = Replace(strValue, Chr(13), "\r")
        strValue = Replace(strValue, Chr(10), "\n")
        strValue = Replace(strValue, Chr(9), "\t")
    End If
    ToJsonStr = strValue
End Function


Public Function GetJsonNodeString(ByVal strNodeName As String, ByVal strValue As String, _
    Optional ByVal intType As JSON_TYPE, Optional ByVal blnZeroToNull As Boolean) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取Json接点串
    '入参:strNodeName-接点名
    '     strValue-值
    '     intType-类型:0-字符;1-数字
    '     blnZeroToEmpty-是否将数值0转换为Null，仅类型为数字时有效
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-09 18:59:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String
    strJson = Chr(34) & strNodeName & Chr(34)
    If intType = Json_Text Then
        strJson = strJson & ":" & Chr(34) & ToJsonStr(strValue) & Chr(34)
    Else
        If strValue = "" Or (blnZeroToNull And Val(strValue) = 0) Then
            strJson = strJson & ":null"
        Else
            strJson = strJson & ":" & IIf(Mid(strValue, 1, 1) = ".", "0", "") & strValue
        End If
    End If
    GetJsonNodeString = strJson
End Function
Public Function GetCollValue(ByVal colValue As Collection, ByVal varRow As Variant, Optional ByVal strElement As String) As Variant
    '功能：获取Json数组返回的集合数据中指定行或指定元素的值
    '参数：
    '  varRow=行索引或行关键字
    '  strElement=元素名
    '返回：
    '  当未传入strElement参数时，返回指定行的集合对象；当传入strElement参数时，返回指定行指定元素的值
    '  失败时返回Nothing或Empty，但不会报错
    If strElement <> "" Then
        GetCollValue = Empty
    Else
        Set GetCollValue = Nothing
    End If
    
    If colValue Is Nothing Then Exit Function
    
    On Error Resume Next
    If strElement <> "" Then
        GetCollValue = colValue(varRow)(strElement)
    Else
        Set GetCollValue = colValue(varRow)
    End If
    Err.Clear: On Error GoTo 0
End Function

Public Function CollectionExitsValue(ByVal coll As Collection, _
    ByVal strKey As String) As Boolean
    '根据关键字判断元素是否存在于集合中
    Dim blnExits As Boolean

    If coll Is Nothing Then Exit Function
    CollectionExitsValue = True
    Err = 0: On Error Resume Next
    blnExits = IsObject(coll(strKey))
    If Err <> 0 Then Err = 0: CollectionExitsValue = False
End Function


Public Function GetNodeString(ByVal strNodeName As String) As String
    GetNodeString = Chr(34) & strNodeName & Chr(34)
End Function


Private Function GetServiceCall(ByRef objServiceCall_Out As Object, Optional blnShowErrMsg As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取公共服务对象
    '出参:objServiceCall_Out-返回公共服务对象
    '返回:获取成功，返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-08 18:49:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strErrMsg As String
    If Not mobjServiceCall Is Nothing Then Set objServiceCall_Out = mobjServiceCall: GetServiceCall = True: Exit Function
    
    Err = 0: On Error Resume Next
    Set mobjServiceCall = CreateObject("zlServiceCall.clsServiceCall")
    If Err <> 0 Then
        strErrMsg = "部件【zlServiceCall】丢失，请与系统管理员联系，恢复该部件！"
        If blnShowErrMsg Then
            MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
            Err = 0: On Error GoTo 0
        Else
            Err.Raise Err.Number, Err.Source, strErrMsg: Exit Function
        End If
        
        Err = 0: On Error GoTo 0
        Exit Function
    End If
    
    On Error GoTo ErrHandle
    If mobjServiceCall.InitService(gcnOracle, gstrDBUser, 0, 0) = False Then Set mobjServiceCall = Nothing: Exit Function
    Set objServiceCall_Out = mobjServiceCall
    GetServiceCall = True
    Exit Function
ErrHandle:
    If blnShowErrMsg = False Then
        Err.Raise Err.Number, Err.Source, Err.Description: Exit Function
    End If
    
    MsgBox Err.Description, vbInformation, gstrSysName
End Function

Public Function zlExseSvr_UpdRgstArrangeMent(ByVal int操作类型 As Integer, ByVal lng医生ID As Long, _
                Optional ByVal str撤档时间 As String, Optional ByRef strErrMsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:调整号源、有效的安排、有效的出诊记录中的医生姓名。
    '入参:int操作类型-1-修改姓名,2-停用人员,3-启用人员
    '     str撤档时间-停用和启用时传入，启用时传入原撤档时间
    '出参:strErrMsg_Out
    '返回:获取成功返回True，获取失败返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objServiceCall As Object
    Dim intReturn As Integer
    Dim strJson As String
    
    On Error GoTo ErrHandler
    If GetServiceCall(objServiceCall) = False Then
        MsgBox "连接费用服务失败！", vbInformation, gstrSysName
        Exit Function
    End If
    
'    Zl_ExseSvr_UpdRgstArrangement
'    --功能：调整号源、有效的安排、有效的出诊记录中的医生姓名。
'    --入参
'    --input      调整号源、有效的安排、有效的出诊记录中的医生姓名
'    --  oper_type     N  1  操作方式：1-修改姓名,2-停用人员,3-启用人员
'    --  rgst_dr_id      N  1  病人id
'    --  revoke_time   C         撤档时间
'    --出参
'    --output
'    --  code          C    1  应答码：0-失败；1-成功
'    --  message         C  1  应答消息：成功时返回成功信息，失败时返回具体的错误信息

    strJson = ""
    strJson = strJson & "" & GetJsonNodeString("oper_type", int操作类型, Json_num)
    strJson = strJson & "," & GetJsonNodeString("rgst_dr_id", lng医生ID, Json_num)
    If str撤档时间 <> "" Then
        strJson = strJson & "," & GetJsonNodeString("revoke_time", str撤档时间, Json_Text)
    End If
    strJson = "{""input"":{" & strJson & "}}"
    
    If objServiceCall.CallService("zl_ExseSvr_UpdRgstArrangement", strJson, , "", 0, False) = False Then Exit Function
    intReturn = Val(objServiceCall.GetJsonNodeValue("output.code"))
    If intReturn <> 1 Then
        strErrMsg_Out = objServiceCall.GetJsonNodeValue("output.message")
        If strErrMsg_Out <> "" Then strErrMsg_Out = "更新挂号安排失败！"
        Exit Function
    End If
    
    zlExseSvr_UpdRgstArrangeMent = True
    Exit Function
ErrHandler:
    strErrMsg_Out = Err.Description
End Function
