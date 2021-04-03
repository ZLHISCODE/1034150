VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'##############################################################################
'#���ܣ�ZLHISϵͳ������������
'�������򣺺�����������Ϊͨ���׶���Ӣ�ļ���
'�ϼ��ࣺclsDatabase
'##############################################################################
Private mobjFIle As New FileSystemObject

Public Function IsMovedByNO(ByVal strTableName As String, ByVal strNO As String, Optional ByVal strIF As String = "��¼���� IN", Optional ByVal strIFValue As String = "", Optional ByVal strFormCaption As String = "ZL9ComLib.NOMoved", Optional ByVal blnOnline As Boolean = False) As Boolean
'����:����ָ���ı��������ݺš���¼����,�ж�ָ���ĵ��ݺ��Ƿ���ת���������ݱ���
'����:NO������:F0000007����:Between 'F0000001' And 'F0000007',��:In('F0000001','F0000002')
'       strIFValue�ڷ���ϵͳ��ָ��¼����,������1,2������ʽ
'       strIF,strIFValue:�����ݺ���ĸ�������������ֵ,Ĭ��Ϊ����ϵͳ������,���������������ʹ��
'����:
'     ע��:������NOMoved�е�NO��ʾ���ݺ�NO,������no,��Ҫ����Ϊû�б�ת��,��ȷ�ĺ�����:ָ���ĵ��ݱ�ת��,��Ϊ��,����Ϊ��,����Ҳ���ؼ�.
    IsMovedByNO = gobjComLib.zlDatabase.NOMoved(strTableName, strNO, strIF, strIFValue, strFormCaption, blnOnline)
End Function

Public Function IsMovedByDate(ByVal strDate As String, Optional ByVal int��� As Integer = 1, Optional ByVal intϵͳ As Integer = 100, Optional ByVal strFormCaption As String = "ZL9ComLib.DateMoved") As Boolean
'����:����ָ�������ڡ���š�ϵͳ�ж�ָ�����ڵ������Ƿ���ת���������ݱ���
     IsMovedByDate = gobjComLib.zlDatabase.DateMoved(strDate, int���, intϵͳ, strFormCaption)
End Function

Public Function IsMovedByTable(ByVal strTableName As String, ByVal strIF As String, ByVal strIFValue As String, Optional ByVal strFormCaption As String = "ZL9ComLib.NOMoved", Optional ByVal blnOnline As Boolean = False) As Boolean
'����:����ָ���ı������ж�ת���������ݱ���
'       strIF,strIFValue:�����ݺ���ĸ�������������ֵ,Ĭ��Ϊ����ϵͳ������,���������������ʹ��
'����:
    IsMovedByTable = gobjComLib.zlDatabase.TableDataMoved(strTableName, strIF, strIFValue, strFormCaption, blnOnline)
End Function

Public Function ReturnMovedClinic(ByVal intPatientID As Integer, ByVal strTimes As String, Optional ByVal bytFlag As Byte = 1, Optional ByVal strFormCaption As String = "ZL9ComLib.ReturnMovedClinic") As Boolean
'����:��ѡ���ز���ĳ������סԺҽ������(Oracle����:Zl_Retu_Clinic)
'����:  Times,�Һŵ��Ż�סԺ��ҳid
'       Flag,�����סԺ��־:0-����,1-סԺ
    ReturnMovedClinic = gobjComLib.zlDatabase.ReturnMovedClinic(intPatientID, strTimes, bytFlag, strFormCaption)
End Function

Public Function ReturnMovedExes(ByVal strNO As String, Optional ByVal bytType As Byte = 1, Optional ByVal strFormCaption As String = "ZL9ComLib.ReturnMovedExes") As Boolean
'����:��ѡ����ָ�����õ��ݺŵ�����(Oracle����:Zl_Retu_Exes)
'����:NO���ݺ�,bytType��ʾ��������,ֵ::1-�շ�,2-����,3-�Զ�����,4-�Һ�,5-���￨,6-Ԥ��,7-���ʣ�
    ReturnMovedExes = gobjComLib.zlDatabase.ReturnMovedExes(strNO, bytType, strFormCaption)
End Function

Public Function ReadLob(ByVal lngSys As Long, ByVal Action As Long, ByVal KeyWord As String, _
                        Optional ByVal strFile As String, Optional ByVal bytFunc As Byte = 0, _
                        Optional bytMoved As Byte = 0) As String
'���ܣ���ָ����LOB�ֶθ���Ϊ��ʱ�ļ�
'������
'lngSys:ϵͳ���
'Action:�������ͣ����������ǲ����ĸ�����
'---ϵͳ100,Zl_Lob_Append
'0-�������ͼ��;1-�����ļ���ʽ;2-�����ļ�ͼ��;3-�������ĸ�ʽ;4-��������ͼ��;
'5-���Ӳ�����ʽ;6-���Ӳ���ͼ��;7-����ҳ���ʽ(ͼ��)��8-���Ӳ�������;9-�����ص����
'10-�ٴ�·���ļ�,11-�ٴ�·��ͼ��;14-��Ա֤���¼;15-��Ա��;16-��Ա��Ƭ;
'17-ҩƷ���(ʹ��˵��);18-ҩƷ���(ͼƬ);23-��Ӧ��ͼƬ
'---ϵͳ2400,Zl24_Lob_Append
'���鳣��ͼ��,��Action
'---ϵͳ2100,Zl21_Lob_Append
'1-�������͵���;2-���������(��ͼƬֻ�ж�ȡ��û�б���);3-����걨��¼;4-���������Ա,5-���������
'---ϵͳ2600,Zl26_Lob_Append
'14-����ؼ�Ŀ¼,15-������ԴĿ¼
'      KeyWord:ȷ�����ݼ�¼�Ĺؼ��֣����Ϲؼ����Զ��ŷָ�(��5-���Ӳ�����ʽΪ����)
'      strFile:�û�ָ����ŵ��ļ�������ָ��ʱ���Զ�ȡ��ʱ�ļ���
'bytFunc-0-BLOB,1-CLOB
'bytMoved=0������¼,1��ȡת���󱸱���¼
'���أ�������ݵ��ļ�����ʧ���򷵻��㳤��""
    Const conChunkSize As Integer = 10240
    
    Dim rsLob As ADODB.Recordset
    Dim lngFileNum As Long, lngCount As Long, lngBound As Long
    Dim aryChunk() As Byte, strText As String
    Dim strSQL As String
    Dim objFile As New FileSystemObject
    
    Err = 0: On Error GoTo Errhand
    Select Case lngSys \ 100
        Case 1
            strSQL = "Select Zl_Lob_Read([1],[2],[3],[4],[5]) as Ƭ�� From Dual"
        Case 24
            strSQL = "Select Zl24_Lob_Read([2],[3]) as Ƭ�� From Dual"
        Case 21
            strSQL = "Select Zl21_Lob_Read([1],[2],[3]) as Ƭ�� From Dual"
        Case 26
            strSQL = "Select Zl26_Lob_Read([1],[2],[3]) as Ƭ�� From Dual"
    End Select
    If strSQL = "" Then strFile = "": Exit Function
    If bytFunc = 0 Then 'BLOB
        If strFile = "" Then
            strFile = objFile.GetSpecialFolder(TemporaryFolder) & "\" & objFile.GetTempName
        End If
        lngFileNum = FreeFile
        Open strFile For Binary As lngFileNum
        lngCount = 0
        Do
            Set rsLob = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "zllobRead", Action, KeyWord, lngCount, bytMoved, bytFunc)
            If rsLob.EOF Then Exit Do
            If IsNull(rsLob.Fields(0).value) Then Exit Do
            strText = rsLob.Fields(0).value
            
            ReDim aryChunk(Len(strText) / 2 - 1) As Byte
            For lngBound = LBound(aryChunk) To UBound(aryChunk)
                aryChunk(lngBound) = CByte("&H" & Mid(strText, lngBound * 2 + 1, 2))
            Next
            
            Put lngFileNum, , aryChunk()
            lngCount = lngCount + 1
        Loop
        Close lngFileNum
        If lngCount = 0 Then Kill strFile: strFile = ""
    Else  'CLOB
        lngCount = 0
        strFile = ""
        Do
            Set rsLob = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "zllobRead", Action, KeyWord, lngCount, bytMoved, bytFunc)
            If rsLob.EOF Then Exit Do
            If IsNull(rsLob.Fields(0).value) Then Exit Do
            strText = rsLob.Fields(0).value
            strFile = strFile & strText
            lngCount = lngCount + 1
        Loop
    End If
    ReadLob = strFile
    Exit Function
Errhand:
    If bytFunc = 0 Then
        Close lngFileNum
        Kill strFile: ReadLob = ""
    End If
    Err.Clear
End Function

Public Function SaveLob(ByVal lngSys As Long, ByVal Action As Long, ByVal KeyWord As String, ByVal strFile As String, Optional ByVal bytFunc As Byte = 0) As Boolean
'���ܣ�����ָ�����ļ���ָ������¼BLOB�ֶ�
'������
'lngSys:ϵͳ���
'Action:�������ͣ����������ǲ����ĸ�����
'---ϵͳ100,Zl_Lob_Append
'0-�������ͼ��;1-�����ļ���ʽ;2-�����ļ�ͼ��;3-�������ĸ�ʽ;4-��������ͼ��;
'5-���Ӳ�����ʽ;6-���Ӳ���ͼ��;7-����ҳ���ʽ(ͼ��)��8-���Ӳ�������;9-�����ص����
'10-�ٴ�·���ļ�,11-�ٴ�·��ͼ��;14-��Ա֤���¼;15-��Ա��;16-��Ա��Ƭ;
'17-ҩƷ���(ʹ��˵��);18-ҩƷ���(ͼƬ);23-��Ӧ��ͼƬ
'---ϵͳ2400,Zl24_Lob_Append
'���鳣��ͼ��,��Action
'---ϵͳ2100,Zl21_Lob_Append
'1-�������͵���;2-���������(��ͼƬֻ�ж�ȡ��û�б���);3-����걨��¼;4-���������Ա,5-���������
'---ϵͳ2600,Zl26_Lob_Append
'14-����ؼ�Ŀ¼,15-������ԴĿ¼
'      KeyWord:ȷ�����ݼ�¼�Ĺؼ��֣����Ϲؼ����Զ��ŷָ�(��5-���Ӳ�����ʽΪ����)
'      strFile:�û�ָ����ŵ��ļ�������ָ��ʱ���Զ�ȡ��ʱ�ļ���
'      KeyWord:ȷ�����ݼ�¼�Ĺؼ��֣����Ϲؼ����Զ��ŷָ�
'      strFile:�û�ָ��Ҫ��ŵ��ļ���
'      arySql():�ڸ����ݵĻ�������չ���ӱ����SQL���
'      bytFunc=0-BLOB,1-CLOB
'���أ��ɹ�����True��ʧ�ܷ���False
    Dim arrSQL() As String
    Dim i As Long
    
    If GetLobSql(lngSys, Action, KeyWord, strFile, arrSQL, bytFunc) Then
        Call gobjComLib.zlDatabase.ExecuteProcedureBeach(arrSQL, "zllobSave", False, False)
    Else
        SaveLob = False
    End If
    SaveLob = True
    Exit Function
Errhand:
    Err.Clear
    SaveLob = False
End Function

Public Function GetLobSql(ByVal lngSys As Long, ByVal Action As Long, ByVal KeyWord As String, _
                ByVal strFile As String, ByRef arySql() As String, Optional ByVal bytFunc As Byte = 0) As Boolean
'���ܣ���������ָ�����ļ���ָ������¼BLOB/CLOB�ֶε�SQL���
'������
'lngSys:ϵͳ���
'Action:�������ͣ����������ǲ����ĸ�����
'---ϵͳ100,Zl_Lob_Append
'0-�������ͼ��;1-�����ļ���ʽ;2-�����ļ�ͼ��;3-�������ĸ�ʽ;4-��������ͼ��;
'5-���Ӳ�����ʽ;6-���Ӳ���ͼ��;7-����ҳ���ʽ(ͼ��)��8-���Ӳ�������;9-�����ص����
'10-�ٴ�·���ļ�,11-�ٴ�·��ͼ��;14-��Ա֤���¼;15-��Ա��;16-��Ա��Ƭ;
'17-ҩƷ���(ʹ��˵��);18-ҩƷ���(ͼƬ);23-��Ӧ��ͼƬ
'---ϵͳ2400,Zl24_Lob_Append
'���鳣��ͼ��,��Action
'---ϵͳ2100,Zl21_Lob_Append
'1-�������͵���;2-���������(��ͼƬֻ�ж�ȡ��û�б���);3-����걨��¼;4-���������Ա,5-���������
'---ϵͳ2600,Zl26_Lob_Append
'14-����ؼ�Ŀ¼,15-������ԴĿ¼
'      KeyWord:ȷ�����ݼ�¼�Ĺؼ��֣����Ϲؼ����Զ��ŷָ�(��5-���Ӳ�����ʽΪ����)
'      strFile:�û�ָ����ŵ��ļ�������ָ��ʱ���Զ�ȡ��ʱ�ļ���
'      KeyWord:ȷ�����ݼ�¼�Ĺؼ��֣����Ϲؼ����Զ��ŷָ�
'      strFile:BLOBʱ,�û�ָ��Ҫ��ŵ��ļ���;CLOBʱ,��Ҫ�洢���ı�����
'      arySql():�ڸ����ݵĻ�������չ���ӱ����SQL���
'      bytFunc=0-BLOB,1-CLOB
'���أ��ɹ�����True��ʧ�ܷ���False
    Dim conChunkSize As Integer
    Dim lngFileSize As Long, lngCurSize As Long, lngModSize As Long
    Dim lngBlocks As Long, lngFileNum As Long
    Dim lngCount As Long, lngBound As Long
    Dim aryChunk() As Byte, aryHex() As String, strText As String
    Dim strSQLRoot As String
    Dim lngLBound As Long, lngUbound As Long    '�����������С����±�
    Dim strSubTxt As String
    
    Err = 0: On Error Resume Next
    lngLBound = LBound(arySql): lngUbound = UBound(arySql)
    If Err <> 0 Then lngLBound = 0: lngUbound = -1
    Err = 0: On Error GoTo 0
    Select Case lngSys \ 100
        Case 1
            strSQLRoot = "Zl_Lob_Append(" & Action & ",'" & KeyWord & "'"
        Case 24
            strSQLRoot = "Zl24_Lob_Append('" & KeyWord & "'"
        Case 21
            strSQLRoot = "Zl21_Lob_Append(" & Action & ",'" & KeyWord & "'"
        Case 26
            strSQLRoot = "Zl26_Lob_Append(" & Action & ",'" & KeyWord & "'"
    End Select
    If strSQLRoot = "" Then GetLobSql = False: Exit Function
    If bytFunc = 0 Then  'BLOB
        lngFileNum = FreeFile
        Open strFile For Binary Access Read As lngFileNum
        lngFileSize = LOF(lngFileNum)
        
        Err = 0: On Error GoTo Errhand
        conChunkSize = 2000
        lngModSize = lngFileSize Mod conChunkSize
        lngBlocks = lngFileSize \ conChunkSize - IIf(lngModSize = 0, 1, 0)
        ReDim Preserve arySql(lngLBound To lngUbound + lngBlocks + 1)
        For lngCount = 0 To lngBlocks
            If lngCount = lngFileSize \ conChunkSize Then
                lngCurSize = lngModSize
            Else
                lngCurSize = conChunkSize
            End If
            
            ReDim aryChunk(lngCurSize - 1) As Byte
            ReDim aryHex(lngCurSize - 1) As String
            Get lngFileNum, , aryChunk()
            For lngBound = LBound(aryChunk) To UBound(aryChunk)
                aryHex(lngBound) = Hex(aryChunk(lngBound))
                If Len(aryHex(lngBound)) = 1 Then aryHex(lngBound) = "0" & aryHex(lngBound)
            Next
            strText = Join(aryHex, "")
            arySql(lngUbound + lngCount + 1) = strSQLRoot & ",'" & strText & "'," & IIf(lngCount = 0, 1, 0) & ")"
        Next
        Close lngFileNum
    Else  'CLOB
        conChunkSize = 2000
        strText = strFile
        lngCount = 0
        Do
            strSubTxt = Left(strText, conChunkSize)
            strText = Mid(strText, conChunkSize + 1)
            ReDim Preserve arySql(lngCount)
            arySql(lngCount) = strSQLRoot & ",'" & strSubTxt & "'," & IIf(lngCount = 0, 1, 0) & ",1)"
            lngCount = lngCount + 1
        Loop While Len(strText) > 0
    
    End If
    GetLobSql = True
    Exit Function
Errhand:
    Err.Clear
    Close lngFileNum
    GetLobSql = False
End Function

Public Function ReadXML(ByVal strTable As String, ByVal strXMLColumn As String, ByVal strWhereFilter As String, ByRef strErrInfo As String, ParamArray arrInput() As Variant) As String
'���ܣ���ȡXML�ֶβ������ַ���
'������
'strTable:XML�ֶ����ڵı�
'strXMLColumn:XML�ֶ���
'strWhereFilter:Where�����������а���������SQL���,������ʽΪ"[x]"
'             x>=1Ϊ�Զ��������,"[]"֮�䲻���пո�
'             ͬһ�������ɶദʹ��,�����Զ���ΪADO֧�ֵ�"?"����ʽ
'             ʵ��ʹ�õĲ����ſɲ�����,������Ĳ���ֵ��������(��SQL���ʱ��һ��Ҫ�õ��Ĳ���)
'      arrInput=���������Ĳ���ֵ,��������˳�����δ���,��������ȷ����
'               ��Ϊʹ�ð󶨱���,�Դ�"'"���ַ�����,����Ҫʹ��"''"��ʽ��
'strErrInfo=���صĴ�����Ϣ��
'           ���м�¼ʱ�����ش���
'���أ���ȡ��XML�ֶ�
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim arrPars() As Variant, intLoop As Long
    Dim lngLength As Long
    
    On Error GoTo errH
    ReadXML = ""
    If strWhereFilter <> "" Then strWhereFilter = " Where " & strWhereFilter
    arrPars = arrInput
    strSQL = "Select length(" & strXMLColumn & ".getclobval()) As ValueLength" & vbNewLine & _
            "  From " & strTable & strWhereFilter
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecordByArray(strSQL, "��ȡXML�ֶ�ֵ", arrPars)
    If rsTmp.RecordCount = 1 Then
        lngLength = Val(rsTmp!ValueLength & "")
        If lngLength > 2000 Then
            For intLoop = 1 To lngLength / 2000 + 1
                strSQL = "select To_Char(substr(" & strXMLColumn & ".getclobval(), " & CDbl(intLoop) * 2000 - 1999 & ", 2000)) as PartValue" & vbNewLine & _
                        "  from " & strTable & strWhereFilter
                Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecordByArray(strSQL, "��ȡXML�ֶ�ֵ", arrPars)
                If Not rsTmp.EOF Then
                    ReadXML = ReadXML & rsTmp!PartValue
                End If
            Next
        Else
            strSQL = "select " & strXMLColumn & ".getstringval() as PartValue" & vbNewLine & _
                    "  from " & strTable & strWhereFilter
            Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecordByArray(strSQL, "��ȡXML�ֶ�ֵ", arrPars)
            If Not rsTmp.EOF Then
                ReadXML = ReadXML & rsTmp!PartValue
            End If
        End If
    ElseIf rsTmp.RecordCount > 1 Then '��������XML�ֶ�ֻ�ܵ�������
        strErrInfo = "���������ض�������"
    End If
    Exit Function
errH:
    strErrInfo = Err.Description
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Function


Public Function RowValue(ByVal strTable As String, Optional ByVal arrValues As Variant, Optional ByVal strGetFields As String, Optional ByVal strWhereField As String = "ID") As Variant
'���ܣ���ȡָ����ָ���ֶ���Ϣ
'������strTable=��ȡ���ݵı�
'          arrValues=����ֵ�����Դ����飬Ҳ���Դ�����ֵ��Ҳ���Բ�����������ȡȫ��
'          strGetField=��ȡ���ֶ�,����ֶ��Զ��ŷָͬSQL��д��ȡ�ֶ�һ��
'          strWhereField=���˵��ֶΣ����ֶ�Ϊ�򵥵���ֵ���ַ����ͻ��������ͣ����������޷�֧��
'���أ�
'ֻ������һ����������ض���һ��ֵ��δ����NULLֵ����
'      strGetField=�����ֶ�
'      arrValues=Ϊ����ֵ���򲻸���һ��Ԫ�ص�����
'������������ؼ�¼��

    Dim rsTmp As New ADODB.Recordset, blnReturnRec As Boolean
    Dim strSQL As String
    Dim strWhere As String
    Dim arrPars As Variant
    Dim i As Long, strPars As String
    
    On Error GoTo errH
    blnReturnRec = True
    If TypeName(arrValues) = "Variant()" Then
        arrPars = arrValues
        For i = LBound(arrValues) To UBound(arrValues)
            strPars = strPars & ",[" & i + 1 & "]"
        Next
        If strGetFields <> "" Then '�������Ԫ�ز�����һ��,�һ�ȡ����Ԫ�أ��򲻷��ؼ�¼��
            If UBound(arrValues) - LBound(arrValues) + 1 <= 1 And Not strGetFields Like "*,*" Then blnReturnRec = False
        End If
        If strPars <> "" Then
            strWhere = " Where " & strWhereField & " In (" & strPars & ")"
        End If
    ElseIf TypeName(arrValues) <> "Error" Then
        '����ֵʱ������ȡ�����ֶΣ��򲻷��ڼ�¼��
         If strGetFields <> "" And Not strGetFields Like "*,*" Then blnReturnRec = False
        arrPars = Array(arrValues)
        strWhere = " Where " & strWhereField & "=[1]"
    Else
        strWhere = ""
    End If
    
    If strGetFields = "" Then strGetFields = "*"
    strSQL = "Select " & strGetFields & " From " & strTable & strWhere
    If strWhere <> "" Then
        Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "RowValue", arrPars)
    Else
        Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "RowValue")
    End If
    If blnReturnRec Then
        Set RowValue = rsTmp
    Else
        If Not rsTmp.EOF Then
            RowValue = rsTmp.Fields(strGetFields).value
        Else '��ȡ��ֵʱ��û�л�ȡ����ֵ���򷵻�Ĭ��ֵ
            If IsType(rsTmp.Fields(strGetFields).type, adVarChar) Then
                RowValue = ""
            ElseIf IsType(rsTmp.Fields(strGetFields).type, adInteger) Then
                RowValue = 0
            Else
                RowValue = Null
            End If
        End If
    End If
    Exit Function
errH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Function

Public Function IsZLHIS10() As Boolean
'���ܣ��жϱ�׼���Ƿ���10�汾��ϵͳ
    IsZLHIS10 = gobjComLib.zlDatabase.IsZLHIS10
End Function

Public Function GetUserInfo() As ADODB.Recordset
'���ܣ���ȡ��ǰ�û��Ļ�����Ϣ
'���أ�����Ado��¼��
    Set GetUserInfo = gobjComLib.zlDatabase.GetUserInfo
End Function

Public Function Currentdate() As Date
'���ܣ���ȡ�������ϵ�ǰ����
'������
'���أ�����Oracle���ڸ�ʽ�����⣬����
    Currentdate = gobjComLib.zlDatabase.Currentdate
End Function

Public Function IP(Optional ByVal strErr As String) As String
    '******************************************************************************************************************
    '����:ͨ��oracle��ȡ�ļ������IP��ַ
    '���:strDefaultIp_Address-ȱʡIP��ַ
    '����:
    '����:����IP��ַ
    '******************************************************************************************************************
    Dim rsTmp As ADODB.Recordset
    Dim strIp_Address As String
    Dim strSQL As String
        
    On Error GoTo Errhand
    
    strSQL = "Select Sys_Context('USERENV', 'IP_ADDRESS') as Ip_Address From Dual"
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "��ȡIP��ַ")
    If rsTmp.EOF = False Then
        strIp_Address = gobjComLib.NVL(rsTmp!Ip_Address)
    End If
    If strIp_Address = "" Then strIp_Address = gobjComLib.OS.IP(strErr)
    If Replace(strIp_Address, " ", "") = "0.0.0.0" Then strIp_Address = ""
    IP = strIp_Address
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
Errhand:
    strErr = strErr & IIf(strErr = "", "", "|") & Err.Description
    Err.Clear
End Function

Public Function GetNextNo(ByVal int��� As Integer, Optional ByVal lng����ID As Long, Optional ByVal strTag As String, Optional ByVal intStep As Integer = 1) As Variant
    GetNextNo = gobjComLib.zlDatabase.GetNextNo(int���, lng����ID, strTag, intStep)
End Function

Public Function NextId(strTable As String) As Long
'���ܣ���ȡָ��������Ӧ������(���淶������������Ϊ��������_id��)����һ��ֵ
'������
'   strTable��������
'���أ�
    NextId = gobjComLib.zlDatabase.GetNextId(strTable)
End Function

Public Function MaxCode(ByVal strTable As String, ByVal strField As String, ByVal intLength As Integer, Optional ByVal strWhere As String) As String
'���ܣ���ȡָ�����ı�����������ֵ
'������strTable  ����;
'      strField  �ֶ���;
'      intLength �ֶγ���
'���أ��ɹ����� �¼�������; ���߷��� 0
    MaxCode = gobjComLib.zlDatabase.GetMax(strTable, strField, intLength, strWhere)
End Function

Public Function DeptHaveProperty(ByVal lng����ID As Long, ByVal str���� As String, Optional ByRef blnOutDept As Boolean, Optional ByVal blnClearCache As Boolean) As Boolean
'���ܣ����ָ�������Ƿ����ָ����������
'˵������Ϊ��������һ�㲻�䶯���ִ���ʹ�ã����û����ȡ
'        blnClearCache=�Ƿ��������
'���أ�blnOutDept=�Ƿ�Ϊ������������Ĳ���
    Static rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long, blnNew As Boolean
    Dim rsDept As ADODB.Recordset
    blnOutDept = False
    
    If blnClearCache Then Set rsTmp = Nothing '��ջ���
    If rsTmp Is Nothing Then
        blnNew = True
    Else
        blnNew = (rsTmp.State = adStateClosed)
    End If
    strSQL = "Select ����ID,��������,������� From ��������˵�� where ����ID=[1]"
    On Error GoTo errH
    
    '��ȡ����Դ
    If blnNew Then
        Set rsTmp = gobjComLib.Rec.CopyNew(gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "DeptHaveProperty", lng����ID))
    Else
        rsTmp.Filter = "����ID=" & lng����ID
        If rsTmp.EOF Then
            Call gobjComLib.Rec.Append(rsTmp, gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "DeptHaveProperty", lng����ID))
        End If
    End If
    
    '�����ж�
    rsTmp.Filter = "����ID=" & lng����ID & " And ��������='" & str���� & "'"
    DeptHaveProperty = Not rsTmp.EOF
    If rsTmp.RecordCount > 0 Then
        rsTmp.Filter = "����ID=" & lng����ID & " And ��������='" & str���� & "' And �������<>1"
        blnOutDept = rsTmp.RecordCount = 0
    End If
    Exit Function
errH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Function

Public Function SystemVersion(Optional ByVal lngSys As Long = 100) As String
'���ܣ���ȡָ��ϵͳ�汾�ţ�Ĭ��ΪҽԺϵͳ��׼��
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo Errhand
    
    strSQL = " select �汾�� from zlsystems where ���=[1]"
    Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "��ȡϵͳ�汾��", lngSys)
    SystemVersion = rsTemp!�汾��
    Exit Function
Errhand:
    If gobjComLib.ErrCenter = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Function

Public Function SystemOwner(Optional ByVal lngSys As Long = 100) As String
'���ܣ���ȡָ��ϵͳ��������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL  As String
    
    On Error GoTo errH
    strSQL = "Select ������ From zlSystems Where ���=[1]"
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "SystemOwner", lngSys)
    If Not rsTmp.EOF Then
        SystemOwner = rsTmp!������
    End If
    Exit Function
errH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Function

Public Function SystemShareWith(ByVal lngSysShareNO As Long, Optional ByVal lngSysMainNO As Long = 100) As Boolean
'���ܣ���׼ϵͳ������ϵͳ�Ƿ��ǹ�����װ
'������lngSysShareNO= ������װ��ϵͳ
'           lngSysMainNO=��ϵͳ
    Dim lngShareNum As Long
    Dim strSQL As String
    Dim rsTmp As Recordset
'Select * From (Select * From zlSystems Start With ��� = 100 Connect By Prior ��� = ������) Where ��� = 300
'Select * From (Select * From zlSystems Start With ��� = 300 Connect By Prior ��� = ������) Where ��� = 100
    strSQL = "Select s.���" & vbNewLine & _
            "From zlSystems S" & vbNewLine & _
            "Where s.������װ = 1 And s.���  = [1] And s.������ = [2]"
    On Error GoTo errH
    '���ڴ��ڶ������������׼������ױ��100��101������������199����������ж�
    '�����ײ��ܹ�����װ
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, gstrSysName, (lngSysShareNO \ 100) * 100, (lngSysMainNO \ 100) * 100)
    SystemShareWith = rsTmp.RecordCount > 0
    Exit Function
errH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Function

Public Function IsSysSetUp(ByVal lngSys As Long) As Boolean
'���ܣ��Ƿ�װĳ��ϵͳ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select 1 From zlSystems Where ���=[1]"
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "IsSysSetUp", lngSys)
    IsSysSetUp = Not rsTmp.EOF
    Exit Function
errH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Function

Public Function FieldsLength(strTable As String, strColumns As String) As Variant
'���ܣ���ȡһ�������ֶζ��峤��
'������strTable=����
'          strColumns=����������֮���Զ��ŷָ�
'���أ�����ֱ�ӷ����г��ȣ����з��س����Զ��ŷָ�,��ĳһ��δ��ȡ�����򷵻�-1
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim arrTmp As Variant, i As Long
    Dim strReturn As String
    strSQL = "Select Column_Name,Nvl(Data_Precision, Data_Length) Collen ,Decode(Owner, User, 1, 2) Sort" & vbNewLine & _
                "From All_Tab_Columns" & vbNewLine & _
                "Where Table_Name = [1] And Column_Name In(Select Column_Value From Table(Cast(f_Str2list([2]) As zlTools.t_Strlist)))"
    On Error GoTo errH
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "FieldsLength", strTable, UCase(strColumns))
    arrTmp = Split(UCase(strColumns), ",")
    For i = LBound(arrTmp) To UBound(arrTmp)
        rsTmp.Filter = "Column_Name='" & arrTmp(i) & "'": rsTmp.Sort = "Sort"
        If Not rsTmp.EOF Then
            strReturn = strReturn & "," & Val("" & rsTmp!collen)
        Else
            strReturn = strReturn & ",-1"
        End If
    Next
    If strReturn <> "" Then strReturn = Mid(strReturn, 2)
    If strReturn Like "*,*" Then
        FieldsLength = strReturn
    Else
        FieldsLength = Val(strReturn)
    End If
    Exit Function
errH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Function

Public Function GetPrivFuncByUser(lngSys As Long, lngProgId As Long, ByVal strDBUser As String) As String
'���ܣ������û����е�ָ������Ĺ��ܴ�,����¼�û�
'������lngSys     ����ǹ̶�ģ�飬��Ϊ0
'      lngProgId  �������
'      strDBUser=���ݿ��û�
'���أ��ֺż���Ĺ��ܴ�,Ϊ�ձ�ʾû��Ȩ��

    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim strTmp As String
    
    On Error GoTo errH
    '��ǰ�û����⴦����ֱ�ӵ���gobjRegister.zlRegFunc
    If UCase(strDBUser) = UCase(gstrDBUser) Then
        GetPrivFuncByUser = gobjComLib.GetPrivFunc(lngSys, lngProgId)
        Exit Function
    End If
    '�ǵ�ǰ�û������ȡ���ű����򵥷���
    '��ɫ��Ȩʱֻ������ȨZlregFunc�д��ڵĹ���Ȩ�ޣ���˿��Բ�������
    '����������ģ������Ϊģ�鹦�ܣ�����������̨��Ϊ����ģ�飬���ǵ���Ȩ����Zlrolegrant�С�
    strSQL = "Select Distinct b.����" & vbNewLine & _
            "From Zluserroles a, Zlrolegrant b" & vbNewLine & _
            "Where a.�û� = [3] And a.��ɫ = b.��ɫ And b.ϵͳ = [1] And b.��� = [2]"
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "��ȡ������Ա����Ȩ��", lngSys, lngProgId, strDBUser)
    Do While Not rsTmp.EOF
        strTmp = strTmp & ";" & rsTmp!����
        rsTmp.MoveNext
    Loop
    strTmp = Mid(strTmp, 2)
    
    GetPrivFuncByUser = strTmp
    Exit Function
errH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Function
