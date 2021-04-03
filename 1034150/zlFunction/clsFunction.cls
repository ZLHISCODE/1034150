VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFunction"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Sub About(Optional frmParent As Object)
'���ܣ���ʾ���ڴ���
    Call ShowAbout(frmParent)
End Sub

Public Function GetFunScript(ByVal cnLink As ADODB.Connection, ByVal strExp As String) As String
'���ܣ���ȡָ������������
'������strFunc="ZLHIS.ZL_FUN_XXXX(X,X,X,X,)"
'���أ�"������:XXX~������:XXX~˵��:XXX||������:XXX~������:XXX~ֵ:XXX|������:XXX~....."
    Dim strSQL As String, strScript As String
    Dim rsTmp As New ADODB.Recordset

    Dim lngSys As Long, strOwner As String, strFunc As String
    Dim strVal As String, strVals As String
    Dim objPars As FuncPars, tmpPar As FuncPar
    Dim i As Integer, j As Integer, strTmp As String
    
    If cnLink Is Nothing Then Exit Function
    If cnLink.State = adStateClosed Then Exit Function
    Set gcnOracle = cnLink
    gstrDBUser = GetDBUser
    
    Call SplitFunc(strExp, strOwner, strFunc, strVals)
    
    On Error GoTo errH
    
    strSQL = "Select * From zlFunctions Where Upper(������)='" & UCase(strFunc) & "' And ϵͳ IN(Select ��� From zlSystems Where Upper(������)='" & UCase(strOwner) & "')"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡָ������������")
    If rsTmp.EOF Then Exit Function
    lngSys = rsTmp!ϵͳ
    
    Set objPars = ReadFuncPars(lngSys, rsTmp!������)
    For i = 1 To objPars.Count
        objPars(i).����SQL = Replace(objPars(i).����SQL, "[ϵͳ]", lngSys)
        objPars(i).��ϸSQL = Replace(objPars(i).��ϸSQL, "[ϵͳ]", lngSys)
    Next
        
    i = 0: strScript = ""
    For Each tmpPar In objPars
        '��ǰֵ
        strVal = ""
        If UBound(Split(strVals, "|")) >= i Then
            strVal = Split(strVals, "|")(i)
            If Left(strVal, 1) = "'" And Right(strVal, 1) = "'" Then
                strVal = Mid(strVal, 2, Len(strVal) - 2)
            End If
        Else
            Exit For
        End If
        
        i = i + 1
        
        strScript = strScript & "|������:" & tmpPar.���� & "~������:" & tmpPar.������ & "~ֵ:"
        
        If UCase(strVal) = "[ZLBEGINTIME]" Or UCase(strVal) = "[ZLENDTIME]" Then
            strScript = strScript & "��̬ȷ��"
        Else
            If tmpPar.ȱʡֵ = "�̶�ֵ�б���" Then
                For j = 0 To UBound(Split(tmpPar.ֵ�б�, "|"))
                    '��ֵ��ͬ��ֵ
                    If strVal = Split(Split(tmpPar.ֵ�б�, "|")(j), ",")(1) And strVal <> "" Then
                        strScript = strScript & Split(Split(tmpPar.ֵ�б�, "|")(j), ",")(0)
                        Exit For
                    End If
                Next
            ElseIf tmpPar.ȱʡֵ = "ѡ�������塭" Then
                '���ݰ�ֵ��ֵ
                If strVal <> "" Then
                    strTmp = ""
                    If tmpPar.ֵ�б� Like "*|*" Then
                        strTmp = Split(tmpPar.ֵ�б�, "|")(1)
                        If (strVal = strTmp) Or (UCase(strVal) = "NULL" And Trim(strTmp) = "") Then
                            strScript = strScript & Split(tmpPar.ֵ�б�, "|")(0)
                            strTmp = "OK"
                        Else
                            strTmp = ""
                        End If
                    End If
                    
                    If strTmp = "" Then
                        If InStr(tmpPar.����, "|") > 0 Then strTmp = Split(tmpPar.����, "|")(0)
                        strTmp = SQLOwner(Replace(RemoveNote(tmpPar.��ϸSQL), "[*]", ""), strTmp)
                        strTmp = GetBalndValue(strTmp, tmpPar.��ϸ�ֶ�, strVal)
                        If strTmp <> "" Then strScript = strScript & Split(strTmp, "|")(0)
                    End If
                End If
            Else
                If tmpPar.���� = 2 Then
                    '�Ǻ����ڲŸ�ֵ
                    If Left(tmpPar.ȱʡֵ, 1) <> "&" And strVal <> "" Then
                        If UCase(strVal) Like "TO_DATE('*','*')" Then
                            strScript = strScript & Format(GetDate(strVal), "yyyy-MM-dd")
                        End If
                    End If
                Else
                    '��ֵ
                    If strVal <> "" Then strScript = strScript & strVal
                End If
            End If
        End If
    Next
    
    strScript = "������:" & rsTmp!������ & "~������:" & rsTmp!������ & "~˵��:" & IIf(IsNull(rsTmp!˵��), "", rsTmp!˵��) & "||" & Mid(strScript, 2)
    GetFunScript = strScript
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function FuncGuide(ByVal cnMain As ADODB.Connection, frmParent As Object, Optional strExp As String) As String
'���ܣ�ѡ��������ȡ��ʽ
'������
'     cnMain:����������ݿ�����
'     frmParent:������
'���أ�ѡ�����ĵ��ù�ʽ
    If cnMain Is Nothing Then Exit Function
    If cnMain.State = adStateClosed Then Exit Function
    Set gcnOracle = cnMain
    gstrDBUser = GetDBUser
    Set gfrmMain = frmParent
    
    gblnDBA = HaveDBA()
    
    If grsObject Is Nothing Then Set grsObject = UserObject
    If grsObject Is Nothing Then Exit Function
    If grsObject.State = adStateClosed Then Set grsObject = UserObject
       
    frmParInput.mstrExp = strExp
    Load frmParInput
    
    Err.Clear
    On Error Resume Next
    frmParInput.Show 1, frmParent
    If Err.Number <> 0 Then
        Err.Clear
        frmParInput.Show 1
    End If
    If gblnOK Then
        FuncGuide = frmParInput.mstrExp
        frmParInput.mstrExp = ""
    End If
End Function

Public Function FuncManage(ByVal cnMain As ADODB.Connection, Optional frmParent As Object) As Boolean
'���ܣ���������
'������
'     cnMain:����������ݿ�����
'     frmParent:������
'���أ������Ƿ�ɹ�
    If cnMain Is Nothing Then Exit Function
    If cnMain.State = adStateClosed Then Exit Function
    Set gcnOracle = cnMain
    gstrDBUser = GetDBUser
    Set gfrmMain = frmParent
    '49074
    InitCommon gcnOracle
    gblnDBA = HaveDBA()
    
    If grsObject Is Nothing Then Set grsObject = UserObject
    If grsObject Is Nothing Then Exit Function
    If grsObject.State = adStateClosed Then Set grsObject = UserObject
    
    If Not frmParent Is Nothing Then
        Load frmMain
        Err.Clear
        On Error Resume Next
        frmMain.Show , frmParent
        If Err.Number <> 0 Then
            Err.Clear
            frmMain.Show 1
        End If
    Else
        frmMain.Show
    End If
    
    FuncManage = True
End Function

Private Sub Class_Initialize()
    Call RegFuncFile
End Sub

Private Sub Class_Terminate()
    Set gcnOracle = Nothing
    Set grsObject = Nothing
End Sub

Public Function CloseWindows() As Boolean
'����:�ر������Ӵ���
    Dim frmThis As Form
    
    On Error Resume Next
    
    For Each frmThis In Forms
        Unload frmThis
    Next
    CloseWindows = (Forms.Count = 0)
End Function
