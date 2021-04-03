VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBusiness"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--------------------------------------------------------------------------------
'˵��: Ϊ����ҵ��ϵͳ���ŵĽӿ���
'--------------------------------------------------------------------------------
Option Explicit

Public Enum RA_Side
    s_�ⲿ = 1
    s_�ڲ� = 2
End Enum

Public Enum RA_AutoAuditReturn
    aar_�쳣 = 0
    aar_δ������ = 1
    aar_δ�������� = 2
    aar_����ҩƷ = 3
    aar_�������� = 4
    aar_ȫ���ϸ� = 5
    aar_�����������δ���� = 6
End Enum

Private mblnInit As Boolean                                 'True�ѵ��ó�ʼ��������Init����Falseδ����
Private mblnOutPatient As Boolean                           'True���ﲡ�ˣ�FalseסԺ���ˣ����ⲿ�������Init��������
Private mrsRAItems As ADODB.Recordset                       '
Private mrsAuditSend As ADODB.Recordset                     '
Private mstrAutoAuditID As String                           '

Public Sub Init(ByVal cnMain As ADODB.Connection, ByVal blnOutPatient As Boolean, Optional ByVal lngSys As Long = 100)
'���ܣ���ʼ����ģ��
'������
'  cnMain��Oracle���ݿ����Ӷ���
'  blnOutPatient��True���FalseסԺ
    
    Dim blnParaBasis As Boolean
    Dim strSQL As String
    
    If cnMain Is Nothing Then
        MsgBox "���ݿ����Ӷ���δ������", vbInformation, gstrSysName
        Exit Sub
    End If
    If cnMain.State <> adStateOpen Then
        MsgBox "���ݿ����Ӷ���δ�򿪣�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    'clsBusiness����clsRecipeAudit��Ӧ����ͬһ�����ݿ�����ӣ����ò�ͬ�����Ӷ���
    'Set gcnBusiness = cnMain
    
    glngSys = lngSys
    mblnOutPatient = blnOutPatient
    If gcnOracle Is Nothing Then Set gcnOracle = cnMain
    
    If Not GetUserInfo Then
        MsgBox "��ǰ�û�δ���ö�Ӧ����Ա��Ϣ,����ϵͳ����Ա��ϵ,�ȵ��û���Ȩ���������ã�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '�ٴ���������
    On Error Resume Next
    Set gobjPubAdvice = New zlPublicAdvice.clsPublicAdvice
    If Not gobjPubAdvice Is Nothing Then
        Call gobjPubAdvice.InitCommon(gcnOracle, glngSys)
    End If
    Err.Clear: On Error GoTo 0
    
    '����Ŀ����
    blnParaBasis = (Val(zlDatabase.GetPara("�����������", glngSys)) = 1)    '1-���ݡ��������������淶��28�2-���ݡ����������취��7��
    
    '��ʼ��mrsRAItems
    On Error GoTo errHandle
    strSQL = "Select ID, ���, ����, ���, PASS��� " & vbCr & _
             "From ���������Ŀ " & vbCr & _
             "Where (����ʱ�� Is Null Or To_Char(����ʱ��, 'yyyy') = '3000') And ��� In ([1], 3, 4) " & vbCr & _
             "  And " & IIf(blnOutPatient, " �Ƿ��������� = 1 ", " �Ƿ�סԺ���� = 1 ")
    Set mrsRAItems = zlDatabase.OpenSQLRecord(strSQL, IIf(blnOutPatient, "��ȡ���������Ŀ", "��ȡסԺ�����Ŀ"), IIf(blnParaBasis, 2, 1))
    
    '�ɹ���ʼ��
    mblnInit = True
    
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume Next
End Sub

Public Function AutoAudit( _
    ByVal frmOwner As Object, _
    ByVal bytMode As RA_Side, _
    ByVal strSubmitID As String, _
    ByVal lngClinicID As Long, _
    ByVal bytPatientMode As Byte, _
    ByVal lngPatientID As Long, _
    ByVal lngBillID As Long, _
    Optional ByRef strErr As String _
) As RA_AutoAuditReturn
'���ܣ��Զ����
'������
'  frmOwner�������������
'  'bytClass�����õ�ҵ��1-����ҵ��2-סԺҵ��
'  bytMode��1-�ⲿ���ã�2-�ڲ����ã������Ƿ���ʾ��ʾ�Ի���Ͳ�����������
'  strSubmitID�����ύ���ĸ�ҩ;��ҽ��ID����ʽ����ҩ;��ҽ��ID[,��ҩ;��ҽ��ID[...]]��
'  lngClinicID���ٴ�����ID
'  bytPatientMode��������Դ��0-���1-סԺ
'  lngPatientID������ID
'  lngBillID����ҳID��Һ�ID
'  strErr����ѡʵ�Σ����쳣��Ϣ
'���أ�
'  0-�쳣���ⲿ������ֹ���̻��������
'  1-δ�����󷽣��ⲿ�����������
'  2-�����󷽣���δ�������ϸ����ѣ��ⲿ�����������
'  3-�����󷽣��������ϸ����ѣ�ҽ������������ҩƷ�����ⲿ������ֹ����
'  4-�����󷽣��������ϸ����ѣ�ҽ���������������ѡ����ⲿ�����������
'  5-�����󷽣��������ϸ����ѣ����û�в��ϸ����Ŀ���ⲿ�����������
'  6-���￪���󷽣��������ϸ����ѣ������������������δ���κ����ã��ⲿ�����������

    Dim frmWarning As frmRAWarning
    Dim blnParaRA As Boolean, blnParaWarning As Boolean
    Dim lngTmp As Long
    Dim strTmp As String, strIDs As String, strNG As String
    Dim strMedicalID As String
    Dim bytResult As Byte
    Dim arrSubmit As Variant, arrID As Variant
    
    gstrErrInfo = ""
    mstrAutoAuditID = ""
    
    '���²������͵ı���ֵ�������ٴ��õ�ȫ�ֶ��������
    mblnOutPatient = (bytPatientMode = 0)
    
    If mblnInit = False Then
        strErr = "���ȵ��ýӿڳ�ʼ���ķ���"
        AutoAudit = RA_AutoAuditReturn.aar_�쳣
        Exit Function
    End If
    
    '��鴦����������0-�����סԺ�������ã�1-�������ã�סԺ�����ã�2-���ﲻ���ã�סԺ���ã�3-�����סԺ������
    lngTmp = Val(zlDatabase.GetPara("�������", glngSys))
    If mblnOutPatient Then
        '����ҵ��
        blnParaRA = (lngTmp = 1 Or lngTmp = 3)
        strTmp = "δ�������ﴦ����飬������ñ��ӿڣ�"
    Else
        'סԺҵ��
        blnParaRA = (lngTmp = 2 Or lngTmp = 3)
        strTmp = "δ����סԺ������飬������ñ��ӿڣ�"
    End If
    If blnParaRA = False Then
        strErr = strTmp
        AutoAudit = RA_AutoAuditReturn.aar_δ������
        Exit Function
    End If
    
    On Error GoTo errHandle
    
    '�Զ���鲻�ϸ����Ѳ���
    If mblnOutPatient Then
        '����
        blnParaWarning = (Val(zlDatabase.GetPara("��������ҽ�����ϸ�ҽ��", glngSys)) = 1)
    Else
        'סԺ
        blnParaWarning = (Val(zlDatabase.GetPara("����סԺҽ�����ϸ�ҽ��", glngSys)) = 1)
    End If
    
'    'ȡ�ύ��ҽ��ID
'    arrSubmit = Split(strSubmitID, ",")
'    For lngTmp = LBound(arrSubmit) To UBound(arrSubmit)
'        strIDs = strIDs & Val(arrSubmit(lngTmp)) & IIf(lngTmp >= UBound(arrSubmit), "", ",")
'    Next
    
    If strSubmitID = "" Then
        strErr = "δ����ҽ��ID��Ϣ"
        AutoAudit = RA_AutoAuditReturn.aar_�쳣
        Exit Function
    End If
    
    '׼����¼�������ݽṹ
    SetRecordsetStructure 1, mrsAuditSend
    
    If mblnOutPatient Then
        '��������ദ���������
        lngTmp = IsOutpatientControl(strSubmitID)
        If lngTmp = 0 Or lngTmp = 2 Then
            AutoAudit = RA_AutoAuditReturn.aar_�����������δ����
            Exit Function
        End If
    End If
    
    '�����Ŀ
    With mrsRAItems
        If .RecordCount > 0 Then .MoveFirst
        Do While .EOF = False
            '��鴦��
            bytResult = AuditDrug(.Fields, lngPatientID, IIf(mblnOutPatient, 1, 2), lngBillID, strSubmitID, strMedicalID)
            
            '��װ����
            If strMedicalID <> "" Then
                '��ҽ��ID
                arrID = Split(strMedicalID, ";")
                For lngTmp = LBound(arrID) To UBound(arrID)
                    mrsAuditSend.AddNew
                    mrsAuditSend!�����ĿID = !ID
                    mrsAuditSend!���� = !����
                    mrsAuditSend!��� = !���
                    mrsAuditSend!����� = bytResult
                    mrsAuditSend!ҽ��ID = Val(arrID(lngTmp))
                    mrsAuditSend.Update
                Next
            Else
                '��ҽ��ID����ʾ����ҽ��
                mrsAuditSend.AddNew
                mrsAuditSend!�����ĿID = !ID
                mrsAuditSend!���� = !����
                mrsAuditSend!��� = !���
                mrsAuditSend!����� = bytResult
                mrsAuditSend!ҽ��ID = 0
                mrsAuditSend.Update
            End If

            .MoveNext
        Loop
    End With
    
    '�������ĸ�ҩ;��ҽ��ID
    mstrAutoAuditID = strSubmitID
    
    '��鲻�ϸ�����
    If blnParaWarning And bytMode = RA_Side.s_�ⲿ Then
        '�������ѣ������ⲿ����
        With mrsAuditSend
            strNG = ""
            If .RecordCount > 0 Then .MoveFirst
            Do While .EOF = False
                If !����� = 2 Then       'ֻ�������ϸ�
                    strNG = strNG & mdlDefine.FormatString("[1],[2]|", !�����ĿID, zlCommFun.NVL(!ҽ��ID, "0"))
                End If
                .MoveNext
            Loop
            If strNG <> "" Then strNG = Left(strNG, Len(strNG) - 1)
        End With
        
        If strNG <> "" Then
            '��ʾ���ϸ���Ϣ
            Set frmWarning = New frmRAWarning
            If frmWarning.ShowMe(strNG, frmOwner) Then
                AutoAudit = RA_AutoAuditReturn.aar_����ҩƷ
            Else
                AutoAudit = RA_AutoAuditReturn.aar_��������
            End If
        Else
            AutoAudit = RA_AutoAuditReturn.aar_ȫ���ϸ�       '���û�в��ϸ����Ŀ
        End If
    Else
        '�ر����ѻ��ڲ�������������
        AutoAudit = RA_AutoAuditReturn.aar_δ��������
    End If
    
    Exit Function
    
errHandle:
    strErr = gstrErrInfo
    AutoAudit = 0
End Function

Public Function BuildData(ByVal strSubmitID As String, ByVal lngClinicID As Long, _
    ByVal bytPatientMode As Byte, ByVal lngPatientID As Long, ByVal lngBillID As Long, _
    ByRef strNotSendIDs As String, _
    Optional ByRef strErr As String _
) As Boolean
'���ܣ��ٴ�����վ���ñ����������ɴ����ģ�����/סԺ������
'������
'  strSubmitID�����ύ���ĸ�ҩ;��ҽ��ID����ʽ����ҩ;��ҽ��ID[,��ҩ;��ҽ��ID[...]]��
'  lngClinicID���ٴ�����ID
'  bytPatientMode��0-���1-סԺ
'  lngPatientID������ID
'  lngBillID����ҳID��Һ�ID
'  strNotSendIDs��ʵ�Σ������ز��ܡ����͡��ĸ�ҩ;��ҩ��ID����ʽ����ҩ;��ҽ��ID,...��
'  strErr��ʵ�Σ����쳣��Ϣ
'���أ�True�ɹ���Falseʧ��

    Dim typAudit As RA_AutoAuditReturn
    Dim strSQL As String, strIDs As String
    Dim strTmp As String
    Dim lngRelationID As Long
    Dim rsTemp As ADODB.Recordset
    Dim i As Integer
    Dim lngStoreID As Long, lngAuditID As Long
    Dim arrID As Variant
    Dim colSQL As New Collection
    Dim strCannotSend As String, strSubmitNot As String
    Dim intPara As Integer
    
    '���²������͵ı���ֵ�������ٴ��õ�ȫ�ֶ��������
    mblnOutPatient = (bytPatientMode = 0)
    
    If mblnInit = False Then
        strErr = "���ȵ��ýӿڳ�ʼ���ķ���"
        BuildData = False
        MsgBox strErr
        Exit Function
    End If
    
    If Trim(strSubmitID) = "" Then
        strErr = "�ύ�����IDΪ��"
        BuildData = False
        MsgBox strErr
        Exit Function
    End If
    
    '����������
    intPara = Val(zlDatabase.GetPara("�������", glngSys, , "0"))
    If mblnOutPatient Then
        If Not (intPara = 1 Or intPara = 3) Then
            'δ����������
            BuildData = True
'            MsgBox "δ����������"
            Exit Function
        End If
    Else
        If Not (intPara = 2 Or intPara = 3) Then
            'δ����סԺ��
            BuildData = True
'            MsgBox "δ����סԺ��"
            Exit Function
        End If
    End If
    
    '����Ĵ������ʱ��������1-��������ǰ��2-�����䷢ҩǰ
    intPara = Val(zlDatabase.GetPara("������ʱ��", glngSys, , "1"))
    
    '�����������ǰ��������ʱ��δ�����󷽼�¼״̬��ҩ���䷢ҩǰ��������
    If mblnOutPatient And intPara = 1 Then
        Call AuditSend(lngClinicID, strSubmitID, strTmp)
    End If
    
    On Error GoTo hErr
    
    '�ȷ������ύ��δ�ύ�����ID
    strSQL = "Select Distinct d.Column_Value ���id, c.״̬, c.�����, c.Id ��id " & vbNewLine & _
             "From ����ҽ����¼ A, ���������ϸ B, ��������¼ C, Table(f_Num2list([1], ',')) D " & vbNewLine & _
             "Where d.Column_Value = a.���id(+) And a.Id = b.ҽ��id(+) And b.����ύ(+) = 1 And b.��id = c.Id(+) " & vbNewLine & _
             "    And a.������� In ('5', '6', '7') "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "�������ID", strSubmitID)
    With rsTemp
        Do While .EOF = False
            If IsNull(!��id) Then
                'δ�ύ�����id
                strSubmitNot = strSubmitNot & "," & CStr(!���id)
            Else
'                '����-����ǰ�󷽡�
'                If intPara = 1 And mblnOutPatient Then
                    Select Case zlCommFun.NVL(!״̬, 0)
                    Case 0          '����
                        strCannotSend = strCannotSend & "," & CStr(!���id)
                    Case 1          '����
                        If zlCommFun.NVL(!�����, 0) = 2 Then   '2-���ϸ�
                            strCannotSend = strCannotSend & "," & CStr(!���id)
                        End If
                    Case 2, 3       '��ʱ����
                    Case 11         '���󱻳���
                        strCannotSend = strCannotSend & "," & CStr(!���id)
                    End Select
'                End If
            End If
            
            .MoveNext
        Loop
        .Close
    End With
    'δ�ύ�����ID
    If strSubmitNot <> "" Then strSubmitNot = Mid(strSubmitNot, 2)
    '���ύ����δ������鲻�ϸ�����ID
    If strCannotSend <> "" Then strCannotSend = Mid(strCannotSend, 2)
    
    '����δ�ύ��ҽ��
    If strSubmitNot <> "" Then
        '�󷽹���վ�Ƿ���������
        If mblnOutPatient Then
            '����
'            strSQL = "Select Count(1) Rec " & vbNewLine & _
'                     "From ���������� " & vbNewLine & _
'                     "Where Nvl(�������, 0) = 0 And ',' || ��Դ���� || ',' Like '%,' || [1] || '%,' And �Ƿ����� = 1 "
            strSQL = "Select Count(1) Rec " & vbNewLine & _
                     "From ���������� " & vbNewLine & _
                     "Where Nvl(�������, 0) = 0 And �Ƿ����� = 1 "
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ������鹤��վ����������") ', lngClinicID)
            If rsTemp!Rec <= 0 Then
                '�󷽹���վδ�������񣬲����ɴ�������¼
                rsTemp.Close
                strSubmitNot = ""   'δ�ύ��ҽ�������ظ��ϲ�ҵ��
                GoTo makOne
            End If
            rsTemp.Close
        Else
            'סԺ������Ƿ���������
        End If
        
        typAudit = AutoAudit(Nothing, s_�ڲ�, strSubmitNot, lngClinicID, bytPatientMode, lngPatientID, lngBillID, strErr)
        '�ڲ����÷��ء�δ�������ѡ�ֵ
        If typAudit <> aar_δ�������� Then
            '������������¼
            strErr = "�����ύ���"
            strSubmitNot = ""       'δ�ύ��ҽ�������ظ��ϲ�ҵ��
            MsgBox strErr
            GoTo makOne
        End If
        
        '�������
        If mrsAuditSend Is Nothing Then
            strErr = "�������δ����"
            strSubmitNot = ""       'δ�ύ��ҽ�������ظ��ϲ�ҵ��
            MsgBox strErr
            GoTo makOne
        End If
        
        If mrsAuditSend.State <> adStateOpen Then mrsAuditSend.Open
        
        If mrsAuditSend.RecordCount <= 0 Then
            strErr = "���������"
            strSubmitNot = ""       'δ�ύ��ҽ�������ظ��ϲ�ҵ��
            MsgBox strErr
            GoTo makOne
        End If
        
        '��ҩ��ID�ֱ��ύ������¼
    '    strSQL = "Select a.ִ�п���id, a.Id, a.���id, b.ҽ��id, c.�����, " & vbNewLine & _
    '             "    Max(c.״̬) Over(Partition By a.���id) ״̬, " & vbNewLine & _
    '             "    Count(a.���id) Over(Partition By a.���id) as ҩ����, " & vbNewLine & _
    '             "    Sum(Decode(nvl(b.ҽ��id,0), 0, 0, 1)) Over(Partition By a.���id) as �ύ�� " & vbNewLine & _
    '             "From ����ҽ����¼ A, ���������ϸ B, ��������¼ C, Table(f_Num2list([1], ',')) D " & vbNewLine & _
    '             "Where a.���id = d.Column_Value And a.Id = b.ҽ��id(+) And b.��id = c.Id(+) " & _
    '             "    And a.������� In ('5', '6', '7') And b.����ύ(+) = 1 " & vbNewLine & _
    '             "Order By a.ִ�п���id, a.���id, a.ID "
        
        strSQL = "Select a.ִ�п���id, a.Id " & vbNewLine & _
                 "From ����ҽ����¼ A, Table(f_Num2list([1], ',')) B " & vbNewLine & _
                 "Where a.���id = b.Column_Value " & vbNewLine & _
                 "Order By a.ִ�п���ID, a.���Id, a.ID "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡҽ����Ϣ", strSubmitNot)
        With rsTemp
            If .EOF = False Then lngStoreID = !ִ�п���ID
            Do While .EOF = False
                '����Ƿ�����ύ��ҽ��ID
                    
                ''׼��ִ�й���
                If lngStoreID = !ִ�п���ID Then
                    strIDs = strIDs & mdlDefine.FormatString("[1],", !ID)
                Else
makSQL:
                    '��ҩ��ID�ֱ��ύ����¼
                    lngAuditID = zlDatabase.GetNextId("��������¼")
                    If lngAuditID <= 0 Then
                        strErr = "��������¼�����ж��󲻴��ڣ��޷��ύ������¼"
                        strSubmitNot = ""
                        MsgBox strErr, vbInformation, gstrSysName
                        Exit Function
                    End If
                    
                    If strIDs <> "" Then strIDs = Left(strIDs, Len(strIDs) - 1)
                    
                    '��ҩʦ���ҽ����ϸ
                    strSQL = mdlDefine.FormatString("ZL_�������_INSERT( [1],[2],[3],[4],[5],'[6]',[7],'[8]' )", _
                                lngAuditID, _
                                lngPatientID, _
                                IIf(mblnOutPatient = False, "Null", lngBillID), _
                                IIf(mblnOutPatient = False, lngBillID, "Null"), _
                                lngClinicID, _
                                UserInfo.����, _
                                lngStoreID, _
                                strIDs)
                    
                    'SQL���뼯�϶���
                    AddArray colSQL, strSQL
                    
                    '�Զ������
                    ''�������ҽ��ID����ҩ��ID�ֱ��ύ
                    strTmp = ""
                    mrsAuditSend.Filter = "(�����=2 and ҽ��ID<=0) Or (�����=1 and ҽ��ID<=0)"
                    'If mrsAuditSend.RecordCount > 0 Then mrsAuditSend.MoveFirst
                    Do While mrsAuditSend.EOF = False
                        strTmp = strTmp & mdlDefine.FormatString("[1],0|", "" & mrsAuditSend!����� & "_" & mrsAuditSend!�����ĿID)
                        mrsAuditSend.MoveNext
                    Loop
                    If strTmp <> "" Then
                        strTmp = Left(strTmp, Len(strTmp) - 1)
                        strSQL = mdlDefine.FormatString("ZL_�������_AUTO( [1], '[2]' )", _
                                    lngAuditID, _
                                    strTmp)
                        AddArray colSQL, strSQL
                    End If
                    
                    ''�������ҽ��ID��ֻ�ύһ��
                    strTmp = ""
                    arrID = Split(strIDs, ",")
                    For i = LBound(arrID) To UBound(arrID)
                        strTmp = strTmp & mdlDefine.FormatString("(�����=2 and ҽ��id=[1]) or ", Val(arrID(i)))
                    Next
                    If strTmp <> "" Then
                        mrsAuditSend.Filter = Left(strTmp, Len(strTmp) - 4)     '���˳��в��ϸ��ҽ��ID
                        strTmp = ""
                        If mrsAuditSend.RecordCount > 0 Then mrsAuditSend.MoveFirst
                        Do While mrsAuditSend.EOF = False
                            strTmp = strTmp & mdlDefine.FormatString("2_[1],[2]|", mrsAuditSend!�����ĿID, mrsAuditSend!ҽ��ID)
                            mrsAuditSend.MoveNext
                        Loop
                        If strTmp <> "" Then
                            strTmp = Left(strTmp, Len(strTmp) - 1)
                            strSQL = mdlDefine.FormatString("ZL_�������_AUTO( [1], '[2]' )", _
                                        lngAuditID, _
                                        strTmp)
                            AddArray colSQL, strSQL
                        End If
                    End If
                    
                    If .EOF = False Then
                        '���´洢��ǰ��¼��ID
                        strIDs = mdlDefine.FormatString("[1],", !ID)
                    End If
                End If
                
                If .EOF = False Then
                    lngStoreID = !ִ�п���ID
                    .MoveNext
                    If .EOF Then GoTo makSQL
                End If
                    
            Loop
        End With
        
        'ִ�д洢����
        Err = 0: On Error GoTo errHandle
        ExecuteProcedureArray colSQL, "�ύ������¼"
        
    End If
    
makOne:
    
    On Error GoTo hErr
    
    '�������ύ��ҽ���Ƿ���������
    If strCannotSend <> "" Then
        strSQL = "Select f_List2Str(Cast(Collect(Cast(a.���id as Varchar2(20))) as t_StrList), ',') ���id " & vbNewLine & _
                 "From ����ҽ����¼ A, ���������ϸ B, " & vbNewLine & _
                 "     (Select /*+ RULE*/ b.��id " & vbNewLine & _
                 "      From ����ҽ����¼ A, ���������ϸ B, ��������¼ C, Table(f_Num2list([1], ',')) D " & vbNewLine & _
                 "      Where d.Column_Value = a.���id And a.Id = b.ҽ��id And b.����ύ = 1 And b.��id = c.Id) C " & vbNewLine & _
                 "Where a.Id = b.ҽ��id And b.��id = c.��id And a.���id Is Not Null "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�����ύ�����ID", strCannotSend)
        If rsTemp.EOF = False Then
            strNotSendIDs = zlCommFun.NVL(rsTemp!���id)
        End If
    End If

    If intPara = 1 And mblnOutPatient Then
        '�ٽ����ύ��˵����ID����
        If strSubmitNot <> "" Then
            If strNotSendIDs <> "" Then
                strNotSendIDs = strNotSendIDs & "," & strSubmitNot
            Else
                strNotSendIDs = strSubmitNot
            End If
        End If
    End If
    
    BuildData = True
    Exit Function
    
hErr:
    If zl9ComLib.ErrCenter() = 1 Then
        Resume
    Else
        strNotSendIDs = strCannotSend   '������쳣��δ�ύ��ҽ�����ύ��飨���סԺ��
    End If
    Exit Function
    
errHandle:
    gcnOracle.RollbackTrans
    strErr = Err.Description
    Call zl9ComLib.ErrCenter
    strNotSendIDs = strCannotSend       '������쳣��δ�ύ��ҽ�����ύ��飨���סԺ��
    BuildData = False
End Function

Private Function GetGoodID(ByVal lngAuditID As Long) As String
'���ܣ���ȡ�ϸ���������ID��Ӧ��ҽ��ID
'������
'  lngAuditID����ID
'���أ�ҽ��ID�ַ�������ʽ��ҽ��ID,...��

    Dim strSQL As String, strReturn As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    strSQL = "Select a.ҽ��ID From ���������ϸ A, ��������¼ B " & vbCr & _
             "Where a.��ID = b.ID And (b.����� = 1 or b.����� is Null And Nvl(b.״̬, 0) between 1 and 3) And a.��ID = [1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ��ID��Ӧ��ҽ��ID", lngAuditID)
    Do While rsTemp.EOF = False
        strReturn = strReturn & mdlDefine.FormatString(",[1]", rsTemp!ҽ��ID)
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    
    If strReturn <> "" Then strReturn = Mid(strReturn, 2)
    GetGoodID = strReturn
    
    Exit Function
    
errHandle:
    If ErrCenter = 1 Then Resume
End Function

Public Function AuditSend(ByVal lngClinicDeptID As Long, ByVal strSubmitID As String, _
    ByRef strSendID As String, _
    Optional ByRef strErr As String) As Boolean
'���ܣ��������/סԺ���͵�ҽ��ID�Ƿ����ϸ�
'������
'  lngClinicDeptID���ٴ���������ID
'  strSubmitID�����ĸ�ҩ;��ҽ��ID����ʽ����ҩ;��ҽ��ID,...��
'  strSendID���ɷ��͵�ҽ��ID
'  strErr��ʧ�ܵĳ�����Ϣ
'���أ�True�ɹ���Falseʧ��

    Dim strSQL As String, strReturn As String, strTemp As String
    Dim rsTemp As ADODB.Recordset
    Dim blnTrans As Boolean
    
    If mblnInit = False Then
        strErr = "���ȵ��ýӿڳ�ʼ���ķ���"
        AuditSend = False
        Exit Function
    End If
    
    On Error GoTo errHandle

    '������״̬
    strSQL = "Select Distinct b.Id, b.�����, b.�����, b.״̬, b.�ύʱ�� " & vbNewLine & _
             "From ���������ϸ A, ��������¼ B, ����ҽ����¼ C, Table(f_Num2list([1], ',')) D " & vbNewLine & _
             "Where a.��id = b.Id And a.ҽ��id = c.Id And c.���Id = d.Column_Value " & vbNewLine & _
             "    And a.����ύ = 1 And c.�������  in ('5','6','7') "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����¼��״̬", strSubmitID)
    If rsTemp.RecordCount <= 0 Then
        rsTemp.Close
        strErr = "��������¼������"
        Exit Function
    End If
    
    '�������ID
    Do While rsTemp.EOF = False
        If mblnOutPatient Then
            '����ҵ��
            If Val(zlCommFun.NVL(rsTemp!�����)) = 0 And zlCommFun.NVL(rsTemp!״̬) = 0 Then
                'δ���
                '�Գ�ʱ��δ���ļ�¼״̬�������洢�������ж��Ƿ�ʱ��
                If blnTrans = False Then
                    gcnOracle.BeginTrans
                    blnTrans = True
                End If
                
                strSQL = mdlDefine.FormatString("ZL_�������_UPDATE(1, [1], [2])", lngClinicDeptID, rsTemp!ID)
                Call zlDatabase.ExecuteProcedure(strSQL, "������������¼״̬")
            End If
        End If
            
        '����������
        strTemp = GetGoodID(rsTemp!ID)
        If strTemp <> "" Then
            strReturn = strReturn & strTemp & ","
        End If
        
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    
    If strReturn <> "" Then strReturn = Left(strReturn, Len(strReturn) - 1)
    
    If blnTrans Then gcnOracle.CommitTrans
    
    AuditSend = True
    Exit Function
    
errHandle:
    If ErrCenter = 1 Then Resume
    If blnTrans Then gcnOracle.RollbackTrans
End Function

Public Function Auditting(ByVal strSubmitID As String, ByRef strAudittingID As String) As Boolean
'���ܣ����ҽ���Ƿ����������
'������
'  strSubmitID�����ĸ�ҩ;��ҽ��ID����ʽ����ҩ;��ҽ��ID[,��ҩ;��ҽ��ID...]��
'  strAudittingID����������ҽ��ID
'���أ�True�ɹ���Falseʧ��

    Dim strSQL As String, strReturn As String
    Dim rsTemp As ADODB.Recordset
    
    If mblnInit = False Then
        Auditting = False
        Exit Function
    End If
    
    If strSubmitID = "" Then Exit Function
    
    On Error GoTo errHandle
    
'    strSQL = "Select A2.ҽ��id " & vbNewLine & _
'             "From ���������ϸ A1, ���������ϸ A2, ��������¼ B, Table(f_Num2list([1], ',')) C " & vbNewLine & _
'             "Where A1.��id = b.Id And A1.ҽ��id = c.Column_Value And b.Id = A2.��id And A1.����ύ = 1 " & vbNewLine & _
'             "  And (b.�����û� Is Not Null Or b.�����û� <> '') " & vbNewLine & _
'             "Order By a2.ҽ��id "
    strSQL = "Select b.ҽ��id " & vbNewLine & _
             "From ����ҽ����¼ A, ���������ϸ B, ��������¼ C, Table(f_Num2list([1], ',')) D " & vbNewLine & _
             "Where a.���ID = d.Column_Value And a.Id = b.ҽ��ID And b.��Id = c.ID And b.����ύ = 1 " & vbNewLine & _
             "  And (c.�����û� Is Not Null Or c.�����û� <> '') " & vbNewLine & _
             "Order By b.ҽ��id "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ��������ҽ��ID", strSubmitID)
    Do While rsTemp.EOF = False
        strReturn = strReturn & mdlDefine.FormatString("[1],", rsTemp!ҽ��ID)
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    
    If strReturn <> "" Then strReturn = Left(strReturn, Len(strReturn) - 1)
    strAudittingID = strReturn
    
    Auditting = True
    Exit Function
    
errHandle:
    If ErrCenter = 1 Then Resume
End Function

Public Sub ShowResult(ByVal lngMedicalID As Long, ByVal frmOwner As Object)
'���ܣ��鿴ĳҽ��ID��Ӧ�󷽵������
'������
'  lngSubmitID����ҩ;��ҽ��ID
'  frmOwner�������������

    Dim frmResult As frmRAResult
    
    If mblnInit = False Then Exit Sub
    
    Set frmResult = New frmRAResult
    frmResult.ShowMe lngMedicalID, frmOwner
    Set frmResult = Nothing
End Sub

Public Function CancelData(ByVal strSubmitID As String, ByRef strLockID As String, Optional ByRef strErr As String) As Boolean
'���ܣ�1.��δ��顢δ�������Ĵ�������ݽ��г�����2.������飬��δ���͵Ĵ�������ݽ��г���
'������
'  strSubmitID�����������ĸ�ҩ;��ҽ��ID����ʽ����ҩ;��ҽ��ID,...��
'  strLockID����������ҽ��ID
'  strErr��ʧ�ܵĳ�����Ϣ
'���أ�True�ɹ���Falseʧ��

    Dim cmdSQL As New ADODB.Command
    
    If mblnInit = False Then
        strErr = "���ȵ��ýӿڳ�ʼ���ķ���"
        CancelData = False
        Exit Function
    End If
    
    On Error GoTo errHandle
    
    '�����г��Σ���Щδ�ù����������ô洢����
    With cmdSQL
        .ActiveConnection = gcnOracle
        .CommandText = "ZL_�������_CANCEL"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("P1", adVarChar, adParamInput, 4000, strSubmitID)
        .Parameters.Append .CreateParameter("P2", adVarChar, adParamOutput, 4000)
        .Execute
        strLockID = .Parameters("P2").Value
    End With
    
    CancelData = True
    Exit Function
    
errHandle:
    strErr = Err.Description
End Function

Public Function CloseWindows() As Boolean
'���ܣ��رղ��������еĴ���
'���أ�True�ɹ���Falseʧ��

    Dim frmThis As Form
    
    For Each frmThis In Forms
        Unload frmThis
    Next
    CloseWindows = (Forms.Count = 0)
End Function

Private Sub Class_Initialize()
    mblnInit = False
    gintHoursRecipe = 24
End Sub

Private Sub Class_Terminate()
    Call CloseWindows
    If Not gobjRecipeAuditEx Is Nothing Then
        gobjRecipeAuditEx.CloseWindows
        Set gobjRecipeAuditEx = Nothing
    End If
    If Not gobjPubAdvice Is Nothing Then Set gobjPubAdvice = Nothing
End Sub

Private Function IsOutpatientControl(ByVal strMedicalIDs As String) As Byte
'���ܣ�������ﴦ��ҩ���Ƿ�Ϊ����������������ļ�ط�Χ
'������
'  strMedicalIDs����ҩ;��ҩ��ID�ַ���
'���أ�0-δ�������ﴦ��������������޼�����Ĵ�����1-��Ҫ���Ĵ�����2-�������Ĵ�����

    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errHandle
    
    If strMedicalIDs = "" Then Exit Function

    '���Ȳ�조��������������Ƿ��м�¼
    strSQL = "Select Count(1) Rec From ����������� "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��鴦����������ļ�¼��")
    If rsTemp!Rec <= 0 Then
        rsTemp.Close
        Exit Function
    End If
    
    '1-�ٴ�����
    strSQL = "Select Count(1) Rec " & vbNewLine & _
             "From ����ҽ����¼ A, Table(f_Num2list([1])) B, ����������� C " & vbNewLine & _
             "Where a.Id = b.Column_Value And a.��������id = c.����id And c.��� = [2] "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "�����������-�ٴ�����", strMedicalIDs, 1)
    If rsTemp!Rec > 0 Then
        rsTemp.Close
        IsOutpatientControl = 1
        Exit Function
    End If
    
    '2-ҽ��
    strSQL = "Select Count(1) Rec " & vbNewLine & _
             "From ����ҽ����¼ A, Table(f_Num2list([1])) B, ��Ա�� C, ����������� D " & vbNewLine & _
             "Where a.Id = b.Column_Value And a.����ҽ�� = c.���� And c.Id = d.ҽ��id And d.��� = [2] "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "�����������-ҽ��", strMedicalIDs, 2)
    If rsTemp!Rec > 0 Then
        rsTemp.Close
        IsOutpatientControl = 1
        Exit Function
    End If
    
    '˵����������ϼ�¼.ҽ��id �� ����ҽ����¼.���id
    '3-���
    strSQL = "Select Count(1) Rec" & vbNewLine & _
             "From (Select 1 " & vbNewLine & _
             "      From ������ϼ�¼ A, Table(f_Num2list([1])) B, ����������� C" & vbNewLine & _
             "      Where a.ҽ��id = b.Column_Value And a.���id = c.���id And c.��� = [2] And Rownum < 2" & vbNewLine & _
             "      Union All" & vbNewLine & _
             "      Select 1" & vbNewLine & _
             "      From ������ϼ�¼ A, �������ҽ�� B, Table(f_Num2list([1])) C, ����������� D" & vbNewLine & _
             "      Where a.Id = b.���id And b.ҽ��id = c.Column_Value And a.���id = d.���id And d.��� = [2] And Rownum < 2" & vbNewLine & _
             ")"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "�����������-���", strMedicalIDs, 3)
    If rsTemp!Rec > 0 Then
        rsTemp.Close
        IsOutpatientControl = 1
        Exit Function
    End If
    
    '4-����
    strSQL = "Select Count(1) Rec" & vbNewLine & _
             "From (Select 1 " & vbNewLine & _
             "      From ������ϼ�¼ A, Table(f_Num2list([1])) B, ����������� C" & vbNewLine & _
             "      Where a.ҽ��id = b.Column_Value And a.����id = c.����id And c.��� = [2] And Rownum < 2" & vbNewLine & _
             "      Union All" & vbNewLine & _
             "      Select 1" & vbNewLine & _
             "      From ������ϼ�¼ A, �������ҽ�� B, Table(f_Num2list([1])) C, ����������� D" & vbNewLine & _
             "      Where a.Id = b.���id And b.ҽ��id = c.Column_Value And a.����id = d.����id And d.��� = [2] And Rownum < 2" & vbNewLine & _
             ")"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "�����������-����", strMedicalIDs, 4)
    If rsTemp!Rec > 0 Then
        rsTemp.Close
        IsOutpatientControl = 1
        Exit Function
    End If
    
    '5-ҩƷ
    strSQL = "Select Count(1) Rec " & vbNewLine & _
             "From ����ҽ����¼ A, Table(f_Num2list([1])) B, ����������� C " & vbNewLine & _
             "Where a.���Id = b.Column_Value And a.������Ŀid = c.ҩ��id And c.��� = [2] "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "�����������-ҩ��", strMedicalIDs, 5)
    If rsTemp!Rec > 0 Then
        rsTemp.Close
        IsOutpatientControl = 1
        Exit Function
    End If
    
    rsTemp.Close
    
    IsOutpatientControl = 2
    Exit Function
    
errHandle:
    If ErrCenter = 1 Then Resume
End Function