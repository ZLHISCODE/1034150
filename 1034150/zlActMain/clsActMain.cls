VERSION 1.0 CLASS
BEGIN
  MultiUse = 0   'False
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsActMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private blnReady As Boolean

Public Function BHCodeMan(ByVal lngSys As Long, ByVal lngModul As Long, ByVal lngMain As Long, _
    ByVal strDBUser As String, ByVal strPassword As String, ByVal strServer As String) As Boolean
'-----------------------------------------------------------------
'���ܣ�����������ָ�����ܣ�����ִ����س���
'������
'   lngSys,lngModul:��Ҫִ�й��ܵ�ϵͳ�ż�ģ���
'   lngMain:��������
'   strDbUser,strPassword,strServer:���ݿ������û���������(���������ת���������ݿ�����)��������
'���أ�ģ������Ƿ�ɹ�����;����һ���˳��򷵻�ʧ��
'-----------------------------------------------------------------
    Dim strTitle As String                  '��Ʒ����
    Dim strTag As String                    '�콢���־
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim strClass As String, objClass As Object
    Dim intCount As Integer, blnLoad As Boolean
    Dim strVersion As String, strCompareVersion As String
    Dim ArrayVersion As Variant
    Dim objRis As Object
    Dim intAtom As Integer, strCommon As String
    Dim strModul As String, strFuncs As String
    
    On Error GoTo ErrMain
    '��ʼ��Login����
    If gclsLogin Is Nothing Then
        Set gclsLogin = New clsRelogin
    End If
    gstrCommand = CStr(Command())
    
    gclsLogin.DBUser = UCase(strDBUser)
    gclsLogin.InputPwd = strPassword
    gclsLogin.InputUser = strDBUser
    gclsLogin.IsTransPwd = False
    gclsLogin.ServerName = strServer
    Screen.MousePointer = 11
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "��ʾ", "")
    '�ж��Ƿ��Ѿ���ģ̬����򿪣���������²������ٴ�ģ��
    strCommon = ExistModalWindows(lngMain)
    If strCommon <> "" Then
        Screen.MousePointer = 0
        'MsgBox """" & strCommon & """�����ж�ռ��ģ̬����򿪣����ȹر���ִ�е�ǰ������", vbInformation, gstrSysName
        Exit Function
    End If
    glngMain = lngMain
    '�����쳣����
    'Call zlKillHISPID
    '��¼��ش�����ͬһ����ֻ��¼һ��
    If gcnOracle.State <> adStateOpen Then
        'BH�ú�
        Load frmTmr
        Call frmTmr.SetTimr(True)
        On Error GoTo ErrMain
        '�������ݿ�(�̶���ת�����룬��Ϊ���������ת���������ݿ�����)
        If Not OraDataOpen(strServer, strDBUser, strPassword) Then Call frmTmr.SetTimr(False): Exit Function
        'д�뱾�����������EXE�ļ���
        Call SaveSetting("ZLSOFT", "����ȫ��", "ִ���ļ�", App.EXEName & ".exe")
        SaveSetting "ZLSOFT", "����ȫ��", "����·��", App.Path & "\" & App.EXEName & ".exe"
    
        '�Ƿ�����ʹ�ñ�����վ
        If Not CheckAllowByTerminal() Then Call frmTmr.SetTimr(False): Exit Function
        zl9ComLib.SetDbUser gclsLogin.DBUser
        zl9ComLib.SetNodeNo gclsLogin.NodeNo
        '��ʼ����������(������ע����֤)
        Call InitCommon(gcnOracle)
        On Error GoTo ErrMain
        '��ʼ����Ȩ��֤
        If Not RegCheck Then Call frmTmr.SetTimr(False): Exit Function
        '���ע����Ϣд��ע���
        Select Case zlRegInfo("��Ȩ����")
        Case "1" '��ʽ
            SaveSetting "ZLSOFT", "ע����Ϣ", "Kind", ""
        Case "2" '����
            SaveSetting "ZLSOFT", "ע����Ϣ", "Kind", "����"
        Case "3" '����
            SaveSetting "ZLSOFT", "ע����Ϣ", "Kind", "����"
        Case Else '����
            Screen.MousePointer = 0
            MsgBox "��Ȩ���ʲ���ȷ���������˳���", vbInformation, gstrSysName
            Call frmTmr.SetTimr(False)
            Exit Function
        End Select
    
        gstrSysName = zlRegInfo("��Ʒ����") & "����"
        SaveSetting "ZLSOFT", "ע����Ϣ", "��ʾ", gstrSysName
        SaveSetting "ZLSOFT", "ע����Ϣ", UCase("gstrSysName"), gstrSysName
        
        strTitle = zlRegInfo("��Ʒ����")
        If strTitle <> "" Then
            If InStr(strTitle, "-") > 0 Then
                If Split(strTitle, "-")(1) = "Ultimate" Then
                    strTag = "�콢��"
                ElseIf Split(strTitle, "-")(1) = "Professional" Then
                    strTag = "רҵ��"
                End If
            End If
        End If
        strTitle = Split(strTitle, "-")(0)
        SaveSetting "ZLSOFT", "ע����Ϣ", "��Ʒȫ��", strTitle
        SaveSetting "ZLSOFT", "ע����Ϣ", "��Ʒϵ��", strTag
        
        SaveSetting "ZLSOFT", "ע����Ϣ", "��Ʒ����", zlRegInfo("��Ʒ����")
        SaveSetting "ZLSOFT", "ע����Ϣ", "��λ����", zlRegInfo("��λ����", , -1)
        SaveSetting "ZLSOFT", "ע����Ϣ", "����֧����", zlRegInfo("����֧����", , -1)
        SaveSetting "ZLSOFT", "ע����Ϣ", "������", zlRegInfo("��Ʒ������", , -1)
        SaveSetting "ZLSOFT", "ע����Ϣ", "WEB֧���̼���", zlRegInfo("֧���̼���")
        SaveSetting "ZLSOFT", "ע����Ϣ", "WEB֧��EMAIL", zlRegInfo("֧����MAIL")
        SaveSetting "ZLSOFT", "ע����Ϣ", "WEB֧��URL", zlRegInfo("֧����URL")
        
        SaveSetting "ZLSOFT", "ע����Ϣ", UCase("gstrAviPath"), App.Path & "\�����ļ�"
        
        '�����׼�⣺
        '  ��֧�ֶ����ף���֧�ַǹ�����װ
        '  ָ����ϵͳ�ţ������ڶ����׵�����
        '###
        
        'ָ������̨���壬���ڹ��������ص������Բ�ָ��
        'Call SetMainForm(??)
        
        '��ʼ����ӡ�������漰Ȩ�޶�ȡ��Ҫ����Ȩ��֤֮��
        Call IniPrintMode(gcnOracle, gclsLogin.DBUser)
        '��ʼ����RISϵͳ
        On Error Resume Next
        Set objRis = CreateObject("zl9XWInterface.clsHISInner")
        Err.Clear: On Error GoTo 0
        If Not objRis Is Nothing Then
            '�̶���ת�����룬��Ϊ���������ת���������ݿ�����
            Call objRis.SaveDBConnectInfo(UCase(strDBUser), strPassword, strServer, False)
        End If
    End If
    Call frmTmr.SetTimr(True)
    blnReady = True
    '----------------------------------------------------------------------------------
    '��鲢��ȡģ���Ӧ�Ĳ���
    strSQL = "Select ����,���� From zlPrograms Where Nvl(ϵͳ,0) = [1] And ��� = [2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, App.Title, lngSys, lngModul)
    If rsTemp.EOF Then
        Screen.MousePointer = 0
        MsgBox "û�а�װ��ģ�飡", vbExclamation, gstrSysName
        Call frmTmr.SetTimr(False)
        Exit Function
    End If
    If IsNull(rsTemp!����) Then
        Screen.MousePointer = 0
        MsgBox "��ģ��Ĳ�����Ϣ��������", vbExclamation, gstrSysName
        Call frmTmr.SetTimr(False)
        Exit Function
    End If
    strClass = rsTemp!����
    strModul = rsTemp!����
    
    '��ȡ�Ѿ������Ĳ���
    blnLoad = False
    If Trim(Join(gstrClass)) <> "" Then
        For intCount = 0 To UBound(gstrClass)
            If gstrClass(intCount) = strClass Then
                Set objClass = gobjClass(intCount)
                blnLoad = True: Exit For
            End If
        Next
    End If
    'ʹ���²�������
    If UCase(strClass) = UCase("zl9EmrInterface") And blnLoad = False Then
        If gclsLogin.mobjEmr Is Nothing Then
            MsgBox "�����������ʧ�ܣ����鲢���µ�¼��", vbInformation, gstrSysName
            Call frmTmr.SetTimr(False)
            Exit Function
        ElseIf gclsLogin.mobjEmr.IsInited = False Then
            MsgBox "�������δ�ܳ�ʼ��," & gclsLogin.mobjEmr.GetError(), vbInformation, gstrSysName
            Call frmTmr.SetTimr(False)
            Exit Function
        End If
        If Trim(Join(gstrClass)) <> "" Then
            intCount = UBound(gstrClass)
        Else
            intCount = -1
        End If
        intCount = intCount + 1
        ReDim Preserve gstrClass(intCount)
        gstrClass(intCount) = strClass
        ReDim Preserve gobjClass(intCount)
        Set gobjClass(intCount) = gclsLogin.mobjEmr
        
        Dim strSpecify As String 'Ƭ�Σ�����Ȩ�޹̶��ڵ���ǰ����
        If Not gclsLogin.mobjEmr.HasInjectAuthorization(2201) Then
            strSpecify = GetPrivFunc(lngSys, 2201)
            Call gclsLogin.mobjEmr.InjectAuthorization(2201, strSpecify)
        End If
        If Not gclsLogin.mobjEmr.HasInjectAuthorization(2203) Then
            strSpecify = GetPrivFunc(lngSys, 2203)
            Call gclsLogin.mobjEmr.InjectAuthorization(2203, strSpecify)
        End If
        blnLoad = True
    End If
    '���û�д������´���
    If blnLoad = False Then
        On Error Resume Next
        Set objClass = CreateObject(strClass & ".cls" & Mid(strClass, 4))
        If Err.Number <> 0 Then
            Screen.MousePointer = 0
            MsgBox "���� " & strClass & " �����������������鰲װ�Ƿ���ȷ��", vbExclamation, gstrSysName
            Err.Clear
            Call frmTmr.SetTimr(False)
            Exit Function
        End If
        On Error GoTo ErrMain
        
        '���ò����İ汾�Ƿ�����ϵͳ����(���汾-3;�ΰ汾-3;���汾-3)
        If UCase(strClass) <> "ZL9REPORT" And UCase(strClass) <> "ZL9DOC" Then
            strSQL = "Select Nvl(���汾,1) As ���汾,Nvl(�ΰ汾,0) As �ΰ汾,Nvl(���汾,0) As ���汾,����" & _
                    " From zlComponent Where Upper(RTrim(����))=[1] And ϵͳ=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, App.Title, UCase(strClass), lngSys)
            If rsTemp.EOF Then
                Screen.MousePointer = 0
                MsgBox "ϵͳ�����ݲ�����������ϵͳ����Ա��ϵ��", vbExclamation, gstrSysName
                Call frmTmr.SetTimr(False)
                Exit Function
            End If
        
            With rsTemp
                '��װ�汾��Ϊ��λ���汾����λ�ΰ汾����λ���汾
                strCompareVersion = String(3 - Len(!���汾), "0") & !���汾 & "." & _
                                    String(3 - Len(!�ΰ汾), "0") & !�ΰ汾 & "." & _
                                    String(3 - Len(!���汾), "0") & !���汾
                ArrayVersion = Split(objClass.Version, ".")
                strVersion = String(3 - Len(ArrayVersion(0)), "0") & ArrayVersion(0) & "." & _
                             String(3 - Len(ArrayVersion(1)), "0") & ArrayVersion(1) & "." & _
                             String(3 - Len(ArrayVersion(2)), "0") & ArrayVersion(2)
                If strVersion < strCompareVersion Then
                    Screen.MousePointer = 0
                    MsgBox !���� & "(" & strClass & ")�İ汾�Ѳ�������ϵͳ����������ϵͳ����Ա��ϵ��", vbExclamation, gstrSysName
                    Call frmTmr.SetTimr(False)
                    Exit Function
                End If
            End With
        End If
        
        '���벿������
        If Trim(Join(gstrClass)) <> "" Then
            intCount = UBound(gstrClass)
        Else
            intCount = -1
        End If
        intCount = intCount + 1
        ReDim Preserve gstrClass(intCount)
        gstrClass(intCount) = strClass
        ReDim Preserve gobjClass(intCount)
        Set gobjClass(intCount) = objClass
    End If
    
    'ִ�в����е�ģ�鹦��
    If UCase(strClass) = "ZL9REPORT" Then
        '����������̨���Զ��屨��
        objClass.ReportOpen gcnOracle, lngSys, lngModul, Nothing
    ElseIf UCase(strClass) = UCase("zl9EmrInterface") Then
        On Error GoTo ErrMain
        strFuncs = GetPrivFunc(lngSys, lngModul)
        Call gclsLogin.mobjEmr.CodeMain(lngModul, strModul, lngMain, gclsLogin.InputUser, "[DBPASSWORD]" & gclsLogin.InputPwd, strFuncs)
    Else
        '������ͻ��˽��ͬ���
        Call CreateSynonyms(lngSys, lngModul)
        
        'ΪͨѶԭ�Ӹ�ֵ
        strCommon = Format(Now, "yyyyMMddHHmm")
        strCommon = TranPasswd(strCommon) & "||" & AnalyseComputer
        '����ͨѶԭ��
        intAtom = GlobalAddAtom(strCommon)
        Call SaveSetting("ZLSOFT", "����ȫ��", "����", intAtom)
        On Error Resume Next
        objClass.BHCodeMan lngSys, lngModul, gcnOracle, lngMain, gclsLogin.DBUser, gclsLogin
        '����Ĳ������������Է���,��ЩBHComdeManû�����һ��gclsLogin�������ᱨ�ô���
        If Err.Number = 450 Or Err.Number = 450 - 2146828288 Then
            Err.Clear
            On Error GoTo ErrMain
            objClass.BHCodeMan lngSys, lngModul, gcnOracle, lngMain, gclsLogin.DBUser
        Else
            If Err.Number <> 0 Then
                GoTo ErrMain
            Else
                On Error GoTo ErrMain
            End If
        End If
        Call GlobalDeleteAtom(intAtom)
        '��ҽ������ֻ��CodeMan()���ܻ�ȡϵͳ�ţ��ڶ�ȡ����ʱ����֪��ϵͳ�ţ���д��ע��������ҽ��������Ĭ��Ϊ 100
        Call SaveSetting("ZLSOFT", "����ȫ��", "ϵͳ��", lngSys)
    End If
    Call frmTmr.SetTimr(False)
    Screen.MousePointer = 0
    BHCodeMan = True
    Exit Function
ErrMain:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call frmTmr.SetTimr(False)
End Function

Public Function CloseWindows() As Boolean
    Dim intCount As Integer, objClass As Object
    Dim frmThis As Form
    
    CloseWindows = True
    If Trim(Join(gstrClass)) <> "" Then
        Err = 0: On Error Resume Next
        For intCount = 0 To UBound(gstrClass)
            Set objClass = gobjClass(intCount)
            If Not objClass.CloseWindows() Then CloseWindows = False
        Next
    End If
    For Each frmThis In Forms
        Unload frmThis
    Next
    If Not gclsLogin.mobjEmr Is Nothing Then
        gclsLogin.mobjEmr.Dispose
        Set gclsLogin.mobjEmr = Nothing
    End If
    Set gclsLogin = Nothing
End Function

Public Function IsReady() As Boolean
    IsReady = blnReady
End Function

Private Sub Class_Terminate()
    Call Me.CloseWindows
End Sub
