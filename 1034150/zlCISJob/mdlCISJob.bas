Attribute VB_Name = "mdlCISJob"
Option Explicit
Public gblnShowInTaskBar As Boolean         '�Ƿ���ʾ��������������
Public gobjRichEPR As New cRichEPR          '�������Ĳ���
Public gobjKernel As New zlPublicAdvice.clsPublicAdvice       '�ٴ����Ĳ���
Public gobjPath As New zlPublicPath.clsPublicPath             '�ٴ�·������
Public gobjRegist As Object
Public gobjCommunity As Object              '���������ӿڶ���
Public gclsInsure As New clsInsure          'ҽ������
Public gcnOracle As ADODB.Connection        '�������ݿ����ӣ��ر�ע�⣺��������Ϊ�µ�ʵ��
Public gstrPrivs As String                  '��ǰ�û����еĵ�ǰģ��Ĺ���
Public gcolPrivs As Collection              '��¼�ڲ�ģ���Ȩ��
Public gstrSysName As String                'ϵͳ����
Public glngSys As Long
Public glngModul As Long
Public gstrDBUser As String                 '��ǰ���ݿ��û�
Public gstrUnitName As String               '�û���λ����
Public gstrProductName As String            'OEM��Ʒ����
Public gfrmMain As Object                   '����̨����
Public gblnOK As Boolean
Public gobjCISBase As Object                '��ʿվ��ҽ��վ���������շѶ���
Public gobjPlugIn As Object                 '��ҹ��ܶ���
Public gblnKSSStrict As Boolean             '����ҩ���ϸ����
Public gbln�����ּ����� As Boolean  '�Ƿ����������ּ�����
Public gbln��Ѫ�ּ����� As Boolean  '�Ƿ�������Ѫ�ּ�����
Public gblnѪ��ϵͳ As Boolean  '�Ƿ�װѪ��ϵͳ
Public gobjEmr  As Object                   '�°没������
Public gbln���������Һ���Ч���� As Boolean   '�������������Һ���Ч�����Ĳ���
Public gobjLIS As Object                    'LIS���벿��
Public gobjPublicPacs As Object                  'PACS��������
Public gobjRis As Object 'RIS�ӿ�
Public gobjVitualExpense As Object
Public gobjPublicBlood As Object             'Ѫ�⹫������

'����ǩ��
Public gintCA As Integer '����ǩ����֤����
Public gstrESign As String '����ǩ�����Ƴ���
Public grsSign As Recordset  '����ǩ�����ò���

Public gbln��Ѫ����������� As Boolean  '��Ѫ�����������
'������ҩ�ӿ�����,0-δʹ��,1-����,2-��ͨ,3-̫Ԫͨ
Public gbytPass As Byte
'0-ҽ��ѡ��1-��ҩƷĿ¼���룬2-������Դ����
Public gint����������Դ As Integer
'̫Ԫͨ�ӿڶ���
Public gobjPass As Object

Public glngPreHWnd As Long '����֧�������ֹ���

Public Enum ���ó���
    E������� = 1
    EסԺ���� = 2
End Enum

'ϵͳ����
Public gstrLike As String   '�����˫��ƥ�䣬��Ϊ%
Public gint���� As Integer  '����ƥ�䷽ʽ��0-ƴ��,1-���
Public gbytDec As Byte '���ý���С����λ��
Public gstrDec As String '��С��λ������ĸ�ʽ����,��"0.0000"

Public gbytCardLen As Byte '���￨�ų���
Public gblnCardHide As Boolean '���￨��������ʾ

Public gbytBillOpt As Byte '���ѽ��ʵļ��ʵ��ݵĲ���Ȩ��:0-����,1-����,2-��ֹ��
Public gint��ͨ�Һ����� As Integer '��ͨ�Һŵ���Ч����
Public gint����Һ����� As Integer '����Һŵ���Ч����

Public gbln������֤ As Boolean '����һ��ͨ���Ѽ���ʣ����ʱ�Ƿ���Ҫ��֤
Public gblnִ��ǰ�Ƚ��� As Boolean '����һ��ͨ,��Ŀִ��ǰ�������շѻ��ȼ������

Public gblnִ�з��� As Boolean 'ִ��֮�������Զ�����
Public gblnִ�к���� As Boolean    'ִ�к��Զ���˻��۵�
Public gstrҽ���˶� As String    '��ѪƤ��ҽ����Ҫ�˶� ��λ��ȡ11����һλΪ ��Ѫҽ�����ڶ�λΪ Ƥ��ҽ��
Public gstr��Һ�������� As String          '��-�����ã���������
Public gblnDo As Boolean  '�Ƿ�ʹ�ø��Ի�����
Public gintҽ��ִ����Ч���� As Integer '�����޸�n���ڵǼǵ�ҽ��ִ�м�¼
Public gbln����Ӱ����Ϣϵͳ�ӿ� As Boolean
Public gbln����Ӱ����ϢϵͳԤԼ As Boolean

'�ڲ�Ӧ��ģ��Ŷ���
Public Enum Enum_Inside_Program
    p���Ӳ������� = 2250
    p�°�סԺ���� = 2252
    p�°����ﲡ�� = 2251
    p���ﲡ������ = 1250
    pסԺ�������� = 1251
    p����ҽ���´� = 1252
    pסԺҽ���´� = 1253
    pסԺҽ������ = 1254
    p������¼���� = 1255
    p�ٴ�·��Ӧ�� = 1256
    pҽ�����ѹ��� = 1257
    p���Ʊ������ = 1258
    p���Ӳ������� = 1259
    p����ҽ��վ = 1260
    pסԺҽ��վ = 1261
    pסԺ��ʿվ = 1262
    pҽ������վ = 1263
    P�°滤ʿվ = 1265
    p������ϲο� = 1270
    pҩƷ���Ʋο� = 1271
    p���˲������� = 1273
    p��Ƭ���߹��� = 1289
    p������� = 1132
    pסԺ���� = 1133
    p���ò�ѯ = 1139
    p���������� = 1113
    p�Ŷӽк�����ģ�� = 1160
    p������ҩ��� = 1266
    p������˹��� = 1267
    p���Ӳ������ = 1560
    p��Ѫ��˹��� = 1268
    p����ӿ� = 2425
    p������Ȩ���� = 1080
End Enum

Public Type TYPE_USER_INFO
    ID As Long
    ����ID As Long
    ��� As String
    ���� As String
    ���� As String
    �û��� As String
    ��ҩ���� As Long
End Type
Public UserInfo As TYPE_USER_INFO

'�����ļ���
Public Declare Function WNetAddConnection2 Lib "mpr.dll" Alias "WNetAddConnection2A" (lpNetResource As NETRESOURCE, ByVal lpPassword As String, ByVal lpUserName As String, ByVal dwFlags As Long) As Long
Public Declare Function WNetCancelConnection2 Lib "mpr.dll" Alias "WNetCancelConnection2A" (ByVal lpName As String, ByVal dwFlags As Long, ByVal fForce As Long) As Long
Public Declare Function WNetGetLastError Lib "mpr.dll" Alias "WNetGetLastErrorA" (lpError As Long, ByVal lpErrorBuf As String, ByVal nErrorBufSize As Long, ByVal lpNameBuf As String, ByVal nNameBufSize As Long) As Long
Public Type NETRESOURCE ' ������Դ
    dwScope As Long
    dwType As Long
    dwDisplayType As Long
    dwUsage As Long
    lpLocalName As String
    lpRemoteName As String
    lpComment As String
    lpProvider As String
End Type
Public Const RESOURCE_PUBLICNET = &H2
Public Const RESOURCETYPE_ANY = &H0
Public Const RESOURCEDISPLAYTYPE_GENERIC = &H0
Public Const RESOURCEUSAGE_CONNECTABLE = &H1
Public Const CONNECT_UPDATE_PROFILE = &H1

Public Function GetUserInfo() As Boolean
'���ܣ���ȡ��½�û���Ϣ
    Dim rsTmp As ADODB.Recordset
    
    UserInfo.�û��� = gstrDBUser
    UserInfo.���� = gstrDBUser
    Set rsTmp = zlDatabase.GetUserInfo
    If Not rsTmp Is Nothing Then
        If Not rsTmp.EOF Then
            UserInfo.ID = rsTmp!ID
            UserInfo.��� = rsTmp!���
            UserInfo.����ID = Nvl(rsTmp!����ID, 0)
            UserInfo.���� = Nvl(rsTmp!����)
            UserInfo.���� = Nvl(rsTmp!����)
                        UserInfo.�û��� = rsTmp!user & ""
            GetUserInfo = True
        End If
    End If
End Function

Public Function GetInsidePrivs(ByVal lngProg As Enum_Inside_Program, Optional ByVal blnLoad As Boolean, Optional ByVal lngSys As Long) As String
'���ܣ���ȡָ���ڲ�ģ���������е�Ȩ��
'������blnLoad=�Ƿ�̶����¶�ȡȨ��(���ڹ���ģ���ʼ��ʱ,�����û�ͨ��ע���ķ�ʽ�л���)
'      lngSys=ָ��ϵͳ���ڲ�ģ��Ȩ�ޣ���0�򲻴���Ĭ���ǵ�ǰϵͳ
    Dim strPrivs As String
    
    If gcolPrivs Is Nothing Then
        Set gcolPrivs = New Collection
    End If
    
    If lngSys = 0 Then lngSys = glngSys
    On Error Resume Next
    strPrivs = gcolPrivs("_" & lngProg)
    If err.Number = 0 Then
        If blnLoad Then
            gcolPrivs.Remove "_" & lngProg
        End If
    Else
        err.Clear: On Error GoTo 0
        blnLoad = True
    End If
    
    If blnLoad Then
        strPrivs = GetPrivFunc(lngSys, lngProg)
        gcolPrivs.Add strPrivs, "_" & lngProg
    End If
    GetInsidePrivs = IIf(strPrivs <> "", ";" & strPrivs & ";", "")
End Function

Public Function InitSysPar() As Boolean
'���ܣ���ʼ��ϵͳ����
'���أ���-�����ɹ�
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, strTmp As String
    
    '55928:������,2012-11-20
    gblnDo = Val(zlDatabase.GetPara("ʹ�ø��Ի����")) <> 0
    
    gstrLike = IIf(zlDatabase.GetPara("����ƥ��") = "0", "%", "")
    gint���� = Val(zlDatabase.GetPara("���뷽ʽ"))
    
    '���ý��С����λ��
    gbytDec = Val(zlDatabase.GetPara(9, glngSys, , 2))
    gstrDec = "0." & String(gbytDec, "0")
    
    '���￨��������ʾ
    strSql = "Select ���ų���, Nvl(��������, 0) �������� From ҽ�ƿ���� Where �ض���Ŀ = '���￨'"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "���￨")
    If rsTmp.RecordCount > 0 Then
        gblnCardHide = rsTmp!�������� <> "0"
        gbytCardLen = Val("" & rsTmp!���ų���)
    Else
        gblnCardHide = False
        gbytCardLen = 8
    End If
    
    
    '�Һ���Ч����
    strTmp = zlDatabase.GetPara(21, glngSys)
    If Len(strTmp) = 1 Then strTmp = strTmp & strTmp
    gint��ͨ�Һ����� = Val(Mid(strTmp, 1, 1))
    gint����Һ����� = Val(Mid(strTmp, 2, 1))
    
    '���ѽ��ʵļ��ʵ��ݵĲ���Ȩ��:0-����,1-����,2-��ֹ��
    gbytBillOpt = Val(zlDatabase.GetPara(23, glngSys))
    
    'һ��ͨ������֤
    gbln������֤ = Val(zlDatabase.GetPara(28, glngSys)) <> 0
    
    'ִ��֮�������Զ�����
    gblnִ�з��� = Val(zlDatabase.GetPara(33, glngSys)) <> 0
    
    'ִ�к��Զ����
    gblnִ�к���� = Val(zlDatabase.GetPara(81, glngSys)) <> 0
    
    '����һ��ͨ,��Ŀִ��ǰ�������շѻ��ȼ������
    gblnִ��ǰ�Ƚ��� = Val(zlDatabase.GetPara(163, glngSys)) <> 0
    
    '����ǩ����֤����
    gintCA = Val(zlDatabase.GetPara(25, glngSys))
    
    '����ǩ�����Ƴ���
    gstrESign = zlDatabase.GetPara(26, glngSys)
    
    '��ȡ������������
    If glngModul = p����ҽ��վ Or glngModul = pסԺҽ��վ Or glngModul = pסԺ��ʿվ Or glngModul = pҽ������վ Or _
        glngModul = P�°滤ʿվ Or glngModul = p������ҩ��� Then
        '��ȡ������������
        Set grsSign = New ADODB.Recordset
        grsSign.Fields.Append "����ID", adBigInt
        grsSign.Fields.Append "����", adBigInt
        grsSign.Fields.Append "�Ƿ�����", adBigInt
        grsSign.CursorLocation = adUseClient
        grsSign.LockType = adLockOptimistic
        grsSign.CursorType = adOpenStatic
        grsSign.Open
    End If
    
    '��Ѫ��Ƥ��ҽ��ִ�к���Ҫ�˶�
    gstrҽ���˶� = zlDatabase.GetPara(186, glngSys)
    
    '����ҩ��ּ�����
    gblnKSSStrict = Val(zlDatabase.GetPara(187, glngSys)) <> 0
    
    '�Ƿ����������ּ�����
    gbln�����ּ����� = Val(zlDatabase.GetPara(209, glngSys)) <> 0
    
    '�Ƿ�������Ѫ�ּ�����
    gbln��Ѫ�ּ����� = Val(zlDatabase.GetPara(216, glngSys)) <> 0
    
    '�Ƿ�װѪ��ϵͳ
    gblnѪ��ϵͳ = (IsSetUp(2200) And Val(zlDatabase.GetPara(236, glngSys)) <> 0)
    
    '�������������Һ���Ч�����Ĳ���
    gbln���������Һ���Ч���� = Val(zlDatabase.GetPara(210, glngSys)) <> 0
    
    '61762:������,2012-05-20
    gstr��Һ�������� = Get��Һ��������

    '��Ѫ�����������
    gbln��Ѫ����������� = Val(zlDatabase.GetPara(218, glngSys)) <> 0
    
    '�����޸�n���ڵǼǵ�ҽ��ִ�м�¼
    gintҽ��ִ����Ч���� = Val(zlDatabase.GetPara(220, glngSys))
    '������ҩ�ӿ����ͣ�0-δ���ã�1-�Ĵ���ͨ��2-��ͨ��3-̫Ԫͨ
    gbytPass = Val(zlDatabase.GetPara(30, glngSys))
    
    '����������Դ����
    gint����������Դ = Val(zlDatabase.GetPara(224, glngSys))
    
    gbln����Ӱ����Ϣϵͳ�ӿ� = Val(zlDatabase.GetPara(255, glngSys)) = 1
    
    InitSysPar = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function IsSetUp(ByVal lngϵͳ��� As Long) As Boolean
'���ܣ��Ƿ�װĳ��ϵͳ
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    
    strSql = "Select 1 From zlSystems Where ���=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlCISKernel", lngϵͳ���)
    IsSetUp = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function Get�Һ�ID(ByVal strNO As String) As Long
'���ܣ����ݹҺŵ���ȡ�Һ�ID
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    
    strSql = "Select ID From ���˹Һż�¼ Where NO=[1] And ��¼����=1 And ��¼״̬=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "Get�Һ�ID", strNO)
    If Not rsTmp.EOF Then Get�Һ�ID = rsTmp!ID
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetFullNO(ByVal strNO As String, ByVal intNum As Integer) As String
'���ܣ����û�����Ĳ��ݵ��ţ�����ȫ���ĵ��š�
'������intNum=��Ŀ���,Ϊ0ʱ�̶��������
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, intType As Integer
    Dim dtCurDate As Date, strMaxNo As String
    Dim strYearStr As String
    
    err = 0: On Error GoTo errH:
    If Len(strNO) >= 8 Then
        GetFullNO = Right(strNO, 8)
        Exit Function
    ElseIf Len(strNO) = 7 Then
        GetFullNO = PreFixNO & strNO
        Exit Function
    End If
'    ElseIf intNum = 0 Then
'        GetFullNO = PreFixNO & Format(Right(strNO, 7), "0000000")
'        Exit Function
'    End If
    GetFullNO = strNO
    
    strSql = "Select ��Ź���,Sysdate as ����,������ From ������Ʊ� Where ��Ŀ���=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, App.ProductName, intNum)
    dtCurDate = Date
    If Not rsTmp.EOF Then
        intType = Val("" & rsTmp!��Ź���)
        dtCurDate = rsTmp!����
        strMaxNo = Nvl(rsTmp!������)
    End If
    strYearStr = PreFixNO
    If strMaxNo = "" Then strMaxNo = strYearStr & "000001"
    If intType = 1 Then
        '���ձ��
        strSql = Format(CDate(Format(dtCurDate, "YYYY-MM-dd")) - CDate(Format(dtCurDate, "YYYY") & "-01-01") + 1, "000")
        GetFullNO = PreFixNO & strSql & Format(Right(strNO, 4), "0000")
        Exit Function
    End If
    '������
    If Len(strNO) = 6 Then
        GetFullNO = Left(strMaxNo, 2) & strNO: Exit Function
    End If
    GetFullNO = Left(strMaxNo, 2) & zlLeftPad(Right(strNO, 6), 6, "0")
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zlLeftPad(ByVal strCode As String, lngLen As Long, Optional strChar As String = " ") As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ָ���������ƿո�
    '����:�����ִ�
    '����:���˺�
    '����:2012-02-22 17:58:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngTmp As Long
    Dim strTmp As String
    strTmp = strCode
    lngTmp = LenB(StrConv(strCode, vbFromUnicode))
    If lngTmp < lngLen Then
        strTmp = String(lngLen - lngTmp, strChar) & strTmp
    ElseIf lngTmp > lngLen Then  '���ڳ���ʱ,�Զ��ض�
        strTmp = zlSubstr(strCode, 1, lngLen)
    End If
    zlLeftPad = Replace(strTmp, Chr(0), strChar)
End Function

Private Function zlSubstr(ByVal strInfor As String, ByVal lngStart As Long, ByVal lngLen As Long) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ���ִ���ֵ,�ִ��п��԰�������
    '���:strInfor-ԭ��
    '         lngStart-ֱʼλ��
    '         lngLen-����
    '����:�Ӵ�
    '����:���˺�
    '����:2012-02-22 18:00:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTmp As String, i As Long
    err = 0: On Error GoTo ErrHand:
    zlSubstr = StrConv(MidB(StrConv(strInfor, vbFromUnicode), lngStart, lngLen), vbUnicode)
    zlSubstr = Replace(zlSubstr, Chr(0), " ")
    Exit Function
ErrHand:
    zlSubstr = ""
End Function

Public Function GetAdvicePause(ByVal lngҽ��ID As Long) As String
'���ܣ���ȡָ��ҽ������ͣʱ��μ�¼
'���أ�"��ͣʱ��,��ʼʱ��;...."
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, i As Long
    Dim strTmp As String
    
    On Error GoTo errH
    
    strSql = "Select ��������,����ʱ�� From ����ҽ��״̬" & _
        " Where �������� IN(6,7) And ҽ��ID=[1]" & _
        " Order by ����ʱ��"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlCISWork", lngҽ��ID)
    For i = 1 To rsTmp.RecordCount
        If rsTmp!�������� = 6 Then
            strTmp = strTmp & ";" & Format(rsTmp!����ʱ��, "yyyy-MM-dd HH:mm:ss") & ","
        ElseIf rsTmp!�������� = 7 Then
            '���õ���һ�벻����ͣ�ķ�Χ֮��
            strTmp = strTmp & Format(DateAdd("s", -1, rsTmp!����ʱ��), "yyyy-MM-dd HH:mm:ss")
        End If
        rsTmp.MoveNext
    Next
    GetAdvicePause = Mid(strTmp, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPatiDiagnose(ByVal lng����ID As Long, ByVal lng����ID As Long, ByVal int��Դ As Integer) As String
'���ܣ���ȡ����ָ���ξ�����������
'������lng����ID=�Һ�ID����ҳID
'      int��Դ=1-����,2-סԺ
'���أ���"��"�ŷָ��Ķ����ϴ�
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    
    strSql = "Select ��¼��Դ,�������,��ϴ���,�������,�Ƿ�����,Mod(�������,10) as ���� From ������ϼ�¼" & _
        " Where ����ID=[1] And ��ҳID=[2] And NVL(�������,1) = 1 And ������� IN(" & IIf(int��Դ = 1, "1,11", "1,2,3,11,12,13") & ")" & _
        " Order by ��¼��Դ,�������,��ϴ���"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "GetPatiDiagnose", lng����ID, lng����ID)
    
    '�Ȱ���Դ����˳�����
    rsTmp.Filter = "��¼��Դ=3" '��ҳ����
    If rsTmp.EOF Then rsTmp.Filter = "��¼��Դ=2" '��Ժ�Ǽ�
    If rsTmp.EOF Then rsTmp.Filter = "��¼��Դ=1" '����
    If rsTmp.EOF Then rsTmp.Filter = "��¼��Դ=4" '������¼��
    
    'סԺ�ٰ���������˳�����
    If Not rsTmp.EOF And int��Դ = 2 Then
        strSql = rsTmp.Filter
        rsTmp.Filter = strSql & " And ����=3"
        If rsTmp.EOF Then rsTmp.Filter = strSql & " And ����=2"
        If rsTmp.EOF Then rsTmp.Filter = strSql & " And ����=1"
    End If
    
    strSql = ""
    Do While Not rsTmp.EOF
        If Not IsNull(rsTmp!�������) Then
            strSql = strSql & "��" & rsTmp!������� & IIf(Nvl(rsTmp!�Ƿ�����, 0) = 1, "������", "")
        End If
        rsTmp.MoveNext
    Loop
    
    GetPatiDiagnose = Mid(strSql, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function is����(ByVal lng����ID As Long) As Boolean
    Dim rsTmp As ADODB.Recordset, strSql As String
     
    On Error GoTo errH
    strSql = "Select �������� From ��������˵�� Where ��������='����' And ����ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, App.ProductName, lng����ID)
    is���� = rsTmp.RecordCount > 0
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ReadAdviceSignSource(ByVal int�������� As Integer, _
    ByVal lng����ID As Long, ByVal varTime As Variant, strIDs As String, _
    ByVal lngǩ��ID As Long, ByVal blnMoved As Boolean, strSource As String, _
    Optional ByVal lngǰ��ID As Long, Optional ByVal colSomeTime As Collection) As Integer
'���ܣ���ȡ�������ڵ���ǩ��/��֤��ҽ��Դ������
'������
'  int��������=Ҫǩ��/��֤ǩ����ҽ��״̬
'  ǩ��ʱ���룺
'    lng����ID
'    varTime=���˹Һŵ��Ż���ҳID
'    strIDs=ָ��Ҫǩ����ҽ��ID����(��ID)
'    lngǰ��ID=�¿�ҽ��Ҫǩ����ҽ����Դ(�Ƿ�ҽ��)
'    colSomeTime=ĳҽ����ʱ�����ݣ���ֹͣҽ��ǩ��ʱ���������ҽ��ִ����ֹʱ������ݣ�У��ʱ����У��ʱ������
'  ��֤ǩ��ʱ��
'    lngǩ��ID=ǩ����¼��ID
'    blnMoved=�Ƿ�ҽ��������ת��
'���أ�ǩ��/��֤ǩ����Դ�����ɹ���
'      strIDs=ǩ��/��֤ǩ����ҽ��ID����(ÿ����ϸID)
'      strSource=ǩ��/��֤ǩ����ҽ��Դ��
    Dim rsTmp As New ADODB.Recordset
    Dim str��IDs As String, strSql As String, i As Long
    Dim arrField As Variant, strField As String
    Dim strLine As String, intRule As Integer
    
    On Error GoTo errH
    
    str��IDs = strIDs
    strSource = "": strIDs = ""
    intRule = 1 '�������µ�ҽ��ǩ��Դ�����ɹ�����
    
    If lngǩ��ID = 0 Then
        'ǩ��ʱ
        If int�������� = 1 Then
            '���¿���ҽ������ǩ�������ξ���/סԺ��ǰҽ�����´��δǩ��ҽ��
            strSql = _
                " Select /*+ Rule*/ A.* From ����ҽ����¼ A,����ҽ��״̬ B" & _
                " Where A.ID=B.ҽ��ID And B.ǩ��ID is Null And B.��������=1" & _
                " And A.ҽ��״̬=1 And Nvl(A.ǰ��ID,0)=[5]" & _
                " And Decode(A.��˱��,1,Substr(A.����ҽ��,1,Instr(A.����ҽ��,'/')-1),Substr(A.����ҽ��,Instr(A.����ҽ��,'/')+1))=[3]" & _
                " And Exists(Select M.���� From ��Ա�� M,ִҵ��� N" & _
                "       Where M.����=Decode(A.��˱��,1,Substr(A.����ҽ��,1,Instr(A.����ҽ��,'/')-1),Substr(A.����ҽ��,Instr(A.����ҽ��,'/')+1))" & _
                "         And M.ִҵ���=N.���� And N.���� IN('ִҵҽʦ','ִҵ����ҽʦ')" & _
                "   )" & _
                IIf(TypeName(varTime) = "String", " And A.����ID+0=[1] And A.�Һŵ�=[2]", " And A.����ID=[1] And A.��ҳID=[2]") & _
                IIf(str��IDs <> "", " And Nvl(A.���ID,A.ID) IN(Select Column_Value From Table(f_Num2list([4])))", "") & _
                " Order by A.Ӥ��,Nvl(A.���ID,A.ID),A.���"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlCISKernel", lng����ID, varTime, UserInfo.����, str��IDs, lngǰ��ID)
        Else
            '��Ҫ���ϡ�ֹͣ��У�Ե�ҽ������ǩ�����¿�ʱǩ������ָ��ҽ������һ���ǵ�ǰҽ���´�
            strSql = _
                " Select /*+ Rule*/ A.* From ����ҽ����¼ A,����ҽ��״̬ B" & _
                " Where A.ID=B.ҽ��ID And B.ǩ��ID is Not Null And B.��������=1" & _
                IIf(TypeName(varTime) = "String", " And A.����ID+0=[1] And A.�Һŵ�=[2]", " And A.����ID=[1] And A.��ҳID=[2]") & _
                IIf(str��IDs <> "", " And Nvl(A.���ID,A.ID) IN(Select Column_Value From Table(f_Num2list([3])))", "") & _
                " Order by A.Ӥ��,Nvl(A.���ID,A.ID),A.���"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlCISKernel", lng����ID, varTime, str��IDs)
        End If
    Else
        '��֤ǩ��ʱ:�ȶ�ȡǩ��ʱ��Դ�����ɹ���
        strSql = "Select ǩ������ From ҽ��ǩ����¼ Where ID=[1]"
        If blnMoved Then
            strSql = Replace(strSql, "ҽ��ǩ����¼", "Hҽ��ǩ����¼")
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlCISKernel", lngǩ��ID)
        If Not rsTmp.EOF Then intRule = Nvl(rsTmp!ǩ������, 1)
        '--
        strSql = _
            " Select A.* From ����ҽ����¼ A,����ҽ��״̬ B" & _
                " Where A.ID=B.ҽ��ID And B.ǩ��ID=[1] Order by A.Ӥ��,Nvl(A.���ID,A.ID),A.���"
        If blnMoved Then
            strSql = Replace(strSql, "����ҽ����¼", "H����ҽ����¼")
            strSql = Replace(strSql, "����ҽ��״̬", "H����ҽ��״̬")
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlCISKernel", lngǩ��ID)
    End If
    
    'ҽ��Դ�ĵĲ�ͬ���ɹ���
    If intRule = 1 Then
        If int�������� = 3 Then
            strField = "ID,���ID,����,�Ա�,����,Ӥ��,ҽ����Ч,��ʼִ��ʱ��,ҽ������,�걾��λ,��������,�ܸ�����," & _
                "ҽ������,ִ��Ƶ��,Ƶ�ʴ���,Ƶ�ʼ��,�����λ,ִ��ʱ�䷽��,У��ʱ��,ִ������,������־,����ҽ��,����ʱ��"
        ElseIf int�������� = 8 Then
            strField = "ID,���ID,����,�Ա�,����,Ӥ��,ҽ����Ч,��ʼִ��ʱ��,ҽ������,�걾��λ,��������,�ܸ�����," & _
                "ҽ������,ִ��Ƶ��,Ƶ�ʴ���,Ƶ�ʼ��,�����λ,ִ��ʱ�䷽��,ִ����ֹʱ��,ִ������,������־,����ҽ��,����ʱ��"
        Else
            strField = "ID,���ID,����,�Ա�,����,Ӥ��,ҽ����Ч,��ʼִ��ʱ��,ҽ������,�걾��λ,��������,�ܸ�����," & _
                "ҽ������,ִ��Ƶ��,Ƶ�ʴ���,Ƶ�ʼ��,�����λ,ִ��ʱ�䷽��,ִ������,������־,����ҽ��,����ʱ��"
        End If
    End If
    arrField = Split(strField, ",")
        
    '����ҽ��ǩ��Դ��
    Do While Not rsTmp.EOF
        strLine = ""
        For i = 0 To UBound(arrField)
            If lngǩ��ID = 0 And int�������� = 3 And arrField(i) = "У��ʱ��" Then
                'У��ҽ��ǩ��ʱ,��У��ʱ�����⴦������������ִ�й���֮ǰȡǩ��Դ��,��ʱ��δд�����ݿ�
                strLine = strLine & vbTab & colSomeTime("_" & Nvl(rsTmp!���ID, rsTmp!ID))
            ElseIf lngǩ��ID = 0 And int�������� = 8 And arrField(i) = "ִ����ֹʱ��" Then
                'ֹͣҽ��ǩ��ʱ,����ֹʱ�����⴦������������ִ�й���֮ǰȡǩ��Դ��,��ʱ��δд�����ݿ�
                strLine = strLine & vbTab & colSomeTime("_" & Nvl(rsTmp!���ID, rsTmp!ID))
            Else
                If IsNull(rsTmp.Fields(arrField(i)).Value) Then
                    strLine = strLine & vbTab & ""
                Else
                    If IsType(rsTmp.Fields(arrField(i)).Type, adDBTimeStamp) Then
                        strLine = strLine & vbTab & Format(rsTmp.Fields(arrField(i)).Value, "yyyy-MM-dd HH:mm:ss")
                    Else
                        strLine = strLine & vbTab & rsTmp.Fields(arrField(i)).Value
                    End If
                End If
            End If
        Next
        strSource = strSource & vbCrLf & Mid(strLine, 2)
        strIDs = strIDs & "," & rsTmp!ID
        rsTmp.MoveNext
    Loop
    
    strSource = Mid(strSource, 3)
    strIDs = Mid(strIDs, 2)
    
    ReadAdviceSignSource = intRule
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPatiDept(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal bytMode As Byte) As Long
'���ܣ���ȡ���˵�ǰ�����Ϳ���
'������bytMode=0-�����,1=�鲡��
    Dim rsTmp As ADODB.Recordset, strSql As String
    
    On Error GoTo errH
    strSql = "Select " & IIf(bytMode = 0, "��ǰ����id", "��ǰ����id") & " as ����ID" & vbNewLine & _
            "From ������Ϣ" & vbNewLine & _
            "Where ����id = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, App.ProductName, lng����ID, lng��ҳID)
    If rsTmp.RecordCount > 0 Then GetPatiDept = Val("" & rsTmp!����ID)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPatiLog(lng����ID As Long, lng��ҳID As Long) As ADODB.Recordset
'���ܣ���ȡ���˱䶯��¼
    Dim strSql As String
    
    On Error GoTo errH
    strSql = "Select ��ֹԭ��,��ֹʱ��,��ʼԭ��,Decode(��ʼԭ��, 1, '��Ժ��ס', 2, '��ס', 3," & _
            " Decode(��ʼʱ��, Null, 'ת��', 'ת����ס'), 4, '����', 5, '��λ�ȼ��䶯', 6, '�����ȼ��䶯', 7," & vbNewLine & _
            "               '����ҽʦ�ı�', 8, '���λ�ʿ�ı�', 9, 'תΪסԺ����', 10, 'Ԥ��Ժ', 11, '����ҽʦ�䶯'," & _
            " 12, '����ҽʦ�䶯', 13, '�����䶯',14,'תҽ��С��',15,Decode(��ʼʱ��, Null, 'ת����', 'ת������ס')) ����" & vbNewLine & _
            "From ���˱䶯��¼" & vbNewLine & _
            "Where Nvl(���Ӵ�λ, 0) = 0 And ����id = [1] And ��ҳid = [2]" & vbNewLine & _
            "Order By ��ֹʱ�� Desc, ��ʼʱ�� Desc"
    Set GetPatiLog = zlDatabase.OpenSQLRecord(strSql, App.ProductName, lng����ID, lng��ҳID)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPati������Ϣ(ByVal lng����ID As Long, lng��ҳID As Long) As String
'���ܣ���ȡ��ǰ���˵ķ�����Ϣ
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    Dim lng�������� As Long
    
    strSql = "Select �������� From ������ҳ Where ����ID=[1] And ��ҳID=[2]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "GetPati������Ϣ", lng����ID, lng��ҳID)
    If Not rsTmp.EOF Then
        lng�������� = Val(rsTmp!�������� & "")
    End If
    strSql = _
        " Select �������,Ԥ�����,0 as Ԥ�����,0 as ������ From ������� Where ����=1 And ����ID=[1] And ���� = [3]" & _
        " Union ALL" & _
        " Select 0,0,0, Sum(������) as ������ From ���˵�����¼ Where ����id = [1] And ��ҳid = [2] And ɾ����־ = 1 And (Sysdate <= ����ʱ�� Or ����ʱ�� Is Null)" & _
        " Union ALL" & _
        " Select 0,0,Sum(���),0 From ����ģ����� A,������ҳ B" & _
        " Where A.����ID=B.����ID And A.��ҳID=B.��ҳID And B.���� Is Not Null And A.����ID=[1] And A.��ҳID=[2]"
    strSql = "Select Sum(�������) as �������,Sum(Ԥ�����) as Ԥ�����,Sum(Ԥ�����) as Ԥ�����,sum(������) as ������ From (" & strSql & ")"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "GetPati������Ϣ", lng����ID, lng��ҳID, IIf(lng�������� = 1, 1, 2))
    If Not rsTmp.EOF Then
        GetPati������Ϣ = _
            "Ԥ�����:" & FormatEx(Nvl(rsTmp!Ԥ�����, 0), 2) & ",δ�����:" & FormatEx(Nvl(rsTmp!�������, 0), 2) & _
            IIf(Nvl(rsTmp!Ԥ�����, 0) <> 0, ",Ԥ�����:" & FormatEx(Nvl(rsTmp!Ԥ�����, 0), 2), "") & _
            ",ʣ���:" & FormatEx(Nvl(rsTmp!Ԥ�����, 0) - Nvl(rsTmp!�������, 0) + Nvl(rsTmp!Ԥ�����, 0), 2) & ",������:" & Nvl(rsTmp!������, 0)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetסԺ����ҩռ��(ByVal lng����ID As Long, lng��ҳID As Long) As String
'���ܣ���ȡ��ǰ���˵�סԺ����ҩռ��
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    
    strSql = "Select /*+ RULE */" & vbNewLine & _
            " c.���� As ����, Sum(Decode(a.�շ����, '5', a.ʵ�ս��, 0)) As ��ҩ��, Sum(Decode(a.�շ����, '6', a.ʵ�ս��, 0)) As ��ҩ��," & vbNewLine & _
            " Sum(Decode(a.�շ����, '7', a.ʵ�ս��, 0)) As ��ҩ��, Sum(Decode(a.�շ����, '5', 0, '6', 0, '7', 0, a.ʵ�ս��)) As ��ҩ��," & vbNewLine & _
            " Sum(a.ʵ�ս��) As ���з�" & vbNewLine & _
            "From סԺ���ü�¼ A, Table(f_Num2list2([1])) B, ������Ϣ C" & vbNewLine & _
            "Where a.����id = b.C1 And a.��ҳid = b.C2 And b.C1 = c.����id And a.��¼״̬ <> 0 Having Sum(a.ʵ�ս��) > 0" & vbNewLine & _
            "Group By c.����" & vbNewLine & _
            "Order By ����"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "GetסԺ����ҩռ��", lng����ID & ":" & lng��ҳID)
    If Not rsTmp.EOF Then
        GetסԺ����ҩռ�� = ",ҩռ��:" & Format((Val(rsTmp!��ҩ��) + Val(rsTmp!��ҩ��) + Val(rsTmp!��ҩ��)) / Val(rsTmp!���з�) * 100, "0.0") & "%"
    Else
        GetסԺ����ҩռ�� = ",ҩռ��:0.0%"
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    GetסԺ����ҩռ�� = ",ҩռ��:0.0%"
End Function


Public Sub LoadOldData(ByVal strOld As String, ByRef txt���� As TextBox, ByRef cbo���䵥λ As ComboBox)
'����:�����ݿ��б�������䰴�淶�ĸ�ʽ���ص�����,���淶��ԭ����ʾ
    Dim strTmp As String, lngidx As Long
    
    If Trim(strOld) = "" Then Exit Sub

    If InStr(strOld, "��") > 0 Then
        If InStr(strOld, "��") = Len(strOld) Then
            strTmp = Mid(strOld, 1, InStr(strOld, "��") - 1)
            lngidx = 0
        End If
    ElseIf InStr(strOld, "��") > 0 Then
        If InStr(strOld, "��") = Len(strOld) Then
            strTmp = Mid(strOld, 1, InStr(strOld, "��") - 1)
            lngidx = 1
        End If
    ElseIf InStr(strOld, "��") > 0 Then
        If InStr(strOld, "��") = Len(strOld) Then
            strTmp = Mid(strOld, 1, InStr(strOld, "��") - 1)
            lngidx = 2
        End If
    ElseIf IsNumeric(strOld) Then
        strTmp = strOld
        lngidx = 0
    End If
    
    If IsNumeric(strTmp) Then
        txt����.Text = strTmp
    Else
        txt����.Text = strOld
        lngidx = -1
    End If
    
    If cbo���䵥λ.ListCount > 0 Then Call zlControl.CboSetIndex(cbo���䵥λ.hwnd, lngidx)
    If lngidx = -1 Then
        cbo���䵥λ.Visible = False
    Else
        If cbo���䵥λ.Visible = False Then cbo���䵥λ.Visible = True
    End If
End Sub

Public Function LoadPatiAllergy(ByVal lng����ID As Long, Optional ByRef objCbo As Object, Optional ByRef rsAller As ADODB.Recordset) As Boolean
'���ܣ���ȡ���˵Ĺ�����¼����������
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    Dim strTmp As String
        
    strSql = "Select Distinct B.����ʱ�� as �Һ�ʱ��,D.���� as �Һſ���,C.��ҳID,E.���� as סԺ����,A.����ID," & _
        " A.ҩ����,Nvl(A.����ʱ��,A.��¼ʱ��) as ����ʱ��,B.NO as �Һŵ�,A.ҩ��ID,A.����Դ����,A.������Ӧ,(max(lengthB(a.ҩ����)) over()-lengthB(a.ҩ����)+4) AS �ո񳤶�" & _
        " From ���˹�����¼ A,���˹Һż�¼ B,������ҳ C,���ű� D,���ű� E" & _
        " Where A.����ID=B.����ID(+) And A.��ҳID=B.ID(+) And B.��¼����(+)=1 And B.��¼״̬(+)=1" & _
        " And A.����ID=C.����ID(+) And A.��ҳID=C.��ҳID(+)" & _
        " And B.ִ�в���ID=D.ID(+) And C.��Ժ����ID=E.ID(+)" & _
        " And A.���=1 And ҩ���� is Not NULL And A.����ID=[1] And Not Exists" & vbNewLine & _
        " (Select ҩ��id" & vbNewLine & _
        "       From ���˹�����¼" & vbNewLine & _
        "       Where (Nvl(ҩ��id, 0) = Nvl(a.ҩ��id, 0) Or Nvl(ҩ����, 'Null') = Nvl(a.ҩ����, 'Null')) And Nvl(���, 0) = 0 And" & vbNewLine & _
        "             ��¼ʱ��>A.��¼ʱ�� And ����id = [1])" & _
        " Order by Nvl(A.����ʱ��,A.��¼ʱ��) Desc"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "LoadPatiAllergy", lng����ID)
        
    If Not objCbo Is Nothing Then
        objCbo.Clear
        Do While Not rsTmp.EOF
'            If Not IsNull(rsTmp!�Һ�ʱ��) Then
'                strTmp = Format(rsTmp!����ʱ��, "yyyy-MM-dd") & "," & Nvl(rsTmp!ҩ����) & ",�������:" & Nvl(rsTmp!�Һſ���)
'            Else
'                strTmp = Format(rsTmp!����ʱ��, "yyyy-MM-dd") & "," & Nvl(rsTmp!ҩ����) & ",��" & rsTmp!��ҳID & "��סԺ:" & Nvl(rsTmp!סԺ����)
'            End If
            strTmp = Nvl(rsTmp!ҩ����) & String(Val(rsTmp!�ո񳤶�), " ") & Format(rsTmp!����ʱ��, "yyyy-MM-dd") & String(4, " ")

            If Not IsNull(rsTmp!������Ӧ) Then strTmp = strTmp & IIf(Nvl(rsTmp!������Ӧ) = "", "", "������Ӧ:" & rsTmp!������Ӧ)

            objCbo.AddItem strTmp
            
            rsTmp.MoveNext
        Loop
        If objCbo.ListCount = 0 Then
            objCbo.AddItem "�޼�¼"
        End If
        objCbo.ListIndex = 0
        objCbo.ForeColor = vbRed
    End If
    
    If Not rsAller Is Nothing Then Set rsAller = rsTmp
        
    LoadPatiAllergy = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetOwner(ByVal lngSys As Long) As String
'���ܣ���ȡָ��ϵͳ��������
    Dim rsTmp As New ADODB.Recordset
    Dim strSql  As String
    
    On Error GoTo errH
    strSql = "Select ������ From zlSystems Where ���=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "GetOwner", lngSys)
    If Not rsTmp.EOF Then
        GetOwner = rsTmp!������
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetRefuseReason(ByVal lng����ID As Long, lng��ҳID As Long) As String
'���ܣ���ȡ���˵Ĳ����ύ��������
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    
    '�Ըô�סԺ���һ�α��ܵ�Ϊ׼
    strSql = "Select �������� From (Select �������� From �����ύ��¼ Where ����ID=[1] And ��ҳID=[2] And ��¼״̬=2 Order by ID Desc) Where Rownum=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "GetRefuseReason", lng����ID, lng��ҳID)
    If Not rsTmp.EOF Then GetRefuseReason = Nvl(rsTmp!��������)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function PatiMedRecHaveSubmit(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
'���ܣ����ָ�����˵Ĳ����Ƿ��Ѿ��ύ(ͨ���ύ��¼)
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    
    '�Ըô�סԺ���һ�α��ܵ�Ϊ׼
    strSql = "Select 1 From �����ύ��¼ Where ����ID=[1] And ��ҳID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "PatiMedRecHaveSubmit", lng����ID, lng��ҳID)
    PatiMedRecHaveSubmit = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ReadCaseMap(lngID As Long) As StdPicture
'���ܣ����ݱ��ͼID����ͼ�ζ���
    Dim lngFileSize As Long, arrBin() As Byte
    Dim strFile As String, intFile As Integer
    
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    
    strSql = "Select ͼ�� From �������ͼ Where Ԫ��ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlCISWork", lngID)
        
    If rsTmp.EOF Then Exit Function
    If IsNull(rsTmp!ͼ��) Then Exit Function
    
    On Error GoTo 0
    
    intFile = FreeFile
    strFile = CurDir & "\zlNewPicture" & Timer & ".pic"
    
    Open strFile For Binary As intFile
    
    lngFileSize = rsTmp.Fields("ͼ��").ActualSize
    ReDim arrBin(lngFileSize - 1) As Byte
    arrBin() = rsTmp.Fields("ͼ��").GetChunk(lngFileSize)
    Put intFile, , arrBin()
    Close intFile
    
    Set ReadCaseMap = VB.LoadPicture(strFile)
    Kill strFile
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub ViewImage(ByVal lngҽ��ID As Long, frmParent As Object, Optional ByVal blnMoved As Boolean = False, Optional ByVal strPrivs As String = "")
'���ܣ����ù�Ƭվ
    Dim objPacsCore As Object
    Dim strFtpHost As String
    Dim strSql As String, rsTmp As ADODB.Recordset
    Dim strSDPath As String
    Dim strSDUser As String
    Dim strSDPwd As String
    Dim intImageLocation As Long
    Dim lng����ID As Long
    Dim blnCanViewImage As Boolean  '��ҽ���ı��滹û�����(û����ʽǩ�������ִ��)ʱ���Ƿ���Թ�Ƭ
    Dim lngִ�п���ID As Long
    
    On Error GoTo DBError
    
    '��ѯͼ����RISPACS����������PACS,�Լ�ִ�п���ID
    strSql = "Select ͼ��λ��,ִ�п���ID from Ӱ�����¼ where ҽ��ID =[1]"

    If blnMoved Then
        strSql = Replace(strSql, "Ӱ�����¼", "HӰ�����¼")
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ѯͼ�����ڵ�λ��", lngҽ��ID)
    
    If rsTmp.RecordCount <> 0 Then
        intImageLocation = Nvl(rsTmp!ͼ��λ��, 0)
        lngִ�п���ID = Val(Nvl(rsTmp!ִ�п���ID, 0))
    End If
    
    If lngִ�п���ID > 0 Then
        strSql = "Select ����ֵ from Ӱ�����̲��� where ����ID = [1] and ������=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡ����", lngִ�п���ID, "��ͼ��ҽ��վ���ɹ�Ƭ")
        If rsTmp.RecordCount > 0 Then blnCanViewImage = Val(Nvl(rsTmp!����ֵ, 0)) = 1
    Else
        blnCanViewImage = False
    End If
    
    '���û�й�ѡ���� ��ͼ��ҽ��վ���ɹ�Ƭ�������жϸ�ҽ���ı����Ƿ����
    If Not blnCanViewImage Then
        strSql = "Select ����ID From ����ҽ������ Where ҽ��ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "CheckEPRReport", lngҽ��ID)
        If Not rsTmp.EOF Then lng����ID = rsTmp!����id
        
        If CheckEPRReport(lngҽ��ID, IIf(lng����ID = 0, -1, lng����ID)) <> 1 Then
            MsgBox "��ҽ���ı��滹û�����(û����ʽǩ�������ִ��)�����ڲ��ܲ鿴ͼ��", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    'ͼ����RIS���ݿ⣬�����RIS��WEB���
    If intImageLocation = 1 Then
        Call XWWebViewerOpen(lngҽ��ID)
        Exit Sub
    End If
    
    
    '���ж��Ƿ����ͼ��û��ͼ������ʾ���˳�
    strSql = "Select A.���UID,Count(B.����UID) as �������� From Ӱ�����¼ A,Ӱ�������� B Where A.���UID=B.���UID And A.ҽ��ID=[1] Group by A.���UID"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��Ƭ����", lngҽ��ID)
    If rsTmp.EOF Then
        MsgBox "û�п����ڹ�Ƭ�ı���ͼ��", vbInformation, gstrSysName
        Exit Sub
    End If

    strFtpHost = ""
    
    '������Ҫ�򿪵�����ͼ����Ϣ
    strSql = "Select /*+RULE*/ D.IP��ַ As Host1,d.�豸�� as �豸��1," & _
        "Decode(C.��������,Null,'',to_Char(C.��������,'YYYYMMDD')||'\')" & _
        "||C.���UID||'\' As Path,E.IP��ַ As Host2,e.�豸�� as �豸��2, " & _
        "D.����Ŀ¼ AS ����Ŀ¼1, E.����Ŀ¼ AS ����Ŀ¼2,D.����Ŀ¼�û��� as ����Ŀ¼�û���1, " & _
        "E.����Ŀ¼�û��� AS ����Ŀ¼�û���2,D.����Ŀ¼���� AS ����Ŀ¼����1,E.����Ŀ¼���� AS ����Ŀ¼����2 " & _
        "From Ӱ�����¼ C,Ӱ���豸Ŀ¼ D,Ӱ���豸Ŀ¼ E " & _
        "Where C.λ��һ=D.�豸��(+) And C.λ�ö�=E.�豸��(+) And C.ҽ��ID=[1] "
        
    '�����ת����־�����ȡת������ʷ��
    If blnMoved Then
        strSql = Replace(strSql, "Ӱ�����¼", "HӰ�����¼")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡ����Ŀ¼��Ϣ", lngҽ��ID)
    
    If rsTmp.RecordCount > 0 Then
        '�������صĻ���Ŀ¼����Ҫ�ڵ��ù�Ƭվ֮ǰ�ȴ������Ŀ¼����Ƭվ��ֻ�����أ����������ػ���Ŀ¼
        MkLocalDir App.Path & "\TmpImage\" & rsTmp("Path")
        ClearCacheFolder App.Path & "\TmpImage\"
        
        '��ȡFTP�����������û��������룬IP��ַ��
        If rsTmp("�豸��1") <> "" Then
            strFtpHost = rsTmp("Host1")
            strSDPath = Nvl(rsTmp("����Ŀ¼1"))
            strSDUser = Nvl(rsTmp("����Ŀ¼�û���1"))
            strSDPwd = Nvl(rsTmp("����Ŀ¼����1"))
        ElseIf Nvl(rsTmp("�豸��2")) <> "" Then
            strFtpHost = rsTmp("Host2")
            strSDPath = Nvl(rsTmp("����Ŀ¼2"))
            strSDUser = Nvl(rsTmp("����Ŀ¼�û���2"))
            strSDPwd = Nvl(rsTmp("����Ŀ¼����2"))
        End If
        
        '�жϹ���Ŀ¼�Ƿ��Ѿ����ӣ����û�����ӣ����������
        On Error Resume Next
        If strSDPath <> "" Then
            Call funcConnectShardDir("\\" & strFtpHost & "\" & strSDPath, strSDUser, strSDPwd)
        End If
        
        Set objPacsCore = CreateObject("zl9PacsCore.clsViewer")
        objPacsCore.CallOpenViewer "", lngҽ��ID, frmParent, gcnOracle, blnMoved, False
    End If

    Exit Sub
DBError:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Function XWWebViewerOpen(lngOrderID As Long) As Long
''--------------------------------------------
''���ܣ� ��RIS��WEB Viewer
'           lngOrderID -- ҽ��ID
''���أ�0-�ɹ�;1-����
''--------------------------------------------
    Dim strIp As String
    Dim strURL As String
    
    On Error GoTo err
    
    strIp = zlDatabase.GetPara("XWWEB������IP", glngSys, 1288, "")
    
    If strIp <> "" Then
        strURL = "C:\Program Files\Internet Explorer\iexplore.exe http://" & strIp & ":8080/imageweb/imageAction.action?ColID0=22&ColValue0=" & lngOrderID
        
        Shell strURL, vbMaximizedFocus
        XWWebViewerOpen = 0
    Else
        MsgBox "WEB������IP��ַΪ�գ��������ú�WEB��������", vbOKOnly, "��ʾ��Ϣ"
        XWWebViewerOpen = 1
    End If
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Function funcConnectShardDir(strShareRemoteDir As String, strUserName As String, strPassWord As String) As Long
    '����������Դ
    Dim NetR As NETRESOURCE
    Dim lngResult As Long
    
    NetR.dwType = RESOURCETYPE_ANY
    NetR.lpLocalName = vbNullString
    NetR.lpRemoteName = strShareRemoteDir
    NetR.lpProvider = vbNullString
    lngResult = WNetAddConnection2(NetR, strPassWord, strUserName, 0)
    
    If lngResult <> 0 Then
        MsgBox "��������ʧ�ܣ��������������Ƿ���ȷ��"
    End If
    funcConnectShardDir = lngResult
End Function

Public Sub MkLocalDir(ByVal strDir As String)
'���ܣ���������Ŀ¼
    Dim objFile As New Scripting.FileSystemObject
    Dim aNestDirs() As String, i As Integer
    Dim strPath As String
    On Error Resume Next
    
    '��ȡȫ����Ҫ������Ŀ¼��Ϣ
    ReDim Preserve aNestDirs(0)
    aNestDirs(0) = strDir
    
    strPath = objFile.GetParentFolderName(strDir)
    Do While Len(strPath) > 0
        ReDim Preserve aNestDirs(UBound(aNestDirs) + 1)
        aNestDirs(UBound(aNestDirs)) = strPath
        strPath = objFile.GetParentFolderName(strPath)
    Loop
    '����ȫ��Ŀ¼
    For i = UBound(aNestDirs) To 0 Step -1
        MkDir aNestDirs(i)
    Next
End Sub

Public Sub ClearCacheFolder(ByVal strCacheFolder As String)
'���ܣ���ָ��Ŀ¼�Ĵ�С�ﵽһ���ٷֱ�ʱ����ո�Ŀ¼
    Dim objFile As New Scripting.FileSystemObject
    Dim objCurFolder As Scripting.Folder, objCurFile As Scripting.File, objFiles As Scripting.Files
    Dim strDriver As String
    
    On Error Resume Next
    strDriver = objFile.GetDriveName(strCacheFolder)
    Set objCurFolder = objFile.GetFolder(strCacheFolder)
    If objCurFolder.Size / objFile.GetDrive(strDriver).FreeSpace > 0.2 Then
        objCurFolder.Delete True
    End If
End Sub
Public Function zlGetFeeFields(Optional strTableName As String = "������ü�¼", Optional blnReadDatabase As Boolean = False) As String
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���ȡָ������ֵ
    '��Σ�strTableName:��:������ü�¼;סԺ���ü�¼;....
    '      blnReadDatabase-�����ݿ��ж�ȡ
    '���Σ�
    '���أ��ֶμ�
    '���ƣ����˺�
    '���ڣ�2010-03-10 10:41:42
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSql As String, strFileds As String
    
    err = 0: On Error GoTo ErrHand:
    If blnReadDatabase Then GoTo ReadDataBaseFields:
    Select Case strTableName
    Case "������ü�¼"
        zlGetFeeFields = "" & _
        "Id, ��¼����, No, ʵ��Ʊ��, ��¼״̬, ���, ��������, �۸񸸺�, ���ʵ�id, ����id, ҽ�����, �����־, ���ʷ���, " & _
        "����, �Ա�, ����, ��ʶ��, ���ʽ, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ����, ��ҩ����, ����, " & _
        "�Ӱ��־, ���ӱ�־, Ӥ����, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ������, ��������id, ������, " & _
        "����ʱ��, �Ǽ�ʱ��, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, ����, ����Ա���, ����Ա����, ����id, ���ʽ��, " & _
        "���մ���id, ������Ŀ��, ���ձ���, ��������, ͳ����, �Ƿ��ϴ�, ժҪ, �Ƿ���"
        Exit Function
    Case "סԺ���ü�¼"
        zlGetFeeFields = "" & _
         " Id, ��¼����, No, ʵ��Ʊ��, ��¼״̬, ���, ��������, �۸񸸺�, �ಡ�˵�, ���ʵ�id, ����id, ��ҳid, ҽ�����, " & _
         " �����־, ���ʷ���, ����, �Ա�, ����, ��ʶ��, ����, ���˲���id, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, " & _
         " ����, ��ҩ����, ����, �Ӱ��־, ���ӱ�־, Ӥ����, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ������, " & _
         " ��������id, ������, ����ʱ��, �Ǽ�ʱ��, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, ����, ����Ա���, ����Ա����, " & _
         " ����id , ���ʽ��, ���մ���ID, ������Ŀ��, ���ձ���, ��������, ͳ����, �Ƿ��ϴ�, ժҪ, �Ƿ���"
         Exit Function
    Case "���˽��ʼ�¼"
        zlGetFeeFields = "Id, No, ʵ��Ʊ��, ��¼״̬, ��;����, ����id, ����Ա���, ����Ա����, �շ�ʱ��, ��ʼ����, ��������, ��ע"
        Exit Function
    Case "����Ԥ����¼"
        zlGetFeeFields = "" & _
        " Id, ��¼����, No, ʵ��Ʊ��, ��¼״̬, ����id, ��ҳid, ����id, �ɿλ, ��λ������, ��λ�ʺ�, ժҪ, ���, " & _
        " ���㷽ʽ, �������, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ�, �Ҳ�"
        Exit Function
    Case "��Ա��"
        zlGetFeeFields = "" & _
        "Id, ���, ����, ����, ����֤��, ��������, �Ա�, ����, ��������, �칫�ҵ绰, �����ʼ�, ִҵ���, ִҵ��Χ, " & _
        "����ְ��, רҵ����ְ��, Ƹ�μ���ְ��, ѧ��, ��ѧרҵ, ��ѧʱ��, ��ѧ����, ������ѵ, ���п���, ���˼��, ����ʱ��, " & _
        "����ʱ��, ����ԭ��, ����, վ��"
        Exit Function
    End Select
ReadDataBaseFields:
    err = 0: On Error GoTo ErrHand:
    strSql = "Select  column_name From user_Tab_Columns Where Table_Name = Upper([1]) Order By Column_ID;"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "��ȡ����Ϣ", strTableName)
    strFileds = ""
    With rsTemp
        Do While Not .EOF
            strFileds = strFileds & "," & Nvl(!column_name)
            .MoveNext
        Loop
        If strFileds <> "" Then strFileds = Mid(strFileds, 2)
    End With
    If strFileds = "" Then strFileds = "*"
    zlGetFeeFields = strFileds
    Exit Function
ErrHand:
  zlGetFeeFields = "*"
End Function

Public Function zlGetFullFieldsTable(Optional strTableName As String = "������ü�¼", Optional bytHistory As Byte = 2, _
    Optional strWhere As String = "", Optional blnSubTable As Boolean = True, Optional strAliasName As String = "A", Optional blnReadDatabaseFields As Boolean = False)
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���ȡһ�����ݱ��е��ֶ�.������Select Id,....
    '��Σ�bytHistory-0-��������ʷ����,1-��������ʷ����,2-����������( select * from tablename Union select * from Htablename)
    '      strWhere-����
    '      blnSubTable-�Ƿ��ӱ�
    '      strAliasName-����
    '���Σ�
    '���أ�select ID ... From tableName Union ALL
    '���ƣ����˺�
    '���ڣ�2010-03-10 11:19:11
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim strFields As String, strSql As String
    
    strFields = zlGetFeeFields(Trim(strTableName), blnReadDatabaseFields)
    Select Case bytHistory
    Case 0 '��
        strSql = "  Select  " & strFields & " From " & strTableName & " " & strWhere
    Case 1 '����ʷ
        strSql = " Select  " & strFields & " From H" & Trim(strTableName) & " " & strWhere
    Case Else '���߶�����
        strSql = " Select  " & strFields & " From " & Trim(strTableName) & " " & strWhere & " UNION ALL " & " Select  " & strFields & " From H" & Trim(strTableName) & " " & strWhere
    End Select
    If blnSubTable Then strSql = " (" & strSql & ") " & strAliasName
    zlGetFullFieldsTable = strSql
End Function

 
Public Function zlGetLocaleComputerNamePara(ByVal varPara As Variant, Optional ByVal lngSys As Long, Optional ByVal lngModual As Long, Optional ByVal strDeFault As String, _
        Optional strComputerName As String = "") As String
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���ȡָ����������
    '��Σ�varPara=�����Ż�������������ֻ��ַ����ʹ�������
    '      lngSys=ʹ�øò�����ϵͳ��ţ���100
    '      lngModual=ʹ�øò�����ģ��ţ���1230
    '      strDefault=�����ݿ���û�иò���ʱʹ�õ�ȱʡֵ(ע�ⲻ��Ϊ��ʱ)
    '     strComputerName-��ȡָ������������
    '���Σ�
    '���أ�����ֵ���ַ�����ʽ
    '���ƣ����˺�
    '���ڣ�2010-06-07 13:56:22
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, i As Integer, rsPara As ADODB.Recordset, rsUserPara As ADODB.Recordset
    Dim blnNew As Boolean, blnEnabled As Boolean
    
    On Error GoTo errH
    
    strSql = "Select ID,Nvl(����ֵ,ȱʡֵ) as ����ֵ,SYS_CONTEXT('USERENV','TERMINAL') as MName From zlParameters where ģ��=[1] and ϵͳ=[2]"
    If TypeName(varPara) = "String" Then
        strSql = strSql & " and ������=[3]"
    Else
        strSql = strSql & " and ������=[4]"
    End If
    Set rsPara = zlDatabase.OpenSQLRecord(strSql, "GetPara", lngModual, lngSys, CStr(varPara), Val(varPara))
    If rsPara.EOF = False Then
        strSql = _
            "   Select ����ֵ " & _
            "   From zlUserParas Where ����ID=[1]  and  ������=[2]"
        Set rsUserPara = zlDatabase.OpenSQLRecord(strSql, "GetPara", Val(Nvl(rsPara!ID)), IIf(strComputerName = "", CStr(Nvl(rsPara!MName)), strComputerName))
         If Not rsUserPara.EOF Then
                zlGetLocaleComputerNamePara = Nvl(rsUserPara!����ֵ, strDeFault)
         Else
                zlGetLocaleComputerNamePara = Nvl(rsPara!����ֵ, strDeFault)
         End If
    Else
        zlGetLocaleComputerNamePara = strDeFault
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function CheckDoctorPatisIsValid() As Byte
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ����ҽ�����������Ƿ���Ч
    '���أ�0-����̨�������;1-ҽ����������;2-�ȷ���̨���ƽ�,����ҽ������
    '���ƣ����˺�
    '���ڣ�2010-06-07 14:32:47
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim blnValid As Boolean, strComputerName As String

    '���˺�:Ӧ�����Ŷӽкŵĺ����˴�:��Ҫ��Ϸ���̨ģ����Ŷӽк�ģʽΪ���������ŶӺ���վ��=1ʱ��Ч
     
     '��Ҫ����Ƿ�Ϊҽ���������з�ʽ
     '�ŶӽкŴ���ģʽ:1.��������̨������л�ҽ����������;2-�ȷ������,��ҽ�����о���.0-���Ŷӽк�
     blnValid = Val(zlDatabase.GetPara("�Ŷӽк�ģʽ", glngSys, p����������)) = 1
    If blnValid Then
         '����Ҫ���:�ŶӺ���վ��=1
         '�ŶӺ���վ��: 0-��������̨�������;1-����ҽ����������
         strComputerName = zlDatabase.GetPara("Զ�˺���վ��", glngSys, p�Ŷӽк�����ģ��)
        blnValid = Val(zlGetLocaleComputerNamePara("�ŶӺ���վ��", glngSys, p����������, "0", strComputerName)) = 1
    End If
    CheckDoctorPatisIsValid = blnValid
End Function

Public Sub PrintInMedRec(ByRef objClsMedRec As zlMedRecPage.clsInOutMedRec, ByVal intType As Integer, ByVal lng����ID As Long, ByVal lng��ҳID As Long, _
        ByRef objReport As Object, ByVal lng����ID As Long, ByRef objForm As Object, Optional intPage As Integer)
'���ܣ���ҳ��ӡ��Ԥ��
'������intType=2����ӡ����=1��Ԥ����0=����
'     mobjReport-��ӡ������lng����ID-���˿��ң�mobjForm-������
'     intPage=1-4��ӡ��ҳ������ʽ��=5��ӡ����+��ҳ1��=6��ӡ����+��ҳ2
'    If lng����ID <> 0 Then
'        If objClsMedRec Is Nothing Then
'            Set objClsMedRec = New clsInOutMedRec
'            Call objClsMedRec.InitMedRec(gcnOracle, glngSys, gobjCommunity, gclsInsure)
'        End If
'        Call objClsMedRec.PrintOrPriviewInMedRec(intType, lng����ID, lng��ҳID, objReport, lng����ID, objForm, intPage)
'    End If
'    Exit Sub
    Dim strName As String
    Dim lngPage As Long
    
    If lng����ID <> 0 Then
        If objReport Is Nothing Then Set objReport = New clsReport
        Select Case Val(zlDatabase.GetPara("������ҳ��׼", glngSys, pסԺҽ��վ, "0"))
    
            Case 0 '��������׼
                If Have��������(lng����ID, "��ҽ��") Then
                    strName = "ZL1_INSIDE_1261_4"
                Else
                    strName = "ZL1_INSIDE_1261_1"
                End If
            Case 1    '�Ĵ�ʡ��׼
                If Have��������(lng����ID, "��ҽ��") Then
                    strName = "ZL1_INSIDE_1261_6"
                Else
                    strName = "ZL1_INSIDE_1261_5"
                End If
            Case 2    '����ʡ��׼
                If Have��������(lng����ID, "��ҽ��") Then
                    strName = "ZL1_INSIDE_1261_8"
                Else
                    strName = "ZL1_INSIDE_1261_7"
                End If
            Case 3    '����ʡ��׼
                If Have��������(lng����ID, "��ҽ��") Then
                    strName = "ZL1_INSIDE_1261_10"
                Else
                    strName = "ZL1_INSIDE_1261_9"
                End If
        End Select
        If GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\zl9Report\LocalSet\" & strName, "AllFormat", 0) = 0 And intPage = 0 Then
            Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\zl9Report\LocalSet\" & strName, "AllFormat", 1)
        End If
        If intType = 0 Then
            Call ReportPrintSet(gcnOracle, glngSys, strName, objForm)
        Else
            If intPage = 5 Then
                lngPage = 1
            ElseIf intPage = 6 Then
                lngPage = 2
            Else
                lngPage = intPage
            End If
            Call objReport.ReportOpen(gcnOracle, glngSys, strName, objForm, "����ID=" & lng����ID, "��ҳID=" & lng��ҳID, IIf(intPage <> 0, "ReportFormat=" & lngPage, ""), intType)
            If intPage > 4 Then
                Call objReport.ReportOpen(gcnOracle, glngSys, strName, objForm, "����ID=" & lng����ID, "��ҳID=" & lng��ҳID, IIf(intPage <> 0, "ReportFormat=" & lngPage + 2, ""), intType)
            End If
        End If
    End If
End Sub

Public Function CheckDiseaseFile(ByRef frmParent As Object, ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal intCurDeptID As Long, _
ByVal str����IDs As String, ByVal str���IDs As String, Optional ByRef lngFileID As Long, Optional ByVal blnOnlyCheck As Boolean, Optional ByRef blnNo As Boolean) As Boolean

'���ܣ���鲡����Щ����֤������û����д����ʾ������д
'����:frmParent    ������
'     lng����ID    ����ID
'     lng��ҳID    ���ﴫ�Һ�ID��סԺ����ҳID
'     intCurDeptID ��д��������ID
'     lngҽ��ID    ҽ��ID�����ڼ�鱨�棩
'     blnOnlyCheck true-ֻ���δ��д���������������б�,false-�����δ��д�����򵯳��б�
'     blnNo        �Ƿ�Ҫ��д��Ⱦ�����濨
   Dim rsTmp As ADODB.Recordset
   
   On Error GoTo errH
   
    If str����IDs = "" And str���IDs = "" Then
        CheckDiseaseFile = True
        Exit Function
    End If
    Dim strSql As String
    If str����IDs <> "" Then
        strSql = strSql & " Union Select �ļ�ID From ��������ǰ�� Where ����ID IN (Select Column_Value From Table(f_Num2list([3])))"
    End If
    If str���IDs <> "" Then
        strSql = strSql & " Union Select �ļ�ID From ��������ǰ�� Where ���ID IN (Select Column_Value From Table(f_Num2list([4])))"
    End If
    On Error GoTo errH
    strSql = "(" & Mid(strSql, 8) & ") Minus Select �ļ�ID From ���Ӳ�����¼ Where ����ID=[1] And ��ҳID=[2] And ��������=5"
    strSql = "Select /*+ Rule*/" & vbNewLine & _
            " a.Id, a.����, a.���, a.����, a.˵��" & vbNewLine & _
            "From �����ļ��б� A ,(" & strSql & ") B Where A.ID=B.�ļ�ID  And" & vbNewLine & _
            "(a.ͨ�� = 1 Or a.ͨ�� = 2 And Exists (Select 1 From ����Ӧ�ÿ��� C Where c.�ļ�id = a.Id And c.����id = [5]))" & vbNewLine & _
            "Order By a.���"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "CheckDiseaseFile", lng����ID, lng��ҳID, str����IDs, str���IDs, intCurDeptID)
    blnNo = False
    If rsTmp.RecordCount = 0 Then
        CheckDiseaseFile = True
        Exit Function
    Else
        strSql = ""
        rsTmp.MoveFirst
        Do While Not rsTmp.EOF
            strSql = strSql & vbCrLf & "��" & rsTmp!���� & "��"
            rsTmp.MoveNext
        Loop
    End If

    If rsTmp.RecordCount = 1 Then
        If blnOnlyCheck Then
            If MsgBox("���ݲ��˵������Ϣ�����¼���֤�����滹û����д��" & vbCrLf & vbCrLf & Mid(strSql, 3) & vbCrLf & vbCrLf & "Ҫ������", _
                    vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then blnNo = True: Exit Function
        Else
            If MsgBox("���ݲ��˵������Ϣ�����¼���֤�����滹û����д��" & vbCrLf & vbCrLf & Mid(strSql, 3) & vbCrLf & vbCrLf & "Ҫ������", _
                    vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                rsTmp.MoveFirst
                lngFileID = Val(rsTmp!ID & "")
            Else
                blnNo = True
            End If
        End If
    ElseIf rsTmp.RecordCount > 1 Then
        If blnOnlyCheck Then
            If MsgBox("���ݲ��˵������Ϣ�����¼���֤�����滹û����д��" & vbCrLf & vbCrLf & Mid(strSql, 3) & vbCrLf & vbCrLf & "Ҫ������", _
                    vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then blnNo = True: Exit Function
        Else
            If MsgBox("���ݲ��˵������Ϣ�����¼���֤�����滹û����д��" & vbCrLf & vbCrLf & Mid(strSql, 3) & vbCrLf & vbCrLf & "Ҫ������", _
                    vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                If frmDiseaseFileList.ShowMe(frmParent, rsTmp, lngFileID) = False Then
                    lngFileID = 0
                    CheckDiseaseFile = True
                    Exit Function
                End If
            Else
                blnNo = True
            End If
        End If
    End If
    CheckDiseaseFile = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function OpenEPRDoc(ByRef EPRDoc As Object, ByRef frmParent As Object, ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal intCurDeptID As Long, ByRef str����ID As String, ByRef str���ID As String, Optional ByVal lngPatiFrom As Long, Optional ByVal lngҽ��ID As Long = 0, _
 Optional ByVal blnMoved As Boolean, Optional ByVal blnOnCkeck As Boolean, Optional ByRef blnNo As Boolean)
'���ܣ����ݲ��˻�����Ϣ��򿪲����༭��
'����:EPRDoc       �����ļ��༭������
'     frmParent    ������
'     lngFileID       Ҫ�򿪵Ĳ����ļ�ID
'     lng����ID    ����ID
'     lng��ҳID    ���ﴫ�Һ�ID��סԺ����ҳID
'     intCurDeptID ��д��������ID
'     lngҽ��ID    ҽ��ID�����ڼ�鱨�棩
'     lngPatiFrom     ������Դ   ����Ϊ1��סԺΪ2
'     blnMoved     �����Ƿ�ת��
'     blnOnCkeck   �Ƿ�ֻ���м��
'     blnNO        �Ƿ�Ҫ��д��Ⱦ�����濨
    Dim lngFileID As Long
    If CheckDiseaseFile(frmParent, lng����ID, lng��ҳID, intCurDeptID, str����ID, str���ID, lngFileID, blnOnCkeck, blnNo) = False Then OpenEPRDoc = False: Exit Function
    
    If lngFileID = 0 Then OpenEPRDoc = False: Exit Function
    On Error GoTo errH
    If EPRDoc Is Nothing Then
        Set EPRDoc = New zlRichEPR.cEPRDocument
    End If
    Call EPRDoc.InitEPRDoc(0, 2, lngFileID, lngPatiFrom, lng����ID, lng��ҳID, , intCurDeptID, lngҽ��ID, blnMoved)
    Call EPRDoc.ShowEPREditor(frmParent)
    OpenEPRDoc = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub Set�ٴ��Թ�ҩ(objFrom As Object)
     On Error Resume Next
    If gobjCISBase Is Nothing Then
        Set gobjCISBase = CreateObject("zl9CISBase.clsCISBase")
        If gobjCISBase Is Nothing Then
            MsgBox "���ƻ�������(ZLCISBase)û����ȷ��װ���ù����޷�ִ�С�", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    err.Clear: On Error GoTo 0
    
    Call gobjCISBase.SetMedList(objFrom, gcnOracle, glngSys, gstrDBUser)
End Sub

Public Sub SetPublicRTFFont(ByRef objRTF As Object, ByVal bytSize As Byte)
'���ܣ���RichTextBox������������
'������objRTF RichTextBox����
'      bytSize 0-С����,1-������
    With objRTF
        .SelStart = 0
        .SelLength = Len(.Text)
        .SelFontSize = IIf(bytSize = 0, 9, 12)
        .SelLength = 0
    End With
End Sub

Public Function CheckMecRed(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal strfrmCation As String, Optional ByVal strOperateName As String) As Boolean
'���ܣ���鲡���Ƿ��Ѿ���Ŀ,�����Ƿ��ڴ������������(��ʱ��ҳ��������״̬���������޸�)
'       lng����ID:��ǰ����ID
'       lng��ҳID:��ǰ������ҳID
'       strfrmCation:���øú����Ĵ�������
'       strOperateName:���øú����Ĳ������ơ�strOperateNameΪ��ʱ����������ʾ
    Dim strSql As String, rsTmp As Recordset
    Dim int����״̬ As Integer
    Dim strMsg As String
    
    On Error GoTo errH
    '��ȡ����״̬
    strSql = "Select Nvl(����״̬, 0) ����״̬ From ������ҳ Where ����id = [1] And ��ҳid = [2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, strfrmCation, lng����ID, lng��ҳID)
    rsTmp.MoveFirst
    int����״̬ = rsTmp!����״̬
    '��ҳ���������ж�
    Select Case int����״̬
        Case 1 '�ȴ����
            strMsg = "�ò����ȴ������,����"
        Case 3 '�������
            strMsg = "�ò������������,����"
        Case 5 '���鵵
            strMsg = "�ò����Ѿ����鵵,����"
        Case 10 '���մ���
            strMsg = "�ò����ڽ��մ�����,����"
        Case Else '2-�ܾ����4-��鷴��;6-�������;13-���ڳ��;14-��鷴��;16-�������
            strMsg = ""
    End Select
    
    If strMsg = "" Then
        strSql = "Select ��Ŀ���� from ������ҳ where ����ID=[1] And ��ҳID=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, strfrmCation, lng����ID, lng��ҳID)
        If Not IsNull(rsTmp!��Ŀ����) Then
            strMsg = "�ò��˵Ĳ����Ѿ���Ŀ������"
        End If
    End If
    
    If strMsg <> "" Then  '������ҳ
        If strOperateName <> "" Then
            MsgBox strMsg & strOperateName & "��", vbInformation, gstrSysName
        End If
        Exit Function
    End If
    
    CheckMecRed = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CanUnExec(ByVal datExec As Date, Optional ByVal datNow As Date) As Boolean
'���ܣ�����ִ�м�¼��ִ��ʱ���ж��ܷ�ȡ��ִ�л�ȡ�����
'������datExec=ִ�м�¼��ִ��ʱ��
'      datNow =��ǰʱ��
'���أ�CanUnExec=true-����ȡ��ִ�У�false-������ȡ��ִ��

    Dim lngDatDiff As Long
    If datExec <> CDate(Format("0", "yyyy-MM-dd HH:mm")) Then
        If datNow = CDate(0) Then
            datNow = zlDatabase.Currentdate
        End If
        lngDatDiff = DateDiff("D", datExec, datNow)
        CanUnExec = lngDatDiff <= gintҽ��ִ����Ч����
    Else
        CanUnExec = True
    End If
    
End Function

Public Function GetPatiDiagnoseByDept(ByVal lng����ID As Long, Optional ByVal intType As Integer = 1) As ADODB.Recordset
'���ܣ���ȡָ��������Ժ���˵������������
'������
'      lng����id=����id/����id
'      intType=0-��������ʾ��1-��������ʾ,Ĭ�ϰ�������ʾ
'���أ���¼��
    Dim strSql As String
    
    strSql = " Select A.����ID,A.�������, A.�������" & _
        " From ������ϼ�¼ A,������ҳ B,������Ϣ C,��Ժ���� R" & _
        " Where a.������� In (1, 2, 3, 11, 12, 13) And NVL(A.�������,1) = 1 And A.����ID=B.����ID And A.��ҳID=B.��ҳID And B.����ID=C.����ID And C.��ҳID=B.��ҳID And C.����ID=R.����ID And C.��ǰ����ID=R.����ID " & _
        " And ��ϴ���=1 And" & IIf(intType = 1, " (R.����ID=[1] Or b.Ӥ������ID=[1])", " (r.����id = [1] Or b.Ӥ������id = [1])") & _
        " Order by A.����ID asc,A.��¼��Դ desc,A.������� desc"
    On Error GoTo errH
    Set GetPatiDiagnoseByDept = zlDatabase.OpenSQLRecord(strSql, "mdlCISJob", lng����ID)
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Sub InitObjLis(ByVal lngProgram As Long)
'�ж�����°�LIS����Ϊ�վͳ�ʼ��
    Dim strErr As String
    If gobjLIS Is Nothing Then
        On Error Resume Next
        Set gobjLIS = CreateObject("zl9LisInsideComm.clsLisInsideComm")
        If Not gobjLIS Is Nothing Then
            If gobjLIS.InitComponentsHIS(glngSys, lngProgram, gcnOracle, strErr) = False Then
                If strErr <> "" Then MsgBox "LIS������ʼ������" & vbCrLf & strErr, vbInformation, gstrSysName
                Set gobjLIS = Nothing
            End If
        End If
        err.Clear: On Error GoTo 0
    End If
End Sub

Public Function ISPassShowCard() As Boolean
'���ܣ��Ƿ�������ʾ���￨��
'����:True ������ʾ,False ������
    Dim rsTemp As New ADODB.Recordset
    Dim strSql As String
    Dim blnPassShowCard As Boolean
    
    On Error GoTo errHandle
    strSql = "Select �������� From ҽ�ƿ���� where ����='���￨' and �Ƿ�̶�=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "ҽ�ƿ����")
    If Not rsTemp.EOF Then
        blnPassShowCard = Nvl(rsTemp!��������) <> ""
    End If
    
    ISPassShowCard = blnPassShowCard
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ReadPatPricture(ByVal lng����ID As Long, ByRef imgPatient As Image, Optional ByRef strFile As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������Ƭ
    '������lng����ID=��ȡָ�����˵���Ƭ
    '           imgPatient=��Ƭ����λ��
    '           strFile=��Ƭ�ı���·��
    '74421,������,2014-07-04,��ȡ������Ƭ��Ϣ
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String
    Dim strTmp As String
    Dim rsData As Recordset
    Dim blnReRead As Boolean
    On Error GoTo ErrHand
    imgPatient.Picture = Nothing
    strFile = ""
    strSql = "Select ����id,��Ƭ From ������Ƭ Where ����id=[1]"
    Set rsData = zlDatabase.OpenSQLRecord(strSql, "���ز�����Ƭ", lng����ID)
    If rsData.BOF = False Then
        strTmp = zlDatabase.ReadPicture(rsData, "��Ƭ", strFile)
        imgPatient.Picture = LoadPicture(strFile)
        ReadPatPricture = True
        If strFile <> "" Then Kill strFile
    End If
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function FlexScroll(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'���ܣ�֧�ֹ��ֵĹ���
    Select Case wMsg
    Case WM_MOUSEWHEEL
        Select Case wParam
        Case -7864320  '���¹�
            zlCommFun.PressKey vbKeyPageDown
        Case 7864320   '���Ϲ�
            zlCommFun.PressKey vbKeyPageUp
        End Select
    End Select
    FlexScroll = CallWindowProc(glngPreHWnd, hwnd, wMsg, wParam, lParam)
End Function

Public Function Get����ҽ����ӡ(ByVal lng����ID As Long, ByVal lng��ҳID) As Integer
'���ܣ��ж�ĳ�����˵�ҽ���Ƿ��ӡ��
'���أ�0-δ��ӡ��1-���ִ�ӡ��2-ȫ����ӡ
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    Dim lngҽ��ID As Long
    Dim dat���� As Date
    Dim bytPrint As Byte
    Dim blnDo As Boolean
    Dim arrӤ�� As Variant
    Dim strӤ�� As String
    Dim lngPrintType As Long
    Dim blnKey As Boolean
    Dim lng��� As Long
    Dim i As Long, j As Long
    
    On Error GoTo errH
    
    strSql = "select count(1) as ��ӡ from ����ҽ����ӡ a where a.����id=[1] and a.��ҳid=[2] and a.��ӡʱ�� is not null"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlCISJob", lng����ID, lng��ҳID)
    If Not rsTmp.EOF Then
        If (rsTmp!��ӡ & "") = 0 Then
            Get����ҽ����ӡ = 0
            Exit Function
        End If
    End If
    
    strSql = "select 1 from ����ҽ����ӡ a where a.����id=[1] and a.��ҳid=[2] and a.��ӡʱ�� is not null and Exists" & _
            " (select 1 from ����ҽ����ӡ where ����id=[1] and ��ҳid=[2] and ��ӡʱ�� is null and rownum<2) and rownum<2"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlCISJob", lng����ID, lng��ҳID)
    If Not rsTmp.EOF Then
        Get����ҽ����ӡ = 1
        Exit Function
    End If
   
    Get����ҽ����ӡ = 1
    
    '�ж��ǲ���ȫ���Ѿ���ӡ
    '��ȡ�������ҽ����������
    lngPrintType = Val(zlDatabase.GetPara("ҽ������ӡģʽ", glngSys, pסԺҽ���´�))
    dat���� = CDate("1900-01-01")
    strSql = "Select ҽ������ʱ�� as ʱ�� From ������ҳ Where ����ID=[1] And ��ҳID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlCISJob", lng����ID, lng��ҳID)
    If Not IsNull(rsTmp!ʱ��) Then dat���� = CDate(rsTmp!ʱ�� & "")
    
    strSql = "Select ���,Ӥ������ From ������������¼ Where ����ID=[1] And ��ҳID=[2] Order by ���"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlCISJob", lng����ID, lng��ҳID)
    strӤ�� = "0"
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            strӤ�� = strӤ�� & "," & rsTmp!���
            rsTmp.MoveNext
        Next
    End If
    arrӤ�� = Split(strӤ��, ",")
    
    For i = 0 To 1 '��������
        For j = 0 To UBound(arrӤ��) 'Ӥ��
            'ͣ����ӡ��ֻ��Ҫ�ж�һ��
            If i = 0 Then
                strSql = "Select 1 From ����ҽ����ӡ A, ����ҽ����¼ B" & _
                    " Where A.ҽ��id=B.ID And A.��Ч = 0 And A.����id=[1] And A.��ҳid=[2] And Nvl(A.Ӥ��,0)=[3] And a.��ӡʱ�� is not null And (B.ȷ��ͣ��ʱ�� Is Not Null And" & _
                    " Not Exists (Select 1 From ����ҽ����ӡ S Where S.ҽ��id = A.ҽ��id And S.��ӡ��� = 2) " & _
                    IIf(lngPrintType = 1, " Or B.ִ����ֹʱ�� Is Not Null And Not exists(Select 1 From ����ҽ����ӡ S Where S.ҽ��id = A.ҽ��id And S.��ӡ��� > 0)", "") & ") And Rownum<2"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlCISJob", lng����ID, lng��ҳID, Val(arrӤ��(j)))
                If Not rsTmp.EOF Then
                    blnKey = True
                    Exit For
                End If
            End If
        
            'δ��ӡ��Ҳû���� ����ҽ����ӡ �в���
            lngҽ��ID = 0
            lng��� = 0
            strSql = "Select ҽ��id From (Select ҽ��id From ����ҽ����ӡ Where ����id =[1] And ��ҳid =[2] And Nvl(Ӥ��, 0)=[3] And ��Ч =[4]" & _
            " And ��ӡʱ�� + 0 >= [5] Order By ҳ�� Desc, �к� Desc) Where Rownum<2"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlCISJob", lng����ID, lng��ҳID, Val(arrӤ��(j)), i, dat����)
            If Not rsTmp.EOF Then lngҽ��ID = Val(rsTmp!ҽ��ID & "")
            ' lngҽ��id=0 ֻ������������һ��Ҳû��
            If lngҽ��ID <> 0 Then
                strSql = "Select Nvl(Max(���), 0) as ��� From (Select ��� From ����ҽ����¼ Where (���id =[1] Or ID =[1])" & _
                    " Union All Select b.��� From ����ҽ����¼ A, ����ҽ����¼ B Where a.������� In ('5', '6') And a.Id =[1] And a.���id = b.Id)"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlCISJob", lngҽ��ID)
                If Not rsTmp.EOF Then lng��� = Val(rsTmp!��� & "")
            End If
            
            If dat���� = CDate("1900-01-01") Then
                strSql = "Select 1 From ����ҽ����¼ A, ������ĿĿ¼ B Where a.����id =[1] And a.��ҳid =[2] And Nvl(a.Ӥ��, 0) =[3] And" & vbNewLine & _
                        " a.������Ŀid = b.Id(+) And a.ҽ��״̬ Not In (-1, 2) and a.ҽ����Ч =[4] And" & vbNewLine & _
                        " ([6] = 1 And a.ҽ��״̬ = 1 Or a.ҽ��״̬ <> 1) And Nvl(a.���δ�ӡ, 0) = 0 And" & vbNewLine & _
                        " Not Exists (Select 1 From ����ҽ����¼ Where ������� = 'F' And ID = a.ǰ��id) And a.��� >[5] And a.������Դ = 2 and rownum<2"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlCISJob", lng����ID, lng��ҳID, Val(arrӤ��(j)), i, lng���, lngPrintType)
                If Not rsTmp.EOF Then
                    blnKey = True
                    Exit For
                End If
            Else
                strSql = "Select 1 From ����ҽ����¼ A, ������ĿĿ¼ B Where a.����id =[1] And a.��ҳid =[2] And Nvl(a.Ӥ��, 0) =[3] And" & vbNewLine & _
                        " a.������Ŀid = b.Id(+) And a.ҽ��״̬ Not In (-1, 2) and a.ҽ����Ч =[4] And" & vbNewLine & _
                        " ([6] = 1 And a.ҽ��״̬ = 1 Or a.ҽ��״̬ <> 1) And Nvl(a.���δ�ӡ, 0) = 0 And" & vbNewLine & _
                        " Not Exists (Select 1 From ����ҽ����¼ Where ������� = 'F' And ID = a.ǰ��id) And a.��� >[5] And a.������Դ = 2 and" & _
                        " Exists (Select 1 From ����ҽ��״̬ C Where a.Id = c.ҽ��id And c.����ʱ�� >=[7]) and rownum<2"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlCISJob", lng����ID, lng��ҳID, Val(arrӤ��(j)), i, lng���, lngPrintType, dat����)
                If rsTmp.EOF Then
                    blnKey = True
                    Exit For
                End If
            End If
        Next
        If blnKey Then Exit For
    Next
    
    If Not blnKey Then Get����ҽ����ӡ = 2
 
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
        Call SaveErrLog
End Function

Public Function Get������Һ��(ByVal lng����ID As Long, ByVal lng��ҳID) As String
'���ܣ���ȡָ�����˽�����������Һ������Ӧ���ز���  �ö��ŷָ� "200,300"
'˵����ҩƷ������λΪml����ý����ҩ;��Ϊ��Һ�� �������ѷ���+δ���ͣ���δ�����к��¿��ģ�
'      �����¿��Ļ��δ���͵�����ҽ����������㣺Ƶ��Ϊ��Ҫʱ�ĳ�����Ƶ��Ϊ��Ҫʱ��һ���Ե�����
'      �ݴ棬У�����ʣ������� ҽ��������ͳ�ƣ�����ʹ�ù���ͣ/���ù��ܵĲ��������ǣ�һ�ɵ���û����ͣ��
    Dim rsTmp As ADODB.Recordset
    Dim rsִ��ʱ�� As ADODB.Recordset
    Dim strSql As String, str�ֽ�ʱ�� As String, strҽ��IDs As String
    Dim dblToday As Double, dblTomorrow As Double, dblTmp As Double
    Dim datCur As Date, datBegin As Date, datEnd As Date
    Dim lng���� As Long
    Dim i As Long, j As Long
    Dim varArr As Variant
    
    strSql = "Select a.��������,a.�״�����,a.��ʼִ��ʱ��,a.�ϴ�ִ��ʱ��,Nvl(a.ִ����ֹʱ��,[4]) as ִ����ֹʱ��,a.Ƶ�ʼ��,a.ִ��ʱ�䷽��,a.Ƶ�ʴ���,a.�����λ,a.ִ��Ƶ��," & vbNewLine & _
        "     a.ҽ����Ч,a.ͣ��ʱ��,a.����,nvl(a.�ɷ����,d.סԺ�ɷ����) as ����,a.�ܸ�����,d.����ϵ��,a.���id" & vbNewLine & _
        "From ����ҽ����¼ A,������ĿĿ¼ B,ҩƷ���� C,ҩƷ��� D" & vbNewLine & _
        "Where a.������Ŀid = b.Id And b.Id = c.ҩ��id And a.�շ�ϸĿid=d.ҩƷid(+) And a.������� In ('5','6') and" & vbNewLine & _
        "     Upper(Nvl(b.���㵥λ,'NULL')) = 'ML' And c.��ý=1 And a.����id =[1] And a.��ҳid=[2] And" & vbNewLine & _
        "     a.��ʼִ��ʱ�� <= [4] And a.ҽ��״̬ Not In (-1,2,4) And" & vbNewLine & _
        "     (a.ҽ����Ч = 1 And" & vbNewLine & _
        "     (a.ִ��Ƶ�� = 'һ����' And a.��ʼִ��ʱ�� >= [3] Or" & vbNewLine & _
        "     a.ͣ��ʱ�� Is Null And a.ִ��ʱ�䷽�� Is Not Null Or" & vbNewLine & _
        "     a.ͣ��ʱ�� Is Not Null And a.ִ����ֹʱ�� >= [3] And (a.ִ��ʱ�䷽�� Is Not Null Or a.ִ��Ƶ�� = '��Ҫʱ')) Or" & vbNewLine & _
        "     a.ҽ����Ч = 0 And" & vbNewLine & _
        "     (a.�ϴ�ִ��ʱ�� Is Null And a.ִ��ʱ�䷽�� Is Not Null And Nvl(a.ִ����ֹʱ��,[3])>=[3] Or" & vbNewLine & _
        "     a.�ϴ�ִ��ʱ�� >= [3] ))"
    '��ʱ��Ҫ����˳�����7��ҩƷҽ����
    '1.Ƶ��Ϊһ���Ե��������ѷ��ͺ�δ���ͣ�
    '2.Ƶ��Ϊָ��������������δ���ͣ�
    '3.Ƶ��Ϊָ���������������ѷ��ͣ�
    '4.Ƶ��Ϊ��Ҫʱ���������ѷ��ͣ�
    '5.Ƶ��Ϊָ�������ĳ�������δ���ͣ�
    '6.Ƶ��Ϊ��Ҫʱ���������ٷ���һ�Σ�
    '7.Ƶ��Ϊָ�������ĳ��������ٷ���һ�Σ�
    '����������û�б����͹���ҽ����������Ҫʱ��������Ҫʱ��������ҽ�����û�з����򲻲�������㣬SQL��ѯ��Ҳ���ᱻ���˳���
    
    On Error GoTo errH
    datCur = zlDatabase.Currentdate
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "Get������Һ��", lng����ID, lng��ҳID, CDate(Format(datCur, "YYYY-MM-DD 00:00:00")), CDate(Format(datCur + 1, "YYYY-MM-DD 23:59:59")))
    If rsTmp.EOF Then Get������Һ�� = "0,0": Exit Function
    
    '�� ҽ��ִ��ʱ�� ����ȡִ��ʱ���
    For i = 1 To rsTmp.RecordCount
        '3.Ƶ��Ϊָ���������������ѷ��ͣ�'6.Ƶ��Ϊ��Ҫʱ���������ٷ���һ�Σ�
        If Val(rsTmp!ҽ����Ч & "") = 1 And rsTmp!ִ��ʱ�䷽�� & "" <> "" And rsTmp!ͣ��ʱ�� & "" <> "" Or _
           Val(rsTmp!ҽ����Ч & "") = 0 And rsTmp!ִ��ʱ�䷽�� & "" = "" And rsTmp!�ϴ�ִ��ʱ�� & "" <> "" Then
           
            If InStr("," & strҽ��IDs & ",", "," & Val(rsTmp!���ID & "") & ",") = 0 Then strҽ��IDs = strҽ��IDs & "," & Val(rsTmp!���ID & "")
        End If
        rsTmp.MoveNext
    Next
    strҽ��IDs = Mid(strҽ��IDs, 2)
    If strҽ��IDs <> "" Then
        strSql = "select a.ҽ��id,a.Ҫ��ʱ�� from ҽ��ִ��ʱ�� a where a.ҽ��id in (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist))) and a.Ҫ��ʱ��<=[2]"
        Set rsִ��ʱ�� = zlDatabase.OpenSQLRecord(strSql, "Get������Һ��", strҽ��IDs, CDate(Format(datCur + 1, "YYYY-MM-DD 23:59:59")))
    End If
    rsTmp.MoveFirst
    
    '��ʼ����
    For i = 1 To rsTmp.RecordCount
        '1.Ƶ��Ϊһ���Ե��������ѷ��ͺ�δ���ͣ�����ʼʱ�����һ��Ϊ׼ֻ��һ�Σ�����
        If Val(rsTmp!ҽ����Ч & "") = 1 And rsTmp!ִ��Ƶ�� & "" = "һ����" Then
            If Format(rsTmp!��ʼִ��ʱ�� & "", "YYYY-MM-DD") = Format(datCur, "YYYY-MM-DD") Then
                dblToday = dblToday + Val(rsTmp!�������� & "")
            Else
                dblTomorrow = dblTomorrow + Val(rsTmp!�������� & "")
            End If
        '2.Ƶ��Ϊָ��������������δ���ͣ����ȼ�������ٷֽ�ʱ���
        ElseIf Val(rsTmp!ҽ����Ч & "") = 1 And rsTmp!ִ��ʱ�䷽�� & "" <> "" And rsTmp!ͣ��ʱ�� & "" = "" Then
            If Nvl(rsTmp!����, 0) <> 0 And Not IsNull(rsTmp!ִ��Ƶ��) Then
                'һ��Ƶ�����ڵĴ���
                If rsTmp!�����λ = "��" Then
                    lng���� = IntEx(rsTmp!���� * (rsTmp!Ƶ�ʴ��� / 7))
                ElseIf rsTmp!�����λ = "��" Then
                    lng���� = IntEx(rsTmp!���� * (rsTmp!Ƶ�ʴ��� / rsTmp!Ƶ�ʼ��))
                ElseIf rsTmp!�����λ = "Сʱ" Then
                    lng���� = IntEx(rsTmp!���� * (rsTmp!Ƶ�ʴ��� / rsTmp!Ƶ�ʼ��) * 24)
                ElseIf rsTmp!�����λ = "����" Then
                    lng���� = IntEx(rsTmp!���� * (rsTmp!Ƶ�ʴ��� / rsTmp!Ƶ�ʼ��) * (24 * 60))
                End If
            Else
                '�ɷ���ҩƷʱ,�������Ե����ı��������ҩ;���Ĵ���,���ɷ�����һ����ʹ��ҩƷʱ���������ԣ����������ϵ����ֵȡ�����ı��������ҩ;���Ĵ�����
                '����һ��Ƶ�����ڵĴ�������
                If Nvl(rsTmp!����, 0) = 0 And Nvl(rsTmp!��������, 0) <> 0 Then
                    lng���� = IntEx(rsTmp!�ܸ����� * rsTmp!����ϵ�� / rsTmp!��������)
                ElseIf (Nvl(rsTmp!����, 0) = 1 Or Nvl(rsTmp!����, 0) = 2) And Nvl(rsTmp!��������, 0) <> 0 Then
                    lng���� = IntEx(rsTmp!�ܸ����� / IntEx(rsTmp!�������� / rsTmp!����ϵ��))
                Else
                    lng���� = Nvl(rsTmp!Ƶ�ʴ���, 0)
                End If
            End If
            If Not IsNull(rsTmp!ִ��ʱ�䷽��) Or Nvl(rsTmp!�����λ) = "����" Then
                str�ֽ�ʱ�� = Calc�����ֽ�ʱ��(lng����, rsTmp!��ʼִ��ʱ��, CDate(Format(datCur + 1, "YYYY-MM-DD 23:59:59")), "", Nvl(rsTmp!ִ��ʱ�䷽��), rsTmp!Ƶ�ʴ���, rsTmp!Ƶ�ʼ��, rsTmp!�����λ)
            End If
            If str�ֽ�ʱ�� <> "" Then
                varArr = Split(str�ֽ�ʱ��, ",")
                For j = 0 To UBound(varArr)
                    If Between(Format(varArr(j), "YYYY-MM-DD HH:MM:SS"), Format(datCur, "YYYY-MM-DD HH:MM:SS"), Format(datCur, "YYYY-MM-DD 23:59:59")) Then
                        dblToday = dblToday + Val(rsTmp!�������� & "")
                    ElseIf Between(Format(varArr(j), "YYYY-MM-DD HH:MM:SS"), Format(datCur + 1, "YYYY-MM-DD 00:00:00"), Format(datCur + 1, "YYYY-MM-DD 23:59:59")) Then
                        dblTomorrow = dblTomorrow + Val(rsTmp!�������� & "")
                    End If
                Next
            End If
        '3.Ƶ��Ϊָ���������������ѷ��ͣ����� ҽ��ִ��ʱ�� ����ִ��ʱ��㼴��
        ElseIf Val(rsTmp!ҽ����Ч & "") = 1 And rsTmp!ִ��ʱ�䷽�� & "" <> "" And rsTmp!ͣ��ʱ�� & "" <> "" Then
            If Not rsִ��ʱ�� Is Nothing Then
                rsִ��ʱ��.Filter = "ҽ��id=" & Val(rsTmp!���ID & "")
                For j = 1 To rsִ��ʱ��.RecordCount
                    If Between(Format(rsִ��ʱ��!Ҫ��ʱ�� & "", "YYYY-MM-DD HH:MM:SS"), Format(datCur, "YYYY-MM-DD 00:00:00"), Format(datCur, "YYYY-MM-DD 23:59:59")) Then
                        dblToday = dblToday + Val(rsTmp!�������� & "")
                    ElseIf Between(Format(rsִ��ʱ��!Ҫ��ʱ�� & "", "YYYY-MM-DD HH:MM:SS"), Format(datCur + 1, "YYYY-MM-DD 00:00:00"), Format(datCur + 1, "YYYY-MM-DD 23:59:59")) Then
                        dblTomorrow = dblTomorrow + Val(rsTmp!�������� & "")
                    End If
                    rsִ��ʱ��.MoveNext
                Next
            End If
        '4.Ƶ��Ϊ��Ҫʱ���������ѷ��ͣ�������ҽ��ֻ��һ�Σ��ҵ�����Ч��ֱ���ÿ�ʼʱ���жϼ���
        ElseIf Val(rsTmp!ҽ����Ч & "") = 1 And rsTmp!ִ��Ƶ�� & "" = "��Ҫʱ" And rsTmp!ͣ��ʱ�� & "" <> "" Then
            If Format(rsTmp!��ʼִ��ʱ�� & "", "YYYY-MM-DD") = Format(datCur, "YYYY-MM-DD") Then
                dblToday = dblToday + Val(rsTmp!�������� & "")
            Else
                dblTomorrow = dblTomorrow + Val(rsTmp!�������� & "")
            End If
        '6.Ƶ��Ϊ��Ҫʱ���������ٷ���һ�Σ����� ҽ��ִ��ʱ�� ����ִ��ʱ��㣬����������Ҫ����ȡ����Сʱ��㣬���״�ִ��ʱ���
        ElseIf Val(rsTmp!ҽ����Ч & "") = 0 And rsTmp!ִ��ʱ�䷽�� & "" = "" And rsTmp!�ϴ�ִ��ʱ�� & "" <> "" Then
            If Not rsִ��ʱ�� Is Nothing Then
                rsִ��ʱ��.Filter = "ҽ��id=" & Val(rsTmp!���ID & "")
                rsִ��ʱ��.Sort = "Ҫ��ʱ��"
                For j = 1 To rsִ��ʱ��.RecordCount
                    dblTmp = Val(rsTmp!�������� & "")
                    If j = 1 And Val(rsTmp!�״����� & "") <> 0 Then dblTmp = Val(rsTmp!�״����� & "")
                    If Between(Format(rsִ��ʱ��!Ҫ��ʱ�� & "", "YYYY-MM-DD HH:MM:SS"), Format(datCur, "YYYY-MM-DD 00:00:00"), Format(datCur, "YYYY-MM-DD 23:59:59")) Then
                        dblToday = dblToday + dblTmp
                    ElseIf Between(Format(rsִ��ʱ��!Ҫ��ʱ�� & "", "YYYY-MM-DD HH:MM:SS"), Format(datCur + 1, "YYYY-MM-DD 00:00:00"), Format(datCur + 1, "YYYY-MM-DD 23:59:59")) Then
                        dblTomorrow = dblTomorrow + dblTmp
                    ElseIf Format(rsִ��ʱ��!Ҫ��ʱ�� & "", "YYYY-MM-DD HH:MM:SS") > Format(datCur + 1, "YYYY-MM-DD 23:59:59") Or _
                        Format(rsִ��ʱ��!Ҫ��ʱ�� & "", "YYYY-MM-DD HH:MM:SS") > Format(rsTmp!ִ����ֹʱ�� & "", "YYYY-MM-DD HH:MM:SS") Then
                        Exit For
                    End If
                    rsִ��ʱ��.MoveNext
                Next
            End If
        '7.Ƶ��Ϊָ�������ĳ��������ٷ���һ�Σ�
        '5.Ƶ��Ϊָ�������ĳ�������δ���ͣ�
        '7��5��һ���Ĵ�����ʽ�����¼���ֽ�ʱ���
        ElseIf Val(rsTmp!ҽ����Ч & "") = 0 And rsTmp!ִ��ʱ�䷽�� & "" <> "" And rsTmp!�ϴ�ִ��ʱ�� & "" <> "" Or _
            Val(rsTmp!ҽ����Ч & "") = 0 And rsTmp!ִ��ʱ�䷽�� & "" <> "" And rsTmp!�ϴ�ִ��ʱ�� & "" = "" Then
            '����״�������Ϊ0����ʼʱ����ҽ����ʼִ��ʱ��Ϊ׼Ϊ�˼�����״�ִ��ʱ��������ж�
            If Val(rsTmp!�״����� & "") = 0 And Format(datCur, "YYYY-MM-DD 00:00:00") > Format(rsTmp!��ʼִ��ʱ�� & "", "YYYY-MM-DD HH:MM:SS") Then
                datBegin = Format(datCur, "YYYY-MM-DD 00:00:00")
            Else
                datBegin = rsTmp!��ʼִ��ʱ��
            End If
        
            If Format(rsTmp!ִ����ֹʱ�� & "", "YYYY-MM-DD HH:MM:SS") > Format(datCur + 1, "YYYY-MM-DD 23:59:59") Then
                datEnd = CDate(Format(datCur + 1, "YYYY-MM-DD 23:59:59"))
            Else
                datEnd = CDate(Format(rsTmp!ִ����ֹʱ�� & "", "YYYY-MM-DD HH:MM:SS"))
            End If
            
            str�ֽ�ʱ�� = Calc���ڷֽ�ʱ��(datBegin, datEnd, "", Nvl(rsTmp!ִ��ʱ�䷽��), Nvl(rsTmp!Ƶ�ʴ���, 0), Nvl(rsTmp!Ƶ�ʼ��, 0), Nvl(rsTmp!�����λ), rsTmp!��ʼִ��ʱ��)
            If str�ֽ�ʱ�� <> "" Then
                varArr = Split(str�ֽ�ʱ��, ",")
                For j = 0 To UBound(varArr)
                    dblTmp = Val(rsTmp!�������� & "")
                    If j = 0 And Val(rsTmp!�״����� & "") <> 0 Then dblTmp = Val(rsTmp!�״����� & "")
                    If Between(Format(varArr(j), "YYYY-MM-DD HH:MM:SS"), Format(datCur, "YYYY-MM-DD 00:00:00"), Format(datCur, "YYYY-MM-DD 23:59:59")) Then
                        dblToday = dblToday + dblTmp
                    ElseIf Between(Format(varArr(j), "YYYY-MM-DD HH:MM:SS"), Format(datCur + 1, "YYYY-MM-DD 00:00:00"), Format(datCur + 1, "YYYY-MM-DD 23:59:59")) Then
                        dblTomorrow = dblTomorrow + dblTmp
                    End If
                Next
            End If
        End If
        rsTmp.MoveNext
    Next
    Get������Һ�� = dblToday & "," & dblTomorrow
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function HaveOperateAdvice(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal intType As Integer) As Boolean
'���ܣ��ж�ָ�������Ƿ񻹴��ڿ��Բ�����ҽ��
'������intType 0-У�ԣ�1��ȷ��ֹͣ
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    Dim strWhere As String
    
    On Error GoTo errH
    If intType = 0 Then
        If gblnKSSStrict Or gbln�����ּ����� Or gbln��Ѫ�ּ����� Or gblnѪ��ϵͳ Then
            strWhere = " And Nvl(A.���״̬,0) Not in (1,3,7" & IIf(gblnѪ��ϵͳ = True, "", ",4,5") & ")"
        End If
        strSql = "select 1 from ����ҽ����¼ a where A.ҽ��״̬=1 and a.����id=[1] and a.��ҳid=[2]" & strWhere & _
                " And Exists ( Select M.���� From ��Ա�� M,ִҵ��� N" & _
                " Where M.����=Decode(A.��˱��,1,Substr(A.����ҽ��,1,Instr(A.����ҽ��,'/')-1),Substr(A.����ҽ��,Instr(A.����ҽ��,'/')+1))" & _
                " And M.ִҵ���=N.���� And N.���� IN('ִҵҽʦ','ִҵ����ҽʦ')) And Nvl(A.ִ�б��,0)<>-1 And A.������Դ<>3 and Rownum<2"
    ElseIf intType = 1 Then
        strSql = "select 1 from ����ҽ����¼ a where A.ҽ��״̬=8 and Nvl(a.ҽ����Ч,0)=0 And a.����id=[1] and a.��ҳid=[2] And Nvl(A.ִ�б��,0)<>-1 And A.������Դ<>3 and Rownum<2"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlCISJob.HaveOperateAdvice", lng����ID, lng��ҳID)
    HaveOperateAdvice = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub PlugInInSideBar(ByRef cbsMain As Object, ByVal strFuncName As String, Optional ByVal intInSide As Integer)
'���ܣ����ù�������ť����ҿ�Ƭ�����еĹ��ܰ�ť
'intInSide ��ҪҪ���ӹ�������ť��Ĭ����Ҫ����
    Dim objControl As CommandBarControl
    Dim objMenuBar As CommandBarPopup
    Dim objPopup As CommandBarPopup
    Dim varArr As Variant
    Dim strTmp As String
    Dim lngTmp As Long
    Dim objCbs As Object
    Dim lngidx As Long
    Dim i As Long
    Dim strName As String, lngIcon As Long
    
    If strFuncName = "" Then Exit Sub
    varArr = Split(strFuncName, "|")
    
    Set objCbs = cbsMain
    
    '��չ:���ڹ����˵�(���������û��)���ļ��˵�����
    '-----------------------------------------------------
    Set objMenuBar = objCbs.ActiveMenuBar.Controls.Find(, conMenu_ManagePopup)
    If objMenuBar Is Nothing Then Set objMenuBar = objCbs.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
 
    Set objMenuBar = objCbs.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_Tool_PlugIn, "��չ(&A)", objMenuBar.Index + 1, False)
 
    With objMenuBar.CommandBar.Controls
        For i = 0 To UBound(varArr)
            strTmp = varArr(i)
            
            strName = strTmp
            lngIcon = conMenu_Tool_PlugIn
            
            If InStr(strTmp, ",") > 0 Then
                strName = Split(strTmp, ",")(0)
                lngIcon = Val(Split(strTmp, ",")(1))
            End If
 
            Set objControl = .Add(xtpControlButton, conMenu_Tool_PlugIn_Item + 1 + i, strName)
                objControl.IconId = lngIcon
                objControl.ToolTipText = strName
                objControl.Style = xtpButtonIconAndCaption
        Next
    End With
    
    If intInSide = 0 Then
        '����������
        '�ҵ�Ҫ���ӵ�λ��
        For Each objControl In objCbs(2).Controls '�����ǰ������һ��Control
            If Val(Left(objControl.ID, 1)) <> conMenu_FilePopup And Val(Left(objControl.ID, 1)) <> conMenu_ManagePopup Then
                Set objControl = objCbs(2).Controls(objControl.Index - 1)
                lngidx = objControl.Index
                Exit For
            End If
        Next
        
        With objCbs(2).Controls
            For i = UBound(varArr) To 0 Step -1
                strTmp = varArr(i)
                
                strName = strTmp
                lngIcon = conMenu_Tool_PlugIn
                
                If InStr(strTmp, ",") > 0 Then
                    strName = Split(strTmp, ",")(0)
                    lngIcon = Val(Split(strTmp, ",")(1))
                End If
                
                Set objControl = .Add(xtpControlButton, conMenu_Tool_PlugIn_Item + 1 + i, strName, lngidx + 1)
                    objControl.IconId = lngIcon
                    objControl.ToolTipText = strName
                    objControl.Style = xtpButtonIconAndCaption
            Next
        End With
    End If
    cbsMain.RecalcLayout
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Function Get��Ⱦ��״̬(ByVal lng��¼ As Long, ByVal lng��д As Long, ByVal lng״̬ As Long) As String
'���ܣ���ȡ��ǰ���˵Ĵ�Ⱦ��״̬
    Dim strTmp As String
    If lng״̬ <> 0 Then
'       -1-�Ѿܾ� 1-�ѽ���;2-�ѳʱ�;3-���ͨ��;4-��ҽ�����ޣ�5-ҽ���ѷ�����ɴ����
        Select Case lng״̬
        Case -1
            strTmp = "�Ѿܾ�"
        Case 1
            strTmp = "�ѽ���"
        Case 2
            strTmp = "�ѳʱ�"
        Case 3
            strTmp = "���ͨ��"
        Case 4
            strTmp = "��ҽ���޸�"
        Case 5
            strTmp = "ҽ������Ĵ����"
        End Select
    ElseIf lng��д > 0 Then
        strTmp = "����д"
    ElseIf lng��¼ > 0 Then
        strTmp = "�����Խ��"
    End If
    Get��Ⱦ��״̬ = strTmp
End Function

Public Function FuncTraReactionRecord(ByVal frmParent As Object, ByVal lng���� As Long, ByVal lngMoudle As Long) As Boolean
'���ܣ���Ѫ��Ӧ���ýӿ�
    On Error GoTo errH
    If InitObjBlood = False Then
        MsgBox "Ѫ�ⲿ������ʧ�ܡ�", vbInformation, gstrSysName
        Exit Function
    End If
    Call gobjPublicBlood.zlShowBloodReactionRecord(frmParent, glngSys, lngMoudle, lng����)
    FuncTraReactionRecord = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
