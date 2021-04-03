Attribute VB_Name = "mdlPublic"
Option Explicit

Public gcnOracle As ADODB.Connection
Public gobjComlib As Object
Public gobjCommFun As Object
Public gobjControl As Object
Public gobjDatabase As Object
Public gclsInsure As Object
Public gobjLIS As Object
Public gobjKernel As zlCISKernel.clsCISKernel
Public gobjCISJob As Object
Public glngSys As Long
Public glngModule As Long
Public gMainPrivs As String
Public gstrDBUser As String
Public gstrNodeNo As String          '��ǰվ���ţ����δ��������վ�㣬��Ϊ"-"
Public gcolPrivs As Collection              '��¼�ڲ�ģ���Ȩ��
Public lngNumPublicAdvice As Long  '����������¼clsPublicAdvice�౻�����Ĵ���


Public Type TYPE_USER_INFO
    ID As Long
    ��� As String
    ���� As String
    ���� As String
    �û��� As String
    ���� As String
    ����ID As Long
    ������ As String
    ������ As String
    רҵ����ְ�� As String
    ��ҩ���� As Long
End Type

Public Enum Msg_Type '��Ϣ�������
    m�¿� = 1
    m��ͣ = 2
    m�·� = 3
    m���� = 4
    mΣ��ֵ = 5
    m��Һ�ܾ� = 6
    m�������� = 7
    mRISԤԼ = 8
    mRISԤԼ׼�� = 9
    mȡѪ֪ͨ = 10    
End Enum

Public Enum ҽԺҵ��
    support����Ԥ�� = 0
    
    support�����˷� = 1
    supportԤ���˸����ʻ� = 2
    support�����˸����ʻ� = 3
    
    support�շ��ʻ�ȫ�Է� = 4       '�����շѺ͹Һ��Ƿ��ø����ʻ�֧��ȫ�ԷѲ��֡�ȫ�Էѣ�ָͳ�����Ϊ0�Ľ��򳬳��޼۵Ĵ�λ�Ѳ���
    support�շ��ʻ������Ը� = 5     '�����շѺ͹Һ��Ƿ��ø����ʻ�֧�������Ը����֡������Ը�����1-ͳ�������* ���
    
    support�����ʻ�ȫ�Է� = 6       'סԺ���������������Ƿ��ø����ʻ�֧��ȫ�ԷѲ��֡�
    support�����ʻ������Ը� = 7     'סԺ���������������Ƿ��ø����ʻ�֧�������Ը����֡�
    support�����ʻ����� = 8         'סԺ���������������Ƿ��ø����ʻ�֧�����޲��֡�
    
    support����ʹ�ø����ʻ� = 9     '����ʱ��ʹ�ø����ʻ�֧��
    supportδ�����Ժ = 10          '�������˻���δ�����ʱ��Ժ
    
    support���ﲿ�����ֽ� = 11      'ֻ��������ҽ����֧���˷Ѳ�ʹ�ñ�������Ҳ����˵�����ֽ�ʱ�ſ��ǲ�������񣬶��˻ص������ʻ���ҽ�������������˷ѡ�
    support����������ҽ����Ŀ = 12  '�ڽ���ʱ�����Ը��շ�ϸĿ�Ƿ�����ҽ����Ŀ���м��
    
    support������봫����ϸ = 13    '�����շѺ͹Һ��Ƿ���봫����ϸ
    
    support�����ϴ� = 14            'סԺ���ʷ�����ϸʵʱ����
    support���������ϴ� = 15        'סԺ�����˷�ʵʱ����

    support��Ժ���˽������� = 16    '������Ժ���˽�������
    support������Ժ = 17            '�����������˳�Ժ
    support����¼�������� = 18    '������Ժ���Ժʱ������¼�������
    support������ɺ��ϴ� = 19      'Ҫ���ϴ��ڼ��������ύ���ٽ���
    support��Ժ��������Ժ = 20    '���˽���ʱ���ѡ���Ժ���ʣ��ͼ������Ժ�ſ��Խ���
    
    support�Һ�ʹ�ø����ʻ� = 21    'ʹ��ҽ���Һ�ʱ�Ƿ�ʹ�ø����ʻ�����֧��

    support���������շ� = 22        '������������֤�󣬿ɽ��ж���շѲ���
    support�����շ���ɺ���֤ = 23  '�������շ���ɣ��Ƿ��ٴε���������֤
    
    supportҽ���ϴ� = 24            'ҽ����������ʱ�Ƿ�ʵʱ����
    support�ֱҴ��� = 25            'ҽ�������Ƿ����ֱ�
    support��;������������ϴ����� = 26 '�ṩ�����ϴ��������ݵĽ��㹦��
    support���������ѽ��ʵļ��ʵ��� = 27 '�Ƿ������������ʵ��ݣ�����õ����Ѿ�����
    
    support�������ݳ������� = 28
    support��Ժ��ʵ�ʽ��� = 29 '��Ժ�ӿ����Ƿ�Ҫ��ӿ��̽��н���
    support�������� = 35            '�Ƿ������������ʣ�����Ա����Ҫӵ�и������ʵ�Ȩ�ޡ��˲���ȱʡΪ�棬��֧�ֵĽӿ��赥������
    
    supportҽ��ȷ���������� = 48
    supportסԺ���˲�����׼��Ŀ���� = 50            'ͬһ�ֲ�,��סԺʱ����¼�����е���Ŀ
    support���ﲡ�˲�����׼��Ŀ���� = 51            '����������ĳ������¿���¼��������Ŀ
    supportʵʱ��� = 60
    
    support�ϴ����ﵵ�� = 70                    '������ҽ������ʱ���Ƿ����TranElecDossier����������ﲡ�˵��Ӿ���/���ӵ������ϴ�
End Enum

Public UserInfo As TYPE_USER_INFO

Public gobjExpense As Object  '���ù�������

'�������
Public gblnҩƷ�������ҽ�� As Boolean
Public gstr��Һ�������� As String
Public gblnִ��ǰ�Ƚ��� As Boolean 'һ��ִͨ��ǰ���շѻ�������
Public gblnStock As Boolean 'ָ��ҩ��ʱ�Ƿ��޶�����ҩƷ�Ŀ��
Public gstrLike As String  '��Ŀƥ�䷽��,%���
Public gbytCode As Byte '�������뷽ʽ
Public gbytDecPrice As Byte '���õ��۵�С����λ��
Public gstrDecPrice As String '�۸�С��λ������ĸ�ʽ����,��"0.0000"
Public gbln�������������� As Boolean '�Ƿ��ڼ���ҽ������ʱ����������
Public gbytMediOutMode As Byte '����ҩƷ���ⷽʽ��0-�������Ƚ��ȳ���1-��Ч������ȳ�,Ч����ͬ�����ٰ������Ƚ��ȳ�
Public gbytDec As Byte '���ý���С����λ��
Public gstrDec As String '�����С��λ������ĸ�ʽ����,��"0.0000"
Public gstr��̬�ѱ� As String               '������ﵱǰ���ҿ��ö�̬�ѱ�,�ڹ���������ʹ��,ʹ��ʱ�Ÿ�ֵ:CalcDrugPrice,CalcPrice
Public gbln��������ۿ� As Boolean '������Ŀ���ܼ����ۿ�
Public gstrסԺ���ͻ��۵� As String
Public gstr���﷢�ͻ��۵� As String
Public gbln������֤ As Boolean '����һ��ͨ���Ѽ���ʣ����ʱ�Ƿ���Ҫ��֤
Public gblnסԺ�Զ����� As Boolean 'סԺ������ɺ��Ƿ��Զ�����
Public gbln�����Զ����� As Boolean '���������ɺ��Ƿ��Զ�����
Public gblnѪ��ϵͳ As Boolean  '�Ƿ�װѪ��ϵͳ
Public gbln�����������۷��� As Boolean '���ʱ����������۷���

Public Function ZVal(ByVal varValue As Variant, Optional ByVal blnForceNum As Boolean) As String
'���ܣ���0��ת��Ϊ"NULL"��,������SQL���ʱ��
'������blnForceNum=��ΪNullʱ���Ƿ�ǿ�Ʊ�ʾΪ������
    ZVal = IIF(Val(varValue) = 0, IIF(blnForceNum, "-NULL", "NULL"), Val(varValue))
End Function

Public Function Decode(ParamArray arrPar() As Variant) As Variant
'���ܣ�ģ��Oracle��Decode����
    Dim varValue As Variant, i As Integer
    
    i = 1
    varValue = arrPar(0)
    Do While i <= UBound(arrPar)
        If i = UBound(arrPar) Then
            Decode = arrPar(i): Exit Function
        ElseIf varValue = arrPar(i) Then
            Decode = arrPar(i + 1): Exit Function
        Else
            i = i + 2
        End If
    Loop
End Function

Public Function FormatEx(ByVal vNumber As Variant, ByVal intBit As Integer) As String
'���ܣ��������뷽ʽ��ʽ����ʾ����,��֤С������󲻳���0,С����ǰҪ��0
'������vNumber=Single,Double,Currency���͵�����,intBit=���С��λ��
    Dim strNumber As String
            
    If TypeName(vNumber) = "String" Then
        If vNumber = "" Then Exit Function
        If Not IsNumeric(vNumber) Then Exit Function
        vNumber = Val(vNumber)
    End If
            
    If vNumber = 0 Then
        strNumber = 0
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
    FormatEx = strNumber
End Function

Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'���ܣ��൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
    Nvl = IIF(IsNull(varValue), DefaultValue, varValue)
End Function


Public Function InitSysPar() As Boolean
'���ܣ���ʼ��ϵͳ����
'���أ���-�����ɹ�
    Dim strTmp As String
    gstrLike = IIF(gobjComlib.zlDatabase.GetPara("����ƥ��") = "0", "%", "")
    gbytCode = Val(gobjComlib.zlDatabase.GetPara("���뷽ʽ"))
 
    'ָ��ҩ��ʱ���ƿ��
    gblnStock = Val(gobjComlib.zlDatabase.GetPara(18, glngSys)) <> 0
        
    'ҩƷ�������ҽ��
    gblnҩƷ�������ҽ�� = Val(gobjComlib.zlDatabase.GetPara(69, glngSys)) = 1
    
    '��Һ��������(����Ϊ���������ġ���ҩ��)
    gstr��Һ�������� = Get��Һ��������

    '����һ��ͨ,��Ŀִ��ǰ�������շѻ��ȼ������
    gblnִ��ǰ�Ƚ��� = Val(gobjComlib.zlDatabase.GetPara(163, glngSys)) <> 0
    
    gbytDec = Val(gobjComlib.zlDatabase.GetPara(9, glngSys, , 2))
    gstrDec = "0." & String(gbytDec, "0")
    gbytDecPrice = Val(gobjComlib.zlDatabase.GetPara(157, glngSys, , 5))
    gstrDecPrice = "0." & String(gbytDecPrice, "0")
    '����ҽ������ʱ����������
    gbln�������������� = Val(gobjComlib.zlDatabase.GetPara(143, glngSys)) <> 0
    
    '����ҩƷ���ⷽʽ
    gbytMediOutMode = Val(gobjComlib.zlDatabase.GetPara(150, glngSys))
    '������Ŀ���ܼ����ۿ�
    gbln��������ۿ� = Val(gobjComlib.zlDatabase.GetPara(93, glngSys)) <> 0
    
    'ҽ���������ɻ��۵������
    gstrסԺ���ͻ��۵� = gobjComlib.zlDatabase.GetPara(80, glngSys)
    gstr���﷢�ͻ��۵� = gobjComlib.zlDatabase.GetPara(86, glngSys)
    'һ��ͨ������֤
    gbln������֤ = Val(gobjComlib.zlDatabase.GetPara(28, glngSys)) <> 0
    'סԺ�Զ�����
    gblnסԺ�Զ����� = Val(gobjComlib.zlDatabase.GetPara(63, glngSys)) <> 0
    '�����Զ�����
    gbln�����Զ����� = Val(gobjComlib.zlDatabase.GetPara(92, glngSys)) <> 0
    '�Ƿ�װѪ��ϵͳ
    gblnѪ��ϵͳ = IsSetUp(2200)
    '���ʱ����������۷���
    gbln�����������۷��� = Val(gobjComlib.zlDatabase.GetPara(98, glngSys)) <> 0
    
    InitSysPar = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function GetUserInfo() As Boolean
'���ܣ���ȡ��½�û���Ϣ
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    Set rsTmp = gobjComlib.zlDatabase.GetUserInfo
    If Not rsTmp Is Nothing Then
        If Not rsTmp.EOF Then
            UserInfo.ID = rsTmp!ID
            UserInfo.�û��� = rsTmp!User
            UserInfo.��� = rsTmp!���
            UserInfo.���� = Nvl(rsTmp!����)
            UserInfo.���� = Nvl(rsTmp!����)
            UserInfo.����ID = Nvl(rsTmp!����ID, 0)
            UserInfo.������ = Nvl(rsTmp!������)
            UserInfo.������ = Nvl(rsTmp!������)
            UserInfo.���� = Get��Ա����
            UserInfo.רҵ����ְ�� = Getרҵ����ְ��(UserInfo.ID)
            GetUserInfo = True
        End If
    End If
    
    gstrDBUser = UserInfo.�û���
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function zlGetComLib() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ����������ض���
    '����:��ȡ�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-05-15 15:34:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Not gobjComlib Is Nothing Then zlGetComLib = True: Exit Function
    
    Err = 0: On Error Resume Next
    Set gobjComlib = GetObject("", "zl9Comlib.clsComlib")
    Set gobjCommFun = GetObject("", "zl9Comlib.clsCommfun")
    Set gobjControl = GetObject("", "zl9Comlib.clsControl")
    Set gobjDatabase = GetObject("", "zl9Comlib.clsDatabase")
    gstrNodeNo = ""
    If Not gobjComlib Is Nothing Then gstrNodeNo = gobjComlib.gstrNodeNo
    Err = 0: On Error GoTo 0
    If Not gobjComlib Is Nothing Then zlGetComLib = True: Exit Function
    Err = 0: On Error Resume Next
    Set gobjComlib = CreateObject("zl9Comlib.clsComlib")
    Call gobjComlib.InitCommon(gcnOracle)
    Set gobjCommFun = gobjComlib.zlCommFun
    Set gobjControl = gobjComlib.zlControl
    Set gobjDatabase = gobjComlib.zlDatabase
    If Not gobjComlib Is Nothing Then gstrNodeNo = gobjComlib.gstrNodeNo
    Err = 0: On Error GoTo 0
End Function

Public Function Between(X, a, B) As Boolean
'���ܣ��ж�x�Ƿ���a��b֮��
    If a < B Then
        Between = X >= a And X <= B
    Else
        Between = X >= B And X <= a
    End If
End Function

Public Function IntEx(vNumber As Variant) As Variant
'���ܣ�ȡ����ָ����ֵ����С����
    IntEx = -1 * Int(-1 * Val(vNumber))
End Function

Public Function GetItemField(ByVal strTable As String, ByVal lngID As Long, Optional ByVal strField As String) As Variant
'���ܣ���ȡָ����ָ���ֶ���Ϣ
'˵����δ����NULLֵ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    If strField = "" Then
        strSQL = "Select * From " & strTable & " Where ID=[1]"
    Else
        strSQL = "Select " & strField & " From " & strTable & " Where ID=[1]"
    End If
    Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lngID)
    If Not rsTmp.EOF Then
        If strField = "" Then
            Set GetItemField = rsTmp
        Else
            GetItemField = rsTmp.Fields(strField).Value
        End If
    End If
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Private Function IsSetUp(ByVal lngϵͳ��� As Long) As Boolean
'���ܣ��Ƿ�װĳ��ϵͳ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select 1 From zlSystems Where ���=[1]"
    Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lngϵͳ���)
    IsSetUp = Not rsTmp.EOF
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function GetInsidePrivs(ByVal lngProg As Enum_Inside_Program, Optional ByVal blnLoad As Boolean) As String
'���ܣ���ȡָ���ڲ�ģ���������е�Ȩ��
'������blnLoad=�Ƿ�̶����¶�ȡȨ��(���ڹ���ģ���ʼ��ʱ,�����û�ͨ��ע���ķ�ʽ�л���)
    Dim strPrivs As String
    
    If gcolPrivs Is Nothing Then
        Set gcolPrivs = New Collection
    End If
    
    On Error Resume Next
    strPrivs = gcolPrivs("_" & lngProg)
    If Err.Number = 0 Then
        If blnLoad Then
            gcolPrivs.Remove "_" & lngProg
        End If
    Else
        Err.Clear: On Error GoTo 0
        blnLoad = True
    End If
    
    If blnLoad Then
        strPrivs = gobjComlib.GetPrivFunc(glngSys, lngProg)
        gcolPrivs.Add strPrivs, "_" & lngProg
    End If
    GetInsidePrivs = IIF(strPrivs <> "", ";" & strPrivs & ";", "")
End Function

Public Sub InitObjLis(ByVal lngProgram As Long)
'�ж�����°�LIS����Ϊ�վͳ�ʼ��
    Dim strErr As String
    If gobjLIS Is Nothing Then
        On Error Resume Next
        Set gobjLIS = CreateObject("zl9LisInsideComm.clsLisInsideComm")
        If Not gobjLIS Is Nothing Then
            If gobjLIS.InitComponentsHIS(glngSys, lngProgram, gcnOracle, strErr) = False Then
                If strErr <> "" Then MsgBox "LIS������ʼ������" & vbCrLf & strErr, vbInformation, "InitObjLis"
                Set gobjLIS = Nothing
            End If
        End If
        Err.Clear: On Error GoTo 0
    End If
End Sub