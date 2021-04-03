VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLisWork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Version As String

'��ҩ�䷽��������Ҫ�Ĳ�����Ϣ
Public Type TYPE_PatiInfoEx
        lng����ID As Long
        lng��ҳID As Long '��ҳID
        str�Һŵ� As String '�Һŵ���
        str�Ա� As String  '�����Ա�
        intӤ�� As Integer
        int���� As Integer 'ҽ�����˵�����
        blnҽ�� As Boolean '�Ƿ�ҽ������
        lng���˿���ID As Long '����ȷ����ҩ�䷽��ȱʡҩ��
End Type

Public Sub CodeMan(lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, frmMain As Object, ByVal strDBUser As String)
'���ܣ� ����������ָ�����ܣ�����ִ����س���
'������
'   lngModul:��Ҫִ�еĹ������
'   cnMain:����������ݿ�����
'   frmMain:������
'   strDBUser:��ǰ��¼���ݿ��û���

    '------------------------------------------------
    Dim strTmp As String
    Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
    gstrProductName = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒ����", "����")
    
    'ȡ���￨�ų���
'    strTmp = zlDatabase.GetPara(20, lngSys, , "7|7|7|7|7")
'    gbytCardNOLen = Val(Split(strTmp, "|")(4))
    
    gstrDBUser = strDBUser
        
    If Not GetUserInfo Then
        MsgBox "��ǰ�û�δ���ö�Ӧ����Ա��Ϣ,����ϵͳ����Ա��ϵ,�ȵ��û���Ȩ���������ã�", vbInformation, gstrSysName
        Exit Sub
    End If
    
   
    Set gfrmMain = frmMain
    
    gstrPrivs = GetPrivFunc(lngSys, lngModul)
    glngModul = lngModul
    glngSys = lngSys
    
    gstrUnitName = GetUnitName
    gstrMatch = IIf(GetSetting("ZLSOFT", "����ģ��\����", "����ƥ��", 0) = "0", "%", "")
    gblnInsure = (UCase(GetSetting("ZLSOFT", "����ȫ��", "�Ƿ�֧��ҽ��", "")) = UCase("Yes"))
    gintInsure = Val(GetSetting("ZLSOFT", "����ȫ��", "ҽ�����", 0))
    'ִ�к��Զ����
    gblnִ�к���� = Val(zlDatabase.GetPara(81, glngSys)) <> 0
    '-------------------------------------------------
    
     '��ʼ�����Ĳ���
    Call gobjKernel.InitCISKernel(gcnOracle, frmMain, glngSys, gstrPrivs)
    Call gobjRichEPR.InitRichEPR(gcnOracle, frmMain, glngSys, False)

    On Error Resume Next
    '��ʼ���Ŵ�������
    gblnManualPH = True
    gintNumberPH = 9999
    
    Select Case lngModul
    Case 1208   '���鼼ʦ����վ
        If gblnInsure Then
            If Not gclsInsure.InitInsure(gcnOracle) Then gblnInsure = False
        Else
            Call gclsInsure.InitOracle(gcnOracle)
        End If
        frmLabMain.Show , frmMain
    Case 1209   '�ʿ����ݲ�ѯ
        frmQCHistory.Show , frmMain
'        FrmQualityDataQuery.Show , frmMain
    Case 1210   '�����ʿع���
        frmQCTodayList.Show , frmMain
'        frmPatientHistoryQuery.Show , frmMain
    Case 1211   '����ɼ�����վ
        frmLabSampling.Show , frmMain
    Case 1212   '����걾�Ǽ�
        frmLabSampleRegister.Show , frmMain
    Case 1213   '�����ȶ�
        frmQCContrast.Show , frmMain
    End Select
End Sub

Public Sub BHCodeMan(lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, lngMain As Long, ByVal strDBUser As String, Optional ByVal objLogin As Object)
'���ܣ� ����������ָ�����ܣ�����ִ����س���
'������
'   lngModul:��Ҫִ�еĹ������
'   cnMain:����������ݿ�����
'   lngMain:��������
'   strDBUser:��ǰ��¼���ݿ��û���

    '------------------------------------------------
    Dim strTmp As String
    Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
    gstrProductName = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒ����", "����")
    
    'ȡ���￨�ų���
'    strTmp = zlDatabase.GetPara(20, lngSys, , "7|7|7|7|7")
'    gbytCardNOLen = Val(Split(strTmp, "|")(4))
    
    gstrDBUser = strDBUser
        
    If Not GetUserInfo Then
        MsgBox "��ǰ�û�δ���ö�Ӧ����Ա��Ϣ,����ϵͳ����Ա��ϵ,�ȵ��û���Ȩ���������ã�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    Set gfrmMain = Nothing
    
    gstrPrivs = GetPrivFunc(lngSys, lngModul)
    glngModul = lngModul
    glngSys = lngSys
    
    gstrUnitName = GetUnitName
    gstrMatch = IIf(GetSetting("ZLSOFT", "����ģ��\����", "����ƥ��", 0) = "0", "%", "")
    gblnInsure = (UCase(GetSetting("ZLSOFT", "����ȫ��", "�Ƿ�֧��ҽ��", "")) = UCase("Yes"))
    gintInsure = Val(GetSetting("ZLSOFT", "����ȫ��", "ҽ�����", 0))
    'ִ�к��Զ����
    gblnִ�к���� = Val(zlDatabase.GetPara(81, glngSys)) <> 0
    '-------------------------------------------------
    
     '��ʼ�����Ĳ���
    Call gobjKernel.InitCISKernel(gcnOracle, Nothing, glngSys, gstrPrivs)
    Call gobjRichEPR.InitRichEPR(gcnOracle, objLogin, glngSys, False)
    Set gobjEmr = objLogin.mobjEMR
    On Error Resume Next
    '��ʼ���Ŵ�������
    gblnManualPH = True
    gintNumberPH = 9999
    
    Select Case lngModul
    Case 1208   '���鼼ʦ����վ
        If gblnInsure Then
            If Not gclsInsure.InitInsure(gcnOracle) Then gblnInsure = False
        Else
            Call gclsInsure.InitOracle(gcnOracle)
        End If
        zlCommFun.ShowChildWindow frmLabMain.hWnd, lngMain, True
    Case 1209   '�ʿ����ݲ�ѯ
        zlCommFun.ShowChildWindow frmQCHistory.hWnd, lngMain
    Case 1210   '�����ʿع���
        zlCommFun.ShowChildWindow frmQCTodayList.hWnd, lngMain
    Case 1211   '����ɼ�����վ
        zlCommFun.ShowChildWindow frmLabSampling.hWnd, lngMain, True
    Case 1212   '����걾�Ǽ�
        zlCommFun.ShowChildWindow frmLabSampleRegister.hWnd, lngMain
    Case 1213   '�����ȶ�
        zlCommFun.ShowChildWindow frmQCContrast.hWnd, lngMain
    End Select
End Sub

Public Function CloseWindows() As Boolean
    '--------------------------------------
    '����:�ر������Ӵ���
    '--------------------------------------
    Dim frmThis As Form
    
    On Error Resume Next
    
    
    
    For Each frmThis In Forms
        Unload frmThis
    Next
    
    Call gclsInsure.CloseWindows
    Call gobjRichEPR.CloseWindows
    Call gobjKernel.CloseWindows
    CloseWindows = (Forms.Count = 0)
End Function

Private Sub Class_Initialize()
    Version = App.Major & "." & App.Minor & "." & App.Revision
End Sub
