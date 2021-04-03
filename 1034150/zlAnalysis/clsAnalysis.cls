VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAnalysis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Version As String

Private Sub Class_Initialize()
    Version = App.Major & "." & App.Minor & "." & App.Revision
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
    CloseWindows = (Forms.Count = 0)
End Function

Public Sub CodeMan(ByVal lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, frmMain As Object, ByVal strDbUser As String)
    '------------------------------------------------
    '���ܣ� �����ĳ�ʼ����������ִ�г���,��������ı�Ҫ���ݴ�������
    '������
    '   cnMain:����������ݿ�����
    '���أ�
    '------------------------------------------------
    Set gcnOracle = cnMain
    gstrAviPath = GetSetting(appName:="ZLSOFT", Section:="ע����Ϣ", Key:=UCase("gstrAviPath"), Default:="")
    gstrSysName = GetSetting(appName:="ZLSOFT", Section:="ע����Ϣ", Key:=UCase("gstrSysName"), Default:="")
    gstrVersion = GetSetting(appName:="ZLSOFT", Section:="ע����Ϣ", Key:=UCase("gstrVersion"), Default:="")
    gstr��λ���� = GetUnitName()
    
    gstrPrivs = GetPrivFunc(lngSys, lngModul)
    gstr��λ���� = GetUnitName()
    gstrDbUser = strDbUser
    Call GetUserInfo
    If Not CheckValid Then Exit Sub

    '1401    ������Դ״����
    '1402    ����Ӧ�ﰲ���ܱ�
    '1403    �շ���Ŀ���Ŀ
    '
    '1410    ȫԺ�ֽ�����ſ�
    '1411    ȫԺ���˷��û���
    '1412    �ֿƲ��˷������
    '1413    ȫԺδ����û���
    '1414    �ֿ�δ����û���
    '1415    ȫԺԤ�տ���ܱ�
    '1416    ҽԺ�������
    '1417    ���˷ѱ���ܱ�
    '1418    ��Լ���˷��û���
    '
    '1420    �����˴λ��ܷ���
    '1421    ��������ͳ�Ʒ���
    '1422    סԺ������������
    '1423    ��Ժ���˷ֲ�
    '1424    ��ԺΣ�ز���
    '1425    ��Ժ�ػ�����
    '
    '1430    ҽԺ������ͳ�Ʊ�
    '1431    �����������ͳ��
    '1432    ������η��÷���
    '1433    סԺ���η��÷���
    '1434    �����շ���Ŀ����
    '
    '1440    ҩƷֱ����֧����
    '1441    ҩƷ������ȱ����
    '1442    ҩƷЧ�ڱ�������
    '1443    ҩƷ���ñ�������
    '1444    �ٴ���ҩ����
    '1445    ������ٷ���
    Select Case lngModul
        Case 1401
        Case 1402
        Case 1403
            frmPriceQuery.Show , frmMain
    End Select

End Sub

Public Sub BHCodeMan(ByVal lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, lngMain As Long, ByVal strDbUser As String)
    '------------------------------------------------
    '���ܣ� �����ĳ�ʼ����������ִ�г���,��������ı�Ҫ���ݴ�������
    '������
    '   cnMain:����������ݿ�����
    '���أ�
    '------------------------------------------------
    Set gcnOracle = cnMain
    gstrAviPath = GetSetting(appName:="ZLSOFT", Section:="ע����Ϣ", Key:=UCase("gstrAviPath"), Default:="")
    gstrSysName = GetSetting(appName:="ZLSOFT", Section:="ע����Ϣ", Key:=UCase("gstrSysName"), Default:="")
    gstrVersion = GetSetting(appName:="ZLSOFT", Section:="ע����Ϣ", Key:=UCase("gstrVersion"), Default:="")
    gstr��λ���� = GetUnitName()
    
    gstrPrivs = GetPrivFunc(lngSys, lngModul)
    gstr��λ���� = GetUnitName()
    gstrDbUser = strDbUser
    Call GetUserInfo
    If Not CheckValid Then Exit Sub

    '1401    ������Դ״����
    '1402    ����Ӧ�ﰲ���ܱ�
    '1403    �շ���Ŀ���Ŀ
    '
    '1410    ȫԺ�ֽ�����ſ�
    '1411    ȫԺ���˷��û���
    '1412    �ֿƲ��˷������
    '1413    ȫԺδ����û���
    '1414    �ֿ�δ����û���
    '1415    ȫԺԤ�տ���ܱ�
    '1416    ҽԺ�������
    '1417    ���˷ѱ���ܱ�
    '1418    ��Լ���˷��û���
    '
    '1420    �����˴λ��ܷ���
    '1421    ��������ͳ�Ʒ���
    '1422    סԺ������������
    '1423    ��Ժ���˷ֲ�
    '1424    ��ԺΣ�ز���
    '1425    ��Ժ�ػ�����
    '
    '1430    ҽԺ������ͳ�Ʊ�
    '1431    �����������ͳ��
    '1432    ������η��÷���
    '1433    סԺ���η��÷���
    '1434    �����շ���Ŀ����
    '
    '1440    ҩƷֱ����֧����
    '1441    ҩƷ������ȱ����
    '1442    ҩƷЧ�ڱ�������
    '1443    ҩƷ���ñ�������
    '1444    �ٴ���ҩ����
    '1445    ������ٷ���
    Select Case lngModul
        Case 1401
        Case 1402
        Case 1403
            Call zlCommFun.ShowChildWindow(frmPriceQuery.hWnd, lngMain)
    End Select

End Sub