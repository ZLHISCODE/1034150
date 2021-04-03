VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCashBill"
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
    Dim strTitle As String, frmList As Form, frmTemp As Form
    Set gcnOracle = cnMain
    If Not CheckValid Then Exit Sub
    
    gstrAviPath = GetSetting(appName:="ZLSOFT", Section:="ע����Ϣ", Key:=UCase("gstrAviPath"), Default:="")
    gstrSysName = GetSetting(appName:="ZLSOFT", Section:="ע����Ϣ", Key:=UCase("gstrSysName"), Default:="")
    gstrVersion = GetSetting(appName:="ZLSOFT", Section:="ע����Ϣ", Key:=UCase("gstrVersion"), Default:="")
    gstrProductName = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒ����", "����")
    gstrLike = IIf(Val(zlDatabase.GetPara("����ƥ��")) = 0, "%", "")
    gstr��λ���� = GetUnitName()
    
    glngSys = lngSys
    glngModul = lngModul
    gstrPrivs = ";" & GetPrivFunc(lngSys, lngModul) & ";"
    gstr��λ���� = GetUnitName()
    gstrDbUser = strDbUser
    
    If Not GetUserInfo Then
        MsgBox "��ǰ�û�δ���ö�Ӧ����Ա��Ϣ,����ϵͳ����Ա��ϵ,�ȵ��û���Ȩ���������ã�", vbInformation, gstrSysName
        Exit Sub
    End If
    gstrMatchMethod = Val(zlDatabase.GetPara("����ƥ��"))
    '-------------------------------------------------
    '64       �շѲ�����
    '65       Ʊ��ʹ�ü��
    '1502   ������
    '1504   Ʊ��������
    '1506   �տ����ʹ���
    '1507   �������տ����
    Select Case lngModul
    Case 1500
        strTitle = "�շѲ�����"
        For Each frmTemp In Forms
            If InStr(1, frmTemp.Caption, strTitle) > 0 Then
                frmTemp.Show
                frmTemp.ZOrder 0
                Exit Sub
            End If
        Next
        Set frmList = New frmFinanceSupervise
        Call frmList.zlShowFinanceSupervise(frmMain, glngModul, gstrPrivs)
    Case 1501
        frmBillSupervise.Show , frmMain
    Case 1502
        strTitle = "��Ա������"
        For Each frmTemp In Forms
            If InStr(1, frmTemp.Caption, strTitle) > 0 Then
                frmTemp.Show
                frmTemp.ZOrder 0
                Exit Sub
            End If
        Next
        Set frmList = New frmPersonLoanMgr
        frmList.ShowList lngModul, strTitle, frmMain
    Case 1504   'Ʊ��������
        frmBillIn.Show , frmMain
    Case 1506 '�շ����ʹ���
        strTitle = "�շ����ʹ���"
        For Each frmTemp In Forms
            If InStr(1, frmTemp.Caption, strTitle) > 0 Then
                frmTemp.Show
                frmTemp.ZOrder 0
                Exit Sub
            End If
        Next
        Set frmList = New frmChargeRollingCurtainMgr
        Call frmList.zlShowChargeRollingCourtain(frmMain, 1506, gstrPrivs)
    Case 1507  '�������տ����
        frmFeeGroupManage.Show , frmMain
    End Select
End Sub

Public Sub BHCodeMan(ByVal lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, lngMain As Long, ByVal strDbUser As String)
    '------------------------------------------------
    '���ܣ� �����ĳ�ʼ����������ִ�г���,��������ı�Ҫ���ݴ�������
    '������
    '   cnMain:����������ݿ�����
    '   lngMain:��������
    '���أ�
    '------------------------------------------------
    Dim strTitle As String, frmList As Object, frmTemp As Form
    Set gcnOracle = cnMain
    If Not CheckValid Then Exit Sub
    
    gstrAviPath = GetSetting(appName:="ZLSOFT", Section:="ע����Ϣ", Key:=UCase("gstrAviPath"), Default:="")
    gstrSysName = GetSetting(appName:="ZLSOFT", Section:="ע����Ϣ", Key:=UCase("gstrSysName"), Default:="")
    gstrVersion = GetSetting(appName:="ZLSOFT", Section:="ע����Ϣ", Key:=UCase("gstrVersion"), Default:="")
    gstrProductName = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒ����", "����")
    gstrLike = IIf(Val(zlDatabase.GetPara("����ƥ��")) = 0, "%", "")
    gstr��λ���� = GetUnitName()
    
    glngSys = lngSys
    glngModul = lngModul
    gstrPrivs = ";" & GetPrivFunc(lngSys, lngModul) & ";"
    gstr��λ���� = GetUnitName()
    gstrDbUser = strDbUser
    
    If Not GetUserInfo Then
        MsgBox "��ǰ�û�δ���ö�Ӧ����Ա��Ϣ,����ϵͳ����Ա��ϵ,�ȵ��û���Ȩ���������ã�", vbInformation, gstrSysName
        Exit Sub
    End If
    gstrMatchMethod = Val(zlDatabase.GetPara("����ƥ��"))
    '-------------------------------------------------
    '64       �շѲ�����
    '65       Ʊ��ʹ�ü��
    '1502   ������
    '1504   Ʊ��������
    Select Case lngModul
    Case 1500
        strTitle = "�շѲ�����"
        For Each frmTemp In Forms
            If InStr(1, frmTemp.Caption, strTitle) > 0 Then
                frmTemp.Show
                frmTemp.ZOrder 0
                Exit Sub
            End If
        Next
        Set frmList = New frmFinanceSupervise
        Call frmList.BHShowList(lngModul, gstrPrivs, lngMain)
    Case 1501
        zlCommFun.ShowChildWindow frmBillSupervise.hwnd, lngMain
    Case 1502
        strTitle = "��Ա������"
        For Each frmTemp In Forms
            If InStr(1, frmTemp.Caption, strTitle) > 0 Then
                frmTemp.Show
                frmTemp.ZOrder 0
                Exit Sub
            End If
        Next
        Set frmList = New frmPersonLoanMgr
        frmList.BHShowList lngModul, strTitle, lngMain
    Case 1504   'Ʊ��������
        zlCommFun.ShowChildWindow frmBillIn.hwnd, lngMain
    Case 1506 '�շ����ʹ���
        strTitle = "�շ����ʹ���"
        For Each frmTemp In Forms
            If InStr(1, frmTemp.Caption, strTitle) > 0 Then
                frmTemp.Show
                frmTemp.ZOrder 0
                Exit Sub
            End If
        Next
        Set frmList = New frmChargeRollingCurtainMgr
        Call frmList.BHShowList(lngModul, gstrPrivs, lngMain)
    Case 1507  '�������տ����
        zlCommFun.ShowChildWindow frmFeeGroupManage.hwnd, lngMain
    End Select
End Sub