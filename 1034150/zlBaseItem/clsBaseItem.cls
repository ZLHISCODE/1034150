VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBaseItem"
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

    '�ر�ר����ʵ��Ĵ��ڣ���Ϊ������һ�������У�����Ҫ��������
    If Not gobjCustAcc Is Nothing Then
        CloseWindows = CloseWindows And gobjCustAcc.CloseWindows
    End If
End Function

Public Sub CodeMan(ByVal lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, frmMain As Object, ByVal strDbUser As String)
    '------------------------------------------------
    '���ܣ� ����������ָ�����ܣ�����ִ����س���
    '������
    '   lngModul:��Ҫִ�еĹ������
    '   cnMain:����������ݿ�����
    '   frmMain:������
    '���أ�
    '------------------------------------------------
    InitCommon cnMain
    Set gcnOracle = cnMain
    gstrAviPath = GetSetting(appName:="ZLSOFT", Section:="ע����Ϣ", Key:=UCase("gstrAviPath"), Default:="")
    gstrSysName = GetSetting(appName:="ZLSOFT", Section:="ע����Ϣ", Key:=UCase("gstrSysName"), Default:="")
    gstrVersion = GetSetting(appName:="ZLSOFT", Section:="ע����Ϣ", Key:=UCase("gstrVersion"), Default:="")
    gblnMyStyle = zlDatabase.GetPara("ʹ�ø��Ի����") = "1"
    gstr��λ���� = GetUnitName()
    gstrDbUser = strDbUser
    Call GetUserInfo
    
    glngSys = lngSys
    gstrPrivs = ";" & GetPrivFunc(lngSys, lngModul) & ";"
    glngModul = lngModul
    If Not CheckValid Then Exit Sub
    '-------------------------------------------------
    '1000   ����ת�ƹ���
    '1001   ���Ź���
    '1002   ��Ա����
    '1003   ���㷽ʽ����
    '1004   ������Ŀ����
    '1005   �ѱ�ȼ�����
    '1009   �շ�ϸĿ����
    '1010   �ڼ仮��
    '1011   ������������
    '1012   �����ϰల��
    '1013   �����������
    '1014   ҽ��С�����
    '1021   �������
    '1031   ר����ʵ�����
    
    
    Select Case lngModul
    Case 1000
        On Error Resume Next
        frmDataMove.Show , frmMain
    Case 1001
        frmDeptManage.Show , frmMain
    Case 1002
        frmPresManage.Show , frmMain
    Case 1003
        frmBalanceManage.Show , frmMain
    Case 1004
        frmEarnManage.Show , frmMain
    Case 1005
        frmChargeSortGrade.Show , frmMain
    Case 1009
        Call InitSystemPara
        frmChargeManage.Show , frmMain
    Case 1010
        frmCourse.Show , frmMain
    Case 1011
        If glngSys \ 100 = 8 Then
            'ҩ�����в���
            frmDrugstorePara.Show vbModal, frmMain
        Else
            frmSystemPara.Show vbModal, frmMain
        End If
    Case 1012
        frmDeptTime.Show , frmMain
    Case 1013
        frmIllManage.Show , frmMain
    Case 1014
        frmMedicalTeam.Show , frmMain
    Case 1031
        If gobjCustAcc Is Nothing Then
            Set gobjCustAcc = CreateObject("zl9CustAcc.clsCustAcc")
        End If
        If gobjCustAcc Is Nothing Then Exit Sub
        Call gobjCustAcc.CodeMan(glngSys, 1031, gcnOracle, frmMain, strDbUser)
    Case 1021
        frmMedInOutClass.Show , frmMain
    End Select
End Sub

Public Sub BHCodeMan(ByVal lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, lngMain As Long, ByVal strDbUser As String)
    '------------------------------------------------
    '���ܣ� ����������ָ�����ܣ�����ִ����س���
    '������
    '   lngModul:��Ҫִ�еĹ������
    '   cnMain:����������ݿ�����
    '   lngMain:��������
    '���أ�
    '------------------------------------------------
    InitCommon cnMain
    Set gcnOracle = cnMain
    gstrAviPath = GetSetting(appName:="ZLSOFT", Section:="ע����Ϣ", Key:=UCase("gstrAviPath"), Default:="")
    gstrSysName = GetSetting(appName:="ZLSOFT", Section:="ע����Ϣ", Key:=UCase("gstrSysName"), Default:="")
    gstrVersion = GetSetting(appName:="ZLSOFT", Section:="ע����Ϣ", Key:=UCase("gstrVersion"), Default:="")
    gblnMyStyle = zlDatabase.GetPara("ʹ�ø��Ի����") = "1"
    gstr��λ���� = GetUnitName()
    gstrDbUser = strDbUser
    Call GetUserInfo
    
    glngSys = lngSys
    gstrPrivs = ";" & GetPrivFunc(lngSys, lngModul) & ";"
    glngModul = lngModul
    If Not CheckValid Then Exit Sub
    '-------------------------------------------------
    '1000   ����ת�ƹ���
    '1001   ���Ź���
    '1002   ��Ա����
    '1003   ���㷽ʽ����
    '1004   ������Ŀ����
    '1005   �ѱ�ȼ�����
    '1009   �շ�ϸĿ����
    '1010   �ڼ仮��
    '1011   ������������
    '1012   �����ϰల��
    '1013   �����������
    '1014   ҽ��С�����
    '1031   ר����ʵ�����
    
    
    Select Case lngModul
    Case 1000
        On Error Resume Next
        zlCommFun.ShowChildWindow frmDataMove.hwnd, lngMain
    Case 1001
        zlCommFun.ShowChildWindow frmDeptManage.hwnd, lngMain
    Case 1002
        zlCommFun.ShowChildWindow frmPresManage.hwnd, lngMain
    Case 1003
        zlCommFun.ShowChildWindow frmBalanceManage.hwnd, lngMain
    Case 1004
        zlCommFun.ShowChildWindow frmEarnManage.hwnd, lngMain
    Case 1005
        zlCommFun.ShowChildWindow frmChargeSortGrade.hwnd, lngMain
    Case 1009
        Call InitSystemPara
        zlCommFun.ShowChildWindow frmChargeManage.hwnd, lngMain
    Case 1010
        zlCommFun.ShowChildWindow frmCourse.hwnd, lngMain
    Case 1011
        If glngSys \ 100 = 8 Then
            'ҩ�����в���
            zlCommFun.ShowChildWindow frmDrugstorePara.hwnd, lngMain
        Else
            zlCommFun.ShowChildWindow frmSystemPara.hwnd, lngMain
        End If
    Case 1012
        zlCommFun.ShowChildWindow frmDeptTime.hwnd, lngMain
    Case 1013
        zlCommFun.ShowChildWindow frmIllManage.hwnd, lngMain
    Case 1014
        zlCommFun.ShowChildWindow frmMedicalTeam.hwnd, lngMain
    Case 1031
        If gobjCustAcc Is Nothing Then
            Set gobjCustAcc = CreateObject("zl9CustAcc.clsCustAcc")
        End If
        If gobjCustAcc Is Nothing Then Exit Sub
        Call gobjCustAcc.BHCodeMan(glngSys, 1031, gcnOracle, lngMain, strDbUser)
    Case 1021
        zlCommFun.ShowChildWindow frmMedInOutClass.hwnd, lngMain
    End Select
End Sub

Public Function OpenEditWholeSetItem(ByVal frmMain As Object, ByVal cnOracle As ADODB.Connection, _
     ByVal lngSys As Long, ByVal lngModule As Long, ByVal strPrivs As String, ByVal strItems As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ӵ����б���ɳ�����Ŀ����
    '���:frmMain:���õ�������
    '       cnOracle    ADODB.Connection    ����    �������Ӷ���
    '       lngSys  Long    ����    ϵͳ��
    '       lngModule   Long    ����    ģ���
    '       strPrivs    String  ����    Ȩ�޴�
    '       strItems    String  ����    ���,����,�շ�ϸĿID,����,����,ִ�п���|���,����,�շ�ϸĿID,����,����,ִ�п���|��
    '����:
    '����:True:����ɹ�,False:����ʧ��
    '����:���˺�
    '����:2010-09-01 14:06:53
    '˵��: Ŀǰ֧������ģ�鹦�ܵ���:
    '        1.  סԺ����(����)
    '        2.  ҽ������վ(����)
    '       �ýӿ��ڲ���Ҫ�������Ȩ��:
    '       ��"���ӳ�����Ŀ"��Ȩ��ʱ,��ֱ�ӷ���false
    '       ��"ȫԺ���׷���"��Ȩ��ʱ,���ܸ���ʹ�÷�ΧΪȫԺ
    '       ��"���Ƴ��׷���"��Ȩ��ʱ,���ܸ���Ϊָ������,ֻ�������Լ��ĳ�����Ŀ.
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng�ϴη���ID  As Long
    
    If InStr(1, strPrivs, ";���ӳ�����Ŀ;") = 0 Then Exit Function
    
    InitCommon cnOracle
    Set gcnOracle = cnOracle
    gstrAviPath = GetSetting(appName:="ZLSOFT", Section:="ע����Ϣ", Key:=UCase("gstrAviPath"), Default:="")
    gstrSysName = GetSetting(appName:="ZLSOFT", Section:="ע����Ϣ", Key:=UCase("gstrSysName"), Default:="")
    gstrVersion = GetSetting(appName:="ZLSOFT", Section:="ע����Ϣ", Key:=UCase("gstrVersion"), Default:="")
    gstr��λ���� = GetUnitName()
    Call GetUserInfo
    glngSys = lngSys
    glngModul = lngModule
    lng�ϴη���ID = Val(zlDatabase.GetPara("�ϴγ��׷�������", glngSys, lngModule, ""))
    Call InitSystemPara
    OpenEditWholeSetItem = frmChargeWholeSetItemEdit.ShowCard(frmMain, EdI_����, strPrivs, lngModule, lng�ϴη���ID, "", strItems)
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

Private Sub Class_Terminate()
    Set gobjPlugIn = Nothing
    
    If Not gobjRIS Is Nothing Then
        Set gobjRIS = Nothing
    End If
End Sub