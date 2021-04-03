VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStuff"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Version As String

Public Sub TransStuff(frmMain As Object, cnMain As ADODB.Connection, lngSys As Long, ByVal strDBUser As String, ByVal lng����id As Long, ByVal strNo As String, ByVal lng�ⷿID As Long, ByVal strStartDate As String, ByVal strEndDate As String)
'���ܣ�ֱ�ӵ������ķ��Ź���
    
    '��������
    Set gfrmMain = frmMain
    If Not CheckValid Then Exit Sub
    
    Set gcnOracle = cnMain
    glngSys = lngSys
    glngModul = 1723
    gstrDBUser = strDBUser
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
    gstrUnitName = GetUnitName
    gstrPrivs = GetPrivFunc(lngSys, 1723)
                
    '�û���Ϣ
    Call GetUserInfo
    
    '��ʼ�����С��λ��
    Call ��ʼС��λ��
    
    If glngModul <> 1720 Then
        '���˺�:2007/06/28:��ʼ��ϵͳ����
        If InitSystemPara = False Then Exit Sub
    End If
    
    '�������ķ��Ź���
    err.Clear
    On Error Resume Next
    If gfrmMain Is Nothing Then
        frm���ķ��Ź���_New.Show
    Else
        frm���ķ��Ź���_New.ShowList frmMain, lng����id, strNo, lng�ⷿID, strStartDate, strEndDate
    End If
End Sub

Public Sub CodeMan(ByVal lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, frmMain As Object, strDBUser As String)
    '------------------------------------------------
    '���ܣ� ����������ָ�����ܣ�����ִ����س���
    '������
    '   lngSys : ϵͳ���
    '   lngModul:��Ҫִ�еĹ������
    '   cnMain:����������ݿ�����
    '   frmMain:������
    '���أ�
    '------------------------------------------------
    Dim frmBillMgr As Form
    Dim frmTemp As Form
    Dim strTitle As String
    
    Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
    
    gstrDBUser = strDBUser
    
    glngModul = lngModul
    glngSys = lngSys
    
    If Not GetUserInfo Then
        MsgBox "��ǰ�û�δ���ö�Ӧ����Ա��Ϣ,����ϵͳ����Ա��ϵ,�ȵ��û���Ȩ���������ã�", vbInformation, gstrSysName
        Exit Sub
    End If
    Set gfrmMain = frmMain
    
    '�ж�����ϵͳ�Ƿ�������������Ȩ��
    gblnCode = GetCodePrivs
    
    '��ȡ���Ĳ���Ҫʹ�õ�ϵͳ����ֵ
    Call GetSysParms
    
    '��ʼ�����С��λ��
    Call ��ʼС��λ��
    
    If lngModul <> 1720 Then
        '���˺�:2007/06/28:��ʼ��ϵͳ����
        If InitSystemPara = False Then Exit Sub
    End If
    
    gbytSimpleCodeTrans = Val(zlDatabase.GetPara("����ƥ�䷽ʽ�л�", , , 1))
    
    gstrMatchMethod = Val(zlDatabase.GetPara("����ƥ��"))
    gstrPrivs = ";" & GetPrivFunc(lngSys, lngModul) & ";"
    
    gstrUnitName = GetUnitName
    
    '-------------------------------------------------
    On Error Resume Next
    
    Call setNOtExcetePrice '������δִ�еļ۸���ִ����
    
    Select Case lngModul
        Case 1710                         'ϵͳ��������'
            frmSystemParaSet.Show , frmMain
        Case 1711                         '����Ŀ¼����'
            frmStuffMgr.Show , frmMain
        Case 1712                       '�����⹺������
            strTitle = "�����⹺������"
            For Each frmTemp In Forms
                If frmTemp.Caption = strTitle Then
                    frmTemp.Show
                    frmTemp.ZOrder 0
                    Exit Sub
                End If
            Next
            Set frmBillMgr = New frmMainList
            frmBillMgr.ShowList lngModul, strTitle, frmMain
        Case 1713
            strTitle = "��������������"
            For Each frmTemp In Forms
                If frmTemp.Caption = strTitle Then
                    frmTemp.Show
                    frmTemp.ZOrder 0
                    Exit Sub
                End If
            Next
            Set frmBillMgr = New frmMainList
            frmBillMgr.ShowList lngModul, strTitle, frmMain
        Case 1714                        '��������������
            strTitle = "��������������"
            For Each frmTemp In Forms
                If frmTemp.Caption = strTitle Then
                    frmTemp.Show
                    frmTemp.ZOrder 0
                    Exit Sub
                End If
            Next
            Set frmBillMgr = New frmMainList
            frmBillMgr.ShowList lngModul, strTitle, frmMain
        Case 1715
            strTitle = "����۵�������"
            For Each frmTemp In Forms
                If frmTemp.Caption = strTitle Then
                    frmTemp.Show
                    frmTemp.ZOrder 0
                    Exit Sub
                End If
            Next
            Set frmBillMgr = New frmMainList
            frmBillMgr.ShowList lngModul, strTitle, frmMain
        Case 1716
            strTitle = "�����ƿ����"
            For Each frmTemp In Forms
                If frmTemp.Caption = strTitle Then
                    frmTemp.Show
                    frmTemp.ZOrder 0
                    Exit Sub
                End If
            Next
            Set frmBillMgr = New frmMainList
            frmBillMgr.ShowList lngModul, strTitle, frmMain
        
        Case 1717
            strTitle = "�������ù���"
            For Each frmTemp In Forms
                If frmTemp.Caption = strTitle Then
                    frmTemp.Show
                    frmTemp.ZOrder 0
                    Exit Sub
                End If
            Next
            Set frmBillMgr = New frmMainList
            frmBillMgr.ShowList lngModul, strTitle, frmMain
        Case 1718
            strTitle = "���������������"
            For Each frmTemp In Forms
                If frmTemp.Caption = strTitle Then
                    frmTemp.Show
                    frmTemp.ZOrder 0
                    Exit Sub
                End If
            Next
            Set frmBillMgr = New frmMainList
            frmBillMgr.ShowList lngModul, strTitle, frmMain
        Case 1719
           strTitle = "�����̵����"
            For Each frmTemp In Forms
                If frmTemp.Caption = strTitle Then
                    frmTemp.Show
                    frmTemp.ZOrder 0
                    Exit Sub
                End If
            Next
            Set frmBillMgr = New frmCheckMain
            frmBillMgr.ShowList lngModul, strTitle, frmMain
        Case 1720                         '���Ĳ�ۼ���'
           frmDiffPriceRecalCard.Show , frmMain
        Case 1721                         '���Ŀ���ѯ'
            frmStuffQuery.Show , frmMain
        Case 1722                         '��������
            frmRequestStuffList.ShowList frmMain
        Case 1723                         '���ķ��Ź���
             'frm���ķ��Ź���.Show , frmMain
             frm���ķ��Ź���_New.Show , frmMain
        Case 1724               '���ļƻ�����
            strTitle = "���ļƻ�����"
            For Each frmTemp In Forms
                If frmTemp.Caption = strTitle Then
                    frmTemp.Show
                    frmTemp.ZOrder 0
                    Exit Sub
                End If
            Next
            Set frmBillMgr = New frmStuffPlanList
            frmBillMgr.ShowList lngModul, strTitle, frmMain
        Case 1725               '�����깺����
            strTitle = "�����깺����"
            For Each frmTemp In Forms
                If frmTemp.Caption = strTitle Then
                    frmTemp.Show
                    frmTemp.ZOrder 0
                    Exit Sub
                End If
            Next
            Set frmBillMgr = New frmStuffPlanList
            frmBillMgr.ShowList lngModul, strTitle, frmMain
        Case 1726
            frmStuffPriceList.Show , frmMain
        Case 1729
            frmClosingAccount.Show , frmMain
    End Select
End Sub

Public Sub BHCodeMan(ByVal lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, lngMain As Long, strDBUser As String)
    '------------------------------------------------
    '���ܣ� ����������ָ�����ܣ�����ִ����س���
    '������
    '   lngSys : ϵͳ���
    '   lngModul:��Ҫִ�еĹ������
    '   cnMain:����������ݿ�����
    '   lngMain:��������
    '���أ�
    '------------------------------------------------
    Dim frmBillMgr As Form
    Dim frmTemp As Form
    Dim strTitle As String
    
    Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
    
    gstrDBUser = strDBUser
    
    glngModul = lngModul
    glngSys = lngSys
    
    If Not GetUserInfo Then
        MsgBox "��ǰ�û�δ���ö�Ӧ����Ա��Ϣ,����ϵͳ����Ա��ϵ,�ȵ��û���Ȩ���������ã�", vbInformation, gstrSysName
        Exit Sub
    End If
    Set gfrmMain = Nothing
    
    '�ж�����ϵͳ�Ƿ�������������Ȩ��
    gblnCode = GetCodePrivs
    
    '��ʼ�����С��λ��
    Call ��ʼС��λ��
    
    If lngModul <> 1720 Then
        '���˺�:2007/06/28:��ʼ��ϵͳ����
        If InitSystemPara = False Then Exit Sub
    End If
    
    gbytSimpleCodeTrans = Val(zlDatabase.GetPara("����ƥ�䷽ʽ�л�", , , 1))
    
    gstrMatchMethod = Val(zlDatabase.GetPara("����ƥ��"))
    gstrPrivs = ";" & GetPrivFunc(lngSys, lngModul) & ";"
    
    gstrUnitName = GetUnitName
    
    '-------------------------------------------------
    On Error Resume Next
    
    Call setNOtExcetePrice '������δִ�еļ۸���ִ����
    
    Select Case lngModul
        Case 1710                         'ϵͳ��������'
            zlCommFun.ShowChildWindow frmSystemParaSet.hwnd, lngMain
        Case 1711                         '����Ŀ¼����'
            zlCommFun.ShowChildWindow frmStuffMgr.hwnd, lngMain
        Case 1712                       '�����⹺������
            strTitle = "�����⹺������"
            For Each frmTemp In Forms
                If frmTemp.Caption = strTitle Then
                    frmTemp.Show
                    frmTemp.ZOrder 0
                    Exit Sub
                End If
            Next
            Set frmBillMgr = New frmMainList
            frmBillMgr.ShowList lngModul, strTitle, lngMain
        Case 1713
            strTitle = "��������������"
            For Each frmTemp In Forms
                If frmTemp.Caption = strTitle Then
                    frmTemp.Show
                    frmTemp.ZOrder 0
                    Exit Sub
                End If
            Next
            Set frmBillMgr = New frmMainList
            frmBillMgr.ShowList lngModul, strTitle, lngMain
        Case 1714                        '��������������
            strTitle = "��������������"
            For Each frmTemp In Forms
                If frmTemp.Caption = strTitle Then
                    frmTemp.Show
                    frmTemp.ZOrder 0
                    Exit Sub
                End If
            Next
            Set frmBillMgr = New frmMainList
            frmBillMgr.ShowList lngModul, strTitle, lngMain
        Case 1715
            strTitle = "����۵�������"
            For Each frmTemp In Forms
                If frmTemp.Caption = strTitle Then
                    frmTemp.Show
                    frmTemp.ZOrder 0
                    Exit Sub
                End If
            Next
            Set frmBillMgr = New frmMainList
            frmBillMgr.ShowList lngModul, strTitle, lngMain
        Case 1716
            strTitle = "�����ƿ����"
            For Each frmTemp In Forms
                If frmTemp.Caption = strTitle Then
                    frmTemp.Show
                    frmTemp.ZOrder 0
                    Exit Sub
                End If
            Next
            Set frmBillMgr = New frmMainList
            frmBillMgr.ShowList lngModul, strTitle, lngMain
        
        Case 1717
            strTitle = "�������ù���"
            For Each frmTemp In Forms
                If frmTemp.Caption = strTitle Then
                    frmTemp.Show
                    frmTemp.ZOrder 0
                    Exit Sub
                End If
            Next
            Set frmBillMgr = New frmMainList
            frmBillMgr.ShowList lngModul, strTitle, lngMain
        Case 1718
            strTitle = "���������������"
            For Each frmTemp In Forms
                If frmTemp.Caption = strTitle Then
                    frmTemp.Show
                    frmTemp.ZOrder 0
                    Exit Sub
                End If
            Next
            Set frmBillMgr = New frmMainList
            frmBillMgr.ShowList lngModul, strTitle, lngMain
        Case 1719
           strTitle = "�����̵����"
            For Each frmTemp In Forms
                If frmTemp.Caption = strTitle Then
                    frmTemp.Show
                    frmTemp.ZOrder 0
                    Exit Sub
                End If
            Next
            Set frmBillMgr = New frmCheckMain
            frmBillMgr.ShowList lngModul, strTitle, lngMain
        Case 1720                         '���Ĳ�ۼ���'
            zlCommFun.ShowChildWindow frmDiffPriceRecalCard.hwnd, lngMain
        Case 1721                         '���Ŀ���ѯ'
            zlCommFun.ShowChildWindow frmStuffQuery.hwnd, lngMain
        Case 1722                         '��������
            frmRequestStuffList.ShowList lngMain
        Case 1723                         '���ķ��Ź���
             zlCommFun.ShowChildWindow frm���ķ��Ź���_New.hwnd, lngMain
        Case 1724               '���ļƻ�����
            strTitle = "���ļƻ�����"
            For Each frmTemp In Forms
                If frmTemp.Caption = strTitle Then
                    frmTemp.Show
                    frmTemp.ZOrder 0
                    Exit Sub
                End If
            Next
            Set frmBillMgr = New frmStuffPlanList
            frmBillMgr.ShowList lngModul, strTitle, lngMain
        Case 1725               '�����깺����
            strTitle = "�����깺����"
            For Each frmTemp In Forms
                If frmTemp.Caption = strTitle Then
                    frmTemp.Show
                    frmTemp.ZOrder 0
                    Exit Sub
                End If
            Next
            Set frmBillMgr = New frmStuffPlanList
            frmBillMgr.ShowList lngModul, strTitle, lngMain
        Case 1726   '���ĵ��۹���
            zlCommFun.ShowChildWindow frmStuffPriceList.hwnd, lngMain
    End Select
End Sub

Private Sub setNOtExcetePrice()
    '����ѵ�ִ�����ڶ��۸�δִ�У�ִ�м������
        Dim rsTemp As ADODB.Recordset
        Dim lng�շ�ϸĿID As Long
        
        gstrSQL = "Select a.Id From �շѼ�Ŀ A, �������� B Where a.�շ�ϸĿid = b.����id And a.�䶯ԭ�� = 0"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "������ʱ��δ��Ч�ļ۸�")
        
        With rsTemp
            If Not rsTemp.EOF Then
                rsTemp.MoveFirst
                Do While Not rsTemp.EOF
                    If Not IsNull(rsTemp!Id) Then
                        lng�շ�ϸĿID = rsTemp!Id
                        gstrSQL = "zl_�����շ���¼_Adjust(" & lng�շ�ϸĿID & ")"
                        
                        Call zlDatabase.ExecuteProcedure(gstrSQL, "������ʱ��δ��Ч�ļ۸�" & "-�������ϼ۸������¼")
                    End If
                    rsTemp.MoveNext
                Loop
            End If
        End With

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Class_Initialize()
    Version = App.Major & "." & App.Minor & "." & App.Revision
    
End Sub

Public Function CloseWindows() As Boolean
    '--------------------------------------
    '����:�ر������Ӵ���
    '--------------------------------------
    Dim frmThis As Form
    For Each frmThis In Forms
        Unload frmThis
    Next
    CloseWindows = (Forms.Count = 0)
End Function


