VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCustAcc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Public Version As String

Private Sub Class_Initialize()
    Version = App.Major & "." & App.Minor & "." & App.Revision
    Call RegBillFile
End Sub

Public Sub CodeMan(ByVal lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, frmMain As Object, ByVal strDbUser As String, Optional varParameter As Variant)
    '------------------------------------------------
    '���ܣ� ����������ָ�����ܣ�����ִ����س���
    '������
    '   lngModul:��Ҫִ�еĹ������
    '   cnMain:����������ݿ�����
    '   frmMain:������
    '���أ�
    '------------------------------------------------
    Set gcnOracle = cnMain
    gstrAviPath = GetSetting("ZLSOFT", "ע����Ϣ", UCase("gstrAviPath"), "")
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", UCase("gstrSysName"), "")
    gstrVersion = GetSetting("ZLSOFT", "ע����Ϣ", UCase("gstrVersion"), "")
    gstr��λ���� = GetUnitName()
    gstrDbUser = strDbUser
    glngSys = lngSys
    glngModul = lngModul
    
    Set gfrmMain = frmMain
    glngMain = 0
    
    If Not GetUserInfo Then
        MsgBox "��ǰ�û�δ���ö�Ӧ����Ա��Ϣ,����ϵͳ����Ա��ϵ,�ȵ��û���Ȩ���������ã�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    Select Case lngModul
        Case 1031 '���ʵ����
            gstrPrivs = GetPrivFunc(lngSys, lngModul)
            
            Set frmDesign = Nothing
            frmDesign.Show , frmMain
        Case -1
            gblnOK = False
            
            '�������������ǣ�
            '����ID��bytUseType��bytInState��strInNO��lngUnitID��lngDeptID��lng����ID��gstrPrivs��blnViewCancel
            gstrPrivs = varParameter(7)
            
            'ҽ���ӿ�(סԺ����)
            If varParameter(1) <> 3 Then
                Call gclsInsure.InitOracle(gcnOracle)
            End If
            
            If InStr(1, GetInsidePrivs(Enum_Inside_Program.p���ʲ���), "ר�����") = 0 Then
                MsgBox "��û�в���ר����ʵ���Ȩ��!", vbInformation, gstrSysName
                varParameter = gblnOK
                Exit Sub
            End If
            
            With frmCharge
                .mlng����ID = Val(varParameter(0))
                .mbytUseType = varParameter(1)
                .mbytInState = varParameter(2)
                .mstrInNO = varParameter(3)
                .mlngUnitID = Val(varParameter(4))
                .mlngDeptID = Val(varParameter(5))
                .mlng����ID = Val(varParameter(6))
                .mstrPrivs = gstrPrivs
                
                If .mbytInState = sta���� Then
                    .mblnViewCancel = varParameter(8)
                Else
                    .mblnViewCancel = False
                End If
            End With
            Call InitLocPar(CByte(varParameter(1)))
            Call InitSysPar
            frmCharge.MainProc
            varParameter = gblnOK
    End Select
End Sub

Public Sub BHCodeMan(ByVal lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, lngMain As Long, ByVal strDbUser As String, Optional varParameter As Variant)
    '------------------------------------------------
    '���ܣ� ����������ָ�����ܣ�����ִ����س���
    '������
    '   lngModul:��Ҫִ�еĹ������
    '   cnMain:����������ݿ�����
    '   lngMain:��������
    '���أ�
    '------------------------------------------------
    Set gcnOracle = cnMain
    gstrAviPath = GetSetting("ZLSOFT", "ע����Ϣ", UCase("gstrAviPath"), "")
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", UCase("gstrSysName"), "")
    gstrVersion = GetSetting("ZLSOFT", "ע����Ϣ", UCase("gstrVersion"), "")
    gstr��λ���� = GetUnitName()
    gstrDbUser = strDbUser
    glngSys = lngSys
    glngModul = lngModul
    
    Set gfrmMain = Nothing
    glngMain = lngMain
    
    If Not GetUserInfo Then
        MsgBox "��ǰ�û�δ���ö�Ӧ����Ա��Ϣ,����ϵͳ����Ա��ϵ,�ȵ��û���Ȩ���������ã�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    Select Case lngModul
        Case 1031 '���ʵ����
            gstrPrivs = GetPrivFunc(lngSys, lngModul)
            
            Set frmDesign = Nothing
            zlCommFun.ShowChildWindow frmDesign.hwnd, lngMain
        Case -1
            gblnOK = False
            
            '�������������ǣ�
            '����ID��bytUseType��bytInState��strInNO��lngUnitID��lngDeptID��lng����ID��gstrPrivs��blnViewCancel
            gstrPrivs = varParameter(7)
            
            'ҽ���ӿ�(סԺ����)
            If varParameter(1) <> 3 Then
                Call gclsInsure.InitOracle(gcnOracle)
            End If
            
            If InStr(1, GetInsidePrivs(Enum_Inside_Program.p���ʲ���), "ר�����") = 0 Then
                MsgBox "��û�в���ר����ʵ���Ȩ��!", vbInformation, gstrSysName
                varParameter = gblnOK
                Exit Sub
            End If
            
            With frmCharge
                .mlng����ID = Val(varParameter(0))
                .mbytUseType = varParameter(1)
                .mbytInState = varParameter(2)
                .mstrInNO = varParameter(3)
                .mlngUnitID = Val(varParameter(4))
                .mlngDeptID = Val(varParameter(5))
                .mlng����ID = Val(varParameter(6))
                .mstrPrivs = gstrPrivs
                
                If .mbytInState = sta���� Then
                    .mblnViewCancel = varParameter(8)
                Else
                    .mblnViewCancel = False
                End If
            End With
            Call InitLocPar(CByte(varParameter(1)))
            Call InitSysPar
            frmCharge.MainProc
            varParameter = gblnOK
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
    CloseWindows = (Forms.Count = 0)
End Function