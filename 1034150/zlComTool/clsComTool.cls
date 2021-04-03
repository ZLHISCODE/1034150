VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsComTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Function GetLBoundIcon() As Long
'ͼ���������½�
    GetLBoundIcon = glngLBound
End Function

Public Function GetUBoundIcon() As Long
'ͼ���������Ͻ�
    GetUBoundIcon = glngUBound
End Function


Public Sub CodeMan(ByVal lngSysIN As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, frmMain As Object, ByVal strDbUser As String, Optional ByVal strParameter As String)
    '------------------------------------------------
    '���ܣ� ����������ָ�����ܣ�����ִ����س���
    '������
    '   lngModul:��Ҫִ�еĹ������
    '   cnMain:����������ݿ�����
    '   frmMain:������
    '���أ�
    '------------------------------------------------
    Set gcnOracle = cnMain
    glngSys = lngSysIN
    gstrAviPath = GetSetting(appName:="ZLSOFT", Section:="ע����Ϣ", Key:=UCase("gstrAviPath"), Default:="")
    gstrSysName = GetSetting(appName:="ZLSOFT", Section:="ע����Ϣ", Key:=UCase("gstrSysName"), Default:="")
    gstrVersion = GetSetting(appName:="ZLSOFT", Section:="ע����Ϣ", Key:=UCase("gstrVersion"), Default:="")
    gstr��λ���� = zlreginfo("��λ����")
    gstrDbUser = strDbUser
    '������
    gstrOwner = zldatabase.GetOwner(lngSysIN)
    
    Set gfrmMain = frmMain
    
    '-------------------------------------------------
    '1   �ֵ��������
    '2   �ʼ��շ�����
    
    
    Select Case lngModul
    Case 1
        Call GetUserInfo
        gstrPrivs = GetPrivFunc(0, 11)
        frmDictManager.�ֵ����
    Case 2
        Call GetUserInfo
        frmMessageManager.Show , frmMain
    Case 3
        Call GetUserInfo
        gstrPrivs = GetPrivFunc(0, 13)
        gstrMenuSys = strParameter
        frmOptions.Show vbModal, frmMain
    Case 4
        NotifyMessage (strParameter = "Open")
    Case 5
        
        Call GetUserInfo
        ShowWindow frmAlert.hWnd, 0
        Call frmAlert.InitAlert
        
    Case 6
        
        frmAlert.Show , frmMain
        Call frmAlert.InitData
        
    End Select
End Sub

Private Sub NotifyMessage(blnOpen As Boolean)
'���ܣ��Զ���Ϣ֪ͨ����Ŀ���
    If gblnMessageGet = blnOpen Then Exit Sub
    
    gblnMessageGet = blnOpen
    If blnOpen = True Then
        Call GetUserInfo
        '�򿪹ر�
        If gblnMessageShow = True Then
            '�Ѿ���
            Exit Sub
        Else
            Load frmMessageRead
        End If
    Else
        If gblnMessageShow = True Then
            '��Ϣ�շ�����Ҫ��
            Exit Sub
        Else
            Unload frmMessageRead
        End If
    End If
End Sub

Public Function CloseWindows() As Boolean
    '--------------------------------------
    '����:�ر������Ӵ���
    '--------------------------------------
    Dim frmThis As Form
    
    On Error Resume Next
        
    Call RemoveIcon(frmAlert.picNotify.hWnd)
    
    For Each frmThis In Forms
    
        If frmThis.Caption = "������Ϣ" Then
            frmAlert.mblnUnLoad = True
        End If
        
        Unload frmThis
    Next
    CloseWindows = (Forms.Count = 0)
End Function

Public Function GetIcon(ByVal IntIcon As Integer) As IPictureDisp
    Set GetIcon = LoadResPicture(IntIcon, vbResIcon)
End Function

Public Function GetResource(ByVal StrName As String) As IPictureDisp
    Set GetResource = LoadCustomPicture(StrName)
End Function
