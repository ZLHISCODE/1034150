VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRegist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private mfrmRegist As frmRegist
Attribute mfrmRegist.VB_VarHelpID = -1

Public Function zlInitCommon(ByVal lngSys As Long, _
     ByVal cnOracle As ADODB.Connection, Optional ByVal strDbUser As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����ص�ϵͳ�ż��������
    '���:lngSys-ϵͳ��
    '     cnOracle-���ݿ����Ӷ���
    '     strDBUser-���ݿ�������
    '����:��ʼ���ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-03-20 17:55:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    glngSys = lngSys: Set gcnOracle = cnOracle: gstrDBUser = strDbUser
    Call gclsInsure.InitOracle(gcnOracle)
    If zlGetComLib = False Then Exit Function
    
    Call GetUserInfo
    Call InitVar
    Set mfrmRegist = New frmRegist
    zlInitCommon = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlInitData(ByVal bytState As Byte, _
     Optional ByVal strDocs As String, Optional ByVal strDeptIDs As String, Optional ByVal strRooms As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����ص���������
    '���:bytState-Ӧ�ó��ϣ�Ŀǰ֧�� 0-����̨���ã�1-����ҽ��վ���ã�2-סԺҽ��վ����
    '     strDocs-ҽ��ID����ҽ������,�ö��ŷָ�����ʽΪ 147,235,313���������������ģ����塭������ʱ��ʾ����ҽ���ĹҺŰ���
    '     strDeptIDs-����ID,�ö��ŷָ�����ʽΪ 147,235,313��������ʱ��ʾ�ÿ��ҵĹҺŰ���
    '     strRooms-ҽ��վ����ʱ���ɴ������ң��Һ�ʱ�Զ�ȷ�����ҡ�
    '����:��ʼ���ɹ�,����true,���򷵻�False
    '����:������
    '����:2015-04-22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    
    gbytState = bytState
    Select Case bytState
        Case 0
            glngModul = 1113
        Case 1
            glngModul = 1260
        Case 2
            glngModul = 1261
        Case 3
            glngModul = 1115
    End Select
    gstrDocs = strDocs
    gstrDeptIDs = strDeptIDs
    gstrRooms = strRooms
    gstrPrivs = ";" & gobjComlib.GetPrivFunc(glngSys, 9000) & ";"
    gbytRegistMode = Val(Split(gobjDatabase.GetPara("�Һ��Ű�ģʽ", glngSys) & "|", "|")(0))
    If Split(gobjDatabase.GetPara("�Һ��Ű�ģʽ", glngSys) & "|", "|")(1) <> "" Then
        gdatRegistTime = CDate(Format(Split(gobjDatabase.GetPara("�Һ��Ű�ģʽ", glngSys) & "|", "|")(1), "yyyy-mm-dd hh:mm:ss"))
    End If
    
    zlInitData = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Sub zlDefCommandBars(ByVal frmMain As Object, ByVal CommandBars As Object, _
                            Optional ByVal blnAddInTool As Boolean, Optional ByVal MenuControlBefore As Object, _
                            Optional ByVal ToolControlBefore As Object)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������˵��͹�����
    '���:frmMain-���õ�������
    '     CommandBars-���ô����CommandBars����
    '     blnAddInTool-�Ƿ񽫹��ܰ�ť�����ӽ�������
    '     MenuControlBefore-�˵�����Ҫ���ӹ��ܶ���ĺ�һ��Control�������Կ������ӵĹ��ܶ����λ�ã�����ʱ������Ĭ�Ϲ���ָ����ťλ�á�
    '     ToolControlBefore-��������Ҫ���ӹ��ܶ���ĺ�һ��Control�������Կ������ӵĹ��ܶ����λ�ã�����ʱ������Ĭ�Ϲ���ָ����ťλ�á�
    '����:������
    '����:2015-04-22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Call mfrmRegist.zlDefCommandBars(frmMain, CommandBars, blnAddInTool, MenuControlBefore, ToolControlBefore)
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As Object)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ô���Ĳ˵��͹������ĸ���ܵ�Visible���Ի�Eanbeld����
    '���:Control-�˵��ؼ�
    '����:������
    '����:2015-04-22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Call mfrmRegist.zlUpdateCommandBars(Control)
End Sub

Public Sub zlExecuteCommandBars(ByVal frmMain As Object, ByVal Control As Object, ByRef strOutNO As String, Optional ByVal lngPatiID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ�йҺŵ���ع���
    '���:frmMain-���õ�������
    '     Control-�˵��ؼ�
    '����:strOutNo-�ɹ�����ĹҺŵ��ݺ�
    '����:���˺�
    '����:2015-04-22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Call mfrmRegist.zlExecuteCommandBars(frmMain, Control, strOutNO, lngPatiID)
End Sub

Public Function zlOpenStopedPlanBySN(ByVal frmMain As Object, ByVal lngModule As Long, _
    Optional ByVal lng��¼ID As Long, _
    Optional ByVal lngDeptID As Long, Optional ByVal lngDoctorID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '���ܣ�����������ſ��Ʒ�ʱ�ε���ͣ�ﰲ�Ű���ſ��ŹҺ�
    '��Σ�
    '   frmMain ���õ�������
    '   lngModule ����ģ���
    '   lng��¼ID ��¼ID,1114ģ�����ʱ����
    '   lngDeptID ����ID
    '   lngDoctorID ҽ��ID
    '���أ��ɹ�����True��ʧ�ܷ���False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim frmOpenStopedPlanBySN As New frmOpenStopedPlanBySN
    zlOpenStopedPlanBySN = frmOpenStopedPlanBySN.ShowMe(frmMain, lngModule, lng��¼ID, lngDeptID, lngDoctorID)
End Function
