VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPACSWork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public VERSION As String

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

Private mfrmPacsMain As frmPacsMain
Private mfrmVideoMain As frmPacsMain
Private mfrmPatholMain As frmPacsMain

Public Sub CodeMan(lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, frmMain As Object, ByVal strDBUser As String)
'���ܣ� ����������ָ�����ܣ�����ִ����س���
'������
'   lngModul:��Ҫִ�еĹ������
'   cnMain:����������ݿ�����
'   frmMain:������
'   strDBUser:��ǰ��¼���ݿ��û���

    '------------------------------------------------
On Error GoTo errHandle
    
    Set gcnOracle = cnMain
    Set gfrmMain = frmMain
    gstrDBUser = strDBUser
    glngModul = lngModul
    glngSys = lngSys

    '��ȡע������ж��Ƿ����õ�����־
    gblnUseDebugLog = IIf(Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\frmPacsMain", "UseDebugLog", 0)) = 0, False, True)
    Call SaveSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\frmPacsMain", "UseDebugLog", IIf(gblnUseDebugLog, 1, 0))
    
    gstrPrivs = ";" & GetPrivFunc(lngSys, lngModul) & ";"
    If gstrPrivs = "" Then MsgBox "�㲻�߱�ʹ�ø�ģ���Ȩ�ޣ�", vbExclamation, gstrSysName: Exit Sub
    
    
    If Not InitSysPar Then Exit Sub '��ʼ��ϵͳ����
    
    Select Case lngModul
        Case 1056 'Ӱ������Ŀ
            frmRadLists.Show , gfrmMain
        Case 1057 'Ӱ���豸Ŀ¼
            frmPACSDevice.Show , gfrmMain
        Case 1079
            frmPatholSpecimenCfg.Show , gfrmMain
        Case 1290 'Ӱ��ʦ����վ
            Call gobjKernel.InitCISKernel(gcnOracle, frmMain, glngSys, gstrPrivs) '��ʼ��ҽ�����������Ĳ���
            Call gobjRichEPR.InitRichEPR(gcnOracle, frmMain, glngSys, False)

            If mfrmPacsMain Is Nothing Then
                Set mfrmPacsMain = New frmPacsMain
            End If
            
            mfrmPacsMain.Caption = "Ӱ��ҽ������վ"
            mfrmPacsMain.ShowStation lngModul, gfrmMain
        
        Case 1291   'Ӱ��ɼ�����վ
            
            Call gobjKernel.InitCISKernel(gcnOracle, frmMain, glngSys, gstrPrivs) '��ʼ��ҽ�����������Ĳ���
            Call gobjRichEPR.InitRichEPR(gcnOracle, frmMain, glngSys, False)
            
            If mfrmVideoMain Is Nothing Then
                Set mfrmVideoMain = New frmPacsMain
            End If
            
            mfrmVideoMain.Caption = "Ӱ��ɼ�����վ"
            mfrmVideoMain.ShowStation lngModul, gfrmMain
            
        Case 1292   'Ӱ�����̿���
            frmWorkFlow.Show , gfrmMain
        Case 1294   'Ӱ��������վ
            Call gobjKernel.InitCISKernel(gcnOracle, frmMain, glngSys, gstrPrivs) '��ʼ��ҽ�����������Ĳ���
            Call gobjRichEPR.InitRichEPR(gcnOracle, frmMain, glngSys, False)
            
            If mfrmPatholMain Is Nothing Then
                Set mfrmPatholMain = New frmPacsMain
            End If
            
            mfrmPatholMain.Caption = "������鹤��վ"
            mfrmPatholMain.ShowStation lngModul, gfrmMain
            
        Case 1295   '�����鵵����վ
            frmPatholArchivesStation.Show , gfrmMain
        Case 1296   '�����軹����վ
            frmPatholReborrowStation.Show , gfrmMain
        Case 1297   '����������ʧ��������վ
            frmPatholLoseStation.Show , gfrmMain
    End Select

Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Public Sub BHCodeMan(lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, lngMain As Long, ByVal strDBUser As String)
'���ܣ� ����������ָ�����ܣ�����ִ����س���
'������
'   lngModul:��Ҫִ�еĹ������
'   cnMain:����������ݿ�����
'   lngMain:��������
'   strDBUser:��ǰ��¼���ݿ��û���

    '------------------------------------------------
On Error GoTo errHandle
    Set gcnOracle = cnMain
    Set gfrmMain = Nothing
    gstrDBUser = strDBUser
    glngModul = lngModul
    glngSys = lngSys

    gstrPrivs = ";" & GetPrivFunc(lngSys, lngModul) & ";"
    If gstrPrivs = "" Then MsgBox "�㲻�߱�ʹ�ø�ģ���Ȩ�ޣ�", vbExclamation, gstrSysName: Exit Sub

    If Not InitSysPar Then Exit Sub '��ʼ��ϵͳ����

    Select Case lngModul
        Case 1056 'Ӱ������Ŀ
            zlCommFun.ShowChildWindow frmRadLists.hWnd, lngMain
        Case 1057 'Ӱ���豸Ŀ¼
            zlCommFun.ShowChildWindow frmPACSDevice.hWnd, lngMain
        Case 1057 'Ӱ���豸Ŀ¼
            zlCommFun.ShowChildWindow frmPACSDevice.hWnd, lngMain
        Case 1290 'Ӱ��ʦ����վ
            Call gobjKernel.InitCISKernel(gcnOracle, Nothing, glngSys, gstrPrivs) '��ʼ��ҽ�����������Ĳ���
            Call gobjRichEPR.InitRichEPR(gcnOracle, Nothing, glngSys, False)

            If mfrmPacsMain Is Nothing Then
                Set mfrmPacsMain = New frmPacsMain
                mfrmPacsMain.Caption = "Ӱ��ҽ������վ"
            End If
            
            zlCommFun.ShowChildWindow mfrmPacsMain.hWnd, lngMain, True
            mfrmPacsMain.ShowStation lngModul, gfrmMain     '��ʱgfrmMainΪnothing
            
        Case 1291   'Ӱ��ɼ�����վ
            Call gobjKernel.InitCISKernel(gcnOracle, Nothing, glngSys, gstrPrivs) '��ʼ��ҽ�����������Ĳ���
            Call gobjRichEPR.InitRichEPR(gcnOracle, Nothing, glngSys, False)

            If mfrmVideoMain Is Nothing Then
                Set mfrmVideoMain = New frmPacsMain
                mfrmVideoMain.Caption = "Ӱ��ɼ�����վ"
            End If
            
            zlCommFun.ShowChildWindow mfrmVideoMain.hWnd, lngMain, True
            mfrmVideoMain.ShowStation lngModul, gfrmMain    '��ʱgfrmMainΪnothing
            
        Case 1292   'Ӱ�����̿���
            zlCommFun.ShowChildWindow frmWorkFlow.hWnd, lngMain
        Case 1294   'Ӱ��������վ
            Call gobjKernel.InitCISKernel(gcnOracle, Nothing, glngSys, gstrPrivs) '��ʼ��ҽ�����������Ĳ���
            Call gobjRichEPR.InitRichEPR(gcnOracle, Nothing, glngSys, False)

            If mfrmPatholMain Is Nothing Then
                Set mfrmPatholMain = New frmPacsMain
                mfrmPatholMain.Caption = "������鹤��վ"
            End If
            
            zlCommFun.ShowChildWindow mfrmPatholMain.hWnd, lngMain, True
            mfrmPatholMain.ShowStation lngModul, gfrmMain   '��ʱgfrmMainΪnothing
            
        Case 1295   '�����鵵����վ
            zlCommFun.ShowChildWindow frmPatholArchivesStation.hWnd, lngMain, True
        Case 1296   '�����軹����վ
            zlCommFun.ShowChildWindow frmPatholReborrowStation.hWnd, lngMain, True
        Case 1297   '����������ʧ��������վ
            zlCommFun.ShowChildWindow frmPatholLoseStation.hWnd, lngMain, True
    End Select
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Public Function CloseWindows() As Boolean
    '--------------------------------------
    '����:�ر������Ӵ���
    '--------------------------------------
On Error GoTo errHandle
    
    Dim frmThis As Form
    
    For Each frmThis In Forms
        UnLoadForm frmThis
    Next
    
    CloseWindows = (Forms.Count = 0)

errHandle:
End Function

Private Sub UnLoadForm(objForm As Object)
On Error GoTo errHandle
    Unload objForm
Exit Sub
errHandle:
    Debug.Print err.Description
End Sub

Private Sub Class_Initialize()
    VERSION = App.major & "." & App.minor & "." & App.Revision
End Sub

Private Sub Class_Terminate()
On Error GoTo errHandle
    Unload frmPacsMain
    Set frmPacsMain = Nothing
    
errHandle:
End Sub