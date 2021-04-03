VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPacsCapture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


'Implements IWorkMenu




Private Const M_STR_PARAMETER_VIDEO_DOCK As String = "VideoDock"


Private mobjActiveMenuBar As CommandBars        '�˵�����

Private WithEvents mobjVideoWindow As frmWork_Video '��Ƶ�ɼ����ڶ���
Attribute mobjVideoWindow.VB_VarHelpID = -1
Private WithEvents mobjPane As Form                 '��Ƶ��������
Attribute mobjPane.VB_VarHelpID = -1
Private WithEvents mobjDock As Form                 '�������ڶ���
Attribute mobjDock.VB_VarHelpID = -1

Private WithEvents mobjNotifyEvent As clsNotifyEvent
Attribute mobjNotifyEvent.VB_VarHelpID = -1

Private mRegEvent As ICaptureEvent

Private mstrParentWindowKey As String
Private mlngCurPicHeight As Long
Private mblnHasVideo As Boolean             '�жϵ�ǰ�����Ƿ�ӵ����Ƶ�ɼ�ģ��
Private mblnInitFinishState As Boolean
Private mblnAllowEventNotify As Boolean
Private mblnRefreshState As Boolean
Private mobjEvent As Object

Public Event OnStateChange(ByVal lngEventType As TVideoEventType, ByVal lngAdviceId As Long, ByVal lngSendNo As Long, ByVal strOther As String)
Public Event OnDockClose()

'����������
Property Get ContainerHwnd() As Long
    ContainerHwnd = 0
    
    If Not mobjPane Is Nothing Then
        ContainerHwnd = mobjPane.hWnd
    End If
End Property

'��������
Property Get ContainerObj() As Object
    Set ContainerObj = mobjPane
End Property


'�ɼ����ھ��
Property Get VideoHwnd() As Long
    VideoHwnd = mobjVideoWindow.hWnd
End Property


'�ɼ����ڶ���
Property Get VideoWindow() As Object
    Set VideoWindow = mobjVideoWindow
End Property

'�����¼�֪ͨ
Property Get AllowEventNotify() As Boolean
    AllowEventNotify = mblnAllowEventNotify
End Property


Property Let AllowEventNotify(value As Boolean)
    mblnAllowEventNotify = value
End Property


'��ȡ��ǰ����ID
Property Get ProcessId() As Long
    ProcessId = GetCurrentProcessId
End Property

Property Get HasVideo() As Boolean
'�жϸö����Ƿ�ӵ����Ƶ�ɼ�
    HasVideo = mblnHasVideo
End Property

''��ȡ�˵��ӿڶ���
'Property Get zlMenu() As IWorkMenu
'    Set zlMenu = Me
'End Property


'����״̬
Property Get VideoDockState() As Boolean
    VideoDockState = gblnDockingState
End Property


'����ʹ�òɼ�ģ��Ĵ��ڹؼ���
Property Get ParentWindowKey() As String
    ParentWindowKey = mstrParentWindowKey
End Property

Property Let ParentWindowKey(ByVal value As String)
    mstrParentWindowKey = value
End Property



Private Sub DoStateChange(ByVal lngEventType As TVideoEventType, ByVal lngAdviceId As Long, ByVal lngSendNo As Long, ByVal strOther As String)
'ִ��ע����¼�����
On Error GoTo errHandle
BUGEX_ "DoStateChange(clsPacsCapture) 1"
    If Not mobjEvent Is Nothing Then
BUGEX_ "DoStateChange(clsPacsCapture) 2"
        Call mobjEvent.OnStateChange(lngEventType, lngAdviceId, lngSendNo, strOther)
    ElseIf Not mRegEvent Is Nothing Then
BUGEX_ "DoStateChange(clsPacsCapture) 3"
        Call mRegEvent.OnVideoEvent(lngEventType, lngAdviceId, lngSendNo, strOther)
    Else
BUGEX_ "DoStateChange(clsPacsCapture) 4"
        RaiseEvent OnStateChange(lngEventType, lngAdviceId, lngSendNo, strOther)
    End If
BUGEX_ "DoStateChange(clsPacsCapture) End"
Exit Sub
errHandle:
    BUGEX_ "DoStateChange(clsPacsCapture) Err:" & err.Description
End Sub

Private Sub DoDockClose()
On Error GoTo errHandle
BUGEX_ "DoDockClose(clsPacsCapture) 1"
    If Not mobjEvent Is Nothing Then
BUGEX_ "DoDockClose(clsPacsCapture) 2"
        Call mobjEvent.OnDockClose
    ElseIf Not mRegEvent Is Nothing Then
BUGEX_ "DoDockClose(clsPacsCapture) 3"
    Else
BUGEX_ "DoDockClose(clsPacsCapture) 4"
        RaiseEvent OnDockClose
    End If
BUGEX_ "DoDockClose(clsPacsCapture) End"
Exit Sub
errHandle:
    BUGEX_ "DoDockClose(clsPacsCapture) Err:" & err.Description
End Sub


'******************************************************************************************************************************
'������ýӿ�
'******************************************************************************************************************************


'��ʾ������Ƶ����
Public Sub zlShowPopupVideo()
    Dim lngPicHeight As Long
    Dim lngWindowStyle As Long
    
    '��������ɼ����򿪣��������ٴδ򿪲ɼ�����
    If gblnDockingState = dsOpen Then Exit Sub
    
    '�򿪸����ɼ�����ǰ��֪ͨ����ʵ�������Ӧ������״̬
    Call mobjNotifyEvent.NotifyContainerLoad(mobjDock)

    Set mobjDock = New frmVideoDockWindow
    
'    If mobjVideoWindow.LockState = True Then
'        mobjDock.Caption = "��Ƶ�ɼ� [" & mobjVideoWindow.LockPatientName & "]"
'    End If

    '������ڵ���״̬������Ҫ����������ʾ״̬
    If gblnOpenDebug Then
        lngWindowStyle = GetWindowLong(mobjDock.hWnd, -20)
        lngWindowStyle = lngWindowStyle Or WS_EX_APPWINDOW
    
        SetWindowLong mobjDock.hWnd, -20, lngWindowStyle
    End If
    
    mobjDock.Caption = mobjVideoWindow.CaptionEx
    Call mobjDock.Show(0)
        
    lngPicHeight = Val(LoadVideoAreaCfg(M_STR_PARAMETER_VIDEO_DOCK))
    
    Call UpdateVideoContainer(mobjDock, lngPicHeight, True)
    
'    Call SetForegroundWindow(mobjDock.hwnd)
    Call mobjVideoWindow.ucSplitter1.RePaint(False)
    
    '�򿪺���Ҫˢ���ػ���Ƶ��ʾ����
    Call mobjDock.Refresh
    
    gblnDockingState = dsOpen
    
    mblnHasVideo = True
End Sub


'�����ڽ���ActiveWindow״̬ʱ����Ҫ���ô˷���
Public Sub zlRefreshVideoWindow()
On Error GoTo errHandle
    '��������ɼ����򿪣����л���Ƶ���ڵ���ʾ��ˢ��
BUGEX_ "zlRefreshVideoWindow 1"
    If gblnDockingState = dsOpen Then
    
        'ע��򿪵��Ⱥ�˳�򣬵��رո�������ʱ����������ȷ�Ĵ����м���
        Call RegVideoWindowOrder
    
        Exit Sub
    End If
    
    If mblnRefreshState = True Then Exit Sub
    
    mblnRefreshState = True
    
    '��Ҫ�����øñ���ֵΪfalse�������ظ�ˢ�µ�ǰ����ʱ���ظ�����mobjNotifyEvent_OnContainerLoad�¼�������
    mblnHasVideo = False
    
BUGEX_ "zlRefreshVideoWindow 2"
    '֪ͨ������ʵ��������picCapture�߶ȵȲ���
    Call mobjNotifyEvent.NotifyContainerLoad(mobjPane)
    
BUGEX_ "zlRefreshVideoWindow 3"
    '��ȡmobjVideoWindow.picCapture�߶�����
    If mstrParentWindowKey <> "" And mlngCurPicHeight <= 0 Then
        '��һ�μ��ش���ʱ����Ҫ��ȡ�ò���ֵ
        mlngCurPicHeight = Val(LoadVideoAreaCfg(mstrParentWindowKey))
    End If
        
BUGEX_ "zlRefreshVideoWindow 4"
    Call UpdateVideoContainer(mobjPane, mlngCurPicHeight)
    
BUGEX_ "zlRefreshVideoWindow 5 mlngCurPicHeight:" & mlngCurPicHeight & " picCaptureHeight:" & mobjVideoWindow.picCapture.Height
    Call mobjVideoWindow.ucSplitter1.RePaint(False)
    
    '�ڵ���Refresh����ʱ���ᴥ��mobjPane��Paint�¼��������Ҫ�Ƚ�mblnHasVideo����Ϊtrue
    mblnHasVideo = True
    
BUGEX_ "zlRefreshVideoWindow 6"
    '�򿪺���Ҫˢ���ػ���Ƶ��ʾ����
'    Call mobjPane.Refresh
    
BUGEX_ "zlRefreshVideoWindow 7"
    '�����ڲ���ZORDER����
    Call RegVideoWindowOrder
        
    mblnRefreshState = False
    
BUGEX_ "zlRefreshVideoWindow End"

Exit Sub
errHandle:
    mblnRefreshState = False
    err.Raise err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext
End Sub

'��֤�����Ƿ���Ч
Public Function IsValid() As Boolean
'��ActiveExe���̱������˳������з�����������
    IsValid = True
End Function


Public Sub zlInitModule(objDbAccess As Object, _
                        ByVal lngSysNo As Long, _
                        ByVal lngCallModule As Long, _
                        ByVal strPrivs As String, _
                        ByVal lngCurDepartId As Long, _
                        ByVal lngOwnerHwnd As Long, _
                        Optional objOwner As Object = Nothing, _
                        Optional blnForceRefresh As Boolean = False, _
                        Optional blnOpenDebugMode As Boolean = False)
'��ʼ��ģ�����
    Dim blnIsEqualDB As Boolean

BUGEX_ "zlInitModule 1", True
    If objDbAccess Is Nothing Then Exit Sub
    
'BUGEX_ "zlInitModule 2", True
'    '�ж����ݿ������Ƿ���ͬ
'    blnIsEqualDB = False
'    If Not gcnVideoOracle Is Nothing Then
'        blnIsEqualDB = IIf(gcnVideoOracle.ConnectionString = cnOracle.ConnectionString, True, False)
'    End If
'
'BUGEX_ "zlInitModule 3", True
'    If Not blnIsEqualDB Then
'BUGEX_ "zlInitModule 4", True
'        Set gcnVideoOracle = Nothing
'
'        '�����ݿ����Ӹı�ʱ�����´�������
'        Set gcnVideoOracle = New ADODB.Connection
'        'ע����������ActiveExeΪ�����Ľ�����Ŀ����˲���ʹ��cnOracleֱ�Ӷ�gcnVideoOracle����ֵ������������������Ͳ���ȷ,XXX���Ĵ���
'        gcnVideoOracle.ConnectionString = cnOracle.ConnectionString
'
'        '�����ݿ�����
'        gcnVideoOracle.Open
'    End If
    

        
BUGEX_ "zlInitModule 5", True
    '�����������
    Set gobjOwner = objOwner
    
    glngSys = lngSysNo
    glngModule = lngCallModule
    gstrPrivs = strPrivs
    glngDepartId = lngCurDepartId
    gblnOpenDebug = IIf(blnOpenDebugMode = False, gblnOpenDebug, blnOpenDebugMode)
        
BUGEX_ "zlInitModule6:Connect DB.", True
    '��ʼ���û��ؼ������ݿ�����
    Call InitCommonLib(objDbAccess)
    

    
BUGEX_ "zlInitModule 7", True
    glngRootHandle = lngOwnerHwnd
    If lngOwnerHwnd <= 0 Then
        glngRootHandle = GetAncestor(mobjPane.hWnd, GA_ROOT)
    End If
BUGEX_ "zlInitModule 8", True
    Call gobjCapturePar.ReadParameter
    
BUGEX_ "zlInitModule 9", True
    If blnForceRefresh And Not VideoIsNull Then
BUGEX_ "zlInitModule 10", True
        Call mobjVideoWindow.zlInitModule
    End If
    
    mblnInitFinishState = True
    
BUGEX_ "zlInitModule End", True
End Sub


'Private Sub ConfigDeptName(ByVal lngDeptId As String)
''���ÿ�������
'    Dim strSQL As String
'    Dim rsData As ADODB.Recordset
'
'    strSQL = "select ���� from ���ű� where id=[1]"
'    Set rsData = zlCL_GetDBObj.OpenSQLRecord(strSQL, "��ȡ��������", lngDeptId)
'
'    If rsData.RecordCount <= 0 Then Exit Sub
'
'    gstrDeptName = Nvl(rsData!����)
'End Sub


Public Sub zlUpdateStudyInf(ByVal lngAdviceId As Long, _
                            ByVal lngSendNo As Long, _
                            ByVal lngStudyState As Long, _
                            ByVal blnMoved As Boolean)
'���²ɼ�ʱ����ļ����Ϣ
    If VideoIsNull = True Then Exit Sub
    
    Call mobjVideoWindow.zlUpdateAdviceInf(lngAdviceId, lngSendNo, lngStudyState, blnMoved)

    If gblnDockingState = dsOpen Then
        If mobjDock Is Nothing Then Exit Sub
        mobjDock.Caption = mobjVideoWindow.CaptionEx
    End If
End Sub


Public Sub zlCaptureImg()
'ͼ��ɼ�
    If VideoIsNull = True Then Exit Sub
    
    Call mobjVideoWindow.subCaptureImg(True)
End Sub

Public Sub zlCaptureAfterImg()
'��̨�ɼ�
    If VideoIsNull = True Then Exit Sub
    If Not gobjCapturePar.IsUseAfterCapture Then Exit Sub
    
    Call mobjVideoWindow.subCaptureImg(True, "", Nothing, True)
End Sub

Public Sub zlUpdateAfterCaptureInfo()
'���±��
    If VideoIsNull = True Then Exit Sub
    If Not gobjCapturePar.IsUseAfterCapture Then Exit Sub
    
    Call mobjVideoWindow.UpdateAfterCaptureInfo
End Sub

Public Sub zlRefreshData(Optional blnForceRefresh As Boolean = False)
'ˢ��ͼ��ɼ�����
BUGEX_ "zlRefreshData 1"
    If VideoIsNull = True Then Exit Sub
    
BUGEX_ "zlRefreshData 2"
    Call mobjVideoWindow.zlRefreshFace(blnForceRefresh)
    
BUGEX_ "zlRefreshData 3"
End Sub


Public Sub zlNotifyRefresh()
'֪ͨ��������ˢ��
    If VideoIsNull = True Then Exit Sub
    
    Call mobjVideoWindow.NotificationRefresh
End Sub

Public Sub zlNotifyQuit()
'֪ͨ�˳�Active��Ƶģ��
    gblnIsQuitModule = True
End Sub


Public Sub zlShowVideoConfig()
'��ʾ��Ƶ���ô���

BUGEX_ "zlShowVideoConfig 1"
    If VideoIsNull = True Then Exit Sub
    
'    If gblnDockingState = dsOpen Then Exit Sub
    
BUGEX_ "zlShowVideoConfig 2"
    Call mobjVideoWindow.ShowVideoConfig
BUGEX_ "zlShowVideoConfig End"
End Sub



'******************************************************************************************************************************
'�ڲ�˽�д�������
'******************************************************************************************************************************

Private Function VideoIsNull() As Boolean
    VideoIsNull = IIf(gobjVideo Is Nothing, True, False)
End Function


Private Sub RegVideoWindowOrder()
'ע����Ƶ���ڴ�˳��
    If Not gobjZOrder Is Nothing Then
        Call DelCurZOrderInf
        Call gobjZOrder.Add(mobjPane.hWnd, mobjPane.hWnd)
        
'        Debug.Print "RegHwnd:" & mobjPane.hwnd
    End If
End Sub

Private Sub CreateGlobalResource()
'��ʼ����Ƶ�ɼ�����
    
BUGEX_ "CreateGlobalResource 1", True
    If glngInstanceCount <= 0 Then

        gblnOpenDebug = GetSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "IsOpenDebug", 0)
        Call SaveSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "IsOpenDebug", IIf(gblnOpenDebug, 1, 0))
    
        If gblnOpenDebug Then
BUGEX_ "CreateResource Open Debug Successed", True
        Else
BUGEX_ "CreateResource Open Debug Failed", True
        End If
    End If
    
    '����������Դ
    If gobjCapturePar Is Nothing Then
BUGEX_ "CreateGlobalResource gobjCapturePar", True
        Set gobjCapturePar = New clsCaptureParameter
    End If
    
BUGEX_ "CreateGlobalResource 2", True
    If gobjZOrder Is Nothing Then
BUGEX_ "CreateGlobalResource gobjZOrder", True
        Set gobjZOrder = New Scripting.Dictionary
    End If
    

BUGEX_ "CreateGlobalResource 3", True
    If gobjVideo Is Nothing Then
BUGEX_ "CreateGlobalResource gobjVideo", True
        Set gobjVideo = New frmWork_Video
    End If
    
BUGEX_ "CreateGlobalResource 4", True
    If gobjNotifyEvent Is Nothing Then
BUGEX_ "CreateGlobalResource gobjNotifyEvent", True
        Set gobjNotifyEvent = New clsNotifyEvent
    End If
    
BUGEX_ "CreateGlobalResource End", True

End Sub



Private Sub Class_Initialize()
On Error GoTo errHandle
BUGEX_ "Class_Initialize 1", True
    '��ʼ��ģ�����
    mstrParentWindowKey = ""
    mlngCurPicHeight = 0
    mblnHasVideo = False
    mblnInitFinishState = False
    gblnIsQuitModule = False
    mblnAllowEventNotify = False
    
BUGEX_ "Class_Initialize 2"
    
    '����������Դ
    Call CreateGlobalResource
    
BUGEX_ "Class_Initialize 3"
    '��������Դ
    If mobjPane Is Nothing Then
        BUGEX_ "CreateResource mobjPane", True
        Set mobjPane = New frmVideoPaneWindow
    End If
    
BUGEX_ "Class_Initialize 4"
    '��������������
    Set mobjVideoWindow = gobjVideo
    Set mobjNotifyEvent = gobjNotifyEvent
    
    glngInstanceCount = glngInstanceCount + 1
    
BUGEX_ "Class_Initialize End", True
Exit Sub
errHandle:
    BUGEX_ "Class_Initialize Err:" & err.Description, True, True
End Sub


Public Sub FreeGlobalResource()
'�ͷ���Ƶ�ɼ���Դ
        
BUGEX_ "FreeGlobalSource 1", True
        '�ͷ��¼�֪ͨ����
        If Not gobjNotifyEvent Is Nothing Then
            Set gobjNotifyEvent = Nothing
        End If
                
BUGEX_ "FreeGlobalSource 2", True
        'ֻ��ʵ������Ϊһ��ʱ�򣬲��ͷ�ȫ�ֹ�����Դ
        If Not gobjVideo Is Nothing Then
            Unload gobjVideo
            Set gobjVideo = Nothing
        End If
        
BUGEX_ "FreeGlobalSource 3", True
        If Not gobjZOrder Is Nothing Then
            Call gobjZOrder.RemoveAll
            Set gobjZOrder = Nothing
        End If
        
BUGEX_ "FreeGlobalSource 4", True

        '���ò������ö���
        If Not gobjCapturePar Is Nothing Then
            Set gobjCapturePar = Nothing
        End If
        
BUGEX_ "FreeGlobalSource 5", True
        '�ͷ�oracle����
        If Not gcnVideoOracle Is Nothing Then
            Set gcnVideoOracle = Nothing
        End If
        
        '�ͷ�ͼ����Դ
        Unload frmPubIcons
        
BUGEX_ "FreeGlobalSource 6", True
        '�ͷ�oracle����
        If Not gobjComLib Is Nothing Then
            Call zlCL_CloseWindow
            Set gobjComLib = Nothing
        End If
        
    
BUGEX_ "FreeGlobalSource End", True
End Sub

Private Sub Class_Terminate()
'�ͷ���Դ
    Dim i As Integer
On Error GoTo errHandle

BUGEX_ "FreeSource 1", True
    Call DelCurZOrderInf
    
    '�ͷ�����Դ--------------------------------------
    
    '�ͷ��¼�����
    If Not mRegEvent Is Nothing Then
        Set mRegEvent = Nothing
    End If
    
    '�ͷŵ�������
    If Not mobjDock Is Nothing Then
        Unload mobjDock
        Set mobjDock = Nothing
    End If
    
BUGEX_ "FreeSource 2", True
    '�Ͽ���Ƶ�������������ڵ�����
    If glngCurVideoContainerHwnd = mobjPane.hWnd Then
        mobjVideoWindow.Visible = False
        SetParent mobjVideoWindow.hWnd, 0
    End If
    
    '�ͷ���������
    If Not mobjPane Is Nothing Then
        Unload mobjPane
        Set mobjPane = Nothing
    End If
    
    
    '�ͷ�ȫ����Դ--------------------------------------
    If glngInstanceCount = 1 Then Call FreeGlobalResource
    
    
    Set mobjNotifyEvent = Nothing
    Set mobjVideoWindow = Nothing
    
BUGEX "zlPacsCapture Class_Terminate"
    If gblnUseActivexLoad Then
        For i = 0 To UBound(gobjGetImage)
            If Not gobjGetImage(i) Is Nothing Then
                gobjGetImage(i).zlUnloadObj
                Set gobjGetImage(i) = Nothing
            End If
        Next
    End If

    glngInstanceCount = glngInstanceCount - 1
errHandle:
End Sub


'���µ�ǰ��Ƶ���ڴ�С
Private Sub UpdateVideoContainer(objPane As Object, Optional ByVal lngPicHeight As Long = 0, Optional ByVal blnIsDock As Boolean = False)
On Error GoTo errHandle
    Dim lngParentHwnd As Long
    Dim lngActiveHwnd As Long
    
    If VideoIsNull = True Then Exit Sub
    
    lngParentHwnd = GetAncestor(objPane.hWnd, GA_PARENT)
    
BUGEX_ "UpdateVideoContainer 1"

    If objPane Is Nothing Then Exit Sub
    If (lngParentHwnd <= 0 Or lngParentHwnd = GetDesktopWindow) And Not blnIsDock Then Exit Sub
    
BUGEX_ "UpdateVideoContainer 2"

    mobjVideoWindow.Visible = False
    SetParent mobjVideoWindow.hWnd, 0
BUGEX_ "UpdateVideoContainer 3 picCaptureHeight:" & mobjVideoWindow.picCapture.Height

    mobjVideoWindow.Left = 0
    mobjVideoWindow.Top = 0

    mobjVideoWindow.Height = objPane.Height
    mobjVideoWindow.Width = objPane.Width

    If blnIsDock Then
        mobjVideoWindow.Width = mobjVideoWindow.Width - 160
        mobjVideoWindow.Height = mobjVideoWindow.Height - 500
    End If

BUGEX_ "UpdateVideoContainer 4"
    SetParent mobjVideoWindow.hWnd, objPane.hWnd
    mobjVideoWindow.Visible = True
    
BUGEX_ "UpdateVideoContainer 5"
    '���mobjPane�趨��parent����ɽ�����ʾ
    If objPane.Tag <> "1" Then objPane.Visible = True
    
    '��Ҫ����������mobjVideoWindow.picCapture�ĸ߶ȣ�����ں���������ÿ������mobjVideoWindow.picCapture�߶Ȳ�����Ч
    mobjVideoWindow.picCapture.Height = IIf(lngPicHeight > 0, lngPicHeight, mobjVideoWindow.picCapture.Height)

    lngActiveHwnd = GetTopHwnd(objPane.hWnd)
'    Call SetFocusEx(GetTopHwnd(objPane.hwnd))
'    Call SetActiveWindow(lngActiveHwnd)
BUGEX_ "UpdateVideoContainer 6 ActiveHwnd:" & lngActiveHwnd
    
    glngCurVideoContainerHwnd = objPane.hWnd
        
BUGEX_ "UpdateVideoContainer End"
    Exit Sub
errHandle:
    BUGEX_ "UpdateVideoContainer Err:" & err.Description
End Sub


Private Sub mobjDock_Load()
    '�ָ�����״̬
BUGEX_ "mobjDock_Load 1"
    SetWindowPos mobjDock.hWnd, -1, mobjDock.CurrentX, mobjDock.CurrentY, mobjDock.ScaleWidth, mobjDock.ScaleHeight, 3 '�������ö�
    
BUGEX_ "mobjDock_Load 2"
    Call zlCL_RestoreWinState(mobjDock, App.ProductName)
    
BUGEX_ "mobjDock_Load End"
End Sub

Private Sub mobjDock_Resize()
On Error GoTo errHandle
    '�������ڴ�С�ı�ʱ����Ҫ������Ƶ�ɼ����ڴ�С
    '������ڽ�����С��ʱ����������Ƶ���ֵ���
    If Not mblnInitFinishState Then Exit Sub
    If mobjDock.WindowState = 1 Then Exit Sub
    
    Call UpdateVideoContainer(mobjDock, 0, True)
    
    Exit Sub
errHandle:
    BUGEX_ "mobjDock_Resize Err:" & err.Description
End Sub

Private Sub BUGEX_(ByVal strInf As String, Optional ByVal blnIsForce As Boolean = False, Optional ByVal blnRaiseErr As Boolean = False)
'���������Ϣ
    If mobjPane Is Nothing Then
        BUGEX "HWND:" & 0 & " " & strInf, blnIsForce
    Else
        BUGEX "HWND:" & mobjPane.hWnd & " " & strInf, blnIsForce
    End If
    
    '������������
    If blnRaiseErr Then
        err.Raise err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext
    End If
End Sub

Private Sub mobjNotifyEvent_OnContainerLoad(objOldPane As Object)
On Error GoTo errHandle
BUGEX_ "mobjNotifyEvent_OnContainerLoad 1"
    If VideoIsNull = True Then Exit Sub
    
BUGEX_ "mobjNotifyEvent_OnContainerLoad 2"
    '���������������д���Ƶ�ɼ�ҳ����ʾ֮ǰ����Ҫ�ȱ��浱ǰ�ɼ�ģ�������ͼ�߶�
    If mblnHasVideo And gblnDockingState = dsClosed Then
BUGEX_ "mobjNotifyEvent_OnContainerLoad 3"
        mlngCurPicHeight = mobjVideoWindow.picCapture.Height
        
        If mstrParentWindowKey <> "" Then
BUGEX_ "mobjNotifyEvent_OnContainerLoad 4 mlngCurPicHeight:" & mlngCurPicHeight
            Call SaveVideoAreaCfg(mstrParentWindowKey, mlngCurPicHeight)
        End If
        
        mblnHasVideo = False
    End If
    
BUGEX_ "mobjNotifyEvent_OnContainerLoad End"
Exit Sub
errHandle:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub

Private Sub mobjNotifyEvent_OnContainerUnLoad(objOldPane As Object)
    Dim lngParentHwnd As Long
    
    If VideoIsNull = True Then Exit Sub                       '����Ƶ������Чʱ�����л���Ƶ��ʾ
    If gblnDockingState = dsOpen Then Exit Sub          '�����ڸ���������ʾʱ�����л���Ƶ��ʾ
    If objOldPane.hWnd = mobjPane.hWnd Then Exit Sub    '���˳��Ĵ��ھ���뵱ǰ��������ͬʱ�����л���ʾ��Ƶ
    If mobjNotifyEvent.IsChangePane Then Exit Sub       '����Ƶ�Ѿ����л�ʱ������Ҫ���ظ��Ķ���Ƶ�����л�
    
    If gobjZOrder.Count >= 1 Then
BUGEX_ "mobjNotifyEvent_OnContainerUnLoad 1"
        '�����Ѿ����ص���Ƶ���ڻ��߶������ͷ�֮ǰ���������Ƴ���gobjZOrder�еľ����Ϣ��
        '���ֻ���ж��Ƴ���������ص���Ƶ�����Ƿ�Ϊ���б������һ�����ھ������
        If mobjPane.hWnd <> gobjZOrder.Keys(gobjZOrder.Count - 1) Then Exit Sub
BUGEX_ "mobjNotifyEvent_OnContainerUnLoad 2"
'    Else
'BUGEX_ "mobjNotifyEvent_OnContainerUnLoad 3"
'        '�жϵ�ǰ�ĸ����ھ���Ƿ�Ϊ֮ǰ��Ƶ���ڵ���һ�����ھ����
'        If GetTopHwnd(mobjPane.hWnd) <> glngNextVideoContainerHwnd Then Exit Sub
'
'        lngParentHwnd = GetAncestor(mobjPane.hWnd, GA_PARENT)
'
'        '��mobjpaneû�и�����ʱ����ֻ������clspacscapture�࣬��û��ʹ�ø����containerhwnd���Խ��ÿؼ����ӵ������У�
'        '��˵��رյ������ں󣬲���Ҫ����Ƶ�ڸ��ӵ�mobjpanel��
'        If lngParentHwnd <= 0 Or lngParentHwnd = GetDesktopWindow Then Exit Sub
'BUGEX_ "mobjNotifyEvent_OnContainerUnLoad 4"
    End If
    
    lngParentHwnd = GetAncestor(mobjPane.hWnd, GA_PARENT)
    
    '��mobjpaneû�и�����ʱ����ֻ������clspacscapture�࣬��û��ʹ�ø����containerhwnd���Խ��ÿؼ����ӵ������У�
    '��˵��رյ������ں󣬲���Ҫ����Ƶ�ڸ��ӵ�mobjpanel��
    If lngParentHwnd <= 0 Or lngParentHwnd = GetDesktopWindow Or mobjPane.Visible = False Then
        Call gobjZOrder.Remove(gobjZOrder.Keys(gobjZOrder.Count - 1))
        Exit Sub
    End If
        
    Call zlRefreshVideoWindow
    mobjNotifyEvent.IsChangePane = True
    
BUGEX_ "mobjNotifyEvent_OnContainerUnLoad End"
End Sub


Private Sub mobjDock_Unload(Cancel As Integer)
On Error GoTo errHandle
    Dim lngParentHwnd As Long
    
    gblnDockingState = dsClosing
    
BUGEX_ "mobjDock_Unload 1"

    If VideoIsNull = True Then Exit Sub
    
BUGEX_ "mobjDock_Unload 2"
    Call zlCL_SaveWinState(mobjDock, App.ProductName)
    Call SaveVideoAreaCfg(M_STR_PARAMETER_VIDEO_DOCK, mobjVideoWindow.picCapture.Height)
    
    mobjVideoWindow.Visible = False
    SetParent mobjVideoWindow.hWnd, 0

'    glngNextVideoContainerHwnd = GetTopHwnd(mobjPane.hWnd)
    
    lngParentHwnd = GetAncestor(mobjPane.hWnd, GA_PARENT)
    
BUGEX_ "mobjDock_Unload 3"
    '���رյ�������ʱ����ֻ������clsPacsCapture��Ψһʵ����ͬʱ��û�н�mobjPane�������ͨ��ContainerHwnd���Ը��ӵ����������У�
    '�������������£�����ֱ���ͷ���Ƶ�ɼ���Դ
    If glngInstanceCount <= 1 And (lngParentHwnd <= 0 Or lngParentHwnd = GetDesktopWindow) Then
BUGEX_ "mobjDock_Unload 4"
        Unload mobjVideoWindow
        Set mobjVideoWindow = Nothing
        
    Else
BUGEX_ "mobjDock_Unload 5"
        If gblnIsQuitModule = False Then
            '֪ͨ����ʵ��������Ƶģ����ʾ
            Call mobjNotifyEvent.NotifyContainerUnLoad(mobjDock)
        End If
    End If
    
BUGEX_ "mobjDock_Unload 6"
    Set mobjDock = Nothing
    
    gblnDockingState = dsClosed
    
    Call DoDockClose
    
BUGEX_ "mobjDock_Unload End"
Exit Sub
errHandle:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub


Private Sub mobjPane_Activate()
'On Error GoTo errHandle
'    Dim lngParentHwnd As Long
'
'BUGEX_ "mobjPane_Activate 1"
'    If VideoIsNull = True Then Exit Sub
'
'BUGEX_ "mobjPane_Activate 2"
'    lngParentHwnd = GetAncestor(mobjPane.hWnd, GA_PARENT)
'
'    '��������ɼ����򿪣����л���Ƶ���ڵ���ʾ��ˢ��
'    If Not mblnInitFinishState Or lngParentHwnd <= 0 Or lngParentHwnd = GetDesktopWindow Then Exit Sub
'    If gblnDockingState = dsOpen Then Exit Sub      '����������ڴ򿪣���ˢ����Ƶ����
'
'BUGEX_ "mobjPane_Activate 3"
'    If mblnHasVideo = True Then
'BUGEX_ "mobjPane_Activate 4"
'        '���������������Ѿ���������Ƶ����ֻ�����µ�����Ӧ�����С����
'        '�������ڴ�С�ı�ʱ����Ҫ������Ƶ�ɼ����ڴ�С
'        Call UpdateVideoContainer(mobjPane, mlngCurPicHeight)
''        mlngCurPicHeight = mobjVideoWindow.picCapture.Height
'    Else
'BUGEX_ "mobjPane_Activate 5"
'        Call zlRefreshVideoWindow
'    End If
'
'BUGEX_ "mobjPane_Activate End"
'    Exit Sub
'errHandle:
End Sub


Private Sub mobjPane_DblClick()
On Error GoTo errHandle
    Dim lngParentHwnd As Long
    
    If VideoIsNull = True Then Exit Sub
    
    lngParentHwnd = GetAncestor(mobjPane.hWnd, GA_PARENT)
    
    '����������ڴ򿪣��򲻻ָ���ǰ��������Ƶ��ʾ
    If Not mblnInitFinishState Or lngParentHwnd <= 0 Or lngParentHwnd = GetDesktopWindow Then Exit Sub
    If gblnDockingState = dsOpen Then
        Call RegVideoWindowOrder
        
        MsgboxCus "���������Ѵ򿪣������ڴ���ʾ��Ƶ��", vbOKOnly + vbInformation, G_STR_HINT_TITLE
        Exit Sub
    End If
    
    Call zlRefreshVideoWindow
Exit Sub
errHandle:
End Sub

Private Sub DrawBackGround()
'���Ʊ���ͼ��
    Dim i As Integer
    Dim Count As Integer
    Dim wordRect As RECT
    
    Count = 2
    wordRect.Bottom = 45
    wordRect.Right = 200

    If mobjPane.picBackImg.Height * 3 >= mobjPane.Height Then Count = 1

    Call mobjPane.Cls
    
    For i = 0 To Count
        Call mobjPane.PaintPicture(mobjPane.picBackImg.Picture, _
            Round(mobjPane.Width / (i + 1)) - mobjPane.picBackImg.Width + 200, _
            Round((mobjPane.Height / 3) * (i + 1) - mobjPane.picBackImg.Height), _
            mobjPane.picBackImg.Width, mobjPane.picBackImg.Height)

        If mobjVideoWindow Is Nothing Then
            wordRect.Left = mobjPane.ScaleX(Round(mobjPane.Width / (i + 1)) - mobjPane.picBackImg.Width, vbTwips, vbPixels) + 35
        Else
            wordRect.Left = mobjPane.ScaleX(Round(mobjPane.Width / (i + 1)) - mobjPane.picBackImg.Width, vbTwips, vbPixels) + 17
        End If
        
        wordRect.Top = mobjPane.ScaleY(Round((mobjPane.Height / 3) * (i + 1) - mobjPane.picBackImg.Height), vbTwips, vbPixels) - 30

        wordRect.Right = wordRect.Left + 200
        wordRect.Bottom = wordRect.Top + 90

        If mobjVideoWindow Is Nothing Then
            Call DrawText(mobjPane.hdc, "��Ƶδ��ע��" & vbCrLf & "�ѽ�����ƵԴ", 27, wordRect, 0)
        Else
            Call DrawText(mobjPane.hdc, "��Ƶ�ѱ���������ռ��" & vbCrLf & " ˫�����ɻָ���ʾ", 44, wordRect, 0)
        End If
    Next i
End Sub

Private Sub mobjPane_Paint()
On Error GoTo errHandle

BUGEX_ "mobjPane_Paint 1"
    
    If mobjPane.Visible = False Then Exit Sub
    
    '�����ǰ��mobjPaneû��Ƕ����Ƶ�ɼ������ػ汳��ͼ��
    If mblnHasVideo = False Or glngCurVideoContainerHwnd <> mobjPane.hWnd Then
BUGEX_ "mobjPane_Paint 2"
        Call DrawBackGround
    Else
BUGEX_ "mobjPane_Paint 3"
'        Call SetFocusEx(GetTopHwnd(mobjPane.hwnd))
'        Call SetActiveWindow(GetTopHwnd(mobjPane.hwnd))

        '��mobjPane_Raise�в�������Ƶ��ʾ������Ҫ������Ƶ�����С
'        If mobjVideoWindow.Width <> mobjPane.Width Or _
'            mobjVideoWindow.Height <> mobjPane.Height Then
'            Call UpdateVideoContainer(mobjPane, mlngCurPicHeight)
'        End If
        
        mobjVideoWindow.timerRePaint.Enabled = True
    End If
    
BUGEX_ "mobjPane_Paint End"

Exit Sub
errHandle:
    
BUGEX_ "mobjPane_Paint Err:" & err.Description
End Sub

Private Sub mobjPane_Resize()
On Error GoTo errHandle
    Dim lngParentHwnd As Long
    
BUGEX_ "mobjPane_Resize 1"
    If VideoIsNull = True Then Exit Sub
    
BUGEX_ "mobjPane_Resize 2"
    lngParentHwnd = GetAncestor(mobjPane.hWnd, GA_PARENT)
    
    '��������ɼ����򿪣����л���Ƶ���ڵ���ʾ��ˢ��
    If Not mblnInitFinishState Or lngParentHwnd <= 0 Or lngParentHwnd = GetDesktopWindow Then Exit Sub
'    If mobjPane.Visible = False Then Exit Sub       '�����ǰ����δ��ʾ����ˢ����Ƶ����
    If gblnDockingState = dsOpen Then Exit Sub      '����������ڴ򿪣���ˢ����Ƶ����

BUGEX_ "mobjPane_Resize 3"
    If mblnHasVideo = True Then
BUGEX_ "mobjPane_Resize 4"

        '����ǵ��������С����������Ƶҳ�汻���أ��������Сʱ�����Զ���ʾ��Ƶҳ��
        mobjPane.Tag = IIf(mobjPane.Visible = False, "1", "")
        
        '���������������Ѿ���������Ƶ����ֻ�����µ�����Ӧ�����С����
        '�������ڴ�С�ı�ʱ����Ҫ������Ƶ�ɼ����ڴ�С
        Call UpdateVideoContainer(mobjPane, mlngCurPicHeight)
'        mlngCurPicHeight = mobjVideoWindow.picCapture.Height
    Else
'BUGEX_ "mobjPane_Resize 5"
'        Call zlRefreshVideoWindow
    End If
    
BUGEX_ "mobjPane_Resize End"
    Exit Sub
errHandle:
    BUGEX_ "mobjPane_Resize Err:" & err.Description
End Sub

Private Sub DelCurZOrderInf()
'��gobjZOrder��ɾ����ǰ�������ڵľ����Ϣ
    If Not gobjZOrder Is Nothing Then
        If gobjZOrder.Exists(mobjPane.hWnd) Then
            Call gobjZOrder.Remove(mobjPane.hWnd)
'            Debug.Print "RemoveHwnd:" & mobjPane.hwnd & " ItemCount:" & gobjZOrder.Count
            
        End If
    End If
End Sub


Public Sub RegEventInterface(objEvent As ICaptureEvent)
'�����¼����ö���
    Set mRegEvent = objEvent
    
    Set mobjEvent = Nothing
End Sub

Public Sub RegEventObj(objEvent As Object)
'ע���¼�����
    Set mobjEvent = objEvent
    
    Set mRegEvent = Nothing
End Sub


Private Sub mobjPane_Unload(Cancel As Integer)
On Error GoTo errHandle
    If VideoIsNull = True Then Exit Sub
    
'    Debug.Print mobjPane.hwnd
    '�����ǰ�����ؼ��а�����Ƶ�ɼ����棬��������ж��ʱ����Ҫ֪ͨ�������������Ƶ,��֮����Ҫ֪ͨ��
    If glngCurVideoContainerHwnd <> mobjPane.hWnd Then Exit Sub
    If mblnHasVideo = False Then Exit Sub
    
    '��ж�ص�ǰ����ʱ�����������Ƶ���ڣ��򱣴浱ǰ���ڵ���Ƶ�߶�
    If mstrParentWindowKey <> "" Then
        Call SaveVideoAreaCfg(mstrParentWindowKey, mobjVideoWindow.picCapture.Height)
    End If
    
    mobjVideoWindow.Visible = False
    SetParent mobjVideoWindow.hWnd, 0
    
'    Debug.Print mobjPane.hwnd
    '��Ҫ�Ƴ���gobjZOrder�д��ڵľ����Ϣ
    Call DelCurZOrderInf
    
'    glngNextVideoContainerHwnd = GetNextWindowCus(GetTopHwnd(mobjPane.hWnd))

    'ֻ��clsPacsCapture��ʵ��Ψһʱ����ֱ����mobjpanel��unload�¼����ͷ���Ƶ�ɼ���Դ
    If glngInstanceCount <= 1 Then
        Unload mobjVideoWindow
        Set mobjVideoWindow = Nothing
    Else
        If gblnIsQuitModule = False Then
            '֪ͨ����ʵ��������Ƶģ����ʾ
            Call mobjNotifyEvent.NotifyContainerUnLoad(mobjPane)
        End If
    End If
    
Exit Sub
errHandle:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub


Private Sub mobjVideoWindow_OnControlResize(objControl As Object)
On Error GoTo errHandle
BUGEX_ "mobjVideoWindow_OnControlResize 1"

    If mblnHasVideo = False Then Exit Sub
    If gblnDockingState = dsOpen Then Exit Sub
    
    If objControl.Name = "picCapture" Then mlngCurPicHeight = objControl.Height
BUGEX_ "mobjVideoWindow_OnControlResize End"
Exit Sub
errHandle:
    BUGEX_ "mobjVideoWindow_OnControlResize Err:" & err.Description
End Sub


Private Sub mobjVideoWindow_OnImgLoadState(ByVal blnLoadFinish As Boolean, ByVal blnUpLoad As Boolean)
    If mobjEvent Is Nothing Then Exit Sub
BUGEX "mobjVideoWindow_OnImgLoadState"
    Call mobjEvent.OnImgLoadState(blnLoadFinish, blnUpLoad)
End Sub

Private Sub mobjVideoWindow_OnStateChange(ByVal lngEventType As TVideoEventType, ByVal lngAdviceId As Long, ByVal lngSendNo As Long, ByVal strOther As String)
On Error GoTo errHandle
    Dim lngCount As Long
BUGEX_ "mobjVideoWindow_OnStateChange 1"
    lngCount = GetTickCount
    
    If mblnAllowEventNotify = True Then
BUGEX_ "mobjVideoWindow_OnStateChange 2"

        Call DoStateChange(lngEventType, lngAdviceId, lngSendNo, strOther)
    End If
    
    '���¸������ڵ���ʾ����
    If gblnDockingState = dsOpen And Not mobjDock Is Nothing Then
        mobjDock.Caption = mobjVideoWindow.CaptionEx
        
'        If lngEventType = vetLockStudy Then
'            mobjDock.Caption = "��Ƶ�ɼ� [" & strOther & "]"
'        ElseIf lngEventType = vetUnLockStudy Then
'            mobjDock.Caption = "��Ƶ�ɼ�"
'        End If
    End If
    
BUGEX_ "mobjVideoWindow_OnStateChange Total Time:" & GetTickCount - lngCount
BUGEX_ "mobjVideoWindow_OnStateChange End"
Exit Sub
errHandle:
    BUGEX_ "mobjVideoWindow_OnStateChange Err:" & err.Description
End Sub













