VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDockInAdvices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Event Activate() '�Ӵ��弤��
Public Event RequestRefresh(ByVal RefreshNotify As Boolean) 'Ҫ��������ˢ��
Public Event StatusTextUpdate(ByVal Text As String) 'Ҫ�����������״̬������
Public Event ViewEPRReport(ByVal ����ID As Long, ByVal CanPrint As Boolean) 'Ҫ��鿴����
Public Event PrintEPRReport(ByVal ����ID As Long, ByVal Preview As Boolean) 'Ҫ���ӡ����
Public Event ViewPACSImage(ByVal ҽ��ID As Long) 'Ҫ����й�Ƭ
Public Event ExecLogNew(ByVal ҽ��ID As Long, ByVal ���ͺ� As Long, ByVal ����ID As Long, ��� As Boolean) 'ִ������Ǽ�
Public Event ExecLogModi(ByVal ҽ��ID As Long, ByVal ���ͺ� As Long, ByVal ����ID As Long, ByVal ִ��ʱ�� As String, ��� As Boolean) 'ִ������޸�
Public Event EditDiagnose(ParentForm As Object, ByVal ����ID As Long, ByVal ��ҳID As Long, ByVal ����ID As Long, ByVal str���� As String, Succeed As Boolean) '�༭סԺ���
Public Event SetEditState(ByVal blnEditState As Boolean)    '�༭״̬ʱ���ò˵��Ϳ�ת�ƽ���Ĺ���
Public Event DoByAdvice(ByVal lngҽ��ID As Long, ByVal lng���ID As Long, ByVal lngWayID As Long, ByVal strTag As String)

Private WithEvents mfrmInAdvice As frmDockInAdvice
Attribute mfrmInAdvice.VB_VarHelpID = -1
Private mblnSquareCardFromOut As Boolean    'ͨ��ҽ������վ�����һ��ͨ���㲿��(סԺҽ��վ�ڷ���ʱ�ٴ�������ʼ��)

Private Sub Class_Initialize()
    Set mfrmInAdvice = New frmDockInAdvice
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    Unload mfrmInAdvice
    Set grsҽ�Ƹ��ʽ = Nothing
    Set mfrmInAdvice = Nothing
    Set gobjPath = Nothing
    Set gobjPlugIn = Nothing
    
    If Not mblnSquareCardFromOut Then
        Set gobjSquareCard = Nothing
    End If
    err.Clear
End Sub

Public Sub UnloadClass()
    On Error Resume Next
    Unload mfrmInAdvice
    Set grsҽ�Ƹ��ʽ = Nothing
    Set mfrmInAdvice = Nothing
    Set gobjPath = Nothing
    Set gobjPlugIn = Nothing
    
    If Not mblnSquareCardFromOut Then
        Set gobjSquareCard = Nothing
    End If
    err.Clear
End Sub

Private Sub mfrmInAdvice_Activate()
    RaiseEvent Activate
End Sub

Private Sub mfrmInAdvice_DoByAdvice(ByVal lngҽ��ID As Long, ByVal lng���ID As Long, ByVal lngWayID As Long, ByVal strTag As String)
    RaiseEvent DoByAdvice(lngҽ��ID, lng���ID, lngWayID, strTag)
End Sub

Private Sub mfrmInAdvice_SetEditState(ByVal blnEditState As Boolean)
    RaiseEvent SetEditState(blnEditState)
End Sub

Private Sub mfrmInAdvice_RequestRefresh(ByVal RefreshNotify As Boolean)
    RaiseEvent RequestRefresh(RefreshNotify)
End Sub

Private Sub mfrmInAdvice_StatusTextUpdate(ByVal Text As String)
    RaiseEvent StatusTextUpdate(Text)
End Sub

Private Sub mfrmInAdvice_ViewEPRReport(ByVal ����ID As Long, ByVal CanPrint As Boolean)
    RaiseEvent ViewEPRReport(����ID, CanPrint)
End Sub

Private Sub mfrmInAdvice_PrintEPRReport(ByVal ����ID As Long, ByVal Preview As Boolean)
    RaiseEvent PrintEPRReport(����ID, Preview)
End Sub

Private Sub mfrmInAdvice_ViewPACSImage(ByVal ҽ��ID As Long)
    RaiseEvent ViewPACSImage(ҽ��ID)
End Sub

Private Sub mfrmInAdvice_ExecLogNew(ByVal ҽ��ID As Long, ByVal ���ͺ� As Long, ByVal ����ID As Long, ��� As Boolean)
    RaiseEvent ExecLogNew(ҽ��ID, ���ͺ�, ����ID, ���)
End Sub

Private Sub mfrmInAdvice_ExecLogModi(ByVal ҽ��ID As Long, ByVal ���ͺ� As Long, ByVal ����ID As Long, ByVal ִ��ʱ�� As String, ��� As Boolean)
    RaiseEvent ExecLogModi(ҽ��ID, ���ͺ�, ����ID, ִ��ʱ��, ���)
End Sub

Private Sub mfrmInAdvice_EditDiagnose(ParentForm As Object, ByVal ����ID As Long, ByVal ��ҳID As Long, ByVal ����ID As Long, ByVal str���� As String, Succeed As Boolean)
    RaiseEvent EditDiagnose(ParentForm, ����ID, ��ҳID, ����ID, str����, Succeed)
End Sub

Public Function zlGetForm() As Object
'���ܣ���ȡ��ҵ�񴰿�,���������ʼ��ʱֻ����һ�λ�ȡ
    '��ʼ��ҽ������:�ɷ�ɢ�������GetForm��ȥ��ʼ��
    Call gclsInsure.InitOracle(gcnOracle)

    '��ʼ��ȡȨ��,�Լӿ�����ٶ�(���Ǳ�Ҫ��)
    Call GetInsidePrivs(pסԺҽ���´�, True)
    Call GetInsidePrivs(pסԺҽ������, True)
    
    'һЩ�ӿڳ�ʼ��,ֻ��ʼ��һ��
    '���漰����ģ���Ȩ��,��������ʼ��
    
    '����ǩ����ʼ��
    '��Ϊ����ģ�����ͬʱʹ��,�ҹ��ñ���gobjESign����˻���Ҫ�ھ���ģ�����ж�Ȩ��
    If gintCA <> 0 Then
        'If InStr(GetInsidePrivs(pסԺҽ���´�), "ҽ������ǩ��") > 0 And gobjESign Is Nothing Then
        If gobjESign Is Nothing Then
            On Error Resume Next
            Set gobjESign = CreateObject("zl9ESign.clsESign")
            err.Clear: On Error GoTo 0
            If Not gobjESign Is Nothing Then
                Call gobjESign.Initialize(gcnOracle, glngSys)
            End If
        End If
    Else
        Set gobjESign = Nothing
    End If
    
    Set zlGetForm = mfrmInAdvice
End Function

Public Sub zlDefCommandBars(ByVal frmParent As Object, ByVal CommandBars As Object, ByVal int���� As Integer, _
    Optional ByVal blnInsideTools As Boolean = False, Optional ByRef objSquareCard As Object, Optional ByVal blnModalNew As Boolean)
'���ܣ������Ӵ���Ĳ˵��͹�����(����������Ҫʹ�õĲ˵��͹�����)
'������int����=0-ҽ��վ����,1-��ʿվ����,2-ҽ��վ����(PACS/LIS)
'      CommandBars=�����ڲ鿴ʱ���Բ���(����Nothing)
'      blnInsideTools=�Ƿ񽫰�ť��ʾ�������ڲ��������ϣ�����ʿվ����ʱ��Ч��
'      objSquareCard=ͨ��ҽ������վ�����һ��ͨ���㲿��(סԺҽ��վ�ڷ���ʱ�ٴ�������ʼ��)

    Call mfrmInAdvice.zlDefCommandBars(frmParent, CommandBars, int����, blnInsideTools, objSquareCard, blnModalNew)
    
    If Not objSquareCard Is Nothing Then
        mblnSquareCardFromOut = True
    Else
        mblnSquareCardFromOut = False
    End If
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As Object)
    Call mfrmInAdvice.zlUpdateCommandBars(Control)
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As Object)
    Call mfrmInAdvice.zlExecuteCommandBars(Control)
End Sub

Public Sub zlExecuteCommandBarsDirect(ByVal Control As CommandBarControl, ByRef frmParent As Object, ByRef strPrivs As String, _
    ByVal bln���� As Boolean, ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal bytӤ�� As Byte, _
    ByVal lng����ID As Long, ByVal lng����id As Long, ByVal lngǰ��ID As Long, ByVal lng�������ID As Long, ByVal int���� As Integer, _
    ParamArray arrPar() As Variant)
'���ܣ��ṩ��������ҽ�������Ľӿ�
    Call mfrmInAdvice.zlExecuteCommandBarsDirect(Control, frmParent, strPrivs, bln����, lng����ID, lng��ҳID, bytӤ��, _
        lng����ID, lng����id, lngǰ��ID, lng�������ID, int����, arrPar)
End Sub

Public Sub zlCheckPrivs(ByVal Control As CommandBarControl, ByVal int���� As Integer)
'���ܣ����˵���ť��Ȩ�ޣ���������ɼ���
    Call mfrmInAdvice.zlCheckPrivs(Control, int����)
End Sub

Public Sub zlPopupCommandBars(ByVal CommandBar As Object)
    Call mfrmInAdvice.zlPopupCommandBars(CommandBar)
End Sub

Public Sub zlRefresh(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal lng����ID As Long, ByVal lng����id As Long, _
    ByVal int״̬ As Integer, Optional ByVal blnMoved As Boolean, Optional ByVal lngǰ��ID As Long, _
    Optional ByVal intִ��״̬ As Integer, Optional ByVal lng�������ID As Long, Optional ByVal lng·��״̬ As Long = -1, _
    Optional ByVal lngҽ������ID As Long, Optional ByRef objMip As Object, Optional ByVal intӤ�� As Integer = -1)
'���ܣ�ˢ��סԺҽ������
'������int״̬=0-��Ժ����,1-Ԥ��Ժ����,2-��Ժ����,3-�����ﲡ��(��Ժ),4-�ѻ��ﲡ��,5-���ת�Ʋ���(��Ժ),6-ת�ƻ�ת��������ס����(������Ժ����ס)
'      lng����ID��lng����ID=����5-���ת�Ʋ��ˡ�ʱΪ����ԭ������ԭ����
'      lngǰ��ID=����ҽ��վ����ʱ����
'      intִ��״̬=����ҽ��վ����ʱ����,ҽ��ִ����Ŀ��ִ��״̬
'      lng�������ID=�����ǰҽ��վ�ǻ��ﲡ�ˣ���Ϊ�������ID�������ҽ��վ���ã���Ϊҽ������ID
'      lng·��״̬=-1:δ����,0-�����ϵ���������1-ִ���У�2-����������3-�������
'      blnMoved=�ò��˵������Ƿ���ת��
'      lngҽ������ID=ҽ��վ����ID
'      strǰ��IDsҽ��վ���ڵ�ǰ����ִ�е�����ҽ��
'      objMip ��Ϣ���� zl9ComLib.clsMipModule
    Call mfrmInAdvice.zlRefresh(lng����ID, lng��ҳID, lng����ID, lng����id, int״̬, blnMoved, lngǰ��ID, intִ��״̬, lng�������ID, lng·��״̬, lngҽ������ID, objMip, intӤ��)
End Sub

Public Sub LocatedAdviceRow(ByVal lngҽ��ID As Long)
'���ܣ����ݹ���վ�����ҽ��id��ѡ�����Ҷ�λҽ��
    Call mfrmInAdvice.LocatedAdviceRow(lngҽ��ID)
End Sub

Public Sub zlSeekAndViewEPRReport(ByVal lng����ID As Long)
'���ܣ���λ�������Ӧ��ҽ�������򿪱���鿴
    Call mfrmInAdvice.zlSeekAndViewEPRReport(lng����ID)
End Sub

Public Sub zlAdvicePrint(ByVal frmParent As Object, ByVal lng����ID As Long, ByVal lng��ҳID As Long)
'���ܣ���ӡҽ����
    Call frmAdvicePrint.ShowMe(frmParent, lng����ID, lng��ҳID)
End Sub

Public Sub zlInitPath(ByRef clsDockPath As Object)
'���ܣ���ʼ���ٴ�·�����󣬴���ҽ��վ��ʿվ���õĶ���
    Set gobjPath = clsDockPath
End Sub

Public Sub zlInitPlugIn(ByVal objPlugIn As Object)
'����:��������Ҳ������
'53740:������,2012-09-19
    Set gobjPlugIn = objPlugIn
End Sub

Public Sub zlInitMip(ByRef objMip As Object)
'����:������Ϣ����
'������objMip zl9ComLib.clsMipModule
    Call mfrmInAdvice.zlInitMip(objMip)
End Sub

Public Sub SetFontSize(ByVal bytSize As Byte)
'����:����ҽ���嵥�������С
'���:bytSize��0-С(ȱʡ)��1-��
    Call mfrmInAdvice.SetFontSize(bytSize)
End Sub

Public Sub zlCompoundUnpack(ByVal frmParent As Object, ByVal lng����ID As Long, ByVal lng����ID As Long, Optional ByVal lngҽ������ID As Long, _
    Optional ByVal lngӤ������ID As Long, Optional ByVal lngӤ������ID As Long)
'���ܣ��������
    frmCompoundPack.ShowMe 1, frmParent, lng����ID, lng����ID, lngҽ������ID, lngӤ������ID, lngӤ������ID
End Sub
