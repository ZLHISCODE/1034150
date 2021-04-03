VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDockPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Event ViewEPRReport(ByVal ����ID As Long, ByVal CanPrint As Boolean) 'Ҫ��鿴����
Public Event Activate() '�Ӵ��弤��
Public Event RequestRefresh(ByVal lngPathState As Long) 'Ҫ��������ˢ��,-1��ʾδ����״̬
Public Event StatusTextUpdate(ByVal Text As String) 'Ҫ�����������״̬������

Private WithEvents mclsDockPath As zlCISPath.clsDockPath
Attribute mclsDockPath.VB_VarHelpID = -1

Private Sub Class_Initialize()
    Set mclsDockPath = New zlCISPath.clsDockPath
End Sub

Private Sub Class_Terminate()
    Set mclsDockPath = Nothing
End Sub

Private Sub mclsDockPath_ViewEPRReport(ByVal ����ID As Long, ByVal CanPrint As Boolean)
    RaiseEvent ViewEPRReport(����ID, CanPrint)
End Sub

Private Sub mclsDockPath_Activate()
    RaiseEvent Activate
End Sub

Private Sub mclsDockPath_RequestRefresh(ByVal lngPathState As Long)
    RaiseEvent RequestRefresh(lngPathState)
End Sub

Private Sub mclsDockPath_StatusTextUpdate(ByVal Text As String)
    RaiseEvent StatusTextUpdate(Text)
End Sub

Public Function zlGetForm() As Object
'���ܣ���ȡ��ҵ�񴰿�,���������ʼ��ʱֻ����һ�λ�ȡ
    Set zlGetForm = mclsDockPath.zlGetForm
End Function

Public Sub zlDefCommandBars(ByVal frmParent As Object, ByVal CommandBars As Object, ByVal int���� As Integer, Optional ByVal blnInsideTools As Boolean = False)
'���ܣ������Ӵ���Ĳ˵��͹�����(����������Ҫʹ�õĲ˵��͹�����)
'������int����=0-ҽ��վ����,1-��ʿվ����,2-ҽ��վ����(PACS/LIS)
'      CommandBars=�����ڲ鿴ʱ���Բ���(����Nothing)
    Call mclsDockPath.zlDefCommandBars(frmParent, CommandBars, int����, blnInsideTools)
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As Object)
    Call mclsDockPath.zlUpdateCommandBars(Control)
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As Object)
    Call mclsDockPath.zlExecuteCommandBars(Control)
End Sub

Public Sub zlPopupCommandBars(ByVal CommandBar As Object)
    Call mclsDockPath.zlPopupCommandBars(CommandBar)
End Sub

Public Sub zlRefresh(ByVal lng����Id As Long, ByVal lng��ҳID As Long, ByVal lng����ID As Long, ByVal lng����ID As Long, ByVal int����״̬ As Integer, Optional ByVal blnMoved As Boolean, _
     Optional ByVal blnForceRefresh As Boolean, Optional ByVal lngState As Long, Optional ByVal lngҽ������ID As Long, Optional ByRef objMip As Object)
'���ܣ�ˢ��·��������
'������int��������=������ҳ�е�״̬��0-����סԺ��1-��δ��ƣ�2-����ת�ƣ�3-��Ԥ��Ժ
'      blnMoved=�ò��˵������Ƿ���ת��
'      blnForceRefresh=ǿ��ˢ��·����
'      lngState =���˵�״̬
'      lngҽ������ID=ҽ��վ����ID
'      objMip ��Ϣ����
    Call mclsDockPath.zlRefresh(lng����Id, lng��ҳID, lng����ID, lng����ID, int����״̬, blnMoved, blnForceRefresh, lngState, lngҽ������ID, objMip)
End Sub

Public Sub zlRefreshReadOnly(ByVal lng����Id As Long, ByVal lng��ҳID As Long)
'���ܣ��ⲿ�������,��ʾֻ����·������
'ע�⣺�������ҽ��վ��ʿվ���ã�֮ǰ�����ȵ�clsCISPath.InitCISPath
    Call mclsDockPath.zlRefreshReadOnly(lng����Id, lng��ҳID)
End Sub

Public Function zlExecPathSend(Optional ByRef blnIsCancel As Boolean, Optional ByVal lngType As Long) As Boolean
'���ܣ�ִ��·������(ҽ��վ�ͻ�ʿվ����)
'������blnIsCancel��û��·��������ʱ���û��Ƿ�ȡ����������true=ȡ��
'      lngType:1-ҽ���༭������ã��������󲻼������ɣ���Ϊҽ���༭���治���ٵ���ҽ���༭��
    zlExecPathSend = mclsDockPath.zlExecPathSend(blnIsCancel, lngType)
End Function

Public Function zlExecEvaluateCancel(Optional ByVal blnPrompt As Boolean = True, Optional ByVal blnRefresh As Boolean = True) As Boolean
'���ܣ�ִ��ȡ������(ҽ��վ�ͻ�ʿվ����)
'������blnPrompt=�Ƿ񵯳�ѯ����ʾ
    zlExecEvaluateCancel = mclsDockPath.zlExecEvaluateCancel(blnPrompt, blnRefresh)
End Function

Public Function zlExePathAppendItem(ByVal strItemType As String, ByVal strAdviceIDs As String, ByVal lngִ��ID As Long, ByVal datDate As Date) As Boolean
'���ܣ�����·������Ŀ(ҽ��վ�ͻ�ʿվ����)
    zlExePathAppendItem = mclsDockPath.zlExePathAppendItem(strItemType, strAdviceIDs, lngִ��ID, datDate)
End Function

Public Function zlImportPath() As Boolean
    zlImportPath = mclsDockPath.zlImportPath
End Function

Public Function zlShowAltAdvice(ByVal frmParent As Object, ByVal lng·����ĿID As Long, Optional ByVal strSelectedIDs As String, Optional ByVal intӤ�� As Integer) As String
'���ܣ�·������ʱ�����ݵ�ǰ�е�·����ĿID����ʾ��ѡҽ��
'������strSelectedIds���Ѿ�ѡ�еı�ѡҽ��IDs
'���أ�ѡ����ѡ�еı�ѡҽ��IDs
    zlShowAltAdvice = mclsDockPath.zlShowAltAdvice(frmParent, lng·����ĿID, strSelectedIDs, intӤ��)
End Function

Public Sub SetFontSize(ByVal bytSize As Byte)
'����:����·�����嵥�������С
'���:bytSize��0-С(ȱʡ)��1-��
    Call mclsDockPath.SetFontSize(bytSize)
End Sub

Public Sub zlAddOutPathItem(ByVal strAdviceIDs As String, ByVal lng����Id As Long, ByVal lng��ҳID As Long)
'���ܣ�����ҽ��Ϊ"ֹͣ"״̬��ҽ��ʱ�����ݻ���ҽ��ID��������·������Ŀ
'������strAdviceIDs�����˵�ҽ��ID��
'      lng����Id-����ID
'      lng��ҳId-��ҳID
    Call mclsDockPath.zlAddOutPathItem(strAdviceIDs, lng����Id, lng��ҳID)
End Sub

Public Sub zlFuncPathTableOutPut(ByVal bytStyle As Byte, Optional ByVal blnIsSetup As Boolean, Optional ByVal strPDFFile As String, Optional ByVal lng����Id As Long, Optional ByVal lng��ҳID As Long, Optional ByVal strDeviceName As String)
'���ܣ�����ٴ�·������PDF
'������bytStyle=1-��ӡ,2-Ԥ��,3-�����Excel,4-�����PDF
'     blnIsSetup-������ӡ�����д�ӡǰ����
'     ��bytStyle=4ʱ����Ҫ����strPDFFile=PDF���Ĭ��·��,�����ļ�������׺
'     ����ⲿδ����zlRefresh��zlRefreshReadOnly�Ļ�������Ҫ���벡��ID,��ҳID
    Call mclsDockPath.zlFuncPathTableOutPut(bytStyle, blnIsSetup, strPDFFile, lng����Id, lng��ҳID, strDeviceName)
End Sub

