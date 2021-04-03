VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDockDiagReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'˵�������ڹҽӸ����Ӵ���,���Ҫȡ��,ɾ���������ݣ�
'1.ȥ���������ã�zl9CISCore,zl9PrintMode,Microsoft Scripting Runtime
'2.�Ƴ������ļ���frmDockDiagReport,frmPubIcons,clsDockDiagReport,clsFtp
'3.ɾ����������վ��mclsOldDiag,"�ɸ���"��ش���

Public Event Activate() '�Ӵ��弤��
Public Event RequestRefresh(ByVal RefreshNotify As Boolean) 'Ҫ��������ˢ��
Public Event StatusTextUpdate(ByVal Text As String) 'Ҫ�����������״̬������
Private WithEvents mfrmDiagReport As frmDockDiagReport
Attribute mfrmDiagReport.VB_VarHelpID = -1

Private Sub Class_Initialize()
    Set mfrmDiagReport = New frmDockDiagReport
End Sub

Private Sub Class_Terminate()
    Set mfrmDiagReport = Nothing
End Sub

Private Sub mfrmDiagReport_Activate()
    RaiseEvent Activate
End Sub

Private Sub mfrmDiagReport_RequestRefresh(ByVal RefreshNotify As Boolean)
    RaiseEvent RequestRefresh(RefreshNotify)
End Sub

Private Sub mfrmDiagReport_StatusTextUpdate(ByVal Text As String)
    RaiseEvent StatusTextUpdate(Text)
End Sub

Public Function zlGetForm() As Object
'���ܣ���ȡ��ҵ�񴰿�,���������ʼ��ʱֻ����һ�λ�ȡ

    '��ʼ��ȡȨ��,�Լӿ�����ٶ�(���Ǳ�Ҫ��)
    Call GetInsidePrivs(p�����¼����, True)
    
    Set zlGetForm = mfrmDiagReport
End Function

Public Sub zlDefCommandBars(ByVal frmParent As Object, ByVal CommandBars As Object, ByVal int���� As Integer, Optional ByVal bln��ʿվ As Boolean)
'���ܣ������Ӵ���Ĳ˵��͹�����(����������Ҫʹ�õĲ˵��͹�����)
'������int����=1-����,2-סԺ
    Call mfrmDiagReport.zlDefCommandBars(frmParent, CommandBars, int����, bln��ʿվ)
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As Object)
    Call mfrmDiagReport.zlUpdateCommandBars(Control)
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As Object)
    Call mfrmDiagReport.zlExecuteCommandBars(Control)
End Sub

Public Sub zlPopupCommandBars(ByVal CommandBar As Object)
    Call mfrmDiagReport.zlPopupCommandBars(CommandBar)
End Sub

Public Sub zlRefresh(ByVal lng����ID As Long, ByVal var����ID As Variant, ByVal int���� As Integer, Optional ByVal blnMoved As Boolean)
'���ܣ�ˢ��סԺҽ������
'������var����ID=����Ϊ�Һ�NO,סԺΪ��ҳID
'      int���ͣ�
'         ���ﲡ�ˣ�0-���ɱ༭,1-���Ա༭
'         סԺ���ˣ�0-��Ժ����,1-Ԥ��Ժ����,2-��Ժ����,3-�����ﲡ��(��Ժ),4-�ѻ��ﲡ��
'      blnMoved=�ò��˵������Ƿ���ת��
    Call mfrmDiagReport.zlRefresh(lng����ID, var����ID, int����, blnMoved)
End Sub