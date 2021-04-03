VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDockSeating"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Seatings"
Option Explicit

Public Event Activate() '�Ӵ��弤��
Public Event RequestRefresh() 'Ҫ��������ˢ��
Public Event StatusTextUpdate(ByVal Text As String) 'Ҫ�����������״̬������
Private WithEvents mfrmSeating As frmDockSeat
Attribute mfrmSeating.VB_VarHelpID = -1

Private Sub Class_Initialize()
    Set mfrmSeating = New frmDockSeat
End Sub

Private Sub Class_Terminate()
    Unload mfrmSeating
    Set mfrmSeating = Nothing
End Sub

Private Sub mfrmSeating_Activate()
    RaiseEvent Activate
End Sub

Private Sub mfrmSeating_RequestRefresh()
    RaiseEvent RequestRefresh
End Sub

Private Sub mfrmSeating_StatusTextUpdate(ByVal Text As String)
    RaiseEvent StatusTextUpdate(Text)
End Sub

Public Function zlGetForm() As Object
'���ܣ���ȡ��ҵ�񴰿�,���������ʼ��ʱֻ����һ�λ�ȡ

    '��ʼ��ȡȨ��,�Լӿ�����ٶ�(���Ǳ�Ҫ��)
    Call GetInsidePrivs(1264, True)
    
    'һЩ�ӿڳ�ʼ��,ֻ��ʼ��һ��
    '���漰����ģ���Ȩ��,��������ʼ��
    
    '����ǩ����ʼ��:���漰����ģ���Ȩ��,��������ʼ��
    '��Ϊ����ģ�����ͬʱʹ��,�ҹ��ñ���gobjESign����˻���Ҫ�ھ���ģ�����ж�Ȩ��

    Set zlGetForm = mfrmSeating
End Function

Public Sub zlDefCommandBars(ByVal frmParent As Object, ByVal CommandBars As Object, ByVal int���� As Integer)
'���ܣ������Ӵ���Ĳ˵��͹�����(����������Ҫʹ�õĲ˵��͹�����)
'������int����=0-ҽ��վ����,1-��ʿվ����,2-ҽ��վ����(PACS/LIS)
    Call mfrmSeating.zlDefCommandBars(frmParent, CommandBars, int����)
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As Object)
    Call mfrmSeating.zlUpdateCommandBars(Control)
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As Object)
    Call mfrmSeating.zlExecuteCommandBars(Control)
End Sub

Public Sub zlPopupCommandBars(ByVal CommandBar As Object)
    Call mfrmSeating.zlPopupCommandBars(CommandBar)
End Sub

Public Sub zlRefresh(ByVal curSeatings As Seatings, ByVal lng����ID As Long, ByVal ObjPati As cPatient)
'���ܣ�ˢ��סԺҽ������
'������lngǰ��ID=����ҽ��վ����ʱ����
'      blnMoved=�ò��˵������Ƿ���ת��
    mfrmSeating.lng����ID = lng����ID
    Set mfrmSeating.ObjPati = ObjPati
    Call mfrmSeating.zlRefresh(curSeatings)
End Sub
