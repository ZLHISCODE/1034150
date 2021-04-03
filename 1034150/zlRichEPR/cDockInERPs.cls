VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDockInEPRs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'�̳д����¼�
Public Event Activate()
Public Event ClickDiagRef(DiagnosisID As Long, Modal As Byte)

Private WithEvents frmInEPRs As frmDockInEPRs
Attribute frmInEPRs.VB_VarHelpID = -1

Private Sub Class_initialize()
    Set frmInEPRs = New frmDockInEPRs
End Sub

Private Sub Class_Terminate()
    Unload frmInEPRs
    Set frmInEPRs = Nothing
End Sub

Private Sub frmInEPRs_Activate()
    RaiseEvent Activate
End Sub

Private Sub frmInEPRs_ClickDiagRef(DiagnosisID As Long, Modal As Byte)
    RaiseEvent ClickDiagRef(DiagnosisID, Modal)
End Sub
Public Sub SetFontSize(ByVal bytSize As Byte)
'-0-С(ȱʡ)��1-��
    Call frmInEPRs.SetFontSize(bytSize)
End Sub

Public Function zlGetForm() As Object
    Set zlGetForm = frmInEPRs
End Function

Public Sub zlDefCommandBars(ByVal CommandBars As Object, Optional ByVal blnInsideTools As Boolean)
'�Խ�������
    Call frmInEPRs.zlDefCommandBars(CommandBars, blnInsideTools)
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Call frmInEPRs.zlUpdateCommandBars(Control)
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Call frmInEPRs.zlExecuteCommandBars(Control)
End Sub

Public Function zlRefresh(ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngDeptId As Long _
    , Optional ByVal bnEdit As Boolean, Optional ByVal blnMoved As Boolean, Optional ByVal lngAdviceID As Long _
    , Optional blnForce As Boolean, Optional ByVal lngUnitID As Long, Optional ByVal intState As Integer) As Long
    '����:����ˢ��ָ�����˵Ĳ������ݣ�����������ṩ�༭����
    '����:  lngPatiId-����id;
    '       lngPageId-��ҳid;
    '       lngDeptId-��ǰ�������ţ�ע�ⲻ�ǲ��˵�ǰ���ڿ��һ��Ժ���ң���Ҫ���ݸò���ȷ�������ӵĲ��˲�����
    '       lngUnitID��lngDeptId-��intState=��5-���ת�Ʋ��ˡ�ʱΪ����ԭ������ԭ����
    '       blnEdit-�Ƿ������༭��ͨ����ǰ�������Ų��ǲ��˵�ǰ���ڿ��һ��Ժ���ң���Ӧ�ò������༭����ʿվֱ������Ϊ�������༭
    '       blnForce-ǰ�����β���ID,��ҳID��ͬʱ,�Ƿ�ǿ��ˢ��
    '       blnMoved-�����Ƿ�ת��
    '       lngAdviceID-ҽ��ID,ĿǰΪ����ģ�����
    '       intState-��ʿվ��ҽ��վ����Ĳ���״̬TYPE_PATI_State
    '    ps��Ժ = 0
'    psԤ�� = 1
'    ps��Ժ = 2
'    ps���� = 3          'ҽ��վ:�����ﲡ��(��Ժ)
'    ps���� = 4          'ҽ��վ:�ѻ��ﲡ��
'    ps���ת�� = 5      'ҽ��վ:���ת�ƻ�ת�����Ĳ���(��Ժ)
'    ps��ת�� = 6        'ҽ��վ:��ƴ���ס��ת��������������
'ת��ҳ�����intState=5������ҳ���������ѻ������intState=4,δ����intState=3������ҳ�����intState=���벡�˵ĵ�ǰ״̬��Ԥ��1����Ժ0����Ժ2��
'ҽ��վ��������ʾ ,intState=5�����ת��ǰ�Ŀ���ID�Ͳ���ID��������ĵ�ǰ���˿���ID�Ͳ���ID
'ҽ��վ��������ʾ������intstate�Ƕ��٣��̶��������������ң���������ͬ������ʾ

    zlRefresh = frmInEPRs.zlRefresh(lngPatiID, lngPageId, lngDeptId, bnEdit, blnForce, blnMoved, lngAdviceID, intState)
End Function
