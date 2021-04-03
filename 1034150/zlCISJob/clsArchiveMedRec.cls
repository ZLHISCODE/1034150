VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsArchiveMedRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mfrmInMedRec As Object
Attribute mfrmInMedRec.VB_VarHelpID = -1
Private mfrmOutMedRec As frmArchiveOutMedRec

Public Sub InitArchiveMedRec(cnMain As ADODB.Connection, ByVal lngSys As Long, ByVal strDBUser As String)
    Set gcnOracle = cnMain
    
    glngSys = lngSys
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
    gstrProductName = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒ����", "����")
    gstrUnitName = GetUnitName
    gstrDBUser = strDBUser
End Sub

Private Sub Class_Terminate()
    If Not mfrmInMedRec Is Nothing Then
        Unload mfrmInMedRec
        Set mfrmInMedRec = Nothing
    End If
    If Not mfrmOutMedRec Is Nothing Then
        Unload mfrmOutMedRec
        Set mfrmOutMedRec = Nothing
    End If
End Sub

Public Function zlGetForm(ByVal intType As Integer) As Object
'���ܣ���ȡ��ҵ�񴰿�,���������ʼ��ʱֻ����һ�λ�ȡ
'������intType 0-���ﲡ������ 1-סԺ��������
    Dim intMecStandard As Integer  '������ҳ��ʽ 0-��������׼��1-�Ĵ�ʡ��׼��2-����ʡ��׼,3-����ʡ��׼
    
    If intType = 0 Then
        If mfrmOutMedRec Is Nothing Then
            Set mfrmOutMedRec = New frmArchiveOutMedRec
        End If
        Set zlGetForm = mfrmOutMedRec
    Else
        If mfrmInMedRec Is Nothing Then
            '������ҳ��׼
            intMecStandard = Val(zlDatabase.GetPara("������ҳ��׼", glngSys, pסԺҽ��վ, "0"))
            Select Case intMecStandard
                Case 0 '��������׼
                    Set mfrmInMedRec = New frmArchiveInMedRec
                Case 1    '�Ĵ�ʡ��׼
                    Set mfrmInMedRec = New frmArchiveInMedRec_SC
                Case 2    '����ʡ��׼
                    Set mfrmInMedRec = New frmArchiveInMedRec_YN
                Case 3    '����ʡ��׼
                    Set mfrmInMedRec = New frmArchiveInMedRec_HN
            End Select
        End If
        Set zlGetForm = mfrmInMedRec
    End If
End Function

Public Function zlRefresh(ByVal intType As Integer, ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal blnMoved As Boolean) As Boolean
'���ܣ�ˢ�»����ҽ���嵥
'������intType 0-���ﲡ������ 1-סԺ��������
'      lng��ҳID :סԺ��ҳ-��ҳID�� ������ҳ-�Һ�ID
'      blnMoved �Ƿ�ת��
    If intType = 0 Then
        Call mfrmOutMedRec.zlRefresh(lng����ID, lng��ҳID, blnMoved)
    Else
        Call mfrmInMedRec.zlRefresh(lng����ID, lng��ҳID, blnMoved)
    End If
End Function

