VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDisReportCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mfrmReport As New frmReport

Public Sub Init(ByVal cnMain As ADODB.Connection, ByVal lngSys As Long)
'���ܣ�������ʼ��

    On Error GoTo errHand
    Set gcnOracle = cnMain
    Call InitCommon(gcnOracle)
    gstrSysName = GetSetting(AppName:="ZLSOFT", Section:="ע����Ϣ", Key:=UCase("gstrSysName"), Default:="")
    Call GetUserInfo
    glngSys = lngSys
    
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err = 0
End Sub

Public Sub OpenDoc(ByVal frmParent As Object, ByVal bytType As Byte, ByVal lngPatiID As Long, ByVal lngPageID As Long, ByVal bytFrom As Byte, ByVal bytBabyNo As Byte, ByVal lngDeptID As Long, ByVal lngFileId As Long)
'���ܣ����������ĵ�ʱ���ã�ʵ�����ݳ�ʼ���༭���ڴ�
'����˵��:
'frmParent:�༭����������ʾ�ĸ�����
'bytType:�༭��ʽ0-������1-�޸ģ�����������ȡ����
'lngPatiID:����ID
'lngPageID:��ҳID
'bytFrom:������Դ
'lngDeptID:��ǰ����
'lngFileId:bytType=1ʱ�������޸ĵ��ļ�ID,��Դ�ڵ��Ӳ�����¼.ID
    On Error GoTo errHand
    Call frmMain.ShowMe(frmParent, bytType, lngPatiID, lngPageID, bytFrom, bytBabyNo, lngDeptID, lngFileId)
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err = 0
End Sub

Public Sub PrintDoc(ByVal frmParent As Object, ByVal lngPatiID As Long, ByVal lngPageID As Long, ByVal lngFileId As Long, ByVal strPrintDeviceName As String)
'����:��ӡ
'����˵��:
'frmParent:�༭����������ʾ�ĸ�����
'lngPatiID:����ID
'lngPageID:��ҳID
'lngFileId:�ļ�ID
'strPrintDeviceName:����������ӡʱ��ָ����ӡ������ָ��ʱ���������ṩ��ӡ��ѡ��
    Call mfrmReport.PrintReport(frmParent, lngPatiID, lngPageID, lngFileId, strPrintDeviceName)
End Sub

Public Sub zlRefresh(ByVal lngPatiID As Long, ByVal lngPageID As Long, ByVal lngFileId As Long, ByVal blnMoved As Boolean)
'����:ˢ��
'����˵��:
'lngPatiID:����ID
'lngPageID:��ҳID
'lngFileId:�ļ�ID
'blnMoved:�����Ƿ�ת��
    On Error GoTo errHand
    Call mfrmReport.zlRefresh(lngPatiID, lngPageID, lngFileId, blnMoved)
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err = 0
End Sub

Public Function zlGetForm() As Object
    If mfrmReport Is Nothing Then Set mfrmReport = New frmReport
    If mfrmReport.hWnd = 0 Then Set mfrmReport = New frmReport '�ⲿUnLoad
    Set zlGetForm = mfrmReport
End Function

Public Sub SetCaption����֤()
    Call mfrmReport.SetCaption����֤
End Sub

Private Sub Class_Initialize()
    On Error Resume Next
    Set mfrmReport = New frmReport
    mfrmReport.blnHaveStatus = False
    Err.Clear
End Sub

Private Sub Class_Terminate()
    Set mfrmReport = Nothing
End Sub