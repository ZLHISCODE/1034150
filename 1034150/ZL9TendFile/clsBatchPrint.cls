VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBatchPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Option Explicit
'Private mfrmMain As Object
'Private mfrmTemperatureChart As Object
'Private WithEvents mfrmTendEditor As frmTendFileEditor
'
'Public Event AfterTendPrint(ByVal lngFileId As Long)
'
'Public Function zlPrintDocument(ByVal bytKind As Byte, Optional ByVal bytMode As Byte = 2, Optional ByVal strPrintDeviceName As String) As Long
'    '����:  ��ӡ���µ�;bytMode��2-��ӡ
'
'    If bytKind = 1 Then
'        '���µ�(����ֵ:1-�ɹ�;2-��ӡ)
'        zlPrintDocument = mfrmTemperatureChart.zlPrintBody(bytMode, strPrintDeviceName)
'    Else
'        '������¼��
'        Call mfrmTendEditor.zlPrintTend(bytMode, strPrintDeviceName)
'    End If
'End Function
'
'Public Function zlRefresh(ByVal bytKind As Byte, ByVal lngKey As Long, Optional ByVal lngPatiID As Long, Optional ByVal lngPageId As Long, Optional ByVal lngDeptId As Long, Optional ByVal strPeriod As String, Optional ByVal byt�������� As Byte = 3, Optional ByVal intBaby As Integer = 0, Optional ByVal blnDataMoved As Boolean) As Long
'    '����:  ��ʾָ���ĵ��Ӳ���
'    '����:  lngKey-���Ӳ�����ID;
'
'    Call mfrmTendEditor.ShowMe(Nothing, lngKey, lngPatiID, lngPageId, lngDeptId, intBaby, strPeriod, True, byt��������, blnDataMoved)
'End Function
'
'Public Function ParentForm(ByVal frmMain As Object) As Boolean
'    Set mfrmMain = frmMain
'End Function
'
'Public Function zlRefreshTendBody(ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngDeptId As Long, Optional ByVal intBaby As Integer = 0, Optional ByVal blnDataMove As Boolean) As Long
'    '����:  ��ʾ���µ�
'    Call mfrmTemperatureChart.zlRefresh(lngPatiID & ";" & lngPageId & ";" & lngDeptId & ";0;0;" & intBaby)
'End Function
'
'Public Function zlGetFormTendFile() As Object
'    Set zlGetFormTendFile = mfrmTendEditor
'End Function
'
'Public Function zlGetFormTendBody() As Object
'    Set zlGetFormTendBody = mfrmTemperatureChart
'End Function
'
'Private Sub mfrmTendEditor_zlAfterPrint(ByVal lngFileId As Long)
'    RaiseEvent AfterTendPrint(lngFileId)
'End Sub
'
'Private Sub Class_initialize()
'    If Not CreateBodyEditor Then Exit Sub
'
'    Set mfrmTemperatureChart = gobjBodyEditor.GetNewTendBody
'    Set mfrmTendEditor = New frmTendFileEditor
'    '�����²�������Ϊ�ޱ��������Ӵ���
'    Call FormSetCaption(mfrmTendEditor, False, False)
'End Sub
'
'Private Sub Class_Terminate()
'    If Not (mfrmTemperatureChart Is Nothing) Then Unload mfrmTemperatureChart
'    If Not (mfrmTendEditor Is Nothing) Then Unload mfrmTendEditor
'
'    Set mfrmTemperatureChart = Nothing
'    Set mfrmTendEditor = Nothing
'    Set mfrmMain = Nothing
'End Sub