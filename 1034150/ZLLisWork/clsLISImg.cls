VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISImg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function Get_Chart2d_File(ByVal strPath As String, ByVal lngID As Long) As Boolean
    ' ��ȡ���ݿ����ͼ�����ݣ�����Chart2d�ؼ��ܹ�ֱ�ӵ��õ��ļ���
    ' ��Σ�
    '   strPath : ���ɵ�ͼ���ļ������·��
    '   lngID   : ����ͼ������ID�������ͼ���ļ���Ϊ��strpath &"\" & lngID & ".cht"
    Get_Chart2d_File = LoadImageData(strPath, lngID)
End Function

Public Function Get_ReportCode(ByVal lngҽ��ID As Long, ByVal lng���ͺ� As Long, ByRef strCode As String, ByRef strNO As String, ByRef bytMode As Byte, Optional ByVal DataMoved As Boolean = False) As Boolean
    'ȡLIS������ ���
    Get_ReportCode = GetReportCode(lngҽ��ID, lng���ͺ�, strCode, strNO, bytMode, DataMoved)
End Function

Public Sub ShowPatientRptPrint(cnOracle As ADODB.Connection, ByVal lngSys As Long, ByVal lngPatiDeptID As Long, ByVal MainPrivs As String, ByVal frmParent As Object)
    '����������ӡ���鱨��
    'lngPatiDeptID  :����ID
    'frmMain        :����������
    Dim frmRptPrint As New frmPatientRptPrint
    Call frmRptPrint.ShowME(cnOracle, lngSys, frmParent, lngPatiDeptID, MainPrivs)
End Sub