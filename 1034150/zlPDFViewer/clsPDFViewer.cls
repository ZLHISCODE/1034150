VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPDFViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mfrmPDF As New frmPDF

Public Function GetFrm() As Object
'���ܣ����ش���
    Set GetFrm = mfrmPDF
End Function

Public Function LoadFile(ByVal strFile As String) As Boolean
'���ܣ������ļ�
    LoadFile = mfrmPDF.LoadFile(strFile)
End Function

Public Function PrintFile(ByVal intType As Integer) As Boolean
'���ܣ���ӡ
'����: intType ��ӡ��ʽ,0-ֱ�Ӵ�ӡ,1-������ӡ
    PrintFile = mfrmPDF.PrintFile(intType)
End Function

Public Function WaitTime(ByVal lng��� As Long, ByVal strFilePath As String, ByVal strName As String)
'����:��ӡ�ȴ�
'����:strFilePath�ļ�·��
'     strName ��������
    WaitTime = mfrmPDF.WaitTime(lng���, strFilePath, strName)
End Function