VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHisCrust"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function ReceiveCommand(ByVal strCommand As String, Optional ByVal strParas As String, Optional ByRef pbSend As PropertyBag, Optional ByRef pbReturn As PropertyBag) As Boolean
'���ܣ�������ִ�����������չ,���ֽӿڼ���
'strCommand:ִ�е�����
'strParas:ִ������Ĳ���
'pbSend�����Ӵ�������
'pbReturn:���ӷ�������
'���أ������Ƿ�ִ�гɹ�
End Function

Public Sub ExitApp()
'���ܣ���������
    End
End Sub