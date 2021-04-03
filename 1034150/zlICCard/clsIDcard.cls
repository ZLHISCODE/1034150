VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIDcard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements clsCardDev

'-- ��������֤������API
'Private Declare Function CVR_InitComm Lib "termb.dll" (ByVal port As Long) As Integer
'Private Declare Function CVR_CloseComm Lib "termb.dll" () As Integer
'Private Declare Function CVR_Authenticate Lib "termb.dll" () As Integer
'Private Declare Function CVR_Read_Content Lib "termb.dll" (ByVal active As Long) As Integer
'Private Declare Function CVR_Ant Lib "termb.dll" (ByVal mode As Long) As Integer

Private Declare Function CVR_InitComm Lib "termb.dll" Alias "InitComm" (ByVal port As Integer) As Integer
Private Declare Function CVR_InitCommExt Lib "termb.dll" Alias "InitCommExt" () As Integer
Private Declare Function CVR_Authenticate Lib "termb.dll" Alias "Authenticate" () As Integer
Private Declare Function CVR_AuthenticateExt Lib "termb.dll" Alias "AuthenticateExt" () As Integer
Private Declare Function CVR_Read_Content_Path Lib "termb.dll" Alias "Read_Content_Path" (ByVal fileName As String, ByVal active As Integer) As Integer
Private Declare Function CVR_Read_Content Lib "termb.dll" Alias "Read_Content" (ByVal active As Integer) As Integer
Private Declare Function CVR_CloseComm Lib "termb.dll" Alias "CloseComm" () As Integer
Private Declare Function CVR_GetSAMID Lib "termb.dll" Alias "GetSAMID" () As String


Private Function clsCardDev_CloseCard(Optional strInput As String, Optional strOut As String) As Boolean
    On Error GoTo errHandle
    clsCardDev_CloseCard = CVR_CloseComm = 1
    Call WritLog("IDcard.CloseComm", "", IIf(clsCardDev_CloseCard, "�ɹ�!", "ʧ��!"))
    Exit Function
errHandle:
    clsCardDev_CloseCard = False
    gLastErr = "ִ��CloseCard���̳��ִ���" & vbNewLine & Err.Description
End Function

Private Function clsCardDev_GetSpare() As Currency

End Function

Private Function clsCardDev_initCard(Optional strInput As String, Optional strOut As String) As Boolean
    On Error GoTo errHandle
    gLastErr = ""
    If Dir(App.Path & "\termb.dll") = "" Then
         gLastErr = App.Path & "��ȱtermb.dll�ļ���": GoTo errHandle
         
    End If
    If Dir(App.Path & "\sdtapi.dll") = "" Then
         gLastErr = App.Path & "��ȱsdtapi.dll�ļ���": GoTo errHandle
    End If
    
'    If Dir(App.Path & "\CVRApi.dll") = "" Then
'         gLastErr = App.Path & "��ȱCVRApi.dll�ļ���": GoTo errHandle
'    End If

    clsCardDev_initCard = CVR_InitComm(1001) = 1
    Call WritLog("IDcard.initCard", "1001", IIf(clsCardDev_initCard, "�ɹ�!", "ʧ��!"))
    Exit Function
errHandle:
    clsCardDev_initCard = False
    gLastErr = "ִ��initCard���̳��ִ���" & vbNewLine & Err.Description & gLastErr
End Function


Private Function clsCardDev_PaymentSwap(ByVal dblMoney As Double, dblBalance As Double, intCardType As Integer, ByVal lngHOSCode As Long, strCardNO As String, strTransFlow As String, ByVal strBalanceID As String, ByVal strPatiID As String) As Boolean

End Function

Private Function clsCardDev_ReadCard(Optional strInput As String, Optional blnCancel As Boolean = False) As String
    Dim strFileName As String, strLines As String
    Dim intReturn As Integer
    On Error GoTo errHandle
    intReturn = CVR_Authenticate
    Call WritLog("IDcard.Authenticate", "", CStr(intReturn))
   
    intReturn = CVR_Read_Content(4)
    Call WritLog("IDcard.Read_Content", "4", CStr(intReturn))
    
    strFileName = Dir(App.Path & "\wz.txt")
    If strFileName <> "" Then
        Open strFileName For Input As #1
        Line Input #1, strLines '����
        Line Input #1, strLines '�Ա�
        Line Input #1, strLines '����
        Line Input #1, strLines '����
        Line Input #1, strLines '��ַ
        Line Input #1, strLines '����֤��
        clsCardDev_ReadCard = Trim(strLines)
        Close #1
        Kill strFileName
        strFileName = Dir(App.Path & "\zp.bmp")
        If strFileName <> "" Then Kill strFileName
        strFileName = Dir(App.Path & "\zp.wlt")
        If strFileName <> "" Then Kill strFileName
    End If
    
    Exit Function
errHandle:
    clsCardDev_ReadCard = ""
    gLastErr = "ִ��initCard���̳��ִ���" & vbNewLine & Err.Description
End Function

Private Function clsCardDev_ReturnSwap(ByVal strCardNO As String, ByVal lngHOSCode As Long, ByVal strTransFlow As String, ByVal dblMoney As Double) As Boolean

End Function

Private Sub clsCardDev_SetCard(Optional strInput As String, Optional strOut As String)
    '
End Sub

Private Function clsCardDev_UploadSwap(ByVal lng����ID As Long, strOut As String) As Boolean

End Function