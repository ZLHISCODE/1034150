VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsICCardDev_D3IC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements clsCardDev

'2010-02-05 ZHQ �����ǻҽԺD3IC��Ƶ����̬����
'comm function
Private Declare Function add_s Lib "dcrf32.dll" (ByVal i%) As Integer

Private Declare Function dc_init Lib "dcrf32.dll" (ByVal port%, ByVal baud As Long) As Long
Private Declare Function dc_exit Lib "dcrf32.dll" (ByVal icdev As Long) As Integer
Private Declare Function dc_request Lib "dcrf32.dll" (ByVal icdev As Long, ByVal mode%, tagtype As Long) As Integer
Private Declare Function dc_anticoll Lib "dcrf32.dll" (ByVal icdev As Long, ByVal bcnt%, snr As Long) As Integer
Private Declare Function dc_select Lib "dcrf32.dll" (ByVal icdev As Long, ByVal snr As Long, size As Byte) As Integer
Private Declare Function dc_card Lib "dcrf32.dll" (ByVal icdev As Long, ByVal mode%, snr As Long) As Integer
Private Declare Function dc_load_key Lib "dcrf32.dll" (ByVal icdev As Long, ByVal mode%, ByVal secnr%, ByRef nkey As Byte) As Integer
Private Declare Function dc_load_key_hex Lib "dcrf32.dll" (ByVal icdev As Long, ByVal mode%, ByVal secnr%, ByVal nkey As String) As Integer
Private Declare Function dc_authentication Lib "dcrf32.dll" (ByVal icdev As Long, ByVal mode%, ByVal scenr%) As Integer
Private Declare Function dc_read Lib "dcrf32.dll" (ByVal icdev As Long, ByVal adr%, ByVal sdata$) As Integer
Private Declare Function dc_read_hex Lib "dcrf32.dll" (ByVal icdev As Long, ByVal adr%, ByVal sdata$) As Integer
Private Declare Function dc_write Lib "dcrf32.dll" (ByVal icdev As Long, ByVal adr%, ByVal sdata$) As Integer
Private Declare Function dc_write_hex Lib "dcrf32.dll" (ByVal icdev As Long, ByVal adr%, ByVal sdata$) As Integer
Private Declare Function dc_changeb3 Lib "dcrf32.dll" (ByVal adr As Long, ByVal secer As Integer, ByRef KeyA As Byte, ByVal B0 As Integer, ByVal B1 As Integer, ByVal B2 As Integer, ByVal B3 As Integer, ByVal Bk As Integer, ByRef KeyB As Byte) As Integer
Private Declare Function dc_read_allhex Lib "dcrf32.dll" (ByVal icdev As Long, ByVal sdata$) As Integer
Private Declare Function dc_write_allhex Lib "dcrf32.dll" (ByVal icdev As Long, ByVal sdata$) As Integer
Private Declare Function dc_set_autoflag Lib "dcrf32.dll" (ByVal icdev As Long, ByVal flag%) As Integer
Private Declare Function dc_check_writehex Lib "dcrf32.dll" (ByVal icdev As Long, ByVal cardid As Long, ByVal mode As Integer, ByVal adr%, ByVal sdata$) As Integer


Private Declare Function dc_HL_initval Lib "dcrf32.dll" (ByVal icdev As Long, ByVal mode As Integer, ByVal adr%, ByVal value As Long, ByRef snr As Long) As Integer
Private Declare Function dc_HL_increment Lib "dcrf32.dll" (ByVal icdev As Long, ByVal mode As Integer, ByVal adr%, ByVal value As Long, ByVal snr As Long, value As Long, ByRef snr As Long) As Integer
Private Declare Function dc_HL_decrement Lib "dcrf32.dll" (ByVal icdev As Long, ByVal mode As Integer, ByVal adr%, ByVal value As Long, ByVal snr As Long, value As Long, ByRef snr As Long) As Integer

'
Private Declare Function dc_initval Lib "dcrf32.dll" (ByVal icdev As Long, ByVal adr%, ByVal value As Long) As Integer
Private Declare Function dc_readval Lib "dcrf32.dll" (ByVal icdev As Long, ByVal adr%, value As Long) As Integer
Private Declare Function dc_increment Lib "dcrf32.dll" (ByVal icdev As Long, ByVal adr%, ByVal value As Long) As Integer
Private Declare Function dc_decrement Lib "dcrf32.dll" (ByVal icdev As Long, ByVal adr%, ByVal value As Long) As Integer
Private Declare Function dc_restore Lib "dcrf32.dll" (ByVal icdev As Long, ByVal adr%) As Integer
Private Declare Function dc_transfer Lib "dcrf32.dll" (ByVal icdev As Long, ByVal adr%) As Integer
Private Declare Function dc_halt Lib "dcrf32.dll" (ByVal icdev As Long) As Integer
 
'device fuction
Private Declare Function dc_srd_eepromhex Lib "dcrf32.dll" (ByVal icdev As Long, ByVal adr%, ByVal lenth%, ByVal sdata$) As Integer
Private Declare Function dc_swr_eepromhex Lib "dcrf32.dll" (ByVal icdev As Long, ByVal adr%, ByVal lenth%, ByVal sdata$) As Integer

Private Declare Function dc_reset Lib "dcrf32.dll" (ByVal icdev As Long, ByVal msec%) As Integer
Private Declare Function dc_beep Lib "dcrf32.dll" (ByVal icdev As Long, ByVal time1 As Integer) As Integer

Private mintICTarget As Long   '�豸��ʶ��
Dim hexkey As String * 12

Private Function clsCardDev_CloseCard(Optional strInput As String, Optional strOut As String) As Boolean
    Dim i As Integer, blnicexit As Boolean
    
    On Error GoTo errH
    i = dc_exit(mintICTarget)
    clsCardDev_CloseCard = i = 0
  
    Call WritLog("CardDev_CloseCard", "", IIf(clsCardDev_CloseCard, "�ɹ�!", "ʧ��!"))
    
    Exit Function
errH:
    MsgBox Err.Description, vbExclamation
End Function

Private Function clsCardDev_GetSpare() As Currency

End Function

Private Function clsCardDev_initCard(Optional strInput As String, Optional strOut As String) As Boolean
    Dim i As Integer
    
    i = Val(GetSetting("ZLSOFT", "����ȫ��\ICCard\8", "�˿�", 0))
    If i < 0 Or i > 7 Then i = 100  'USB
    
    On Error GoTo errH
    
    '��ʼ�豸
    mintICTarget = dc_init(i, 9600) '�˿�Ϊ0
    clsCardDev_initCard = mintICTarget > 0
    
    Call WritLog("CardDev_initCard", i & ",9600", IIf(clsCardDev_initCard, "�ɹ�!", "ʧ��!"))
    
    'װ������
    hexkey = "ffffffffffff"
    ''װ�ص�1���������뵽��д������ͺ���Ҫ
    i = dc_load_key_hex(mintICTarget, 0, 0, hexkey)
    If i <> 0 Then  'У��1���������������Ӧ��
        MsgBox "����װ���������ϵ��Ƶ���ṩ�̣�", vbOKOnly, "IC������"
        clsCardDev_initCard = False
        Exit Function
    End If
    clsCardDev_initCard = True
    Exit Function
errH:
    MsgBox Err.Description, vbExclamation
End Function

Private Function clsCardDev_PaymentSwap(ByVal dblMoney As Double, dblBalance As Double, intCardType As Integer, ByVal lngHOSCode As Long, strCardNO As String, strTransFlow As String, ByVal strBalanceID As String, ByVal strPatiID As String) As Boolean

End Function

Private Function clsCardDev_ReadCard(Optional strInput As String, Optional blnCancel As Boolean = False) As String
    Dim i As Integer, intState As Integer, intSAddr As Integer, intLen As Integer
    Dim strResult As String * 32    '�ӿں���Ҫ�����Ϊ32�����ַ���
    Dim cardmode As Integer, address As Integer, snr
    Dim lngValue As Long
    Dim sector As Integer

    On Error GoTo errH

    cardmode = 0: address = 1: sector = 0
    'Ѱ������ cardmode=1 �ɶԿ�����������Ϊ0ʱ���ִ����HALT����
    i = dc_card(mintICTarget, cardmode, snr)
    If i <> 0 And i <> 1 Then
        MsgBox "�����豸�����Ƿ�����!", vbInformation, "IC������"
        Exit Function
    End If
    If i = 1 Then
        MsgBox "û���ҵ���Ƭ!", vbInformation, "IC������"
        Exit Function
    End If

    DoEvents

    '�˶Կ�����
    i = dc_authentication(mintICTarget, 0, sector)
    DoEvents

    intSAddr = Val(GetSetting("ZLSOFT", "����ȫ��\ICCard\8", "��ʼ��ַ", 32))
    If intSAddr <= 0 Then intSAddr = 1
    intLen = Val(GetSetting("ZLSOFT", "����ȫ��\ICCard\8", "����", 10))
    If intLen > 16 Then intLen = 16
    If intLen <= 0 Then intLen = 1

    i = dc_beep(mintICTarget, 10)
Read:
    i = dc_read_hex(mintICTarget, address, strResult)
    
    If i = 0 Then
        clsCardDev_ReadCard = Mid(Trim(strResult), intSAddr, intLen)
    Else
       If MsgBox("�������ݴ���,�Ƿ�����?", vbYesNo, "IC������") = vbYes Then
          GoTo Read
       End If
    End If

    Call WritLog("CardDev_ReadCard", "", clsCardDev_ReadCard)

    Exit Function
errH:
    MsgBox Err.Description, vbExclamation
End Function

Private Function clsCardDev_ReturnSwap(ByVal strCardNO As String, ByVal lngHOSCode As Long, ByVal strTransFlow As String, ByVal dblMoney As Double) As Boolean

End Function

Private Sub clsCardDev_SetCard(Optional strInput As String, Optional strOut As String)
    Call frmICSet_Demo.ShowMe(8)
End Sub


Private Function clsCardDev_UploadSwap(ByVal lng����ID As Long, strOut As String) As Boolean

End Function