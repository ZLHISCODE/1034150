VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsICCardDev_JCSRFID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements clsCardDev

'������������
Private Declare Function piccreadex Lib "OUR_MIFARE.dll" (ByVal ctrlword As Byte, ByVal serial As Long, ByVal area As Byte, ByVal keyA1B0 As Byte, ByVal picckey As Long, ByVal piccdata0_2 As Long) As Byte

'д����������
Private Declare Function piccwriteex Lib "OUR_MIFARE.dll" (ByVal ctrlword As Byte, ByVal serial As Long, ByVal area As Byte, ByVal keyA1B0 As Byte, ByVal picckey As Long, ByVal piccdata0_2 As Long) As Byte

'�޸ĵ�����������
Private Declare Function piccchangesinglekey Lib "OUR_MIFARE.dll" (ByVal ctrlword As Byte, ByVal serial As Long, ByVal area As Byte, ByVal keyA1B0 As Byte, ByVal piccoldkey As Long, ByVal piccnewkey As Long) As Byte

'���豸�������캯������
Private Declare Function pcdbeep Lib "OUR_MIFARE.dll" (ByVal xms As Long) As Byte

'��ȡ�豸��ź�������
Private Declare Function pcdgetdevicenumber Lib "OUR_MIFARE.dll" (ByVal devicenumber As Long) As Byte

'�����ֶ���,������ָ��,�����ֵĺ�����鿴����˾��վ�ṩ�Ķ�̬��˵��
Private Const BLOCK0_EN = &H1
Private Const BLOCK1_EN = &H2
Private Const BLOCK2_EN = &H4
Private Const NEEDSERIAL = &H8
'Private Const EXTERNKEY = &H10
Private Const NEEDHALT = &H20
        

Private Function clsCardDev_CloseCard(Optional strInput As String, Optional strOut As String) As Boolean
    clsCardDev_CloseCard = True
End Function

Private Function clsCardDev_GetSpare() As Currency
    
End Function

Private Function clsCardDev_initCard(Optional strInput As String, Optional strOut As String) As Boolean
    clsCardDev_initCard = True
End Function

Private Function clsCardDev_PaymentSwap(ByVal dblMoney As Double, dblBalance As Double, intCardType As Integer, ByVal lngHOSCode As Long, strCardNO As String, strTransFlow As String, ByVal strBalanceID As String, ByVal strPatiID As String) As Boolean
    clsCardDev_PaymentSwap = True
End Function

Private Function clsCardDev_ReadCard(Optional strInput As String, Optional blnCancel As Boolean = False) As String
    Dim status As Byte '��ŷ���ֵ
    
    Dim myareano As Byte '����
    Dim authmode As Byte '�������ͣ���A�����B����
    Dim myctrlword As Byte '������
    Dim mypicckey(0 To 5) As Byte '����
    Dim mypiccserial(0 To 3) As Byte '�����к�
    Dim mypiccdata(0 To 47) As Byte '�����ݻ���
    
    Dim strMyPicdata As String  '���շ����ַ���
    Dim i As Integer   'ѭ������
    Dim strBit As String
    Dim strChar As String
    
    On Err GoTo errH:
    
    '������ָ��,�����ֵĺ�����鿴����˾��վ�ṩ�Ķ�̬��˵��
    myctrlword = BLOCK0_EN + BLOCK1_EN + BLOCK2_EN
    
    'ָ������
    myareano = 8 'ָ��Ϊ��8��
    '��������ģʽ
    authmode = 1 '����0��ʾ��A������֤���Ƽ���A������֤
    
    'ָ������
    mypicckey(0) = &HFF
    mypicckey(1) = &HFF
    mypicckey(2) = &HFF
    mypicckey(3) = &HFF
    mypicckey(4) = &HFF
    mypicckey(5) = &HFF
    
    status = piccreadex(myctrlword, VarPtr(mypiccserial(0)), myareano, authmode, VarPtr(mypicckey(0)), VarPtr(mypiccdata(0)))
    '�������趨�ϵ㣬Ȼ��鿴mypiccserial��mypiccdata��
    '������ piccreadex�����ɶ��������кŵ� mypiccserial�����������ݵ�mypiccdata��
    '������Ա�����Լ�����Ҫ����mypiccserial��mypiccdata �е������ˡ�
    '�������غ���
    If gbln�Զ���ȡ Then
        For i = 17 To 20
         strBit = Hex(mypiccdata(i))
         If Len(strBit) = 1 Then strBit = "0" & strBit
            strChar = strBit
            strMyPicdata = strMyPicdata & UCase(strChar)
        Next
        If strMyPicdata = "00000000" Then
           strMyPicdata = ""
        End If
    Else
        Select Case status
            Case 0:
                For i = 17 To 20
                    strBit = Hex(mypiccdata(i))
                    If Len(strBit) = 1 Then strBit = "0" & strBit
                    strChar = strBit
                    strMyPicdata = strMyPicdata & UCase(strChar)
                Next
                If strMyPicdata = "00000000" Then
                    strMyPicdata = ""
                End If
            Case 1:
                MsgBox "ˢ��̫�죬���¿����ݶ�ȡʧ�ܣ�", vbInformation
            Case 2, 3:
                MsgBox "ˢ��̫�죬���¿����ݶ�ȡ��������", vbInformation
            Case 8:
                MsgBox "�뽫�����ڸ�Ӧ����", vbInformation
            Case 9:
                MsgBox "�ж��ſ��ڸ�Ӧ����Ѱ�������з���ͻʧ�ܣ�", vbInformation
            Case 10:
                MsgBox "�Բ��𣬸ÿ������ߣ��޷�ѡ�У�", vbInformation
            Case 11:
                MsgBox "����װ��ʧ�ܣ�", vbInformation
            Case 12:
                MsgBox "������֤ʧ�ܣ�", vbInformation
            Case 21 'û�ж�̬��
                MsgBox "�Ҳ�����̬��ICUSB.DLL�뽫ICUSB.DLL������HIS������װ���Ŀ¼APPSOFT��", vbInformation
            Case 22:
                MsgBox "��̬���������������쳣��", vbInformation
            Case 24:
                MsgBox "������ʱ,�����ǲ������£�", vbInformation
            Case Else
                MsgBox "δ֪����", vbCritical
        End Select
    End If
    If strMyPicdata <> "" Then
      Call pcdbeep(100)
    End If
    clsCardDev_ReadCard = strMyPicdata
    Exit Function
    '0   �����ɹ���������������Ч��
    '1   0~2�鶼û������������ˢ��̫�顣�������к��ѱ���������
    '2   ��0���ѱ���������1~2���ȡʧ�ܡ������к��ѱ���������
    '3   ��0��1���ѱ���������2���ȡʧ�ܡ������к��ѱ���������
    '8   Ѱ�����󣬸�����û�п��ڸ�Ӧ����*serial��Ч��
    '9   �ж��ſ��ڸ�Ӧ����Ѱ�������з���ͻʧ�ܣ�*serial��Ч��
    '10  �ÿ������ѱ����ߣ��޷�ѡ�У��������к��ѱ�������*serial�����е�������Ч��
    '11  ����װ��ʧ��?
    '12  ������֤ʧ��?
    '21  ��������Ҫ���õĶ�̬��ICUSB.DLL���ڵ�ǰĿ¼��?
    '22  ��̬������������쳣������������˳����򣬰γ�IC����д������װ���������ٲ���IC����д�����ԣ��������¿�����̬��OUR_MIFARE.dll����ȷ��λ�á�
    '24  ������ʱ�������ǵ����ж�����USB֡���ݵ��Ȼ�����������IC����д�������⣬����������������Ի����°β�IC����д����
    '27  USB���䲻�ȶ����´�����ַ���ȫ�������������������Ϊ�������ǲ�������������ġ�
    '28  USB���䲻�ȶ�����CRCУ����������������������Ϊ�������ǲ�������������ġ�
    '���� δ֪����
errH:
    MsgBox Err.Description, vbExclamation
End Function

Private Function clsCardDev_ReturnSwap(ByVal strCardNO As String, ByVal lngHOSCode As Long, ByVal strTransFlow As String, ByVal dblMoney As Double) As Boolean
    clsCardDev_ReturnSwap = True
End Function

Private Sub clsCardDev_SetCard(Optional strInput As String, Optional strOut As String)
    Call frmParameterSet.ShowMe(6)
End Sub


Private Function clsCardDev_UploadSwap(ByVal lng����ID As Long, strOut As String) As Boolean

End Function

