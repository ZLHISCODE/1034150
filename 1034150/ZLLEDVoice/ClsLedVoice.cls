VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsLedVoice"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit 'Ҫ���������
Private blnInit_Dev_MDT_SD01 As Boolean  '����Ҫ���˳�ʱ�ر��豸
Private blnDev_surpass As Boolean

'��ģ̬��������ʾ��ģ̬����
Private Const SW_SHOWNORMAL = 1
Private Const SW_SHOWNOACTIVATE = 4

Private Const HWND_DESKTOP = 0
Private Const HWND_TOP = 0
Private Const HWND_TOPMOST = -1
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2

Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Sub Reset(commTemp As Object)
    '''''''''''''''''''''''''''''''''''''''''''
    '���ܣ���ʼ����
    '������
    '���أ�
    '''''''''''''''''''''''''''''''''''''''''''
    Dim mySys As OSVERSIONINFO
    Dim lngResult As Long
        
    On Error Resume Next
    
    Set ctlComm = commTemp
    'ȡ����ϵͳ
    mySys.dwOSVersionInfoSize = Len(mySys)   'Ӧ���� 148
    lngResult = GetVersionEx(mySys)
    
    Select Case mySys.dwPlatformId
        Case VER_PLATFORM_WIN32s
            gintOS = 0
        Case VER_PLATFORM_WIN32_WINDOWS
            gintOS = 1
        Case VER_PLATFORM_WIN32_NT
            gintOS = 2
    End Select
    
    'ȡע����еĲ���
    gintDevice = Val(GetSetting("ZLSOFT", "����ȫ��", "�豸����", 1))
    gintPort = Val(GetSetting("ZLSOFT", "����ȫ��", "�˿�", 1))
    gstrSpeed = Val(GetSetting("ZLSOFT", "����ȫ��", "������", 9600))
    gblnDDisplay = Val(GetSetting("ZLSOFT", "����ȫ��", "˫����ʾ��", 0)) = 1
    
    gblnHaveBottom = Val(GetSetting("ZLSOFT", "����ȫ��", "�е�����Ϣ", 0)) <> 0
    gstrBottom = GetSetting("ZLSOFT", "����ȫ��", "������Ϣ", "")
    '������(2005-10-12)
    gblnNewDev = Val(GetSetting("ZLSOFT", "����ȫ��", "����SHY-II", 0)) <> 0
    gbln������� = Val(GetSetting("ZLSOFT", "����ȫ��", "���������ʾ", 0)) <> 0
    
    Select Case gintDevice
        Case Dev_SYC_XII
            SycVoice "f"
            'SycVoice "*"
        Case Dev_SYC_Q9
            Call SYC_Q9(gintPort, "*")
        Case Dev_LK822
            With ctlComm
                If .portopen = True Then
                    .portopen = False
                End If
                    
                If .commport = False Then
                    .commport = True
                End If
                '.commport = gintPort
                .Settings = gstrSpeed & "n,8,1"
                .InputLen = 0
                
                .portopen = True
                .output = Chr(27) + "@"
                .output = Chr(27) + "CLR"
            End With
        Case Dev_SHY_II
            SHYVoice "F"
        Case Dev_NJF_VH
            Set gobjLED = CreateObject("CTSVR.Bjq")
            If Not gobjLED Is Nothing Then
                gobjLED.Comport = gintPort
                gobjLED.DispMode = 0
                gobjLED.Display "~ "
                gobjLED.Reset '��������ظ�ִ��ʱ�ܿ��ܻ�����
                Set gobjLED = Nothing
            End If
        Case Dev_TDKJ_BJ
            Call TDKJ_BJ_FUN(gintPort, "&Sc$")
        Case Dev_TDKJ_BJ_2008
            Call TDKJ_BJ_2008(gintPort, "&Sc$")
        Case Dev_MDT_SD01
                If Not blnInit_Dev_MDT_SD01 Then '��γ�ʼ��ʱ,VB��������޷�����
                    InitService
                    InitDevice gintPort
                    blnInit_Dev_MDT_SD01 = True
                End If
                Clear_Screen
        Case Dev_surpass
            If Not blnDev_surpass Then '��γ�ʼ��ʱ,VB��������޷�����
               SetComNo (gintPort)
               blnDev_surpass = True
            End If
        Case Dev_FS_YL01
            Call opencomm(gintPort)
        Case Dev_TDKJ_BJ_IV
            Call TDKJ_BJ_IV(gintPort, "&Sc$")
    End Select
End Sub

Public Sub Init(ByVal strDisplay As String, ByVal lngModule As Long, cnMain As ADODB.Connection)
    '''''''''''''''''''''''''''''''''''''''''''
    '���ܣ���ʼ����ʾ                       �磺    '''''''''''''''''''''''''
    '                                               '                       '
    '                                               '��ӭ����Ժ����         '
    '                                               '     ***�շ�ԱΪ����� '
    '                                               '                       '
    '                                               '''''''''''''''''''''''''
    '
    '������
    '   strDisplay:����ʾ�ĵڶ����ַ�     �磺***�շ�ԱΪ�����
    '   lngModule:�����ߵ�ģ����
    '   cnMain:���ݿ����Ӷ���
    '
    '���أ�
    '''''''''''''''''''''''''''''''''''''''''''
    Dim i As Integer
    Dim str���Ե� As String
    
    Select Case gintDevice
        Case Dev_SYC_XII
            SycVoice "$1"
            SycVoice "# ��ӭ����Ժ���� #"
            SycVoice "$2"
            SycVoice "#" & strDisplay & "#"
        Case Dev_SYC_Q9
            Call SYC_Q9(gintPort, "&Sc$")
            Call SYC_Q9(gintPort, "# ��ӭ����Ժ���� #")
            Call SYC_Q9(gintPort, "$2")
            Call SYC_Q9(gintPort, "#" & strDisplay & "#")
        Case Dev_LK822
            strDisplay = Trim(strDisplay)
            For i = 1 To 15 - Len(strDisplay)
                strDisplay = " " & strDisplay
            Next
            ShowLED "��ӭ����Ժ����", " ף �� �� �� �� �� ", strDisplay, gstrBottom
        Case Dev_SHY_II
        
        Case Dev_NJF_VH
            On Error Resume Next
            Set gobjLED = CreateObject("CTSVR.Bjq")
            If Not gobjLED Is Nothing Then
                gobjLED.Comport = gintPort
                gobjLED.DispMode = 0
                gobjLED.Display "~" & strDisplay & ".^" & "��ӭ����Ժ����,ף�����տ���!"
                Set gobjLED = Nothing
            End If
        Case Dev_TDKJ_BJ
            If strDisplay Like "*�շ�ԱΪ������" Then
                strDisplay = Trim(Replace(strDisplay, "�շ�ԱΪ������", ""))
            End If
            Call TDKJ_BJ_FUN(gintPort, "&Sc$")
            Call TDKJ_BJ_FUN(gintPort, "&C11 ��ӭ����Ժ����$")
            Call TDKJ_BJ_FUN(gintPort, "&C21  ף�����տ���$")
            Call TDKJ_BJ_FUN(gintPort, "&C41�շ�Ա:" & strDisplay & "$")
        
        Case Dev_TDKJ_BJ_2008
            If strDisplay Like "*�շ�ԱΪ������" Then
                strDisplay = Trim(Replace(strDisplay, "�շ�ԱΪ������", ""))
            End If
            Call TDKJ_BJ_2008(gintPort, "&Sc$")
            Call TDKJ_BJ_2008(gintPort, "&C11 ��ӭ����Ժ����$")
            Call TDKJ_BJ_2008(gintPort, "&C21  ף�����տ���$")
            Call TDKJ_BJ_2008(gintPort, "&C41�շ�Ա:" & strDisplay & "$")
        Case Dev_MDT_SD01
            If strDisplay Like "*�շ�ԱΪ������" Then
                strDisplay = Trim(Replace(strDisplay, "�շ�ԱΪ������", ""))
            End If
            Display_Line "���,��ӭ����Ժ����", 4, 0
            Voices "010208"  '���,��ӭ����,ף�����տ���
            Display_Line SetLength("�շ�Ա:" & strDisplay, 20), 4, 1
        Case Dev_surpass
            If strDisplay Like "*�շ�ԱΪ������" Then
                strDisplay = Trim(Replace(strDisplay, "�շ�ԱΪ������", ""))
            End If
            Call AllClear  '����
            str���Ե� = "���ã����Ե�.wav"
            Call LocStringDisplay(2, 2, "���,���Ե�" + Chr(0))
            Call LocStringDisplay(2, 22, "ף�����տ���" + Chr(0))
            Call LocStringDisplay(2, 42, "�շ�Ա:" & strDisplay & "" + Chr(0))
        Case Dev_FS_YL01
            '
        Case Dev_TDKJ_BJ_IV
            '--����ͼƬ
            Dim rsTmp As New ADODB.Recordset
            Dim strsql As String
            
            Err = 0
            On Error GoTo ErrHand
            If gBlnPic = False Then
                Set rsTmp = zlDatabase.GetUserInfo()
                strsql = "select ��Աid,��Ƭ from ��Ա��Ƭ where ��Աid=[1]"
                Set rsTmp = zlDatabase.OpenSQLRecord(strsql, "��ȡ��Ƭ", Val(rsTmp!ID))
                                
                If rsTmp.RecordCount > 0 Then
                    '��ȡ��Ƭ
                    zlCommFun.ShowFlash "���Եȣ����ڼ�����Ա��Ƭ����"
                    Call zlDatabase.ReadPicture(rsTmp, "��Ƭ", "c:\" & rsTmp!��ԱID & ".bmp")
                    
                    Call TDKJ_BJ_IV(gintPort, "&B0,0,c:\" & rsTmp!��ԱID & ".bmp$")
                    zlCommFun.StopFlash
                End If
                gBlnPic = True
            End If
            Call TDKJ_BJ_IV(gintPort, "&Sc$")
''            Call TDKJ_BJ_IV(gintPort, "&C21 ��ӭ����Ժ����$")
''            Call TDKJ_BJ_IV(gintPort, "&C31  ף�����տ���$")
''            Call TDKJ_BJ_IV(gintPort, "&C41  �շ�Ա:" & strDisplay & "$")
    End Select
    Exit Sub
ErrHand:
    MsgBox "�ӿڵ���ʧ��:" & vbCrLf & vbCrLf & Err.Description, vbInformation, gstrSysName
End Sub

Public Sub DisplayPatient(ByVal strPatient As String, Optional ByVal lngPatient As Long)
'����:��ʾ��������
'����:strPatient-��������,���Ϊ��,��ʾ��һ�ε���ʱ��ʾ�����ڿ���,����������(������е����Ĳ���,����:���� �Ա� ����)
'����:lngPatient-�е����Ĳ��˲Ŵ���,����Ϊ0
    Dim arrName As Variant
    If gblnDDisplay Or gintDevice = Dev_DDisplay Then
        With frmDisplay
            If Not .Visible Then
                SetParent .hWnd, HWND_DESKTOP
                ShowWindow .hWnd, SW_SHOWNOACTIVATE
                SetWindowPos .hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
            End If
            
            If strPatient = "" Then '��ʾ���ڿ���
                Call .ShowMain(False)
                Call .ShowFee(False)
                
                .Timer1.Enabled = False
                .lblFree.Visible = True
                .lblPause.Visible = False
            Else
                Call .ShowMain(True)
                
                .lblPatient.Caption = strPatient
                Call .ShowFee(False)
            End If
        End With
    End If
    Select Case gintDevice
        Case Dev_FS_YL01
            If strPatient <> "" And strPatient <> "0" Then
                arrName = Split(strPatient, " ")
                Call Dev_FS_YL01_Voice(arrName(0), 0, 0)
            End If
        Case Dev_TDKJ_BJ
            Call TDKJ_BJ_FUN(gintPort, "&Sc$")
            Call TDKJ_BJ_FUN(gintPort, "&C21 ��ӭ����Ժ����$")
            Call TDKJ_BJ_FUN(gintPort, "&C31  ף�����տ���$")
            Call TDKJ_BJ_FUN(gintPort, "&C41  " & strPatient & "$")
        Case Dev_TDKJ_BJ_IV
            Call TDKJ_BJ_IV(gintPort, "&Sc$")
            Call TDKJ_BJ_IV(gintPort, "&C21 ��ӭ����Ժ����$")
            Call TDKJ_BJ_IV(gintPort, "&C31  ף�����տ���$")
            Call TDKJ_BJ_IV(gintPort, "&C41  " & strPatient & "$")
    End Select
    
End Sub

Public Sub Display(ByVal strName As String, ByVal strType As String, ByVal strUnit As String, _
    ByVal dblPrice As Double, ByVal dblNumber As Double, ByVal curMoney As Currency)
'���ܣ�LED��ʾ����������
'������  strName:��Ŀ  strType: ���   strUnit:������λ
        'dblPrice:���� dblNumber:����  curMoney:���
 '��: ҩ������ù�� 14�q00Ԫ
     '��Ŀ������� 7�q00Ԫ
   
    On Error Resume Next '���ܴ���һ����ʾ���µ����
    Select Case gintDevice
        Case Dev_SYC_XII
            SycVoice "$1"
            SycVoice "#" & "��Ŀ��" & Mid(strName, 1, 16) & "#"
            SycVoice "$2"
            SycVoice "#" & "������" & dblNumber & " " & "��" & Format(curMoney, "0.00") & "Ԫ" & "#"
        Case Dev_SYC_Q9
            Call SYC_Q9(gintPort, "$1")
            Call SYC_Q9(gintPort, "#" & "��Ŀ��" & Mid(strName, 1, 16) & "#")
            Call SYC_Q9(gintPort, "$2")
            Call SYC_Q9(gintPort, "#" & "������" & dblNumber & " " & "��" & Format(curMoney, "0.00") & "Ԫ" & "#")
        Case Dev_LK822
            If Len(strName) > 15 Then
                strName = Mid(strName, 1, 15)
            End If
            ShowLED "��Ŀ��" & strName, "������" & dblNumber & "  ��λ��" & strUnit, "��" & Format(curMoney, "0.00") & "Ԫ", gstrBottom
        Case Dev_NJF_VH
            
            Set gobjLED = CreateObject("CTSVR.Bjq")
            If Not gobjLED Is Nothing Then
                gobjLED.Comport = gintPort
                gobjLED.DispMode = 0
                If dblNumber = 1 And strUnit = "" Then
                    '���շ�ģʽ
                    gobjLED.Display "~��Ŀ:" & strName
                    gobjLED.Display "^���:" & FormatEx(curMoney, 2) & "Ԫ"
                Else
                    gobjLED.Display "~" & strName & "," & FormatEx(dblNumber, 4) & strUnit
                    gobjLED.Display "^" & "����:" & FormatEx(dblPrice, 3) & "Ԫ,���:" & FormatEx(curMoney, 2) & "Ԫ"
                End If
                Set gobjLED = Nothing
            End If
        Case Dev_TDKJ_BJ
            '������ʾ4�У�ÿ��8������
            Call TDKJ_BJ_FUN(gintPort, "&Sc$") '����
            Call TDKJ_BJ_FUN(gintPort, "&C11" & SetLength(strName, 16) & "$")
            Call TDKJ_BJ_FUN(gintPort, "&C21" & SetLength("����:" & FormatEx(dblNumber, 5) & " " & strUnit, 16) & "$")
            Call TDKJ_BJ_FUN(gintPort, "&C31" & SetLength("����:" & FormatEx(dblPrice, 5) & " Ԫ", 16) & "$")
            Call TDKJ_BJ_FUN(gintPort, "&C41" & SetLength("���:" & FormatEx(curMoney, 5) & " Ԫ", 16) & "$")
        Case Dev_TDKJ_BJ_2008
            '������ʾ4�У�ÿ��8������
            Call TDKJ_BJ_2008(gintPort, "&Sc$") '����
            Call TDKJ_BJ_2008(gintPort, "&C11" & SetLength(strName, 16) & "$")
            Call TDKJ_BJ_2008(gintPort, "&C21" & SetLength("����:" & FormatEx(dblNumber, 5) & " " & strUnit, 16) & "$")
            Call TDKJ_BJ_2008(gintPort, "&C31" & SetLength("����:" & FormatEx(dblPrice, 5) & " Ԫ", 16) & "$")
            Call TDKJ_BJ_2008(gintPort, "&C41" & SetLength("���:" & FormatEx(curMoney, 5) & " Ԫ", 16) & "$")
        Case Dev_MDT_SD01
            '������ʾ4��,ÿ��10������
            Clear_Screen
            Display_Line SetLength(strName, 20), 0, 0
            Display_Line SetLength("���:" & strType & " ��λ:" & strUnit, 20), 0, 1
            Display_Line SetLength("����:" & FormatEx(dblPrice, 5) & "Ԫ,����:" & FormatEx(dblNumber, 5), 20), 0, 2
            Display_Line SetLength("���:" & FormatEx(curMoney, 5) & " Ԫ", 20), 0, 3
        Case Dev_surpass
             '������
            'Call AllClear
            'Call StringDisplay("�ϼƣ�" & Format(curMoney, "0.00") & "Ԫ" + Chr(0), 0, 0)
           ' Call StringDisplay("ʵ�գ�" & Format(curMoney, "0.00") & "Ԫ" + Chr(0), 0, 0)
            'Call StringDisplay("������" & Format(curMoney, "0.00") & "Ԫ" + Chr(0), 0, 0)
        Case Dev_TDKJ_BJ_IV
''            '������ʾ4�У�ÿ��8������
            Call TDKJ_BJ_IV(gintPort, "&Sc$") '����
''            Call TDKJ_BJ_IV(gintPort, "&C11" & SetLength(strName, 16) & "$")
''            Call TDKJ_BJ_IV(gintPort, "&C21" & SetLength("����:" & FormatEx(dblNumber, 5) & " " & strUnit, 16) & "$")
''            Call TDKJ_BJ_IV(gintPort, "&C31" & SetLength("����:" & FormatEx(dblPrice, 5) & " Ԫ", 16) & "$")
''            Call TDKJ_BJ_IV(gintPort, "&C41" & SetLength("���:" & FormatEx(curMoney, 5) & " Ԫ", 16) & "$")
    End Select
End Sub

Public Sub Speak(ByVal strSpeak As String)
'���ܣ��������ţ�ͬʱLED��ʾ
'������
'   strSpeak:
'           #+��� �� #+���+�ո�+���    ��ʾ�����ض���Ϣ
'           ~+����                        ��ʾ��������
'��Ŷ�Ӧ���ܣ�
'#0  --����������(��������)
'#1  --����,���Ե�
'#2  --лл
'#3  --�뵱�����, лл!
'#4  --������������
'#5  --������ʾ�ſ�
'#6  --��������ҩ������
'#7  --������X��������
'#8  --������ע������Ƥ��
'#9  --����������칫����˸���
'#10 --�������Һ������������
'#11 --������ʾ����֤��ҽ��ƾ֤
'#12 --������ʾ����֤�͹���ҽ��ƾ֤
'#13 --������ʾҽ��ƾ֤�͹���ҽ��ƾ֤
'#14 --��������ʲô��
'#15 --�������ǳ��ﻹ�Ǹ���
'#16 --��������ר�����ﻹ����ͨ����
'#17 --������Ԥ��, Ȼ���ٹҺ�
'#18 --��������ò�����
'#19 --������ʾ������
'#20 --������B��������
'#21 1234.56   --��������һǧ������ʮ�ĵ�����Ԫ  J
'#22 1234.56   --Ԥ��һǧ������ʮ�ĵ�����Ԫ Y
'#23 1234.56   --����һǧ������ʮ�ĵ�����Ԫ Z

'#24 --�����ʾ�籣��
'#25 1234.56   --��ķ���ΪXXXXԪ
'#26 1234.56   --�������XXXXԪ
'#27 1234.56   --��Ŀ��������븶�ֽ�XXXXԪ
'#28 --��������ҽ�����ݼ���
'#30  ���ʾ���￨::32663

'#50 --������ʾҽ��ƾ֤
'#51 --�����շ�:����������

    If Len(strSpeak) < 2 Then Exit Sub
    If Left(strSpeak, 1) <> "#" And Left(strSpeak, 1) <> "~" Then Exit Sub
    
    Select Case gintDevice
        Case Dev_SYC_XII
            ContrastSYC_XII strSpeak
        Case Dev_LK822
        
        Case Dev_SHY_II
            ContrastSHY_II strSpeak
        Case Dev_NJF_VH
            Contrast_NJF_VH strSpeak
        Case Dev_TDKJ_BJ
            Dev_TDKJ_BJ_Speak strSpeak
        Case Dev_TDKJ_BJ_2008
            Dev_TDKJ_BJ_2008_Speak strSpeak
        Case Dev_MDT_SD01
            Dev_MDT_SD01_Speak strSpeak
        Case Dev_surpass
            Dev_surpass_speak strSpeak
        Case Dev_FS_YL01
            If InStr(1, strSpeak, "#22 ") > 0 Then
                If IsNumeric(Mid(strSpeak, 5)) Then
                    Call Dev_FS_YL01_Voice(Mid(strSpeak, 5), 1, 0)
                End If
            End If
        Case Dev_TDKJ_BJ_IV
            Dev_TDKJ_BJ_IV_Speak strSpeak
        Case Dev_SYC_Q9
            ContrastSYC_Q9 strSpeak
    End Select
    
    If gblnDDisplay Or gintDevice = Dev_DDisplay Then
        If InStr(1, strSpeak, "#21 ") > 0 Then
        'Ӧ�ɻ�ý���ʱ����
            With frmDisplay
                .ShowFee (True)
                .lblCash(0).Caption = Replace(strSpeak, "#21 ", "") & "Ԫ"
                .lblCash(1).Caption = ""
                .lblCash(2).Caption = ""
                .lblDrugWindow.Caption = ""
            End With
        End If
    End If
End Sub

Public Sub DispCharge(ByVal dblӦ�� As Double, ByVal dblʵ�� As Double, ByVal dbl�Ҳ� As Double)
   '����:��ʾӦ��,ʵ��,�Ҳ�,������
   Dim dbl�ϼ� As Double, dbl���� As Double
   Dim strӦ�� As String, strʵ�� As String, str���� As String, str���� As String, str�븶�� As String
    Select Case gintDevice
        Case Dev_SYC_XII
            
        Case Dev_SYC_Q9
            
        Case Dev_LK822
            ShowLED "Ӧ�գ�" & Format(dblӦ��, "0.00") & "Ԫ", "ʵ�գ�" & Format(dblʵ��, "0.00") & "Ԫ", "�Ҳ���" & Format(dbl�Ҳ�, "0.00") & "Ԫ", gstrBottom
        Case Dev_SHY_II
            
        Case Dev_NJF_VH
            On Error Resume Next
            Set gobjLED = CreateObject("CTSVR.Bjq")
            If Not gobjLED Is Nothing Then
                gobjLED.Comport = gintPort
                gobjLED.DispMode = 0
                gobjLED.Display "~�ϼ�:" & Format(dblӦ��, "0.00") & "Ԫ^����:" & Format(dblʵ��, "0.00") & "Ԫ" & IIf(dbl�Ҳ� = 0, "", ",����:" & Format(dbl�Ҳ�, "0.00") & "Ԫ")
                Set gobjLED = Nothing
            End If
        Case Dev_TDKJ_BJ
            '������ʾ4�У�ÿ��8������
             Call TDKJ_BJ_FUN(gintPort, "&Sc$") '����
        Case Dev_TDKJ_BJ_2008
            '������ʾ4�У�ÿ��8������
             Call TDKJ_BJ_2008(gintPort, "&Sc$")  '����
'            Call TDKJ_BJ_FUN(gintPort, "&C11" & SetLength("Ӧ��:" & FormatEx(dblӦ��, 5), 16) & "$")
'            Call TDKJ_BJ_FUN(gintPort, "&C21" & SetLength("ʵ��:" & FormatEx(dblʵ��, 5), 16) & "$")
'            Call TDKJ_BJ_FUN(gintPort, "&C31" & SetLength("�Ҳ�:" & FormatEx(dbl�Ҳ�, 5), 16) & "$")
        Case Dev_MDT_SD01
'            Clear_Screen
'            Price CStr(dblӦ��)
'            GetPrice CStr(dblʵ��)
'            Check CStr(dbl�Ҳ�)
        Case Dev_surpass
             '������
             Call AllClear
             '��ʾ����
             dbl�ϼ� = dblӦ��
             dblʵ�� = dblʵ��
             dbl���� = dbl�Ҳ�
             strӦ�� = "Ӧ��.wav"
             strʵ�� = "Ԥ��.wav"
             str���� = "����.wav"
             str�븶�� = "��������.wav"
             str���� = "�����뵱�����лл.wav"
             'Call StringDisplay("�ϼƣ�" & Format(dblӦ��, "0.00") & "Ԫ" + Chr(0), 0, 0)
             Call LocStringDisplay(2, 2, "Ӧ�գ�" & Format(dblӦ��, "0.00") & "Ԫ" + Chr(0))
             'Call StringDisplay("ʵ�գ�" & Format(dblʵ��, "0.00") & "Ԫ" + Chr(0), 0, 0)
             Call LocStringDisplay(2, 22, "Ԥ�գ�" & Format(dblʵ��, "0.00") & "Ԫ" + Chr(0))
             'Call StringDisplay("������" & Format(dbl�Ҳ�, "0.00") & "Ԫ" + Chr(0), 0, 0)
             Call LocStringDisplay(2, 42, "���㣺" & Format(dbl�Ҳ�, "0.00") & "Ԫ" + Chr(0))
             'Call PlayWaves(str����)
             'Call PlayWaves(App.Path & "\" & inputwave.Text)
'             Call PlayWaves(App.Path & "\" & str�븶��)
'             Call RMB2Wav(dblӦ��)
'             'Call PlayWaves(App.Path & "\" & strӦ��)
'             Call PlayWaves(App.Path & "\" & strʵ��)
'             Call RMB2Wav(dblʵ��)
'             Call PlayWaves(App.Path & "\" & str����)
'             Call RMB2Wav(dbl�Ҳ�)
'             Call PlayWaves(App.Path & "\" & str����)
        Case Dev_FS_YL01
            Call Dev_FS_YL01_Voice(dblӦ��, 1, IIf(dblӦ�� > 1000, 5, 3))
            Call Dev_FS_YL01_Voice(dblʵ��, 2, 3)
            Call Dev_FS_YL01_Voice(dbl�Ҳ�, 3, 0)
        Case Dev_TDKJ_BJ_IV
            '������ʾ4�У�ÿ��8������
''             Call TDKJ_BJ_IV(gintPort, "&Sc$") '����
            Dev_TDKJ_BJ_IV_Speak "#21" & FormatEx(dblӦ��, 5)
'            Call TDKJ_BJ_FUN(gintPort, "&C21" & SetLength("ʵ��:" & FormatEx(dblʵ��, 5), 16) & "$")
'            Call TDKJ_BJ_FUN(gintPort, "&C31" & SetLength("�Ҳ�:" & FormatEx(dbl�Ҳ�, 5), 16) & "$")
    End Select
    
    If gblnDDisplay Or gintDevice = Dev_DDisplay Then
        '��ʾ�շ���Ϣ
        With frmDisplay
            .ShowFee (True)
            .lblCash(0).Caption = Format(dblӦ��, "0.00") & "Ԫ"
            .lblCash(1).Caption = Format(dblʵ��, "0.00") & "Ԫ"
            .lblCash(2).Caption = Format(dbl�Ҳ�, "0.00") & "Ԫ"
            .lblDrugWindow.Caption = ""
            
            .lblThanks.Caption = .lblThanks.Tag
            
            '���õȴ���ʱ,����û��ҩ����Ϣ
            .Timer1.Enabled = True
        End With
    End If
End Sub

Public Sub DisplayBank(ParamArray Bank())
    '���ܣ�������������ʾ�����ַ�
    '������bank() Ϊ��������Ĳ���,���һ��������û�У���UBOUND��bank)Ϊһ
    'û�з���ֵ
    Dim intBound As Integer, i As Integer, arrTmp As Variant, strTmp As String
        
    intBound = UBound(Bank)
    If intBound = -1 Then Exit Sub
        
    On Error Resume Next
    
    Select Case gintDevice
        Case Dev_SYC_XII
            SycVoice "*"            '����
            For i = 0 To intBound
                If i < 2 Then
                    SycVoice "$" & i + 1
                    SycVoice "#" & Bank(i) & "#"
                End If
            Next
        Case Dev_SYC_Q9
            Call SYC_Q9(gintPort, "*")            '����)
            For i = 0 To intBound
                If i < 2 Then
                    Call SYC_Q9(gintPort, "$" & i + 1)
                    Call SYC_Q9(gintPort, "#" & Bank(i) & "#")
                End If
            Next
        Case Dev_LK822
            For i = 0 To intBound
                If Len(Bank(i)) > 15 Then
                    Bank(i) = Mid(Bank(i), 1, 15)
                End If
            Next
            If intBound = 0 Then
                ShowLED Bank(0), "", "", gstrBottom
            ElseIf intBound = 1 Then
                ShowLED Bank(0), Bank(1), "", gstrBottom
            ElseIf intBound >= 2 Then
                ShowLED Bank(0), Bank(1), Bank(2), gstrBottom
            End If
        Case Dev_SHY_II
            
        Case Dev_NJF_VH
            Set gobjLED = CreateObject("CTSVR.Bjq")
            If Not gobjLED Is Nothing Then
                gobjLED.Comport = gintPort
                gobjLED.DispMode = 0
                For i = 0 To IIf(intBound > 1, 1, intBound)
                    If i = 0 Then
                        gobjLED.Display "~" & Bank(i)
                    Else
                        gobjLED.Display "^" & Bank(i)
                    End If
                Next
                Set gobjLED = Nothing
            End If
        Case Dev_TDKJ_BJ
            '������ʾ4�У�ÿ��8������
            Call TDKJ_BJ_FUN(gintPort, "&Sc$") '����
            For i = 0 To intBound
                If i < 4 Then
                    Call TDKJ_BJ_FUN(gintPort, "&C" & i + 1 & "1" & SetLength(CStr(Bank(i)), 16) & "$")
                End If
            Next
        Case Dev_TDKJ_BJ_2008
            '������ʾ4�У�ÿ��8������
            Call TDKJ_BJ_2008(gintPort, "&Sc$")  '����
            For i = 0 To intBound
                If i < 4 Then
                    Call TDKJ_BJ_2008(gintPort, "&C" & i + 1 & "1" & SetLength(CStr(Bank(i)), 16) & "$")
                End If
            Next
        Case Dev_MDT_SD01
            Clear_Screen
           For i = 0 To intBound
                If i < 4 Then
                    Display_Line SetLength(CStr(Bank(i)), 20), 0, i
                End If
           Next
'        Case Dev_surpass
             '������
'             Call AllClear
'             For i = 0 To intBound
'                If i < 4 Then
'                    Call StringDisplay("& SetLength(CStr(Bank(i)) &" + Chr(0), 0, 0)
'                End If
'             Next
        Case Dev_TDKJ_BJ_IV
            
    End Select
    
    If gblnDDisplay Or gintDevice = Dev_DDisplay Then
        With frmDisplay
            '��ʾҩ��������Ϣ
            '�շѵ��ô���:zl9LedVoice.DisplayBank "���úϼ�:" & txt�ϼ�.Text, "ȡҩ����:" & IIf(mstr���� <> "", " " & mstr����, "") & IIf(mstr�ɴ� <> "", " " & mstr�ɴ�, "") & IIf(mstr�д� <> "", " " & mstr�д�, "")
            If UBound(Bank) >= 1 Then
                If InStr(1, Bank(1), "ȡҩ����") > 0 Then
                    If .Timer1.Enabled Then .Timer1.Enabled = False
                    If .lblFree.Visible Then .lblFree.Visible = False
                    If .lblPause.Visible Then .lblPause.Visible = False
                    If .lblWaiting.Visible Then .lblWaiting.Visible = False
                    
                    arrTmp = Split(Replace(Bank(1), "ȡҩ����: ", ""), " ")
                    For i = 0 To UBound(arrTmp)
                        strTmp = strTmp & IIf(strTmp = "", "", vbCrLf) & Choose(i + 1, "������ҩ��", "�����ҩ��", "������ҩ��") & arrTmp(i)
                    Next
                    If strTmp = "" Then
                        .lblDrugWindow.Caption = ""
                    Else
                        .lblDrugWindow.Caption = Replace(.lblDrugWindow.Tag, "&Windows", strTmp)
                    End If
                    
                    '���õȴ���ʱ
                    .Timer1.Enabled = True
                    Exit Sub
                End If
            End If
            
            '����֧���ֽ�
            '�շѵ��ô���:Call zl9LedVoice.DisplayBank("�ϼ�:" & txt�ϼ�.Text & "Ԫ,Ӧ��:" & txtӦ��.Text & "Ԫ","����:" & txt�ɿ�.Text & "Ԫ" & IIf(Val(txt�Ҳ�.Text) = 0, "", ",����:" & txt�Ҳ�.Text & "Ԫ"))
            If UBound(Bank) >= 0 Then
                strTmp = Bank(0)
                If InStr(1, strTmp, "Ӧ��") > 0 Then
                    .ShowFee (True)
                    arrTmp = Split(strTmp, ",")
                    .lblCash(0).Caption = Replace(arrTmp(1), "Ӧ��:", "")
                    .lblCash(1).Caption = Replace(arrTmp(2), "����:", "")
                    If UBound(arrTmp) >= 3 Then
                        .lblCash(2).Caption = Replace(arrTmp(3), "����:", "")
                    Else
                        .lblCash(2).Caption = "0Ԫ"
                    End If
                    .lblDrugWindow.Caption = ""
                End If
            End If
        End With
    End If
End Sub

Public Sub VoiceSetting()
    '''''''''''''''''''''''''''''''''''''''''''
    '���ܣ����ò�������
    '�������޲���
    '���أ�
    '''''''''''''''''''''''''''''''''''''''''''
    frmSetting.Show vbModal
End Sub

Private Sub Class_Initialize()
    blnInit_Dev_MDT_SD01 = False
    blnDev_surpass = False
End Sub

Private Sub Class_Terminate()
    If blnInit_Dev_MDT_SD01 = True Then
        CloseDevice
        CloseService
    End If
    If blnDev_surpass = True Then
        Call AllClear
        SetComNo (0) '�رմ���
    End If
    If gblnDDisplay Or gintDevice = Dev_DDisplay Then
        '�����������������ʾ��ӭ��Ϣ,����ñ�����,Ȼ���ٵ���displaypatient
         Call Unload(frmDisplay)
    End If
End Sub

