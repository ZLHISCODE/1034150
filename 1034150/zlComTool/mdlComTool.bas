Attribute VB_Name = "mdlComTool"
Option Explicit
Public Const HKEY_CURRENT_USER = &H80000001
Public Const SWP_NOMOVE = 2
Public Const SWP_NOSIZE = 1
Public Const SWP_SHOWWINDOW = &H40
Public Const flags = SWP_NOMOVE Or SWP_NOSIZE
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SMTO_ABORTIFHUNG = &H2
Public Const BDR_RAISEDINNER = &H4
Public Const BDR_RAISEDOUTER = &H1
Public Const BDR_SUNKENINNER = &H8
Public Const BDR_SUNKENOUTER = &H2
Public Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Public Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)

Public Const BF_BOTTOM = &H8
Public Const BF_LEFT = &H1
Public Const BF_RIGHT = &H4
Public Const BF_TOP = &H2
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

Public Const LVM_FIRST = &H1000
Public Const LVM_SETCOLUMNWIDTH = LVM_FIRST + 30
Public Const LVSCW_AUTOSIZE = -1
Public Const LVSCW_AUTOSIZE_USEHEADER = -2

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Public Const SW_SHOWNORMAL = 1
Public Const SW_SHOWNOACTIVATE = 4
Public Declare Function SetParent Lib "User32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function ShowWindow Lib "User32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function BringWindowToTop Lib "User32" (ByVal hWnd As Long) As Long
Public Declare Function SetActiveWindow Lib "User32" (ByVal hWnd As Long) As Long
Public Declare Function SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Declare Function DrawEdge Lib "User32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Public Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function LockWindowUpdate Lib "User32" (ByVal hWndLock As Long) As Long
Public Declare Function SetCapture Lib "User32" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseCapture Lib "User32" () As Long
Public Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

'����ϵͳ�п��õ����뷨�����������뷨����Layout,����Ӣ�����뷨��
Public Declare Function GetKeyboardLayoutList Lib "User32" (ByVal nBuff As Long, lpList As Long) As Long
'��ȡĳ�����뷨������
Public Declare Function ImmGetDescription Lib "imm32.dll" Alias "ImmGetDescriptionA" (ByVal hkl As Long, ByVal lpsz As String, ByVal uBufLen As Long) As Long
'�ж�ĳ�����뷨�Ƿ��������뷨
Public Declare Function ImmIsIME Lib "imm32.dll" (ByVal hkl As Long) As Long

Public Declare Function GetWindowRect Lib "User32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public gcnOracle As New ADODB.Connection        '�������ݿ����ӣ��ر�ע�⣺��������Ϊ�µ�ʵ��
Public gstrPrivs As String                   '��ǰ�û����еĵ�ǰģ��Ĺ���

Public gstrSysName As String                'ϵͳ����
Public gstrVersion As String                'ϵͳ�汾
Public gstrAviPath As String                'AVI�ļ��Ĵ��Ŀ¼

Public gstrDbUser As String                 '��ǰ���ݿ��û�
Public glngUserId As Long                   '��ǰ�û�id
Public gstrUserCode As String               '��ǰ�û�����
Public gstrUserName As String               '��ǰ�û�����
Public gstrUserAbbr As String               '��ǰ�û�����

Public glngDeptId As Long                   '��ǰ�û�����id
Public gstrDeptCode As String               '��ǰ�û����ű���
Public gstrDeptName As String               '��ǰ�û���������

Public gstr��λ���� As String
Public gstr֧���̼��� As String
Public gstrSQL As String
Public gstrMenuSys As String                '��ǰ�û�ʹ�õĲ˵�ϵͳ
Public glngSys As Long                      '��ǰϵͳ
Public gstrOwner As String                  '��ǰϵͳ������
'��������ϢϵͳҪ�õ���ȫ�ֱ���
Public gfrmMain As Object                   '����̨���ڣ���Ҫ��������Ϣ�༭���ڵĸ�����
Public gblnMessageShow As Boolean           '˵����Ϣ�������Ƿ��Ѿ���ʾ
Public gblnMessageGet  As Boolean           '˵������̨�Ƿ�Ҫ��֪ͨ���ʼ�

Public Const glngLBound As Long = 99
Public Const glngUBound As Long = 240

'�������ݲ˵�ID����
'********************************************************************
Public Const conMenu_FilePopup = 1              '�ļ�
Public Const conMenu_EditPopup = 3              '�༭
Public Const conMenu_ReportPopup = 4            '����
Public Const conMenu_FormatPopup = 5              '��ʽ
Public Const conMenu_ViewPopup = 7              '�鿴
Public Const conMenu_ActionPopup = 8             '����
Public Const conMenu_HelpPopup = 9              '����

Public Const conTool_System = 50              'ϵͳ
'�ļ��˵�
Public Const conMenu_File_PrintSet = 101        '��ӡ����(&S)��
Public Const conMenu_File_Preview = 102         'Ԥ��(&V)
Public Const conMenu_File_Print = 103           '��ӡ(&P)
Public Const conMenu_File_Excel = 104           '�����&Excel��
Public Const conMenu_File_Send = 151            '����
Public Const conMenu_File_Save = 161            '����
Public Const conMenu_File_SaveAs = 162     '���Ϊ   2006-08-18 add by �¶�

Public Const conMenu_File_Exit = 191            '�˳�(&X)

'�༭�˵�
Public Const conMenu_Edit_AddGroup = 30101       '���ӷ���(&G)
Public Const conMenu_Edit_ModifyGroup = 30102          '�޸ķ���(&U)
Public Const conMenu_Edit_DeleteGroup = 30103          'ɾ������(&E)
Public Const conMenu_Edit_Add = 301       '����(&A)
Public Const conMenu_Edit_Modify = 303          '��(&O)
Public Const conMenu_Edit_Delete = 304          'ɾ��(&D)
Public Const conMenu_Edit_Reuse = 305           '��ԭ(&R)

Public Const conMenu_Edit_Reply = 306               '��
Public Const conMenu_Edit_AllReply = 307              'ȫ����
Public Const conMenu_Edit_Transmit = 308              'ת��

Public Const conMenu_Edit_Cut = 310      '����
Public Const conMenu_Edit_Copy = 311     '����
Public Const conMenu_Edit_plaster = 312  'ճ��

Public Const conMenu_Edit_Clear = 320      '���
Public Const conMenu_Edit_CheckAll = 321      'ȫѡ

Public Const conMenu_Edit_setDefault = 330  '��Ϊȱʡ
'��ʽ�˵�
Public Const conMenu_Format_Font = 501      '����
Public Const conMenu_Format_SIZE = 502      '�����С
Public Const conMenu_FORMAT_BOLD = 503      '����
Public Const conMenu_FORMAT_ITALIC = 504      'б��
Public Const conMenu_FORMAT_UNDERLINE = 505      '�»���


Public Const conMenu_Format_ForeColor = 511      '����ɫ
Public Const conMenu_Format_FillColor = 512      '����ɫ

Public Const conMenu_Format_Sig = 521      '��Ŀ����
Public Const conMenu_Format_Left = 522      '�����
Public Const conMenu_Format_Center = 523     '����
Public Const conMenu_Format_Right = 524      '�Ҷ���

Public Const conMenu_Format_Decrease = 531      '��������
Public Const conMenu_Format_Increase = 532      '��������


'�鿴�˵�
Public Const conMenu_View_ToolBar = 701              '������(&T)
Public Const conMenu_View_ToolBar_Button = 7011         '��׼��ť(&S)
Public Const conMenu_View_ToolBar_Text = 7012           '�ı���ǩ(&T)
Public Const conMenu_View_ToolBar_Size = 7013           '��ͼ��(&B)
Public Const conMenu_View_StatusBar = 702            '״̬��(&S)

Public Const conMenu_View_XP = 705            'xp���
Public Const conMenu_View_OutLook = 706            'OutLook���

Public Const conMenu_View_Expend = 711               'չ��/�۵���(&X)
Public Const conMenu_View_Expend_AllCollapse = 7111     '�۵�������(&L)
Public Const conMenu_View_Expend_AllExpend = 7112       'չ��������(&X)
Public Const conMenu_View_Expend_CurCollapse = 7113     '�۵���ǰ��(&C)
Public Const conMenu_View_Expend_CurExpend = 7114       'չ����ǰ��(&E)

Public Const conMenu_View_PreviewWindow = 705            'Ԥ������(&P)
Public Const conMenu_View_ShowAll = 706            '��ʾ�Ѷ�(&E)
Public Const conMenu_View_Login = 707            'δ���ʼ�����(&W)

Public Const conMenu_View_Find = 722                 '����(&F)
Public Const conMenu_View_FindNext = 723             '������һ��(&F)

Public Const conMenu_View_BigIcon = 730             '��ͼ��(&G)
Public Const conMenu_View_MiniIcon = 731             'Сͼ��(&M)
Public Const conMenu_View_List = 732             '�б�(&L)
Public Const conMenu_view_Report = 733           '��ϸ����(&L)

Public Const conMenu_View_Refresh = 791              'ˢ��(&R)

Public Const conMenu_Action_Hight = 801        '��
Public Const conMenu_Action_Low = 802         '��
'�����˵�
Public Const conMenu_Help_Help = 901        '��������(&H)
Public Const conMenu_Help_Web = 902         '&WEB�ϵ�����
Public Const conMenu_Help_Web_Home = 9021       '������ҳ(&H)
Public Const conMenu_Help_Web_Mail = 9022       '���ͷ���(&M)
Public Const conMenu_Help_Web_Forum = 9023      '������̳(&F)
Public Const conMenu_Help_About = 991       '����(&A)��

'������������
'********************************************************************
'CommandBar���г�������
Public Const XTP_ID_WINDOW_LIST = 35000 '�����б�
Public Const XTP_ID_TOOLBARLIST = 59392 '�������б�
Public Const ID_INDICATOR_CAPS = 59137 '״̬������д��
Public Const ID_INDICATOR_NUM = 59138 '״̬�������֣�
Public Const ID_INDICATOR_SCRL = 59139 '״̬����������

'CommandBar�����ȼ�
Public Const FSHIFT = 4
Public Const FCONTROL = 8
Public Const FALT = 16

'********************************************************************
Public Sub GetUserInfo()
'����:�õ��û�����Ϣ

    Dim rsTemp As New ADODB.Recordset
    Dim strSQL  As String
    
    rsTemp.CursorLocation = adUseClient
    On Error GoTo errHand
    Set rsTemp = zlDatabase.GetUserInfo
    With rsTemp
        If .RecordCount > 0 Then
            glngUserId = .Fields("ID").Value                '��ǰ�û�id
            gstrUserCode = .Fields("���").Value            '��ǰ�û�����
            gstrUserName = .Fields("����").Value            '��ǰ�û�����
            gstrUserAbbr = IIf(IsNull(.Fields("����").Value), "", .Fields("����").Value)          '��ǰ�û�����
            glngDeptId = .Fields("����id").Value            '��ǰ�û�����id
            gstrDeptCode = .Fields("������").Value        '��ǰ�û�
            gstrDeptName = .Fields("������").Value        '��ǰ�û�
        Else
            glngUserId = 0
            gstrUserCode = ""
            gstrUserName = ""
            gstrUserAbbr = ""
            glngDeptId = 0
            gstrDeptCode = ""
            gstrDeptName = ""
        End If
        .Close
    End With
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Err = 0
End Sub

'Public Sub OpenRecordset(rsTemp As ADODB.Recordset, ByVal strFormCaption As String)
''���ܣ��򿪼�¼��ͬʱ����SQL���
'    If rsTemp.State = adStateOpen Then rsTemp.Close
'
'    Call SQLTest(App.ProductName, strFormCaption, gstrSQL)
'    rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
'    Call SQLTest
'End Sub

Private Function SystemImes() As Variant
'���ܣ���ϵͳ�������뷨���Ʒ��ص�һ���ַ���������
'���أ�����������������뷨,�򷵻ؿմ�
    Dim arrIme(99) As Long, arrName() As String
    Dim lngLen As Long, StrName As String * 255
    Dim lngCount As Long, i As Integer, j As Integer

    lngCount = GetKeyboardLayoutList(UBound(arrIme) + 1, arrIme(0))
    For i = 0 To lngCount - 1
        If ImmIsIME(arrIme(i)) = 1 Then 'Ϊ1��ʾ�������뷨
            ReDim Preserve arrName(j)
            lngLen = ImmGetDescription(arrIme(i), StrName, Len(StrName))
            arrName(j) = Mid(StrName, 1, InStr(1, StrName, Chr(0)) - 1)
            j = j + 1
        End If
    Next
    SystemImes = IIf(j > 0, arrName, vbNullString)
End Function

Public Function ChooseIME(cmbIME As Object) As Boolean
    Dim varIME As Variant
    Dim i As Integer
    Dim strIme As String
    
    varIME = SystemImes
    If Not IsArray(varIME) Then
        MsgBox "�㻹û��װ�κκ������뷨������ʹ�ñ����ܡ�" & vbCrLf & _
               "���뷨�İ�װ���ڿ����������ɡ�", vbInformation, gstrSysName
        Exit Function
    End If
    cmbIME.Clear
    cmbIME.AddItem "���Զ�����"
    strIme = zlDatabase.GetPara("���뷨")
    For i = LBound(varIME) To UBound(varIME)
        cmbIME.AddItem varIME(i)
        If strIme = varIME(i) Then cmbIME.ListIndex = i + 1
    Next
    If cmbIME.ListIndex < 0 Then cmbIME.ListIndex = 0
    ChooseIME = True
End Function

Public Function GetControlRect(ByVal lngHwnd As Long) As RECT
'���ܣ���ȡָ���ؼ�����Ļ�е�λ��(Twip)
    Dim vRect As RECT
    Call GetWindowRect(lngHwnd, vRect)
    vRect.Left = vRect.Left * Screen.TwipsPerPixelX
    vRect.Right = vRect.Right * Screen.TwipsPerPixelX
    vRect.Top = vRect.Top * Screen.TwipsPerPixelY
    vRect.Bottom = vRect.Bottom * Screen.TwipsPerPixelY
    GetControlRect = vRect
End Function


Public Function NewClientRecord(ByVal strFilds As String) As ADODB.Recordset
    '����һ���յļ�¼��
    'strFilds:�ֶ���,����,����;�ֶ���,����,����...
    '    ����:�û���,varchar2,30;����,varchar2,30
    
    Dim rs As ADODB.Recordset, i As Integer
    Dim varFilds As Variant
    Dim varFild As Variant
    Dim strTmp As String
    Set rs = New ADODB.Recordset
    
    varFilds = Split(strFilds, ";")
    With rs
        For i = LBound(varFilds) To UBound(varFilds)
            strTmp = varFilds(i)
            varFild = Split(strTmp, ",")
            
            If UCase(varFild(1)) = "VARCHAR2" Then
                .Fields.Append varFild(0), adVarWChar, CLng(varFild(2)), adFldIsNullable
            ElseIf UCase(varFild(1)) = "NUMBER" Then
                .Fields.Append varFild(0), adVarNumeric, CLng(varFild(2)), adFldIsNullable
            End If
        Next
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    Set NewClientRecord = rs
End Function

