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

'返回系统中可用的输入法个数及各输入法所在Layout,包括英文输入法。
Public Declare Function GetKeyboardLayoutList Lib "User32" (ByVal nBuff As Long, lpList As Long) As Long
'获取某个输入法的名称
Public Declare Function ImmGetDescription Lib "imm32.dll" Alias "ImmGetDescriptionA" (ByVal hkl As Long, ByVal lpsz As String, ByVal uBufLen As Long) As Long
'判断某个输入法是否中文输入法
Public Declare Function ImmIsIME Lib "imm32.dll" (ByVal hkl As Long) As Long

Public Declare Function GetWindowRect Lib "User32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public gcnOracle As New ADODB.Connection        '公共数据库连接，特别注意：不能设置为新的实例
Public gstrPrivs As String                   '当前用户具有的当前模块的功能

Public gstrSysName As String                '系统名称
Public gstrVersion As String                '系统版本
Public gstrAviPath As String                'AVI文件的存放目录

Public gstrDbUser As String                 '当前数据库用户
Public glngUserId As Long                   '当前用户id
Public gstrUserCode As String               '当前用户编码
Public gstrUserName As String               '当前用户姓名
Public gstrUserAbbr As String               '当前用户简码

Public glngDeptId As Long                   '当前用户部门id
Public gstrDeptCode As String               '当前用户部门编码
Public gstrDeptName As String               '当前用户部门名称

Public gstr单位名称 As String
Public gstr支持商简名 As String
Public gstrSQL As String
Public gstrMenuSys As String                '当前用户使用的菜单系统
Public glngSys As Long                      '当前系统
Public gstrOwner As String                  '当前系统所有者
'以下是消息系统要用到的全局变量
Public gfrmMain As Object                   '导航台窗口，主要用于作消息编辑窗口的父窗口
Public gblnMessageShow As Boolean           '说明消息主窗口是否已经显示
Public gblnMessageGet  As Boolean           '说明导航台是否要求通知新邮件

Public Const glngLBound As Long = 99
Public Const glngUBound As Long = 240

'公共部份菜单ID定义
'********************************************************************
Public Const conMenu_FilePopup = 1              '文件
Public Const conMenu_EditPopup = 3              '编辑
Public Const conMenu_ReportPopup = 4            '报表
Public Const conMenu_FormatPopup = 5              '格式
Public Const conMenu_ViewPopup = 7              '查看
Public Const conMenu_ActionPopup = 8             '动作
Public Const conMenu_HelpPopup = 9              '帮助

Public Const conTool_System = 50              '系统
'文件菜单
Public Const conMenu_File_PrintSet = 101        '打印设置(&S)…
Public Const conMenu_File_Preview = 102         '预览(&V)
Public Const conMenu_File_Print = 103           '打印(&P)
Public Const conMenu_File_Excel = 104           '输出到&Excel…
Public Const conMenu_File_Send = 151            '发送
Public Const conMenu_File_Save = 161            '保存
Public Const conMenu_File_SaveAs = 162     '另存为   2006-08-18 add by 陈东

Public Const conMenu_File_Exit = 191            '退出(&X)

'编辑菜单
Public Const conMenu_Edit_AddGroup = 30101       '增加分类(&G)
Public Const conMenu_Edit_ModifyGroup = 30102          '修改分类(&U)
Public Const conMenu_Edit_DeleteGroup = 30103          '删除分类(&E)
Public Const conMenu_Edit_Add = 301       '增加(&A)
Public Const conMenu_Edit_Modify = 303          '打开(&O)
Public Const conMenu_Edit_Delete = 304          '删除(&D)
Public Const conMenu_Edit_Reuse = 305           '还原(&R)

Public Const conMenu_Edit_Reply = 306               '答复
Public Const conMenu_Edit_AllReply = 307              '全部答复
Public Const conMenu_Edit_Transmit = 308              '转发

Public Const conMenu_Edit_Cut = 310      '剪切
Public Const conMenu_Edit_Copy = 311     '复制
Public Const conMenu_Edit_plaster = 312  '粘贴

Public Const conMenu_Edit_Clear = 320      '清除
Public Const conMenu_Edit_CheckAll = 321      '全选

Public Const conMenu_Edit_setDefault = 330  '设为缺省
'格式菜单
Public Const conMenu_Format_Font = 501      '字体
Public Const conMenu_Format_SIZE = 502      '字体大小
Public Const conMenu_FORMAT_BOLD = 503      '粗体
Public Const conMenu_FORMAT_ITALIC = 504      '斜体
Public Const conMenu_FORMAT_UNDERLINE = 505      '下划线


Public Const conMenu_Format_ForeColor = 511      '字体色
Public Const conMenu_Format_FillColor = 512      '背景色

Public Const conMenu_Format_Sig = 521      '项目符号
Public Const conMenu_Format_Left = 522      '左对齐
Public Const conMenu_Format_Center = 523     '居中
Public Const conMenu_Format_Right = 524      '右对齐

Public Const conMenu_Format_Decrease = 531      '减少缩进
Public Const conMenu_Format_Increase = 532      '增加缩进


'查看菜单
Public Const conMenu_View_ToolBar = 701              '工具栏(&T)
Public Const conMenu_View_ToolBar_Button = 7011         '标准按钮(&S)
Public Const conMenu_View_ToolBar_Text = 7012           '文本标签(&T)
Public Const conMenu_View_ToolBar_Size = 7013           '大图标(&B)
Public Const conMenu_View_StatusBar = 702            '状态栏(&S)

Public Const conMenu_View_XP = 705            'xp风格
Public Const conMenu_View_OutLook = 706            'OutLook风格

Public Const conMenu_View_Expend = 711               '展开/折叠组(&X)
Public Const conMenu_View_Expend_AllCollapse = 7111     '折叠所有组(&L)
Public Const conMenu_View_Expend_AllExpend = 7112       '展开所有组(&X)
Public Const conMenu_View_Expend_CurCollapse = 7113     '折叠当前组(&C)
Public Const conMenu_View_Expend_CurExpend = 7114       '展开当前组(&E)

Public Const conMenu_View_PreviewWindow = 705            '预览窗格(&P)
Public Const conMenu_View_ShowAll = 706            '显示已读(&E)
Public Const conMenu_View_Login = 707            '未读邮件提醒(&W)

Public Const conMenu_View_Find = 722                 '查找(&F)
Public Const conMenu_View_FindNext = 723             '查找下一处(&F)

Public Const conMenu_View_BigIcon = 730             '大图标(&G)
Public Const conMenu_View_MiniIcon = 731             '小图标(&M)
Public Const conMenu_View_List = 732             '列表(&L)
Public Const conMenu_view_Report = 733           '详细资料(&L)

Public Const conMenu_View_Refresh = 791              '刷新(&R)

Public Const conMenu_Action_Hight = 801        '高
Public Const conMenu_Action_Low = 802         '低
'帮助菜单
Public Const conMenu_Help_Help = 901        '帮助主题(&H)
Public Const conMenu_Help_Web = 902         '&WEB上的中联
Public Const conMenu_Help_Web_Home = 9021       '中联主页(&H)
Public Const conMenu_Help_Web_Mail = 9022       '发送反馈(&M)
Public Const conMenu_Help_Web_Forum = 9023      '中联论坛(&F)
Public Const conMenu_Help_About = 991       '关于(&A)…

'其它常量定义
'********************************************************************
'CommandBar固有常量定义
Public Const XTP_ID_WINDOW_LIST = 35000 '窗体列表
Public Const XTP_ID_TOOLBARLIST = 59392 '工具栏列表
Public Const ID_INDICATOR_CAPS = 59137 '状态栏（大写）
Public Const ID_INDICATOR_NUM = 59138 '状态栏（数字）
Public Const ID_INDICATOR_SCRL = 59139 '状态栏（滚动）

'CommandBar辅助热键
Public Const FSHIFT = 4
Public Const FCONTROL = 8
Public Const FALT = 16

'********************************************************************
Public Sub GetUserInfo()
'功能:得到用户的信息

    Dim rsTemp As New ADODB.Recordset
    Dim strSQL  As String
    
    rsTemp.CursorLocation = adUseClient
    On Error GoTo errHand
    Set rsTemp = zlDatabase.GetUserInfo
    With rsTemp
        If .RecordCount > 0 Then
            glngUserId = .Fields("ID").Value                '当前用户id
            gstrUserCode = .Fields("编号").Value            '当前用户编码
            gstrUserName = .Fields("姓名").Value            '当前用户姓名
            gstrUserAbbr = IIf(IsNull(.Fields("简码").Value), "", .Fields("简码").Value)          '当前用户简码
            glngDeptId = .Fields("部门id").Value            '当前用户部门id
            gstrDeptCode = .Fields("部门码").Value        '当前用户
            gstrDeptName = .Fields("部门名").Value        '当前用户
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
''功能：打开记录。同时保存SQL语句
'    If rsTemp.State = adStateOpen Then rsTemp.Close
'
'    Call SQLTest(App.ProductName, strFormCaption, gstrSQL)
'    rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
'    Call SQLTest
'End Sub

Private Function SystemImes() As Variant
'功能：将系统中文输入法名称返回到一个字符串数组中
'返回：如果不存在中文输入法,则返回空串
    Dim arrIme(99) As Long, arrName() As String
    Dim lngLen As Long, StrName As String * 255
    Dim lngCount As Long, i As Integer, j As Integer

    lngCount = GetKeyboardLayoutList(UBound(arrIme) + 1, arrIme(0))
    For i = 0 To lngCount - 1
        If ImmIsIME(arrIme(i)) = 1 Then '为1表示中文输入法
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
        MsgBox "你还没安装任何汉字输入法，不能使用本功能。" & vbCrLf & _
               "输入法的安装可在控制面板中完成。", vbInformation, gstrSysName
        Exit Function
    End If
    cmbIME.Clear
    cmbIME.AddItem "不自动开启"
    strIme = zlDatabase.GetPara("输入法")
    For i = LBound(varIME) To UBound(varIME)
        cmbIME.AddItem varIME(i)
        If strIme = varIME(i) Then cmbIME.ListIndex = i + 1
    Next
    If cmbIME.ListIndex < 0 Then cmbIME.ListIndex = 0
    ChooseIME = True
End Function

Public Function GetControlRect(ByVal lngHwnd As Long) As RECT
'功能：获取指定控件在屏幕中的位置(Twip)
    Dim vRect As RECT
    Call GetWindowRect(lngHwnd, vRect)
    vRect.Left = vRect.Left * Screen.TwipsPerPixelX
    vRect.Right = vRect.Right * Screen.TwipsPerPixelX
    vRect.Top = vRect.Top * Screen.TwipsPerPixelY
    vRect.Bottom = vRect.Bottom * Screen.TwipsPerPixelY
    GetControlRect = vRect
End Function


Public Function NewClientRecord(ByVal strFilds As String) As ADODB.Recordset
    '创建一个空的记录集
    'strFilds:字段名,类型,长度;字段名,类型,长度...
    '    例如:用户名,varchar2,30;姓名,varchar2,30
    
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

