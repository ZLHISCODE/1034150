VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmClientCopy 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "自动升级"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6900
   ControlBox      =   0   'False
   Icon            =   "frmClientCopy.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   6900
   StartUpPosition =   2  '屏幕中心
   Begin zlHisCrust.UsrProgressBar prgPross 
      Height          =   300
      Left            =   45
      TabIndex        =   4
      Top             =   1245
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   529
      Color           =   12937777
      Value           =   100
   End
   Begin VB.CommandButton cmdLog 
      Caption         =   "查看日志(&C)"
      Height          =   375
      Left            =   3615
      TabIndex        =   3
      Top             =   4545
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.FileListBox File1 
      Height          =   450
      Left            =   5145
      TabIndex        =   2
      Top             =   105
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "完成(&O)"
      Height          =   375
      Left            =   5220
      TabIndex        =   1
      Top             =   4545
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSComctlLib.ListView lvwMan 
      Height          =   2430
      Left            =   60
      TabIndex        =   0
      Top             =   2040
      Width           =   6810
      _ExtentX        =   12012
      _ExtentY        =   4286
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "img2"
      SmallIcons      =   "img2"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   11
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "部件"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "升级信息"
         Object.Width           =   7585
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "现版本号"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "原版本号"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "现修改日期"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "原修改日期"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "业务部件"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "安装路径"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "MD5"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "自动升级"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "强制覆盖"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ImageList img2 
      Left            =   5535
      Top             =   1815
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClientCopy.frx":030A
            Key             =   "Ok"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClientCopy.frx":08A4
            Key             =   "Err"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClientCopy.frx":0E3E
            Key             =   "List"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img1 
      Left            =   4725
      Top             =   1830
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClientCopy.frx":13D8
            Key             =   "OK"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClientCopy.frx":16F2
            Key             =   "Err"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClientCopy.frx":1A0C
            Key             =   "List"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "客户端正在升级,请稍候..."
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1815
      TabIndex        =   6
      Top             =   495
      Width           =   4020
   End
   Begin VB.Label lblInfor 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "正在注册部件"
      Height          =   180
      Left            =   60
      TabIndex        =   5
      Top             =   1710
      Width           =   1080
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   315
      Picture         =   "frmClientCopy.frx":1B66
      Top             =   255
      Width           =   720
   End
End
Attribute VB_Name = "frmClientCopy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim blnFirst As Boolean
Dim lngCount As Long
Dim mintColumn As Integer
Dim blnOk As Boolean
Dim blnAutoRun As Boolean '是否下载批处理文件
Dim strAutoRun As String '批处理文件路径
Dim strAutoRunBat As String


Private Sub cmdLog_Click()
    On Error Resume Next
    Dim ret As Long
    Dim strLogFile As String
    Dim strNotPad As String
    strNotPad = GetWinSystemPath & "\notepad.exe"
    If mobjFile.FileExists(strNotPad) Then
        strLogFile = gstrAppPath & "\ZLUpGradeList.lst"
        ret = ShellExecute(0&, "open", strNotPad, strLogFile, strLogFile, 5)    'SW_SHOW
        If ret = 31 Then
           MsgBox "没有找到适当的程序来打开它,请安装有效的程序!", vbInformation, "客户端自动升级"
        End If
    Else
        MsgBox "本机没有安装记事本程序,不能打开日志文件!" & vbCrLf & "请手工用其它程序打开,记事本路径为:" & vbCrLf & strLogFile, vbInformation, "客户端自动升级"
    End If
End Sub

Private Sub cmdOK_Click()
    '确定是否还存在升级
    
    
    
    If IsUpgrade = True Then
        '--------------------------------------------------------
        '升级不成功
        WriteTxtLog ""
        WriteTxtLog ""
        WriteTxtLog "--" & Format(Now(), "yyyy-mm-dd HH:MM:SS")
        WriteTxtLog "=============================================至少有一个部件升级或收集不成功======================================================================================================================================="
        Call SaveClientLog("至少有一个部件升级或收集不成功")
        Call UpdateCondition(2)
        Call CallHISEXE(False)
    Else
        WriteTxtLog ""
        WriteTxtLog ""
        WriteTxtLog "--" & Format(Now(), "yyyy-mm-dd HH:MM:SS")
        WriteTxtLog "=============================================升级或收集成功======================================================================================================================================="
        '实现自动调用主程序
        '执行HIS程序
        Call SaveClientLog("升级或收集成功")
        Call UpdateCondition(1)
        Call CallHISEXE
    End If
    CloseLogFile
    End
End Sub

Private Sub CallHISEXE(Optional bln用户及密码 As Boolean = True)
    '调用HIS
    Dim strUserName As String, strPassWord As String, mError As String
    Dim strFile As String
    
    '如果是ZLBH融合启动，则不再回调
    If UCase(gstrAppEXE) = UCase("zlActMain.exe") Then
        MsgBox "自动升级完成,请重新执行模块!", vbInformation, "自动升级"
        Exit Sub
    End If
    If gblnPreUpgrade Then Exit Sub
    
    If bln用户及密码 Then
        Call AnalyseUserNameAndPassWord(strUserName, strPassWord)
    End If
    
    '确定文件是否存在
    Err = 0: On Error Resume Next
    If gstrAppEXE <> "" Then
        strFile = gstrAppPath & "\" & gstrAppEXE
    Else
        strFile = gstrAppPath & "\ZLHIS90.exe"
    End If
    If FindFile(strFile) = False Then
        strFile = gstrAppPath & "\ZLHIS+.exe"
        If FindFile(strFile) = False Then
            If gstrAppEXE <> "" Then
                strFile = gstrAppPath & "\ZLHIS90.exe"
            End If
        End If
    End If
    
    If bln用户及密码 Then
        mError = Shell(strFile & " " & IIf(gstrHisCommand <> "", gstrHisCommand, strUserName & "/" & strPassWord), vbNormalFocus)
    Else
        mError = Shell(strFile, vbNormalFocus)
    End If
End Sub


Private Sub Form_Load()
    Dim strTxtFile As String, mError As String
    
'    Dim strSourceFile As String, strDescFile As String
    blnOk = False
    
    Call SetWindowPos(Me.hwnd, HWND_TOP, ((Screen.Width - Me.Width) / 2) / 15, ((Screen.Height - Me.Height) / 2) / 15, 0, 0, SWP_NOSIZE)
    Me.cmdOK.Caption = "取消(&C)"
    blnFirst = True
    If gblnPreUpgrade Then
        Me.Hide
    Else
        Me.Show
    End If
    lblInfor.Caption = "正在连接数据库..."
    Me.Refresh
    DoEvents
    
    '连接数据库
    If OpenOracle = False Then End: Exit Sub
    lblInfor.Caption = "初始参数..."
    
    '初始升级方式
    Call InitUpType
    '初始收集方式
    Call iniGatherTYpe
    '初始化变量
    Call InintVar
    
    
    
    '判断是否为USER组权限升级
    If GetAdmin = False Then
        '如果没有管理权限
        If GetAdministrator = False Then '获取管理员权限
            Unload Me
        End If
        End '强制退出进程
    End If
    
    
    If gblnPreUpgrade Then
        '预升级
        Call OpenLogFile(True)
        WriteTxtLog ""
        WriteTxtLog ""
        WriteTxtLog "=============================================开始进行预升级============================================="
        WriteTxtLog "--" & Format(Now(), "yyyy-mm-dd HH:MM:SS")
    Else
        '    确定是否进行升级
        If IsUpgrade = False Then End: Exit Sub
        
        Call OpenLogFile(False)
        WriteTxtLog ""
        WriteTxtLog ""
        WriteTxtLog "=============================================开始升级或收集============================================="
        WriteTxtLog "--" & Format(Now(), "yyyy-mm-dd HH:MM:SS")
    End If
    
    lblInfor.Caption = "正在连接文件服务器..."
    
    '网络是否联通
    If IIf(gbln收集 = True, gintGatherTYpe = 0, gintUpType = 0) Then
        If IsNetServer = False Then
            MsgBox "无法连接到服务器:" & gstrServerPath & "上,请确认网络是否畅通," & vbCrLf _
            & "或检查管理工具中文件" & IIf(gbln收集 = True, "收集", "升级") & "服务是否设置正确!", vbInformation + vbDefaultButton1, "客户端自动" & IIf(gbln收集 = True, "收集", "升级")
            
            Call SaveClientLog("无法连接到服务器:" & gstrServerPath & "上,请确认网络是否畅通。")
            Call UpdateCondition(2)
            End:
            Exit Sub   '连接共享服务器
        End If
    Else
        If IsFtpServer = False Then
            MsgBox "无法连接到服务器:" & gstrServerPath & "上,请确认FTP服务器是否开启," & vbCrLf _
            & "或检查管理工具中文件" & IIf(gbln收集 = True, "收集", "升级") & "服务是否设置正确!", vbInformation + vbDefaultButton1, "客户端自动" & IIf(gbln收集 = True, "收集", "升级")
            
            Call SaveClientLog("无法连接到服务器:" & gstrServerPath & "上,请确FTP服务器是否畅通。")
            Call UpdateCondition(2)
            End:
            Exit Sub   '连接FTP服务器
        End If
    End If
    
    '首先检查是否有MD5备用部件是否需要升级
    Call isMD5UpGrade
    
    
    '确定是否自升身需升级
    If gBlnHisCrustCompare Then
        lblInfor.Caption = "正在升级自身..."
        If InStrRev(UCase(App.Path), UCase("\Apply"), -1) = 0 Then
            If isHisCurstUpGrade = True Then
                Err = 0: On Error Resume Next
                mError = Shell(gstrAppPath & "\Apply\zlHisCrust.exe" & " " & gcnnOracle.ConnectionString & "||1" & "||" & gstrAppEXE & "||||" & gstrHisCommand, vbNormalFocus)
                '调用外壳程序
                If mError <> 0 Then
                    End
                    Exit Sub
                End If
            Else
            End If
        End If
    End If

    '特殊处理7Z的文件
    Call is7zUpGrade
    
    
    '获取临时存放目录
    gstrTempPath = GetTmpPath
    If gstrTempPath <> "" Then
        gstrTempPath = gstrTempPath & "ZLTEMP\"
    Else
        gstrTempPath = GetWinPath & "\ZLTEMP\"
    End If
    
    '获取预升级临时目录
    gstrPerTempPath = GetTmpPath
    If gstrPerTempPath <> "" Then
        gstrPerTempPath = gstrPerTempPath & "ZLPERTEMP\"
    Else
        gstrPerTempPath = GetWinPath & "\ZLPERTEMP\"
    End If
    
    '是否为定时正式升级
    If gblnOfficialUpgrade Then
        gbln预升完成 = GetPreUpgrad(gstrComputerName)
    End If
    
    '加载升级数据
    If gbln收集 Then
        lblInfor.Caption = "加载上传文件数据..."
        If GetClientFiles = False Then End: Exit Sub
    Else
        lblInfor.Caption = "加载升级数据..."
        If getSeverFiles = False Then End: Exit Sub
    End If
    lngCount = 0
    Timer1.Enabled = True
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    End
End Sub

Private Sub Form_Resize()
        Err = 0
        On Error Resume Next
        With Image1
            .Width = Me.ScaleWidth
        End With
        With cmdOK
            .Top = ScaleHeight - .Height - 50
            .Left = ScaleWidth - .Width - 100
        End With
        
        With cmdLog
            .Top = ScaleHeight - .Height - 50
            .Left = ScaleWidth - .Width - cmdOK.Width - 200
        End With
        
        If gblnPreUpgrade Then
            Me.Hide
        Else
            Me.Show
        End If
'        With prgPross
'            .Top = cmdOK.Top - prgPross.Height - 50
'            .Width = ScaleWidth - 100
'        End With
'
'        With lvwMan
'            .Width = ScaleWidth
'            .Height = prgPross.Top - .Top - 100
'        End With
'        With lblInfor
'            .Left = ScaleLeft + 50
'            .Top = cmdOK.Top + (cmdOK.Height - .Height) / 2
'            .Width = cmdOK.Left
'        End With
End Sub

Private Function FileUpgrade() As Boolean
    '功能:文件升级处理
    Dim lst As ListItem
    Dim n As Integer
    Dim strSourceFile As String
    
    Dim strTargetFile As String
    Dim strSourceVer As String, strSourceDate As String
    Dim strTargetVer As String, strTargetDate As String
    Dim strErrMsg As String
    Dim blnCommon As Boolean
    Dim strUpgradeInfor As String
    Dim strLocalPath As String '本地解压文件路径
    Dim objFile As New FileSystemObject
    Dim strTempFlag As String
    Dim lngDefeated As Long '失败个数
    Dim bln注册 As Boolean
    Dim bln覆盖 As Boolean
    Dim blnSysFile As Boolean
    Dim strSysFile As String

    
    FileUpgrade = False
    lngDefeated = 0
    prgPross.Min = 0
    prgPross.Value = 0
    prgPross.Max = IIf(Me.lvwMan.ListItems.Count = 0, 2, Me.lvwMan.ListItems.Count)
    
    strUpgradeInfor = ""
    Me.lblInfor.Caption = "正在注册文件"
    
'    For Each lst In Me.lvwMan.ListItems
    For n = 1 To Me.lvwMan.ListItems.Count
        Set lst = Me.lvwMan.ListItems(n)
        
        blnCommon = False
        prgPross.Value = prgPross.Value + 1
        lst.Selected = True
        
        If gbln收集 Then
            strSourceFile = gstrAppPath & "\" & lst.Text
            strTargetFile = gstrServerPath & "\" & GetMyCompterName & "_" & lst.Text
            
            If gintGatherTYpe = 0 Then
                '比较文件
                If CompareFile(strSourceFile, strTargetFile, strSourceVer, strSourceDate, strTargetVer, strTargetDate) Then
                    GetCopyAndReg strSourceFile, strTargetFile, strErrMsg, True
                    
                    '写入注册信息
                    If strErrMsg <> "正常升级!" And strErrMsg <> "未装此部件!" And strErrMsg <> "客户端不存在此部件,但还是进行了升级!" Then
                        If strUpgradeInfor = "" Then
                            strUpgradeInfor = "在" & Format(Date, "yyyy-mm-dd hh:mm:ss") & "的升级过程中, " & vbCrLf & " 至少存一个部件升级出错,如:" & lst.Text
                        End If
                        lst.SmallIcon = "Err"
                    Else
                        lst.SmallIcon = "Ok"
                    End If
                Else
                    lst.SmallIcon = "Ok"
                    strErrMsg = "文件相同,没有必要升级!"
                End If
            Else
                strTargetFile = GetMyCompterName & "_" & lst.Text
                If FtpupFile(strSourceFile, strTargetFile) Then
                    lst.SmallIcon = "Ok"
                    strErrMsg = "收集完成!"
                Else
                    lst.SmallIcon = "Err"
                    strErrMsg = "下载错误!"
                End If
            End If
        Else
            '0.获取:获取文件绝对路径:检查业务部件是否安装
            If gintUpType = 0 Then
                strSourceFile = gstrServerPath & "\" & lst.Text & ".7z"
                strTargetFile = GetSetupPath(lst.Text, NVL(lst.SubItems(7), ""), NVL(lst.Tag, ""), gstrAppPath, NVL(lst.SubItems(6), ""))
            Else
                strSourceFile = lst.Text & ".7z"
                strTargetFile = GetSetupPath(lst.Text, NVL(lst.SubItems(7), ""), NVL(lst.Tag, ""), gstrAppPath, NVL(lst.SubItems(6), ""))
            End If
            
            '问题号:68569,删除PUBLIC对应的SYSTEM32下的公共部件
            If UCase(NVL(lst.SubItems(7), "")) = "[PUBLIC]" Then
                strSysFile = GetWinSystemPath & "\" & lst.Text
                If objFile.FileExists(strSysFile) Then
                    On Error Resume Next
                    Call objFile.DeleteFile(strSysFile)
                    Sleep 50
                    If objFile.FileExists(strSysFile) Then
                        WriteTxtLog "删除PUBLIC对应的SYSTEM32下的公共部件:" & strSysFile & "失败!"
                    Else
                        WriteTxtLog "删除PUBLIC对应的SYSTEM32下的公共部件:" & strSysFile & "成功!"
                    End If
                End If
            End If
            
            strSourceVer = GetFileListValue(lst.Text, 1) '或lst.SubItems(1)
            strSourceDate = GetFileListValue(lst.Text, 2) '或lst.SubItems(3)
            
            
            'strTargetFile="" 表示未安装的业务部件
            If strTargetFile = "" And NVL(lst.Tag, "") = "1" Then
                strTargetVer = ""
                strTargetDate = ""
                strErrMsg = "本机没有安装该部件,不需升级"
                lst.SmallIcon = "Ok"
                GoTo zt
            End If
            
            
            '1.检查:检查是否需要下载该文件,比较文件的MD5值
            lblInfor.Caption = "正在检查部件:" & lst.Text
            If CompareMD5Down(strTargetFile, lst.Text) = False Then
                strTargetVer = GetCommpentVersion(strTargetFile)
                strTargetDate = Format(FileDateTime(strTargetFile), "yyyy-MM-DD hh:mm:ss")
                strErrMsg = "文件MD5相同,不需升级!"
                lst.SmallIcon = "Ok"
                GoTo zt
            Else
'测试:
                '如果文件存在获取现版本和修改日期
                If mobjFile.FileExists(strTargetFile) Then
                    strTargetVer = GetCommpentVersion(strTargetFile)
                    strTargetDate = Format(FileDateTime(strTargetFile), "yyyy-MM-DD hh:mm:ss")
                Else
                    strTargetVer = ""
                    strTargetDate = ""
                End If
                
                '2.下载文件
                strLocalPath = lst.Text & ".7z"
                lblInfor.Caption = "正在下载部件:" & lst.Text
                If FileTempDown(strSourceFile, strLocalPath, strErrMsg) = False Then
                    If strErrMsg <> "下载完成！" Then
                        If mobjFile.FileExists(strTargetFile) Then
                            strErrMsg = "文件在服务器目录不存在!"
                        Else
                        
                        
                            If strErrMsg = "文件在服务器目录不存在!" Then
                                lst.SmallIcon = "Err"
                                GoTo zt
                            Else
                                strErrMsg = "文件不需升级!"
                                lst.SmallIcon = "Ok"
                                GoTo zt
                            End If
                        End If
                    Else
                        strErrMsg = "下载文件失败!"
                    End If
                    lst.SmallIcon = "Err"
                    GoTo zt
                End If
                
                
                '如果是预升级就算完成了
                If gblnPreUpgrade Then
                    lst.SmallIcon = "Ok"
                    GoTo zt
                End If
                
                
                '3.解压文件
                lblInfor.Caption = "正在解压部件:" & lst.Text
                If FileDeCompression(strLocalPath, strErrMsg) = False Then
                    strErrMsg = "解压缩文件失败!"
                    lst.SmallIcon = "Err"
                    GoTo zt
                End If
                
                '4.拷贝并注册文件
'                strTargetFile = "C:\Temp\" & lst.Text
                lblInfor.Caption = "正在注册部件:" & lst.Text
                If lst.SubItems(9) = "" Or lst.SubItems(9) = "0" Then
                    bln注册 = False
                Else
                    bln注册 = True
                End If

                If lst.SubItems(10) = "" Or lst.SubItems(10) = "0" Then
                    bln覆盖 = False
                Else
                    bln覆盖 = True
                End If
                
                If NVL(lst.Tag, "") = "5" Then
                    blnSysFile = True
                Else
                    blnSysFile = False
                End If
                
                If GetCopyAndReg(strLocalPath, strTargetFile, strErrMsg, bln注册, blnSysFile, bln覆盖) = False Then
                    strErrMsg = "替换文件失败可能已被其它程序独占!"
                    lst.SmallIcon = "Err"
                    On Error Resume Next
                    Call Kill(strLocalPath)
                    GoTo zt
                Else
                    On Error Resume Next
                    lst.SmallIcon = "Ok"
                    Call Kill(strLocalPath)
                End If
                
                
                '5.检查MD5值是否正确
                If strErrMsg = "忽略本部件升级" Then GoTo zt
                If strErrMsg = "被自身独占" Then GoTo zt
                If CheckSysFile(strTargetFile) Then GoTo zt
                If blnSysFile = True And bln覆盖 = False Then GoTo zt
                lblInfor.Caption = "正在比较部件:" & lst.Text
                If CompareMD5Down(strTargetFile, lst.Text, strErrMsg) Then
                   If strErrMsg = "服务器没有该文件MD5信息!" Then
                     strErrMsg = "服务器没有该文件MD5信息!"
                   Else
                     strErrMsg = "文件下载受损,MD5现值与原值不一致!"
                   End If
                   lst.SmallIcon = "Err"
                   GoTo zt
                End If
                lst.SmallIcon = "Ok"
            End If
        End If
zt:
        If lst.SmallIcon = "Err" Then
            If strUpgradeInfor = "" Then
                strUpgradeInfor = "在" & Format(Date, "yyyy-mm-dd hh:mm:ss") & "的升级过程中, " & vbCrLf & " 至少存一个部件升级出错,如:" & lst.Text
            Else
                strUpgradeInfor = strUpgradeInfor & "," & lst.Text
            End If
            strTempFlag = "[失败]:"
            lngDefeated = lngDefeated + 1
        Else
            strTempFlag = "[成功]:"
        End If
        
        strErrMsg = IIf(gbln收集, Replace(strErrMsg, "升级", "收集"), strErrMsg)
        lst.SubItems(3) = strTargetVer
        lst.SubItems(5) = strTargetDate
        lst.SubItems(1) = strErrMsg
        lst.EnsureVisible
        WriteTxtLog strTempFlag & strSourceFile & "(版本:" & strSourceVer & "   修改日期:" & strSourceDate & ")    ====>    " & vbCrLf & _
                        strTargetFile & "(版本:" & strTargetVer & "   修改日期:" & strTargetDate & ")        升级信息:" & strErrMsg & vbCrLf
        
        DoEvents
    Next
    
    '执行批处理文件
    strAutoRun = gstrAppPath & "\zlAutoRun.ini"
    strAutoRunBat = gstrAppPath & "\zlAutoRun.bat"
    If mobjFile.FileExists(strAutoRun) Or mobjFile.FileExists(strAutoRunBat) Then
        Dim ret As Long
        Name strAutoRun As gstrAppPath & "\zlAutoRun.bat"
        On Error Resume Next
        Call Kill(strAutoRun)
        
        ret = ShellExecute(0&, "open", gstrAppPath & "\zlAutoRun.bat", "", gstrAppPath & "\zlAutoRun.bat", 5) 'SW_SHOW
        If ret = 31 Then
            strErrMsg = "批处理执行失败!"
'           MsgBox "没有找到适当的程序来打开它,请安装有效的程序!", vbInformation, "提示"
            WriteTxtLog "批处理文件执行失败!"
        Else
            WriteTxtLog "批处理文件执行成功!"
        End If
        
        blnAutoRun = True
    Else
        blnAutoRun = False
    End If
    
    '清空7z.exe残余系统进程
    Call fun_KillProcess(PROAPPCTION)
    
    If InStr(1, UCase(App.Path), UCase("APPLY")) <> 0 Then  '移入上级目录
        GetCopyAndReg App.Path & "\zlHisCrust.exe", Replace(App.Path, "\Apply", "") & "\zlHisCrust.exe", strErrMsg
    End If
    
   '需收集部件升级文件
    strSourceFile = gstrAppPath & "\ZLUpGradeList.Lst"
    strTargetFile = gstrServerPath & "\" & GetMyCompterName & "_ZLUpGradeList.LOG"
    If InStr(1, gstr收集类型, "LOG") <> 0 And gbln收集 Then
        '收集本机日志
        If objFile.FileExists(strSourceFile) Then
            GetCopyAndReg strSourceFile, strTargetFile, strErrMsg
        End If
    End If
    
    If lngDefeated = 0 Then
        If gblnPreUpgrade Then
            WriteTxtLog "所有预升级完成!"
            Call SaveClientLog("所有预升级完成")
            Call UpdateCondition(1)
        Else
            WriteTxtLog "所有升级完成!"
            Call SaveClientLog("所有升级完成")
            Call UpdateCondition(1)
        End If
        Me.lblInfor.Caption = IIf(gbln收集, "收集", "升级") & "成功"
        cmdLog.Visible = False
        gblnOk = True
    Else
        '记录错误日志
        If GetErrParameter(3) = "1" Then
            Dim i As Long
            With lvwMan
            For i = 1 To .ListItems.Count
                If .ListItems(i).SmallIcon <> "Ok" Then
                    Call SaveErrLog(.ListItems(i).Text & "-" & .ListItems(i).SubItems(1))
                    Call SaveClientLog(.ListItems(i).Text & "-" & .ListItems(i).SubItems(1))
                End If
            Next
            End With
        End If
        

        '处理显示错误列表
        With lvwMan
            For n = 1 To .ListItems.Count
                If n > .ListItems.Count Then
                    Exit For
                End If
                If .ListItems(n).SmallIcon = "Ok" Then
                    .ListItems.Remove n
                    n = n - 1
                End If
            Next
        End With
        
        Me.Height = 5445
        lblInfor.Caption = "客户端部件升级情况"
        WriteTxtLog "总共:" & lngDefeated & "文件升级失败!"
        Call SaveClientLog("总共:" & lngDefeated & "文件升级失败!")
        Me.lblInfor.Caption = "有" & lngDefeated & "文件升级失败,请核查!"
        cmdLog.Visible = True
        gblnOk = False
    End If
    
    Dim strSQL As String
    Err = 0
    
    '更改升级说明,更改升级和收集标记
     If gblnPreUpgrade = False Then
         If strUpgradeInfor <> "" Then
             If LenB(StrConv(strUpgradeInfor, vbFromUnicode)) > 200 Then
                 strUpgradeInfor = Mid(strUpgradeInfor, 1, 200)
             End If
             'strSQL = "Update zltools.zlclients set 说明='" & strUpgradeInfor & "' where upper(工作站)='" & UCase(gstrComputerName) & "'"
             strSQL = "Zl_Zlclients_Control(9,'" & gstrComputerName & "',Null,Null,Null,Null,Null,Null,Null,Null,'" & strUpgradeInfor & "')"
                  
         Else
             strUpgradeInfor = "在" & Format(Now, "yyyy-mm-dd HH:mm:ss") & "升级了部件"
             If gbln收集 Then
                 'strSQL = "Update zltools.zlclients set 说明='" & strUpgradeInfor & "' ,收集标志=0 where upper(trim(工作站))='" & UCase(gstrComputerName) & "'"
                strSQL = "Zl_Zlclients_Control(10,'" & gstrComputerName & "',Null,Null,Null,Null,Null,Null,Null,Null,'" & strUpgradeInfor & "')"
             Else
                ' strSQL = "Update zltools.zlclients set 说明='" & strUpgradeInfor & "' ,升级标志=0 where upper(trim(工作站))='" & UCase(gstrComputerName) & "'"
                strSQL = "Zl_Zlclients_Control(11,'" & gstrComputerName & "',Null,Null,Null,Null,Null,Null,Null,Null,'" & strUpgradeInfor & "')"
                 '如如有预升级目录存在,就进行删除文件目录
                 If gblnOfficialUpgrade And gbln预升完成 Then
                    If mobjFile.FolderExists(gstrPerTempPath) Then
                       On Error Resume Next
                       Call mobjFile.DeleteFolder(Left(gstrPerTempPath, Len(gstrPerTempPath) - 1))
                    End If
                 End If
             End If
        End If
    Else
        '更改站点的预升级完成状态
        If strUpgradeInfor <> "" Then
            If LenB(StrConv(strUpgradeInfor, vbFromUnicode)) > 200 Then
                 strUpgradeInfor = Mid(strUpgradeInfor, 1, 200)
            End If
            strSQL = "Zl_Zlclients_Control(12,'" & gstrComputerName & "',Null,Null,Null,Null,Null,Null,Null,Null,'" & "预升级出错:" & strUpgradeInfor & "')"
           ' strSQL = "Update zltools.zlclients set 预升完成=0,说明='" & "预升级出错:" & strUpgradeInfor & "' where upper(trim(工作站))='" & UCase(gstrComputerName) & "'"
        Else
            strSQL = "Zl_Zlclients_Control(13,'" & gstrComputerName & "')"
           ' strSQL = "Update zltools.zlclients set 预升完成=1 where upper(trim(工作站))='" & UCase(gstrComputerName) & "'"
        End If
    End If
   
    gcnnOracle.Execute strSQL
    
    '如果为定时升级
    If gblnOfficialUpgrade Then
        'strSQL = "Update zltools.zlclients set 预升时点=Null ,预升完成=Null where upper(trim(工作站))='" & UCase(gstrComputerName) & "'"
        strSQL = "Zl_Zlclients_Control(14,'" & gstrComputerName & "')"
        gcnnOracle.Execute strSQL
    End If
    
    Me.cmdOK.Caption = "完成(&O)"
    Me.cmdOK.Visible = True
    blnOk = True
    FileUpgrade = True

    '断开网络连接
    
    If IIf(gbln收集 = True, gintGatherTYpe = 0, gintUpType = 0) Then
        '关闭Share连接
        CancelNetServer
    Else
        '关闭FTP连接
        CancelFtpServer
    End If
'    End
End Function

Private Function FindHisBrow() As Boolean
    '功能:查找并结束HIS主窗口的相关进程
    '成功:结束成功,返回true,否则返回false
    Dim lngHwnd As Long
    Dim lngZlhisHwnd As Long
    Dim lngVBHwnd As Long
    Dim lngPid As Long
    Dim lngProcess As Long
    Err = 0: On Error GoTo ErrHand:
    
    '如果预升级,就退出。
    If gblnPreUpgrade Then
        Exit Function
    End If
    
    Do While True
         lngHwnd = FindWindow(vbNullString, "导航台")
         If lngHwnd = 0 Then
            lngHwnd = FindWindow(vbNullString, "医院信息系统")
            If lngHwnd = 0 Then
                Exit Do
            End If
         End If
         If lngHwnd <> 0 Then
            '区分是否是VB在调用导航台还是程序直接执行导航台
            lngZlhisHwnd = fun_ExitsProcess("zlhis+.exe")
            If lngZlhisHwnd <> 0 Then
                Call TerminateProcess(lngZlhisHwnd, 1&)
            Else
                lngVBHwnd = fun_ExitsProcess("vb6.exe")
                If lngVBHwnd <> 0 Then
                    If MsgBox("升级程序检测到VB6加载了可能会升级的部件." & vbCrLf & "为了保证系统正常升级,是否关闭VB6进程!", vbQuestion + vbYesNo, "客户端自动升级") = vbYes Then
                        Call GetWindowThreadProcessId(lngHwnd, lngPid)
                        lngProcess = OpenProcess(PROCESS_TERMINATE, 0&, lngPid)
                        Call TerminateProcess(lngProcess, 1&)
                    Else
                        GoTo NoClose
                    End If
                Else
                    Call GetWindowThreadProcessId(lngHwnd, lngPid)
                    lngProcess = OpenProcess(PROCESS_TERMINATE, 0&, lngPid)
                    Call TerminateProcess(lngProcess, 1&)
                End If
            End If
         End If
     Loop
     
     
    Do While True
        '自动关闭zlSvrStudio进程
        lngZlhisHwnd = fun_ExitsProcess("zlSvrStudio.exe")
        If lngZlhisHwnd <> 0 Then
            Call TerminateProcess(lngZlhisHwnd, 1&)
        End If
        
        If lngZlhisHwnd = 0 Then
            Exit Do
        End If
    Loop
    
    Do While True
        '自动关闭Zl9LISComm进程
        lngZlhisHwnd = fun_ExitsProcess("Zl9LISComm.exe")
        If lngZlhisHwnd <> 0 Then
            Call TerminateProcess(lngZlhisHwnd, 1&)
        End If
        
        If lngZlhisHwnd = 0 Then
            Exit Do
        End If
    Loop
    
    Do While True
        '自动关闭zlLisReceiveSend进程
        lngZlhisHwnd = fun_ExitsProcess("zlLisReceiveSend.exe")
        If lngZlhisHwnd <> 0 Then
            Call TerminateProcess(lngZlhisHwnd, 1&)
        End If
        
        If lngZlhisHwnd = 0 Then
            Exit Do
        End If
    Loop
    
    Do While True
        '自动关闭ZlPacsSrv进程
        lngZlhisHwnd = fun_ExitsProcess("ZlPacsSrv.exe")
        If lngZlhisHwnd <> 0 Then
            Call TerminateProcess(lngZlhisHwnd, 1&)
        End If
        
        If lngZlhisHwnd = 0 Then
            Exit Do
        End If
    Loop
    
NoClose:
    FindHisBrow = False
    Exit Function
ErrHand:
End Function

'判断窗口是否符合要求
Function TaskWindow(hwcurr As Long) As Long
    Dim lngStyle As Long, IsTask As Long
    '获取窗口风格，并判断是否符合要求
    lngStyle = GetWindowLong(hwcurr, GWL_STYLE)
    If (lngStyle And IsTask) = IsTask Then
     TaskWindow = True
    End If
End Function

Public Sub CloseWindow(app_name As String)
    Dim app_hwnd As Long
    app_hwnd = FindWindow(vbNullString, app_name)
    SendMessage app_hwnd, WM_CLOSE, 0, 0
End Sub

Private Sub lvwMan_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If blnOk = False Then Exit Sub
    Err = 0
    On Error Resume Next
    If mintColumn = ColumnHeader.Index - 1 Then
        lvwMan.SortOrder = IIf(lvwMan.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        mintColumn = ColumnHeader.Index - 1
        lvwMan.SortKey = mintColumn
        lvwMan.SortOrder = lvwAscending
    End If

End Sub

Private Sub Timer1_Timer()
        Dim strSQL As String
        Timer1.Enabled = False
        If FindHisBrow = False Then
            MousePointer = 11
            If FileUpgrade = False Then
                Me.cmdOK.Caption = "取消(&C)"
                Timer1.Enabled = False
                MousePointer = 0
                Exit Sub
            Else
                If gbln收集 = False Then
                    '确定是否收集部件
                    If Is收集文件 Then
                        '加载收集文件
                        Call InintVar
                        lblInfor.Caption = "正在重新连接上传服务器..."

                        '网络是否联通
                        If IIf(gbln收集 = True, gintGatherTYpe = 0, gintUpType = 0) Then
                            If IsNetServer = False Then
                                Timer1.Enabled = False
                                MousePointer = 0
                                MsgBox "无法连接到服务器:" & gstrServerPath & "上,请确认网络是否畅通," & vbCrLf _
                                & "或检查管理工具中文件" & IIf(gbln收集 = True, "收集", "升级") & "服务是否设置正确!", vbInformation + vbDefaultButton1, "客户端自动" & IIf(gbln收集 = True, "收集", "升级")
                                               
                                Call SaveClientLog("无法连接到服务器:" & gstrServerPath & "上,请确共享服务器是否畅通。")
                                Call UpdateCondition(2)
                                Exit Sub   '连接共享服务器
                            End If
                        Else
                            If IsFtpServer = False Then
                                Timer1.Enabled = False
                                MousePointer = 0
                                MsgBox "无法连接到服务器:" & gstrServerPath & "上,请确认FTP服务器是否开启," & vbCrLf _
                                & "或检查管理工具中文件" & IIf(gbln收集 = True, "收集", "升级") & "服务是否设置正确!", vbInformation + vbDefaultButton1, "客户端自动" & IIf(gbln收集 = True, "收集", "升级")
                                
                                Call SaveClientLog("无法连接到服务器:" & gstrServerPath & "上,请确FTP服务器是否畅通。")
                                Call UpdateCondition(2)
                                Exit Sub   '连接FTP服务器
                            End If
                        End If

                        lblInfor.Caption = "加载上传文件数据..."
                        If GetClientFiles = True Then
                            If FileUpgrade = False Then
                                Me.cmdOK.Caption = "取消(&C)"
                                Timer1.Enabled = False
                                MousePointer = 0
                                Exit Sub
                            End If
                        Else
                            lblInfor.Caption = "没有被收集的文件..."
                            Me.cmdOK.Caption = "完成(&C)"
                        End If
                    End If
                End If
            End If
            MousePointer = 0
            Timer1.Enabled = False
            
            '升级完成需退出
            If gblnOk Then '无错误退出
                Call cmdOK_Click
            End If
        Else
            Timer1.Enabled = True
        End If
End Sub
Private Function Is收集文件() As Boolean
    '   功能：确定是否收集文件
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    Err = 0
    On Error GoTo ErrHand:
    strSQL = "Select * From zlClients Where  收集标志=1 and upper(工作站)='" & gstrComputerName & "'"
    With rsTemp
        .Open strSQL, gcnnOracle
        gbln收集 = .RecordCount <> 0
        Is收集文件 = gbln收集
        .Close
    End With
    Exit Function
ErrHand:
    Is收集文件 = False
End Function

Private Function GetClientFiles() As Boolean
     '收集文件
     Dim i As Long, lngfile As Long, j As Long
     Dim strFileName  As String, strFile As String, str版本号 As String
     Dim strArr, strArr1
     Dim lst As ListItem
     Dim objFile As New FileSystemObject
     
     strArr = Split(gstr收集类型, ";")
     
     strFileName = ""
     
     For i = 0 To UBound(strArr)
        strArr1 = Split(strArr(i), ",")
        For j = 0 To UBound(strArr1)
            If InStr(1, strArr1(j), ".") <> 0 Then
                '如果存在小数点,则表示特定文件
                strFileName = strFileName & ";" & strArr(i)
            Else
                strFileName = strFileName & ";*." & strArr(i)
            End If
        Next
     Next
     If strFileName <> "" Then
        strFileName = Mid(strFileName, 2)
     Else
        strFileName = "*.DLSDKSKS"
     End If
     Err = 0
     On Error GoTo ErrHand:
     i = 1
   
    With File1
        .Path = gstrAppPath
        .FileName = strFileName
        lvwMan.ListItems.Clear
        For lngfile = 0 To .ListCount - 1
            '不传SQLNET.Log文件
            If InStr(1, UCase(.List(lngfile)), "SQLNET.LOG") = 0 Then
                strFile = gstrAppPath & "\" & .List(lngfile)
                str版本号 = GetCommpentVersion(strFile)
                Set lst = lvwMan.ListItems.Add(, "K" & i, .List(lngfile), "List", "List")
                    lst.Tag = objFile.GetExtensionName(.List(lngfile))
                    lst.SubItems(2) = str版本号
                    lst.SubItems(4) = Format(FileDateTime(strFile), "yyyy-MM-DD hh:mm:ss") 'Format(FileDateTime(strFile), "yyyy-MM-dd hh:mm:ss")
            End If
            i = i + 1
        Next
    End With
    GetClientFiles = True
    Exit Function
ErrHand:
    GetClientFiles = False
End Function

Private Function getSeverFiles() As Boolean
    '----------------------------------------------------------------------------------------
    '功能:获取服务器的最新的升级部件
    '参数:
    '返回:填写成功,返回true,否则返回False
    '----------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim lst As ListItem
    Dim rsTmp As New ADODB.Recordset
    getSeverFiles = False
    
    strSQL = "Select 序号, 文件名, 版本号, 修改日期,文件类型,业务部件,安装路径,MD5,自动注册,强制覆盖 From zlfilesupgrade where upper(文件名) not in('ZLHISCRUST.EXE' , '7Z.EXE','7Z.DLL') and MD5 is not null order by 序号"

    lvwMan.ListItems.Clear
    With rsTmp
        .CursorLocation = adUseClient
        .Open strSQL, gcnnOracle
        If .RecordCount = 0 Then
            getSeverFiles = True
            Exit Function
        End If
        Err = 0
        On Error GoTo ErrHand:
        Do While Not .EOF
            If UCase(NVL(!文件名)) = "ZLHISCRUST.EXE" Or UCase(NVL(!文件名)) = "7Z.EXE" Or UCase(NVL(!文件名)) = "7Z.DLL" Or UCase(NVL(!文件名)) = "ZLRUNAS.EXE" Then
                GoTo NotAddFile
            End If
            Set lst = lvwMan.ListItems.Add(, "K" & .AbsolutePosition, !文件名, "List", "List")
            lst.Tag = IIf(IsNull(!文件类型), 0, !文件类型)
            lst.SubItems(2) = IIf(IsNull(!版本号), "", GetVersion(IIf(IsNull(!版本号), 0, !版本号)))
            lst.SubItems(4) = Format(!修改日期, "yyyy-MM-DD HH:mm:ss")
            lst.SubItems(6) = NVL(!业务部件, "")
            lst.SubItems(7) = NVL(!安装路径, "")
            lst.SubItems(8) = NVL(!MD5, "")
            lst.SubItems(9) = NVL(!自动注册, "")
            lst.SubItems(10) = NVL(!强制覆盖)
NotAddFile:
            .MoveNext
        Loop
    End With
    
    getSeverFiles = True
    Exit Function
ErrHand:
    getSeverFiles = False
End Function

Private Function GetPreUpgrad(ByVal strComputerName As String) As Boolean
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    On Error GoTo ErrHand
    
    strSQL = "Select * From zlClients Where  预升完成=1 and upper(工作站)='" & strComputerName & "'"
    
     With rsTmp
        .CursorLocation = adUseClient
        .Open strSQL, gcnnOracle
        If .RecordCount = 1 Then
            GetPreUpgrad = True
            Exit Function
        Else
            GetPreUpgrad = False
            Exit Function
        End If
    End With
    Exit Function
ErrHand:
    GetPreUpgrad = False
End Function

Private Function CompareMD5Down(ByVal strLocalFile As String, ByVal strName As String, Optional strMsg As String) As Boolean
    '功能:比较文件的MD5是否相同
    '入参:strLocalFile 需要比较文件完整路径,strName文件名称,intOption 处理返回值，2种返回方法
    '返回:True 需要下载 Flase 不需要下载
    '编制:祝庆
    '日期:2010/12/15
    On Error GoTo errH
    Dim objFile As New FileSystemObject
    Dim strFileMD5 As String
    Dim strListFileMD5 As String
    If objFile.FileExists(strLocalFile) Then
        strFileMD5 = HashFile(strLocalFile, 2 ^ 27)
        strListFileMD5 = GetFileListValue(strName, 0)
        If strListFileMD5 = "" Then
            strMsg = "服务器没有该文件MD5信息!"
        End If
        
        If strFileMD5 = strListFileMD5 Then
            CompareMD5Down = False
        Else
            CompareMD5Down = True
        End If
    Else
        CompareMD5Down = True
    End If
    Exit Function
errH:
    If Err Then
        CompareMD5Down = True
    End If
End Function

Private Function GetFileListValue(ByVal strFileName As String, ByVal intOption As Integer) As String
    '功能从服务器列表获取文件的信息
    '如参 strFileName 需要获取信息的文件名
    '0:获取MD5值
    '1:获取版本号
    '2:获取修改日期
    On Error GoTo errH
    Dim i As Integer
    Dim lngCurFileIndex As Long
    lngCurFileIndex = -1
    With lvwMan
        For i = 1 To .ListItems.Count
            If UCase(.ListItems(i).Text) = UCase(strFileName) Then
                lngCurFileIndex = i
                Exit For
            End If
        Next
        
        If lngCurFileIndex >= 0 Then
            Select Case intOption
            Case 0
                GetFileListValue = .ListItems(i).SubItems(8) '文件MD5值
            Case 1
                GetFileListValue = IIf(.ListItems(i).SubItems(2) = "0", "", .ListItems(i).SubItems(2)) '版本号
            Case 2
                GetFileListValue = .ListItems(i).SubItems(4) '修改日期
            Case 3
                GetFileListValue = NVL(.ListItems(i).SubItems(10), 0)
            End Select
        Else
            GetFileListValue = ""
        End If
    End With
    Exit Function
errH:
    If Err Then
        GetFileListValue = ""
    End If
End Function

'处理USER权限的权限问题
Private Function GetAdministrator() As Boolean
     Dim strAppRunas As String
     Dim strUser As String
     Dim strPass As String
     Dim strMsg As String
     Dim strSQL As String
     Dim rsTmpUser As New ADODB.Recordset
     Dim rsTmpPass As New ADODB.Recordset
 
        '1.检查本地是否有zlRunas文件.
        strAppRunas = App.Path & "\zlRunas.exe"
        If mobjFile.FileExists(strAppRunas) = False Then
            '不存在就重服务器中下载
            
            
            '网络是否联通
            If IIf(gbln收集 = True, gintGatherTYpe = 0, gintUpType = 0) Then
                If IsNetServer = False Then
                    MsgBox "无法连接到服务器:" & gstrServerPath & "上,请确认网络是否畅通," & vbCrLf _
                    & "或检查管理工具中文件" & IIf(gbln收集 = True, "收集", "升级") & "服务是否设置正确!", vbInformation + vbDefaultButton1, "客户端自动" & IIf(gbln收集 = True, "收集", "升级")
                    
                    Call SaveClientLog("无法连接到服务器:" & gstrServerPath & "上,请确共享服务器是否畅通。")
                    Call UpdateCondition(2)
                    End:
                    Exit Function '连接共享服务器
                End If
            Else
                If IsFtpServer = False Then
                    MsgBox "无法连接到服务器:" & gstrServerPath & "上,请确认FTP服务器是否开启," & vbCrLf _
                    & "或检查管理工具中文件" & IIf(gbln收集 = True, "收集", "升级") & "服务是否设置正确!", vbInformation + vbDefaultButton1, "客户端自动" & IIf(gbln收集 = True, "收集", "升级")
                    
                    Call SaveClientLog("无法连接到服务器:" & gstrServerPath & "上,请确FTP服务器是否畅通。")
                    Call UpdateCondition(2)
                    End:
                    Exit Function '连接FTP服务器
                End If
            End If
            
            Call isRunasUpGrade '下载Runas文件
            
            Sleep 100
        End If
        
        
        
         '2.获取本客户端管理员用户及密码
        '判断是否下载成功,成功才继续执行
        If mobjFile.FileExists(strAppRunas) = False Then
            MsgBox "未能正常下载ZLRUNAS,USER权限执行工具" & vbNewLine & "请检查服务器目录中是否存在该文件.", vbInformation + vbDefaultButton1, "客户端自动升级"
            
            
            Call SaveClientLog("未能正常下载ZLRUNAS,USER权限执行工具,请检查服务器目录中是否存在该文件.")
            GetAdministrator = False
            Exit Function
        End If
          

        strSQL = "Select 项目,内容 From zlRegInfo where 项目 like '管理员账号'"
        With rsTmpUser
            .Open strSQL, gcnnOracle
            If rsTmpUser.RecordCount = 1 Then
                strUser = Trim(NVL(rsTmpUser!内容))
                
                
                strSQL = "Select 项目,内容 From zlRegInfo where 项目 like '管理员密码'"
                rsTmpPass.Open strSQL, gcnnOracle
                If rsTmpPass.RecordCount = 1 Then
                    strPass = decipher(Trim(NVL(rsTmpPass!内容)))
                Else
                    strPass = ""
                End If
                
            Else
                strUser = "Administrator"
                strPass = ""
            End If
        End With
        
        '3.执行zlRunas , 使用管理员权限登录
        strMsg = RunasShell(strUser, strPass)
        If strMsg <> "" Then
            MsgBox strMsg, vbInformation + vbDefaultButton1, "客户端自动升级"
            
            Call SaveClientLog(strMsg & "[ZLRUNAS]")
            
            GetAdministrator = False
            Exit Function
        End If
        
        '4.退出程序
        Unload Me
End Function

Private Function RunasShell(ByVal strUser As String, ByVal strPass As String) As String
    On Error GoTo errH
    Dim strRunas As String
    Dim strApp As String
    Dim strMsg As String
    Dim strShellTxt  As String
    strRunas = App.Path & "\ZLRUNAS.EXE"
    strApp = App.Path & "\ZLHISCRUST.EXE"
    '路径中不能有中文，否则执行不成功
    strShellTxt = strRunas & " -u " & strUser & " -p " & strPass & " -ex """ & strApp & """" & " -lwp"
    strMsg = GetCmdTxt(strShellTxt)
    
    If InStr(strMsg, (1326)) > 0 Then
        RunasShell = "登录失败: 未知的用户名或错误密码。"
        Exit Function
    End If
    
    If InStr(strMsg, (1058)) > 0 Then
        RunasShell = "无法启动服务，原因可能是SecLogon服务被禁用。"
        Exit Function
    End If
    
    If InStr(strMsg, (1717)) > 0 Then
        RunasShell = "'路径中不能有中文，否则执行不成功"
        Exit Function
    End If
    
    RunasShell = ""
    Exit Function
errH:
    If Err Then
        RunasShell = ""
    End If
End Function
