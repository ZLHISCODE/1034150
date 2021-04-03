VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.Unicode.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Begin VB.Form frmMessageManager 
   Caption         =   "消息收发管理"
   ClientHeight    =   7245
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   11265
   Icon            =   "frmMessageManager.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7245
   ScaleWidth      =   11265
   ShowInTaskbar   =   0   'False
   Tag             =   "可变化的"
   Begin XtremeReportControl.ReportControl rpcMain 
      Height          =   2160
      Left            =   6045
      TabIndex        =   7
      Top             =   1605
      Width           =   3660
      _Version        =   589884
      _ExtentX        =   6456
      _ExtentY        =   3810
      _StockProps     =   0
      BorderStyle     =   3
   End
   Begin VB.PictureBox picSplit 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3225
      Left            =   2610
      ScaleHeight     =   3225
      ScaleMode       =   0  'User
      ScaleWidth      =   33.75
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1140
      Width           =   45
   End
   Begin VB.PictureBox picSplitH 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   2220
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   3000
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   4245
      Width           =   3000
   End
   Begin VB.PictureBox picCon 
      Height          =   6015
      Left            =   15
      ScaleHeight     =   5955
      ScaleWidth      =   1950
      TabIndex        =   8
      Top             =   960
      Width           =   2010
      Begin XtremeSuiteControls.TaskPanel tplCon 
         Height          =   4770
         Left            =   -630
         TabIndex        =   9
         Top             =   240
         Width           =   3210
         _Version        =   589884
         _ExtentX        =   5662
         _ExtentY        =   8414
         _StockProps     =   64
         Behaviour       =   1
         ItemLayout      =   3
         HotTrackStyle   =   3
      End
   End
   Begin ComCtl3.CoolBar cbarTool 
      Align           =   1  'Align Top
      Height          =   870
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   11265
      _ExtentX        =   19870
      _ExtentY        =   1535
      BandCount       =   2
      BandBorders     =   0   'False
      _CBWidth        =   11265
      _CBHeight       =   870
      _Version        =   "6.7.9782"
      MinHeight1      =   285
      Width1          =   8370
      Key1            =   "only"
      NewRow1         =   0   'False
      MinHeight2      =   525
      NewRow2         =   -1  'True
      AllowVertical2  =   0   'False
      Begin XtremeCommandBars.ImageManager imgPublic 
         Left            =   1035
         Top             =   345
         _Version        =   589884
         _ExtentX        =   635
         _ExtentY        =   635
         _StockProps     =   0
         Icons           =   "frmMessageManager.frx":39CA
      End
      Begin XtremeCommandBars.CommandBars cbsMain 
         Left            =   240
         Top             =   120
         _Version        =   589884
         _ExtentX        =   635
         _ExtentY        =   635
         _StockProps     =   0
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Bindings        =   "frmMessageManager.frx":1253C
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   6885
      Width           =   11265
      _ExtentX        =   19870
      _ExtentY        =   635
      SimpleText      =   $"frmMessageManager.frx":12550
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmMessageManager.frx":12597
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14790
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtfContent 
      Height          =   1485
      Left            =   2190
      TabIndex        =   4
      Top             =   3180
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   2619
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmMessageManager.frx":12E2B
   End
   Begin VB.PictureBox picTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00848484&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   4590
      ScaleHeight     =   405
      ScaleWidth      =   1485
      TabIndex        =   5
      Top             =   1050
      Width           =   1485
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "收件箱"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   90
         TabIndex        =   6
         Top             =   60
         Width           =   990
      End
   End
   Begin MSComDlg.CommonDialog cdg 
      Left            =   2760
      Top             =   1350
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin XtremeCommandBars.ImageManager imgRptIcon 
      Left            =   9480
      Top             =   1425
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmMessageManager.frx":12EC8
   End
   Begin XtremeCommandBars.ImageManager imgICon 
      Left            =   2115
      Top             =   1260
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmMessageManager.frx":15798
   End
End
Attribute VB_Name = "frmMessageManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnLoad As Boolean   '窗口还未打开时为真

Dim mstrKey As String     '未更新的邮件ID
Dim sngStartY As Single   '移动前鼠标的位置
Dim mblnItem As Boolean   '为真表示单击到ListView某一项上
Dim mintColumn As Integer '用于ListView列排序

Public mlngIndexPre As Long       '表示之前是哪个目录
Public mlngIndex As Long          '表示当前是哪个目录
Public mstrPrivs As String        '只是消息收发的模块的权限
Public mblnShowAll As Boolean     '显示已读
Public mblnLogin As Boolean       '登录时显示提醒未读邮件
Const con草稿 = 0
Const con收件箱 = 1
Const con已发送消息 = 2
Const con已删除消息 = 3


Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)

    On Error Resume Next
    Dim objControl As CommandBarControl
    Dim i As Integer
    Dim lngID As Long, lng会话ID As Long
    
    Select Case Control.ID
    Case conMenu_File_PrintSet
        '打印设置
        Call zlPrintSet
    Case conMenu_File_Preview
        '预览
        Call subPrint(2)
    Case conMenu_File_Print
        '打印
        Call subPrint(1)
    Case conMenu_File_Excel
        '输出到Excel
        Call subPrint(3)
    Case conMenu_File_SaveAs
        '另存为文件
        On Error Resume Next
        If rtfContent.Text = "" Then Exit Sub
        cdg.CancelError = True
        cdg.Filter = "RTF文件(*.RTF)|*.rtf"
        '覆盖时有提示，且不能是只读的
        cdg.flags = cdlOFNOverwritePrompt Or cdlOFNNoReadOnlyReturn
        cdg.ShowSave
        
        If Err = 0 Then
            MousePointer = 11
            rtfContent.SaveFile cdg.FileName
            MousePointer = 0
        Else
            Err.Clear
        End If
    Case conMenu_Edit_Add
        '新增
        frmMessageEdit.OpenWindow "", ""
        Call FillList
    Case conMenu_Edit_Modify
        '打开,修改
        Call Edit_Modify
    Case conMenu_Edit_Delete
        '删除
        Call Edit_Delete
    Case conMenu_Edit_Reuse
        '还原
        Call Edit_Restore
    Case conMenu_Edit_Reply
        '答复
        If rpcMain.FocusedRow Is Nothing Then Exit Sub
        lngID = Val(rpcMain.Rows(rpcMain.FocusedRow.Index).Record(0).Caption)
        lng会话ID = Val(rpcMain.Rows(rpcMain.FocusedRow.Index).Record(1).Caption)
        frmMessageEdit.OpenWindow "", lngID, lng会话ID, 1
        Call FillList
    Case conMenu_Edit_AllReply
        '全部答复
        If rpcMain.FocusedRow Is Nothing Then Exit Sub
        lngID = Val(rpcMain.Rows(rpcMain.FocusedRow.Index).Record(0).Caption)
        lng会话ID = Val(rpcMain.Rows(rpcMain.FocusedRow.Index).Record(1).Caption)
        frmMessageEdit.OpenWindow "", lngID, lng会话ID, 2
        Call FillList
    Case conMenu_Edit_Transmit
        '转发
        If rpcMain.FocusedRow Is Nothing Then Exit Sub
        lngID = Val(rpcMain.Rows(rpcMain.FocusedRow.Index).Record(0).Caption)
        lng会话ID = Val(rpcMain.Rows(rpcMain.FocusedRow.Index).Record(1).Caption)
        frmMessageEdit.OpenWindow "", lngID, lng会话ID, 3
        Call FillList
    Case conMenu_View_ToolBar_Button
        '工具栏
        For i = 2 To cbsMain.Count
            Me.cbsMain(i).Visible = Not Me.cbsMain(i).Visible
            
            cbarTool.Bands.item(2).NewRow = Not cbarTool.Bands.item(2).NewRow
            
            If cbarTool.Bands.item(2).NewRow = True Then
                If Me.cbsMain.Options.LargeIcons = True Then
                    cbarTool.Bands.item(2).MinHeight = 520
                Else
                    cbarTool.Bands.item(2).MinHeight = 425
                End If
            Else
                cbarTool.Bands.item(2).MinHeight = cbarTool.Bands(1).MinHeight
            End If
        Next
        Me.cbsMain.RecalcLayout
        Call Form_Resize
    Case conMenu_View_ToolBar_Text
        '按钮文字
        For i = 2 To cbsMain.Count
            For Each objControl In Me.cbsMain(i).Controls
                objControl.Style = IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
        Next
        Me.cbsMain.RecalcLayout
        Call Form_Resize
    Case conMenu_View_ToolBar_Size
        '大图标
        Me.cbsMain.Options.LargeIcons = Not Me.cbsMain.Options.LargeIcons
        
        If Me.cbsMain.Options.LargeIcons = True Then
            cbarTool.Bands.item(2).MinHeight = 520
        Else
            cbarTool.Bands.item(2).MinHeight = 425
        End If
        Me.cbsMain.RecalcLayout
        Call Form_Resize
    Case conMenu_View_StatusBar
        '状态栏
        Me.stbThis.Visible = Not Me.stbThis.Visible
        Me.cbsMain.RecalcLayout
        Call Form_Resize
    Case conMenu_View_PreviewWindow
        '预览窗格
        Me.rtfContent.Visible = Not Me.rtfContent.Visible
        picSplitH.Visible = Me.rtfContent.Visible
        Me.cbsMain.RecalcLayout
        Call Form_Resize
    Case conMenu_View_ShowAll
        '显示已读
        mblnShowAll = Not mblnShowAll
        Me.cbsMain.RecalcLayout
        mlngIndexPre = -1 '强制刷新
        Call FillList
    Case conMenu_View_Login
       '登录时提醒
        mblnLogin = Not mblnLogin
        Call zlDatabase.SetPara("登录检查邮件消息", IIf(mblnLogin, "1", "0"))
        Me.cbsMain.RecalcLayout
    Case conMenu_View_Find
        '查找相关消息
        lng会话ID = Val(rpcMain.Rows(rpcMain.FocusedRow.Index).Record(1).Caption)
        frmMessageRelate.FillList lng会话ID
    Case conMenu_View_Refresh
        '刷新
        mlngIndexPre = -1 '强制刷新
        Call FillList
    Case conMenu_Help_Help
        '帮助
        Call ShowHelp(App.ProductName, Me.hWnd, "ZL9AppTool\" & Me.Name, 0)
    Case conMenu_Help_Web_Home
        'Web上的中联
        Call zlHomePage(Me.hWnd)
    Case conMenu_Help_Web_Forum
        '中联论坛
        Call zlWebForum(Me.hWnd)
    Case conMenu_Help_Web_Mail
        '发送反馈
         Call zlMailTo(Me.hWnd)
    Case conMenu_Help_About
        '关于
        ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
    Case conMenu_File_Exit
        '退出
        Unload Me
    End Select
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnEnabled As Boolean
    
    blnEnabled = Not (rpcMain.FocusedRow Is Nothing)
    
    '权限控制
    If InStr(mstrPrivs, "发送消息") = 0 Then
        Select Case Control.ID
        Case conMenu_Edit_Add, conMenu_Edit_Reply, conMenu_Edit_AllReply, conMenu_Edit_Transmit
            '增加，答复，全部答复，转发
            Control.Enabled = False
            
        End Select
    End If
    
    '菜单设置
    blnEnabled = Not (rpcMain.FocusedRow Is Nothing)
    Dim lngCount As Long, lngSum As Long, lngMessage As Long
        
    For lngCount = 0 To rpcMain.Rows.Count - 1
        If Not rpcMain.Rows(lngCount).Record Is Nothing Then
            If rpcMain.Rows(lngCount).Record(3).Icon = 1 Or rpcMain.Rows(lngCount).Record(3).Icon = 3 Then
                lngSum = lngSum + 1
            End If
            lngMessage = lngMessage + 1
        End If
    Next
    stbThis.Panels(2).Text = "共有" & lngMessage & "条消息" & IIf(lngSum = 0, "。", "，其中有" & lngSum & "条未读。")

    Select Case Control.ID
    Case conMenu_Edit_Modify, conMenu_Edit_Delete, conMenu_Edit_Reply, conMenu_Edit_AllReply, conMenu_Edit_Transmit, conMenu_View_Find
        '修改，删除，答复，全部答复，转发,查找相关消息
        Control.Enabled = blnEnabled
    Case conMenu_Edit_Reuse
        '还原
        Control.Enabled = (mlngIndex = 3 And Not (rpcMain.FocusedRow Is Nothing))
    Case conMenu_File_Print, conMenu_File_Preview, conMenu_File_Excel
        '打印,预览,输出到Excel
        Control.Enabled = lngMessage > 0
    Case conMenu_File_SaveAs
        '另存为
        Control.Enabled = rtfContent.Text <> ""
    Case conMenu_View_ToolBar_Button '工具栏
        If cbsMain.Count >= 2 Then
            Control.Checked = Me.cbsMain(2).Visible
        End If
    Case conMenu_View_ToolBar_Text '图标文字
        If cbsMain.Count >= 2 Then
            Control.Checked = Not (Me.cbsMain(2).Controls(1).Style = xtpButtonIcon)
        End If
    Case conMenu_View_ToolBar_Size '大图标
        Control.Checked = Me.cbsMain.Options.LargeIcons
    Case conMenu_View_StatusBar '状态栏
        Control.Checked = Me.stbThis.Visible
    Case conMenu_View_PreviewWindow '预览窗格
        Control.Checked = Me.rtfContent.Visible
    Case conMenu_View_ShowAll '显示已读
        Control.Checked = mblnShowAll
    Case conMenu_View_Login '登录提醒
        Control.Checked = mblnLogin
    End Select
    
End Sub

Private Sub Form_Activate()
    If mblnLoad = True Then
        Call Form_Resize '为了使CoolBar自适应高度
        mlngIndexPre = -1 '强制刷新
        Call FillList
    End If
    mblnLoad = False
End Sub

Private Sub Form_Load()
    gblnMessageShow = True
    If gblnMessageGet = False Then
        '导航台并没有打开消息通知窗口，只有自己把它打开
        Load frmMessageRead
    End If
    Call DeleteMessage

    mblnLoad = True
    '-----------
    RestoreWinState Me, App.ProductName
    mblnShowAll = Val(zlDatabase.GetPara("显示已读邮件")) <> 0
    mblnLogin = Val(zlDatabase.GetPara("登录检查邮件消息")) <> 0
    
    mstrPrivs = GetPrivFunc(0, 12) '取权限
    '-----------------------------------------------------------------------------------------------------------
        '新界面
        Call InitCommandBar
        Dim tpGroup As TaskPanelGroup
    
        Set tpGroup = tplCon.Groups.Add(101, "邮箱")
        
        tpGroup.Items.Add(con草稿, "草稿", xtpTaskItemTypeLink, con草稿 + 2).Selected = False
        tpGroup.Items.Add(con收件箱, "收件箱", xtpTaskItemTypeLink, con收件箱 + 2).Selected = False
        tpGroup.Items.Add(con已发送消息, "已发送消息", xtpTaskItemTypeLink, con已发送消息 + 2).Selected = False
        tpGroup.Items.Add(con已删除消息, "已删除消息", xtpTaskItemTypeLink, con已删除消息 + 2).Selected = False
        
        With tplCon
            .SetMargins 1, 2, 0, 2, 2
            .SelectItemOnFocus = True
            .VisualTheme = xtpTaskPanelThemeOfficeXPPlain
            Call .Icons.AddIcons(imgICon.Icons)
            .SetIconSize 32, 32
        End With
        tpGroup.CaptionVisible = False
        tpGroup.Expanded = True
        
        
        Call init_rpcMain
        
    '-----------------------------------------------------------------------------------------------------------
    '设置初始化选中
    mlngIndex = 1
End Sub

Private Sub Form_Resize()
    Dim sngTop As Single, sngBottom As Single
    On Error Resume Next

    sngTop = IIf(cbarTool.Visible, cbarTool.Top + cbarTool.Height, 0)
    sngBottom = Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0)
    
    picCon.Top = sngTop + 30
    picCon.Height = IIf(sngBottom - picCon.Top > 0, sngBottom - picCon.Top, 0)
    picCon.Left = 0
    
    picSplit.Top = sngTop
    picSplit.Height = IIf(sngBottom - picSplit.Top > 0, sngBottom - picSplit.Top, 0)
    picSplit.Left = picCon.Left + picCon.Width
    
    picTitle.Top = sngTop + 30
    rpcMain.Left = picSplit.Left + picSplit.Width
    rpcMain.Top = picTitle.Top + picTitle.Height + 60
    
    If Me.ScaleWidth - rpcMain.Left > 0 Then rpcMain.Width = Me.ScaleWidth - rpcMain.Left
    picTitle.Left = rpcMain.Left
    picTitle.Width = rpcMain.Width
    If rtfContent.Visible = True Then
        rpcMain.Height = (sngBottom - rpcMain.Top) * (rpcMain.Height / (rpcMain.Height + picSplitH.Height + rtfContent.Height))
        
        picSplitH.Left = rpcMain.Left
        picSplitH.Top = rpcMain.Top + rpcMain.Height
        picSplitH.Width = rpcMain.Width
        
        rtfContent.Left = rpcMain.Left
        rtfContent.Top = picSplitH.Top + picSplitH.Height
        rtfContent.Height = sngBottom - rtfContent.Top
        rtfContent.Width = rpcMain.Width
    Else
        rpcMain.Height = sngBottom - rpcMain.Top
    End If
    
'    rpcMain.Top = lvwMain.Top
'    rpcMain.Width = lvwMain.Width / 2
'    rpcMain.Left = lvwMain.Left + lvwMain.Width / 2
'    rpcMain.Height = lvwMain.Height
'    lvwMain.Width = lvwMain.Width / 2
    
    tplCon.Top = 0
    tplCon.Left = picCon.Left
    tplCon.Width = picCon.Width
    tplCon.Height = picCon.ScaleHeight

    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    gblnMessageShow = False
    If gblnMessageGet = False Then
        '导航台并没有打开消息通知窗口，因此退出时把它一并关了
        Unload frmMessageRead
    End If
    
    mstrKey = ""
    mlngIndexPre = 0
    Call zlDatabase.SetPara("显示已读邮件", IIf(mblnShowAll, 1, 0))
    SaveWinState Me, App.ProductName
End Sub


Private Sub Edit_Delete()
    On Error GoTo errHandle
    Dim strKey As String
    Dim intIndex As Long
    Dim rsTemp As New ADODB.Recordset
    
    Dim lngID As Long, lng会话ID As Long
    If rpcMain.FocusedRow Is Nothing Then Exit Sub
    lngID = Val(rpcMain.Rows(rpcMain.FocusedRow.Index).Record(0).Caption)
    lng会话ID = Val(rpcMain.Rows(rpcMain.FocusedRow.Index).Record(1).Caption)
    
    gcnOracle.BeginTrans
    If mlngIndex <> 3 Then
        'Delete_Zlmsgstate(删除,消息ID,类型,用户)
        Call zlDatabase.OpenCursor(Me.Caption, "zltools", "b_ComFunc.Delete_Zlmsgstate", _
                                    1, _
                                    lngID, _
                                    lng会话ID, _
                                    gstrDbUser)
    Else
        If MsgBox("你确认要删除主题为“" & rpcMain.Rows(rpcMain.FocusedRow.Index).Record(3).Caption & "”的消息吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
            gcnOracle.RollbackTrans
            Exit Sub
        End If
        Me.MousePointer = 11
        Call zlDatabase.OpenCursor(Me.Caption, "zltools", "b_ComFunc.Delete_Zlmsgstate", _
                                            2, _
                                            lngID, _
                                            lng会话ID, _
                                            gstrDbUser)
        Me.MousePointer = 0
    End If
    gcnOracle.CommitTrans

    Call FillList
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    gcnOracle.RollbackTrans
    MousePointer = 0
End Sub

Private Sub Edit_Restore()
'还原已删除消息
    On Error GoTo errHandle
'    Dim intIndex As Long
    
    Dim lngID As Long, lng会话ID As Long
    lngID = Val(rpcMain.Rows(rpcMain.FocusedRow.Index).Record(0).Caption)
    lng会话ID = Val(rpcMain.Rows(rpcMain.FocusedRow.Index).Record(1).Caption)
    
    'b_ComFunc.Restore_Zlmsgstate(消息ID,类型,用户)
    Call zlDatabase.OpenCursor(Me.Caption, "zltools", "b_ComFunc.Restore_Zlmsgstate", _
                                lngID, _
                                lng会话ID, _
                                gstrDbUser)
    Call FillList

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub



Private Sub picSplitH_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        sngStartY = Y
    End If
End Sub

Private Sub picSplitH_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sngTemp As Single
    On Error Resume Next

    If Button = 1 Then
        sngTemp = picSplitH.Top + Y - sngStartY
        If sngTemp - rpcMain.Top > 2500 And IIf(stbThis.Visible = True, stbThis.Top, Me.ScaleHeight) - (sngTemp + picSplitH.Height) > 1200 Then
            picSplitH.Top = sngTemp
            rpcMain.Height = picSplitH.Top - rpcMain.Top
            rtfContent.Top = picSplitH.Top + picSplitH.Height
            rtfContent.Height = IIf(stbThis.Visible = True, stbThis.Top, Me.ScaleHeight) - rtfContent.Top
            
            rpcMain.Height = rpcMain.Height
        End If
        rpcMain.SetFocus
    End If
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileExcel_Click()
    subPrint 3
End Sub

Private Sub mnuFilePreview_Click()
    subPrint 2
End Sub

Private Sub mnuFilePrint_Click()
    subPrint 1
End Sub

Private Sub mnufileset_Click()
    zlPrintSet
End Sub

Private Sub rpcMain_GotFocus()
    Call FillText
End Sub

Private Sub rpcMain_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Dim objMenu As CommandBarPopup, menuItem As CommandBarButton
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_EditPopup)
        Set menuItem = objMenu.CommandBar.FindControl(, conMenu_Edit_Modify)
        If menuItem.Enabled Then Edit_Modify
    End If

End Sub

Private Sub rpcMain_MouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Call FillText
End Sub

Private Sub rpcMain_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    
    If Button = 2 Then
        Dim Control As CommandBarControl, objControl As CommandBarControl
        Dim Popup As CommandBar
        
        Set Popup = cbsMain.Add("Popup", xtpBarPopup)
        
        With Popup.Controls
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Add, "增加(&A)")
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, "打开(&O)")
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除(&D)")
            Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)"): objControl.BeginGroup = True
        End With
            
        Popup.ShowPopup
    End If
End Sub

Private Sub rpcMain_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal item As XtremeReportControl.IReportRecordItem)
    If rpcMain.FocusedRow Is Nothing Then Exit Sub
    
    Dim objMenu As CommandBarPopup, menuItem As CommandBarButton
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_EditPopup)
    Set menuItem = objMenu.CommandBar.FindControl(, conMenu_Edit_Modify)
    
    If menuItem.Enabled Then Edit_Modify
End Sub

Private Sub rpcMain_SelectionChanged()
    If rpcMain.FocusedRow.Index = 0 Then Exit Sub
    Call FillText
End Sub

Private Sub subPrint(bytMode As Byte)
'功能:进行打印,预览和输出到EXCEL
'参数:bytMode=1 打印;2 预览;3 输出到EXCEL
    If gstrUserName = "" Then Call GetUserInfo
    Dim objPrint As Object
    
    Set objPrint = New zlPrintLvw
    objPrint.Title.Text = IIf(InStr(lblTitle.Caption, "消息") > 0, lblTitle.Caption, lblTitle.Caption & "里的消息")
   ' Set objPrint.Body.objData = lvwMain
    objPrint.BelowAppItems.Add "打印人：" & gstrUserName
    objPrint.BelowAppItems.Add "打印时间：" & Format(zlDatabase.CurrentDate, "yyyy年MM月dd日")
    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrViewLvw objPrint, 1
          Case 2
              zlPrintOrViewLvw objPrint, 2
          Case 3
              zlPrintOrViewLvw objPrint, 3
      End Select
    Else
        zlPrintOrViewLvw objPrint, bytMode
    End If
End Sub

Public Sub FillList()
'功能:装入消息到lvwMain

    Dim rsTemp As New ADODB.Recordset
    Dim lst As ListItem
    'Dim strKey As String
    Dim strTemp As String
    Dim strICO As String
    
    '如果还是同一个目录，就退出
    'If mlngIndexPre = mlngIndex Then Exit Sub
    mlngIndexPre = mlngIndex
    mstrKey = ""
    '保存当前的选中项
    Select Case mlngIndex
        Case 0
            lblTitle.Caption = "草稿"
        Case 1
            lblTitle.Caption = "收件箱"
        Case 2
            lblTitle.Caption = "已发送消息"
        Case 3
            lblTitle.Caption = "已删除消息"
    End Select
    rsTemp.CursorLocation = adUseClient
    
    'Get_Mail_List(消息类型_In,用户_In,显示已读_In,会话ID)
    Set rsTemp = zlDatabase.OpenCursor(Me.Caption, "zlTools", "b_ComFunc.Get_Mail_List", _
                                        lblTitle.Caption, _
                                        gstrDbUser, _
                                        IIf(mblnShowAll, 0, 1), 0)
'    lvwMain.ListItems.Clear
''---------------------
    Dim rptItem As ReportRecordItem
    Dim rptRecord As ReportRecord
    With rpcMain
        .Records.DeleteAll
        '.SetImageList imgIcon
        
        Do Until rsTemp.EOF
            Set rptRecord = .Records.Add
            
            Set rptItem = rptRecord.AddItem(IIf(IsNull(rsTemp!ID), "0", rsTemp!ID))
            rptItem.Caption = IIf(IsNull(rsTemp!ID), "0", rsTemp!ID)
            
            Set rptItem = rptRecord.AddItem(IIf(IsNull(rsTemp!类型), "", rsTemp!类型))
            rptItem.Caption = IIf(IsNull(rsTemp!类型), "", rsTemp!类型)
            'Item.Icon = 2
            
            strTemp = IIf(IsNull(rsTemp("状态")), "0000", rsTemp("状态"))
            If Mid(strTemp, 4, 1) <> "0" Then
                Set rptItem = rptRecord.AddItem(IIf(Mid(strTemp, 4, 1) = 1, "高", "低"))
                rptItem.Caption = IIf(Mid(strTemp, 4, 1) = 1, "高", "低")
                rptItem.Icon = IIf(Mid(strTemp, 4, 1) = 1, 5, 6)
            Else
                Set rptItem = rptRecord.AddItem("")
                rptItem.Caption = ""
            End If
            
            
            Set rptItem = rptRecord.AddItem(IIf(IsNull(rsTemp!主题), "", rsTemp!主题))
            rptItem.Caption = IIf(IsNull(rsTemp!主题), "", rsTemp!主题)
            
            If rsTemp("类型") = 0 Then
                strICO = "Script"
            Else
                strICO = IIf(Mid(strTemp, 1, 1) = "1", "Read", "New") & IIf(Mid(strTemp, 2, 2) <> "00", "Reply", "")   '已读+已处理
            End If
            
            If strICO = "Script" Then
                rptItem.Icon = 7
                rptItem.Bold = False
            End If
            If strICO = "New" Then
                rptItem.Icon = 1
                rptItem.Bold = True
            End If
            
            If strICO = "Read" Then
                rptItem.Icon = 2
                rptItem.Bold = False
            End If
            If strICO = "NewReply" Then
                rptItem.Icon = 3
                rptItem.Bold = True
            End If
            If strICO = "ReadReply" Then
                rptItem.Icon = 4
                rptItem.Bold = False
            End If
            
            Set rptItem = rptRecord.AddItem(IIf(IsNull(rsTemp!发件人), "", rsTemp!发件人))
            rptItem.Caption = IIf(IsNull(rsTemp!发件人), "", rsTemp!发件人)
            
            Set rptItem = rptRecord.AddItem(IIf(IsNull(rsTemp!收件人), "", Trim(rsTemp!收件人)))
            rptItem.Caption = IIf(IsNull(rsTemp!收件人), "", Trim(rsTemp!收件人))
            
            Set rptItem = rptRecord.AddItem(IIf(IsNull(rsTemp!时间), "", Trim(rsTemp!时间)))
            rptItem.Caption = IIf(IsNull(rsTemp!时间), "", Trim(rsTemp!时间))
            
            Set rptItem = rptRecord.AddItem(IIf(IsNull(rsTemp!会话ID), "0", rsTemp!会话ID))
            rptItem.Caption = IIf(IsNull(rsTemp!会话ID), "0", rsTemp!会话ID)

            
            rsTemp.MoveNext
        Loop
        .Populate
    End With
    '统一调用显示文本
    Call FillText
End Sub

Public Sub FillText()
'功能:把消息的内容装入到RichText中

    Dim rsTemp As New ADODB.Recordset
    
    If rpcMain.FocusedRow Is Nothing Then
        '保留原有键值
        rtfContent.Text = ""
        rtfContent.BackColor = RGB(255, 255, 255)
        'Exit Sub
    ElseIf rpcMain.FocusedRow.Record Is Nothing Then
        rtfContent.Text = ""
        rtfContent.BackColor = RGB(255, 255, 255)
        'Exit Sub
    Else
        mstrKey = rpcMain.FocusedRow.Record(0).Caption
    End If
    
    If mstrKey <> "" Then
        rsTemp.CursorLocation = adUseClient
        'Get_Zlmessage(Id_In)
        Set rsTemp = zlDatabase.OpenCursor(Me.Caption, "zltools", "b_ComFunc.Get_Zlmessage", Val(mstrKey))
        
        rtfContent.BackColor = IIf(IsNull(rsTemp("背景色")), RGB(255, 255, 255), rsTemp("背景色"))
        rtfContent.TextRTF = IIf(IsNull(rsTemp("内容")), "", rsTemp("内容"))
    Else
        rtfContent.Text = ""
        rtfContent.BackColor = RGB(255, 255, 255)
    End If
End Sub



Private Sub DeleteMessage()
'功能：删除过时的消息
    '删除若干天前的消息 ,天数从系统参数表中取得
    Call zlDatabase.OpenCursor(Me.Caption, "zltools", "b_ComFunc.Delete_Zlmessage")
End Sub

Private Sub tplCon_ItemClick(ByVal item As XtremeSuiteControls.ITaskPanelGroupItem)
    mlngIndex = item.Index
    Call FillList
End Sub

Private Sub init_rpcMain()
    With rpcMain
        .Columns.DeleteAll
        .Columns.Add 0, "ID", 100, True
        .Columns.Add 1, "类型", 100, True
        .Columns.Add 2, " ", 18, False
        .Columns.Add 3, "主题", 1200, True
        .Columns.Add 4, "发件人", 500, True
        .Columns.Add 5, "收件人", 800, True
        .Columns.Add 6, "时间", 900, True
        .Columns.Add 7, "会话ID", 800, True
        .ShowGroupBox = True
        '.ShowItemsInGroups = True
        
        .AllowColumnRemove = False
        .Columns(0).Visible = False
        .Columns(1).Visible = False
        .Columns(7).Visible = False
        Set .Icons = imgRptIcon.Icons
        .Columns(2).Icon = 9
        .Columns(3).Icon = 8
        With .PaintManager
                
            .ColumnStyle = xtpColumnShaded
            .MaxPreviewLines = 1
            .GroupForeColor = &HC00000
            '.GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridNoLines
            '.VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "没有可显示的信息..."
        End With
    End With
End Sub

Private Sub InitCommandBar()
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl

    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False
    
    Set cbsMain.Icons = imgPublic.Icons
    
    '菜单定义:包括公共部份
    '    请对xtpControlPopup类型的命令ID重新赋值
    '-----------------------------------------------------
    cbsMain.ActiveMenuBar.Title = "菜单"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    objMenu.ID = conMenu_FilePopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)…")
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "打印预览(&V)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "打印(&P)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Excel, "输出到&Excel…")
        Set objControl = .Add(xtpControlButton, conMenu_File_SaveAs, "另存为(&A)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"): objControl.BeginGroup = True
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    objMenu.ID = conMenu_EditPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Add, "增加(&A)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, "打开(&O)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除(&D)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "还原(&S)"): objControl.BeginGroup = True
        
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Reply, "答复(&R)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_AllReply, "全部答复(&L)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Transmit, "转发(&W)")
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    objMenu.ID = conMenu_ViewPopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_ToolBar, "工具栏(&T)")
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False
            .Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False
            .Add xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False
        End With
        Set objControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
        Set objControl = .Add(xtpControlButton, conMenu_View_PreviewWindow, "预览窗格(&P)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_View_ShowAll, "显示已读(&E)")
        
        Set objControl = .Add(xtpControlButton, conMenu_View_Login, "登录时有未读邮件提醒(&W)"): objControl.BeginGroup = True
        
        Set objControl = .Add(xtpControlButton, conMenu_View_Find, "查找相关信息(&F)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)")
    End With
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    objMenu.ID = conMenu_HelpPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Help_Web, "&WEB上的" & gstr支持商简名)
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Help_Web_Home, gstr支持商简名 & "主页(&H)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Forum, gstr支持商简名 & "论坛(&F)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        End With
        Set objControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)…"): objControl.BeginGroup = True
    End With

    '工具栏定义:包括公共部份
    '-----------------------------------------------------
    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagHideWrap
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "打印")
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "预览")
        
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Add, "增加"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, "打开")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "还原")
        
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Reply, "答复"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Transmit, "转发")
        
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
    End With
    For Each objControl In objBar.Controls
        objControl.Style = xtpButtonIconAndCaption
    Next
    
    '命令的快键绑定:公共部份主界面已处理
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyP, conMenu_File_Print '打印
        
        .Add FCONTROL, vbKeyA, conMenu_Edit_Add '新增
        .Add FCONTROL, vbKeyO, conMenu_Edit_Modify '打开
        .Add 0, vbKeyDelete, conMenu_Edit_Delete '删除
        
        .Add FCONTROL, vbKeyAdd, conMenu_View_Expend_AllExpend '展开所有组
        .Add FCONTROL, vbKeySubtract, conMenu_View_Expend_AllCollapse '折叠所有组
        
        .Add FCONTROL, vbKeyF, conMenu_View_Find '查找
        
        .Add 0, vbKeyF5, conMenu_View_Refresh '刷新
        
        .Add 0, vbKeyF1, conMenu_Help_Help '帮助
    End With
    
    '设置一些公共的不常用命令
    With cbsMain.Options
        .AddHiddenCommand conMenu_File_PrintSet '打印设置
        .AddHiddenCommand conMenu_File_Excel '输出到Excel
    End With
End Sub

Private Sub Edit_Modify()
    Dim lngID As Long, lng会话ID As Long
    
    If rpcMain.FocusedRow Is Nothing Then Exit Sub
    lngID = Val(rpcMain.Rows(rpcMain.FocusedRow.Index).Record(0).Caption)
    lng会话ID = Val(rpcMain.Rows(rpcMain.FocusedRow.Index).Record(1).Caption)
    frmMessageEdit.OpenWindow lngID, "", lng会话ID
    Call FillList
End Sub
