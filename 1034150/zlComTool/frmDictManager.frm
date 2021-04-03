VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Begin VB.Form frmDictManager 
   BackColor       =   &H8000000C&
   Caption         =   "字典管理工具"
   ClientHeight    =   5655
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   8415
   Icon            =   "frmDictManager.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5655
   ScaleWidth      =   8415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picSplit2 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3225
      Left            =   7470
      MousePointer    =   9  'Size W E
      ScaleHeight     =   3225
      ScaleWidth      =   45
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1725
      Visible         =   0   'False
      Width           =   45
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   6300
      Top             =   1620
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDictManager.frx":6852
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDictManager.frx":D0B4
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDictManager.frx":13916
            Key             =   "Group"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDictManager.frx":1A178
            Key             =   "GroupOpen"
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.TaskPanel tplTable 
      Height          =   4485
      Left            =   165
      TabIndex        =   5
      Top             =   720
      Width           =   2325
      _Version        =   589884
      _ExtentX        =   4101
      _ExtentY        =   7911
      _StockProps     =   64
      VisualTheme     =   12
      SelectItemOnFocus=   -1  'True
      ItemLayout      =   2
      HotTrackStyle   =   3
   End
   Begin ComCtl3.CoolBar clbOnly 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   1058
      BandCount       =   2
      BandBorders     =   0   'False
      _CBWidth        =   8415
      _CBHeight       =   600
      _Version        =   "6.7.9782"
      Child1          =   "tlbMain"
      MinHeight1      =   540
      Width1          =   4995
      FixedBackground1=   0   'False
      Key1            =   "Comm"
      NewRow1         =   0   'False
      Child2          =   "cmbSys"
      MinWidth2       =   1500
      MinHeight2      =   300
      Width2          =   1500
      NewRow2         =   0   'False
      Begin VB.ComboBox cmbSys 
         Height          =   300
         Left            =   5160
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   150
         Width           =   3165
      End
      Begin XtremeCommandBars.ImageManager imgPublic 
         Left            =   705
         Top             =   135
         _Version        =   589884
         _ExtentX        =   635
         _ExtentY        =   635
         _StockProps     =   0
         Icons           =   "frmDictManager.frx":209DA
      End
      Begin XtremeCommandBars.CommandBars cbsMain 
         Left            =   210
         Top             =   120
         _Version        =   589884
         _ExtentX        =   635
         _ExtentY        =   635
         _StockProps     =   0
      End
   End
   Begin VB.PictureBox picSplit 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3225
      Left            =   3870
      MousePointer    =   9  'Size W E
      ScaleHeight     =   3225
      ScaleWidth      =   45
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1530
      Width           =   45
   End
   Begin MSComctlLib.ListView lvwMain 
      Bindings        =   "frmDictManager.frx":315F0
      Height          =   2235
      Left            =   3330
      TabIndex        =   0
      Top             =   1245
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   3942
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "img32"
      SmallIcons      =   "img16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      NumItems        =   0
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   5295
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   635
      SimpleText      =   $"frmDictManager.frx":31604
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmDictManager.frx":3164B
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9763
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
   Begin MSComctlLib.ImageList img32 
      Left            =   6285
      Top             =   2250
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDictManager.frx":31EDF
            Key             =   "Item"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwMain 
      Height          =   1620
      Left            =   5925
      TabIndex        =   9
      Top             =   3255
      Visible         =   0   'False
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   2858
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "img16"
      Appearance      =   0
   End
   Begin VB.Label lblTable 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "病案"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   3810
      TabIndex        =   6
      Top             =   810
      Width           =   3570
   End
   Begin XtremeSuiteControls.ShortcutCaption picTable 
      Height          =   495
      Left            =   3600
      TabIndex        =   7
      Top             =   660
      Width           =   4005
      _Version        =   589884
      _ExtentX        =   7064
      _ExtentY        =   873
      _StockProps     =   6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GradientColorLight=   -2147483633
      GradientColorDark=   -2147483632
   End
End
Attribute VB_Name = "frmDictManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim sngStartX As Single  '移动前鼠标的位置
Dim mblnItem As Boolean  '为真表示单击到ListView某一项上
Dim mintListIndex As Integer 'cmbSys的前一个列表索引
Dim mintColumn As Integer    '前一个ListView列头序号

Dim mblnFail As Boolean
Dim mcolSys As New Collection  '保存着所有系统的所有者
Dim mstrOwner As String        '当前选中系统的所有者

Dim mblnModify As Boolean
Dim mblnModifyGroup As Boolean
Dim mLastNode As Node
Dim bln有缺省值 As Boolean

Public Sub 字典管理()
    Dim rsSys As New ADODB.Recordset
    
    If mcolSys.Count > 0 Then
        '已经完成了初始化，可能是第二次显示
        frmDictManager.Show , gfrmMain
        Exit Sub
    End If
    
    Load frmDictManager
    '完成初始化
    gstrSQL = "select A.编号,A.名称,A.所有者 " & _
               " from zlSystems A,zlBasecode B,all_tables C " & _
               " Where A.编号 = B.系统 And upper(B.表名) = C.table_name  and A.所有者=C.OWNER " & _
               " group by A.编号,A.名称,A.所有者 " & _
               " Having Count(A.编号) > 0"
    Call zlDatabase.OpenRecordset(rsSys, gstrSQL, Me.Caption)
    
    mblnFail = False
    If rsSys.EOF Then
        MsgBox "你没有可以管理的数据字典。", vbInformation, gstrSysName
        Unload frmDictManager
        Exit Sub
    End If
    Do While Not rsSys.EOF
        cmbSys.AddItem rsSys("名称") & "（" & rsSys("编号") & "）"
        cmbSys.ItemData(Me.cmbSys.NewIndex) = rsSys("编号")
        mcolSys.Add CStr(rsSys("所有者")), "C" & rsSys("编号")
        rsSys.MoveNext
    Loop
    mintListIndex = -1
    If cmbSys.ListCount > 0 Then cmbSys.ListIndex = 0
    If cmbSys.ListCount = 1 Then cmbSys.Enabled = False
    
    If mblnFail = True Then
        Unload frmDictManager
        Exit Sub
    End If
    
    frmDictManager.Show , gfrmMain
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    On Error Resume Next
    Dim objControl As CommandBarControl
    Dim i As Integer
   
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
    Case conMenu_Edit_Add
        '新增
        Call Edit_Add
    Case conMenu_Edit_Modify
        '修改
        Call Edit_Modify
    Case conMenu_Edit_Delete
        '删除
        Call Edit_Delete
    Case conMenu_Edit_AddGroup
        '增加分类
        Call Edit_AddGroup
    Case conMenu_Edit_ModifyGroup
        '修改分类
        Call Edit_ModifyGroup
    Case conMenu_Edit_DeleteGroup
        '删除分类
        Call Edit_DeleteGroup
    Case conMenu_Edit_setDefault
        '设为默认值
        Call Edit_Default
    Case conMenu_View_ToolBar_Text
        '按钮文字
        For i = 2 To cbsMain.Count
            For Each objControl In Me.cbsMain(i).Controls
                objControl.Style = IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
        Next
        
        Me.cbsMain.RecalcLayout
        'Call Form_Resize
    Case conMenu_View_ToolBar_Size
        '工具栏大图标
        Me.cbsMain.Options.LargeIcons = Not Me.cbsMain.Options.LargeIcons
        
        If Me.cbsMain.Options.LargeIcons = True Then
            clbOnly.Bands.item(1).MinHeight = 520
        Else
            clbOnly.Bands.item(1).MinHeight = 425
        End If
        Me.cbsMain.RecalcLayout
        Call Form_Resize
    Case conMenu_View_BigIcon
        '列表大图标
        lvwMain.View = lvwIcon
    Case conMenu_View_MiniIcon
        lvwMain.View = lvwSmallIcon
    Case conMenu_View_List
        lvwMain.View = lvwList
    Case conMenu_view_Report
        lvwMain.View = lvwReport
    Case conMenu_View_XP
        tplTable.VisualTheme = xtpTaskPanelThemeShortcutBarOffice2003
        Me.Refresh
    Case conMenu_View_OutLook
        tplTable.VisualTheme = xtpTaskPanelThemeListView
        Me.Refresh
    Case conMenu_View_StatusBar
        '状态栏
        Me.stbThis.Visible = Not Me.stbThis.Visible
        Me.cbsMain.RecalcLayout
        Call Form_Resize
    Case conMenu_View_Refresh
        '刷新
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

    Dim blnPrint As Boolean, blnNew As Boolean, blnDelete As Boolean, blnSetDefault As Boolean
    Dim blnNewGroup As Boolean, blnDeleteGroup As Boolean
    Dim objControl As CommandBarControl, i As Integer
    blnPrint = lvwMain.ListItems.Count > 0
    
    If Mid(lblTable.Tag, 1, 1) = "W" Then
    
        blnNew = (InStr(picTable.Tag, "'INSERT") > 0)
        If lvwMain.ListItems.Count = 0 Then
            mblnModify = False
            blnDelete = False
        Else
            mblnModify = (InStr(picTable.Tag, "'DELETE") > 0)
            blnDelete = (InStr(picTable.Tag, "'DELETE") > 0)
        End If
        
        If tvwMain.Visible Then
            If tvwMain.Nodes.Count <= 1 Then
                blnNewGroup = (InStr(picTable.Tag, "'INSERT") > 0)
                blnDeleteGroup = False
                mblnModifyGroup = False
            Else
               blnNewGroup = (InStr(picTable.Tag, "'INSERT") > 0)
                blnDeleteGroup = (InStr(picTable.Tag, "'DELETE") > 0)
                mblnModifyGroup = (InStr(picTable.Tag, "'UPDATE") > 0)
                
                If Not tvwMain.SelectedItem Is Nothing Then
                    If tvwMain.SelectedItem.Key = "Root" Then
                        blnDeleteGroup = False
                        mblnModifyGroup = False
                    End If
                End If
            End If
        Else
            blnNewGroup = False
            blnDeleteGroup = False
            mblnModifyGroup = False
        End If
        
    Else
        blnNew = False
        mblnModify = False
        blnDelete = False
        
        blnNewGroup = False
        mblnModifyGroup = False
        blnDeleteGroup = False
    End If
    
    blnSetDefault = mblnModify And Not lvwMain.SelectedItem Is Nothing And bln有缺省值
    
    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        Control.Enabled = blnPrint
    Case conMenu_Edit_Add
        Control.Enabled = blnNew
    Case conMenu_Edit_Modify
        Control.Enabled = mblnModify
    Case conMenu_Edit_Delete
        Control.Enabled = blnDelete
    Case conMenu_Edit_setDefault
        Control.Enabled = blnSetDefault
    Case conMenu_View_BigIcon
        Control.Checked = lvwMain.View = lvwIcon
    Case conMenu_View_MiniIcon
        Control.Checked = lvwMain.View = lvwSmallIcon
    Case conMenu_View_List
        Control.Checked = lvwMain.View = lvwList
    Case conMenu_view_Report
        Control.Checked = lvwMain.View = lvwReport
    Case conMenu_Edit_AddGroup
        Control.Enabled = blnNewGroup
    Case conMenu_Edit_ModifyGroup
        Control.Enabled = mblnModifyGroup
    Case conMenu_Edit_DeleteGroup
        Control.Enabled = blnDeleteGroup
    Case conMenu_View_ToolBar_Size
        Control.Checked = Me.cbsMain.Options.LargeIcons
    Case conMenu_View_ToolBar_Text
    
        For i = 2 To cbsMain.Count
            For Each objControl In Me.cbsMain(i).Controls
                Control.Checked = objControl.Style = xtpButtonIconAndCaption
                Exit For
            Next
            Exit For
        Next
    End Select
    

    
End Sub

Private Sub clbOnly_HeightChanged(ByVal NewHeight As Single)
    Form_Resize
End Sub


Private Sub cmbSys_Click()
    If mintListIndex = cmbSys.ListIndex Then Exit Sub
    
    mstrOwner = mcolSys("C" & cmbSys.ItemData(cmbSys.ListIndex))
    If FillTable = False And mintListIndex >= 0 Then
        cmbSys.ListIndex = mintListIndex
        mstrOwner = mcolSys("C" & cmbSys.ItemData(cmbSys.ListIndex))
        Exit Sub
    End If
    
    mintListIndex = cmbSys.ListIndex
End Sub

Private Sub Form_Load()
    Dim intView As Integer
    
    intView = GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\界面设置\" & App.ProductName & "\" & Me.Name, "OutlookView", 1)
    If intView <> 0 And intView <> 1 Then
        intView = 1
    End If
    tplTable.Tag = intView
    RestoreWinState Me, App.ProductName
    
    Call InitCommandBar
    clbOnly.Bands.item(2).NewRow = True
    clbOnly.Bands.item(1).MinHeight = 520
    
    Set tplTable.Icons = imgPublic.Icons
    
    tplTable.VisualTheme = xtpTaskPanelThemeListViewOffice2003
    tplTable.Behaviour = xtpTaskPanelBehaviourToolbox
    tplTable.HotTrackStyle = xtpTaskPanelHighlightItem
    
    picTable.GradientColorDark = cbsMain.GetSpecialColor(XPCOLOR_TOOLBAR_FACE)
    picTable.GradientColorLight = cbsMain.GetSpecialColor(XPCOLOR_SPLITTER_FACE)
End Sub

Private Sub Form_Resize()
    Dim sngTop As Single, sngBottom As Single
    
    On Error Resume Next
    sngTop = IIf(clbOnly.Visible, clbOnly.Top + clbOnly.Height, 0)
    sngBottom = Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0)
    
    tplTable.Top = sngTop
    tplTable.Height = IIf(sngBottom - tplTable.Top > 0, sngBottom - tplTable.Top, 0)
    tplTable.Left = 0
    
    picSplit.Top = sngTop
    picSplit.Height = IIf(sngBottom - picSplit.Top > 0, sngBottom - picSplit.Top, 0)
    picSplit.Left = tplTable.Left + tplTable.Width
    
    picTable.Top = sngTop + 45
    picTable.Left = picSplit.Left + picSplit.Width
    If Me.ScaleWidth - picTable.Left > 0 Then picTable.Width = ScaleWidth - picTable.Left
    
    lblTable.Width = picTable.Width - 45
    lblTable.Top = picTable.Top + 125
    lblTable.Left = picTable.Left + 45
    lblTable.Height = picTable.Height - 45
    '-- 10152修改
    If tvwMain.Visible Then
        tvwMain.Left = picTable.Left
        tvwMain.Top = picTable.Top + picTable.Height + 45
        tvwMain.Height = IIf(sngBottom - tvwMain.Top > 0, sngBottom - tvwMain.Top, 0)
        
        picSplit2.Left = tvwMain.Left + tvwMain.Width
        picSplit2.Top = tvwMain.Top
        picSplit2.Height = tvwMain.Height
        
        lvwMain.Left = picSplit2.Left + picSplit2.Width
        lvwMain.Top = tvwMain.Top
        lvwMain.Width = picTable.Width - tvwMain.Width - picSplit2.Width - 45
        lvwMain.Height = tvwMain.Height
    Else
        lvwMain.Left = picTable.Left
        lvwMain.Top = picTable.Top + picTable.Height + 45
        lvwMain.Width = picTable.Width
        lvwMain.Height = IIf(sngBottom - lvwMain.Top > 0, sngBottom - lvwMain.Top, 0)
    End If
    
    Me.Refresh

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mcolSys = Nothing
    
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\界面设置\" & App.ProductName & "\" & Me.Name, "OutlookView", tplTable.Tag)
    SaveWinState Me, App.ProductName
End Sub

Private Sub lvwMain_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If mintColumn = ColumnHeader.Index - 1 Then '仍是刚才那列
        lvwMain.SortOrder = IIf(lvwMain.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        mintColumn = ColumnHeader.Index - 1
        lvwMain.SortKey = mintColumn
        lvwMain.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvwMain_DblClick()
    If mblnItem = True And mblnModify Then Call Edit_Modify
End Sub

Private Sub lvwMain_GotFocus()
    Dim i As Integer
    With lvwMain
        For i = 0 To 3
'            mnuViewIcon(i).Checked = False
        Next
'        mnuViewIcon(.View).Checked = True
    End With

End Sub

Private Sub lvwMain_ItemClick(ByVal item As MSComctlLib.ListItem)
    mblnItem = True
End Sub

Private Sub lvwMain_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If mblnModify Then Call Edit_Modify
    End If
End Sub

Private Sub lvwMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnItem = False
End Sub

Private Sub lvwMain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Dim Control As CommandBarControl, objControl As CommandBarControl
        Dim Popup As CommandBar
        
        Set Popup = cbsMain.Add("Popup", xtpBarPopup)
        
        With Popup.Controls
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Add, "增加(&A)")
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改(&M)")
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除(&D)")
            Set objControl = .Add(xtpControlButton, conMenu_View_BigIcon, "大图标(&G)"): objControl.BeginGroup = True
            .Add xtpControlButton, conMenu_View_MiniIcon, "小图标(&M)"
            .Add xtpControlButton, conMenu_View_List, "列表(&L)"
            .Add xtpControlButton, conMenu_view_Report, "详细资料(&D)"
            
            Set objControl = .Add(xtpControlButton, conMenu_Edit_setDefault, "设为默认项(&F)"): objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)"): objControl.BeginGroup = True
            
        End With
            
        Popup.ShowPopup
    End If
End Sub

Private Sub Edit_Default()
    
    On Error GoTo errHandle
    
    gstrSQL = "Update " & mstrOwner & "." & Mid(lblTable.Tag, 2) & _
        " Set 缺省标志=Decode(编码,'" & Mid(lvwMain.SelectedItem.Key, 2) & "',1,0)"
    
    Call zlDatabase.RunProcedure(Me.Caption, cmbSys.Tag, "ZL_字典管理_execute", gstrSQL)
    Call FillList
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Edit_Delete()
    On Error GoTo errHandle
    Dim lngSystem As Long
    Dim intIndex As Integer
    
    If MsgBox("你确认要删除《" & Mid(lblTable.Tag, 2) & "》中名称为“" & lvwMain.SelectedItem.Text & "”的项目吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
        gstrSQL = "delete from " & mstrOwner & "." & Mid(lblTable.Tag, 2) & _
            " where 编码='" & Mid(lvwMain.SelectedItem.Key, 2) & "'"
        '用过程进行封装
        lngSystem = cmbSys.ItemData(cmbSys.ListIndex) \ 100
        'gstrSQL = "ZL_字典管理_execute('" & Replace(gstrSQL, "'", "''") & "')"
        Call zlDatabase.RunProcedure(Me.Caption, cmbSys.Tag, "ZL_字典管理_execute", gstrSQL)
        
        With lvwMain
            '保存当前项目的索引
            intIndex = .SelectedItem.Index
            .ListItems.Remove .SelectedItem.Key
            If .ListItems.Count > 0 Then
                intIndex = IIf(.ListItems.Count > intIndex, intIndex, .ListItems.Count)
                .ListItems(intIndex).Selected = True
                .ListItems(intIndex).EnsureVisible
            End If
            Call SetMenu
        End With
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub picsplit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error Resume Next
    If Button = 1 Then
        If tplTable.Width + X < 300 Then Exit Sub
        If tvwMain.Visible Then
            If tvwMain.Width - X < 220 Then Exit Sub
        End If
        
        picSplit.Left = picSplit.Left + X
        tplTable.Width = tplTable.Width + X
            
        picTable.Left = picTable.Left + X
        lblTable.Left = picTable.Left + 45
        picTable.Width = picTable.Width - X
        lblTable.Width = picTable.Width - 45
        
        '-- 10152问题
        If tvwMain.Visible Then
            tvwMain.Left = picTable.Left
            tvwMain.Width = tvwMain.Width - X
        Else
            lvwMain.Left = picTable.Left
            lvwMain.Width = picTable.Width
        End If
            
    End If
End Sub


'Private Sub tlbMain_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
'    Dim i As Integer
'    For i = 0 To 3
'        mnuViewIcon(i).Checked = False
'    Next
'    mnuViewIcon(ButtonMenu.Index - 1).Checked = True
'    If Me.ActiveControl Is outTable_S Then
'        outTable_S.View = ButtonMenu.Index - 1
'    Else
'        lvwMain.View = ButtonMenu.Index - 1
'    End If
'End Sub

Private Sub subPrint(bytMode As Byte)
'功能:进行打印,预览和输出到EXCEL
'参数:bytMode=1 打印;2 预览;3 输出到EXCEL
    Dim objPrint As New zlPrintLvw
    objPrint.Title.Text = Mid(lblTable.Tag, 2)
    Set objPrint.Body.objData = lvwMain
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

Private Function FillTable() As Boolean
'功能:装入所有编辑表到outTable_S
    Dim rsTemp As New ADODB.Recordset
    Dim item As OutItem
    Dim strOwner As String, strGroup As String
    
    If cmbSys.ListIndex = -1 Then Exit Function
    
    strOwner = UCase(mcolSys("C" & cmbSys.ItemData(Me.cmbSys.ListIndex)))
    cmbSys.Tag = strOwner
    
    gstrSQL = "select rownum as 编号,A.表名,A.固定,A.说明,A.分类,B.privilege 权限 " & _
            " from zlBasecode A," & _
            "    (select table_name,privilege from all_tab_privs where TABLE_SCHEMA=[1] and privilege in('SELECT','INSERT','DELETE','UPDATE')" & _
            "     union select table_name,'SELECT' from all_tables where owner=[1] and (owner=user or exists(select 1 from session_roles where ROLE='DBA') OR exists(select 1 from USER_SYS_PRIVS where PRIVILEGE='SELECT ANY TABLE'))" & _
            "     union select table_name,'INSERT' from all_tables where owner=[1] and (owner=user or exists(select 1 from session_roles where ROLE='DBA'))" & _
            "     union select table_name,'DELETE' from all_tables where owner=[1] and (owner=user or exists(select 1 from session_roles where ROLE='DBA'))" & _
            "     union select table_name,'UPDATE' from all_tables where owner=[1] and (owner=user or exists(select 1 from session_roles where ROLE='DBA'))) B " & _
            " Where a.表名 = b.table_name and A.系统=[2] order by A.分类"
    
    rsTemp.CursorLocation = adUseClient
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strOwner, Val(cmbSys.ItemData(cmbSys.ListIndex)))
    
    rsTemp.Filter = "权限='SELECT'"
    If rsTemp.RecordCount = 0 Then
        MsgBox "你没有能设置的编码表,不能运行本程序。", vbExclamation, gstrSysName
        mblnFail = True
        Exit Function
    End If
    
    tplTable.LockRedraw = Not tplTable.LockRedraw
    tplTable.Visible = Not tplTable.Visible
    tplTable.Groups.Clear
    strGroup = ""
    
    Dim lngGroupID As Long
    Dim tplGroup As TaskPanelGroup, tplItem As TaskPanelGroupItem, tplItemOne As TaskPanelGroupItem
    Do Until rsTemp.EOF
        If rsTemp("分类") <> strGroup Then
            strGroup = rsTemp("分类")
            lngGroupID = lngGroupID + 1
            Set tplGroup = tplTable.Groups.Add(lngGroupID, rsTemp("分类"))
            If lngGroupID = 1 Then
                tplGroup.Expanded = True
            Else
                tplGroup.Expanded = False
            End If
        End If
        
        If IIf(IsNull(rsTemp("固定")), 0, rsTemp("固定")) = 0 Then
            Set tplItem = tplGroup.Items.Add(rsTemp("编号"), rsTemp("表名"), xtpTaskItemTypeLink, 112)

        Else
            Set tplItem = tplGroup.Items.Add(rsTemp("编号"), rsTemp("表名"), xtpTaskItemTypeLink, 113)
        End If
        If tplItemOne Is Nothing Then
            Set tplItemOne = tplItem
        End If
        rsTemp.MoveNext
    Loop
    
    For Each tplGroup In tplTable.Groups
        For Each tplItem In tplGroup.Items
            rsTemp.Filter = "表名='" & tplItem.Caption & "'"
            If rsTemp.RecordCount > 0 Then
                tplItem.Tag = IIf(IsNull(rsTemp("说明")), "", rsTemp("说明"))
            End If
            Do Until rsTemp.EOF
                tplItem.Tag = tplItem.Tag & "'" & rsTemp("权限")
                rsTemp.MoveNext
            Loop
        Next
    Next
    
    
    tplTable.SetMargins 2, 2, 2, 2, 1
    tplTable.Visible = Not tplTable.Visible
    tplTable.LockRedraw = Not tplTable.LockRedraw
    
    lblTable.Tag = ""
    
    Call tplTable_ItemClick(tplItemOne)
    tplItemOne.Selected = True
    FillTable = True
End Function

Public Sub FillList()
'功能:装入对应编码表的项目到lvwMain
    Dim strTable As String
    Dim rsTemp As New ADODB.Recordset
    Dim strKey As String
    Dim fld As Field
    Dim lst As ListItem
    
    strTable = Mid(lblTable.Tag, 2)
    
    If Not lvwMain.SelectedItem Is Nothing Then
        '保留原有键值
        strKey = lvwMain.SelectedItem.Key
    End If
    
    If strTable = "" Then
        lvwMain.ListItems.Clear
        lvwMain.ColumnHeaders.Clear
        lvwMain.ColumnHeaders.Add , , "请选择数据字典", 2000
        tvwMain.Nodes.Clear
        Call SetMenu
        Exit Sub
    End If
    
    '-- 10152修改 检查是否有末级,有则将上级显示到TreeList中
    gstrSQL = "Select table_name From all_col_comments Where owner = '" & mstrOwner & "' And table_name='" & UCase(strTable) & "' And column_name='上级'"
    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
    If Not rsTemp.EOF Then
        tvwMain.Visible = True
        tvwMain.Tag = strTable
        picSplit2.Visible = True
        Call FillTree(mstrOwner & "." & strTable)
    Else
        tvwMain.Tag = ""
        tvwMain.Visible = False
        picSplit2.Visible = False
    End If
    Call Form_Resize
    
    If Not mLastNode Is Nothing And tvwMain.Tag <> "" Then
        Call ShowList(strTable, Mid(mLastNode.Key, 2))
    Else
        Call ShowList(strTable)
    End If
    ' strTable
    '---------
End Sub

Public Sub ShowList(ByVal strTable As String, Optional ByVal strTreeNodeKey As String)
    
    Dim rsTemp As New ADODB.Recordset
    Dim strKey As String
    Dim fld As Field
    Dim lst As ListItem
    Dim strWhere As String
    rsTemp.CursorLocation = adUseClient

    If tvwMain.Tag <> "" Then
        strWhere = " Where 末级=1"
        If strTreeNodeKey = "" Or strTreeNodeKey = "oot" Then
            strWhere = strWhere & " And Nvl(上级,0)=0"
        Else
            strWhere = strWhere & " And 上级='" & strTreeNodeKey & "'"
        End If
    End If
    gstrSQL = "select * from " & mstrOwner & "." & strTable & strWhere
    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
    
    mintColumn = 0
    
    LockWindowUpdate lvwMain.hWnd
    lvwMain.ColumnHeaders.Clear
    lvwMain.ColumnHeaders.Add , "名称", "名称"
    bln有缺省值 = False
    For Each fld In rsTemp.Fields
        If InStr(",名称,上级,末级,", "," & fld.Name & ",") <= 0 Then lvwMain.ColumnHeaders.Add , fld.Name, fld.Name
       
        If fld.Name = "缺省标志" Then
            bln有缺省值 = True
        End If
    Next
    lvwMain.ListItems.Clear
    Do Until rsTemp.EOF
        If tvwMain.Tag <> "" Then
            Dim strIcon As String
            strIcon = IIf(zlCommFun.NVL(rsTemp("末级"), 0) = 1, "Item", "Group")
            
            Set lst = lvwMain.ListItems.Add(, "C" & rsTemp("编码"), IIf(IsNull(rsTemp("名称")), "", rsTemp("名称")), strIcon, strIcon)
        Else
            Set lst = lvwMain.ListItems.Add(, "C" & rsTemp("编码"), IIf(IsNull(rsTemp("名称")), "", rsTemp("名称")), "Item", "Item")
        End If
        For Each fld In rsTemp.Fields
            '-- 10152修改 加入末级的处理
            If fld.Name = "缺省标志" Or fld.Name Like "是否*" Then
                lst.SubItems(lvwMain.ColumnHeaders(fld.Name).Index - 1) = IIf(fld.Value = 1, "√", "")
            Else
                If InStr(",名称,上级,末级,", "," & fld.Name & ",") <= 0 Then
                    lst.SubItems(lvwMain.ColumnHeaders(fld.Name).Index - 1) = IIf(IsNull(fld.Value), "", fld.Value)
                End If
            End If
            
        Next
        rsTemp.MoveNext
    Loop
    '使列宽自适应
    Dim i As Integer
    For i = 0 To lvwMain.ColumnHeaders.Count - 1
        SendMessage lvwMain.hWnd, LVM_SETCOLUMNWIDTH, i, LVSCW_AUTOSIZE_USEHEADER
        If lvwMain.ColumnHeaders(i + 1).Width < 600 Then lvwMain.ColumnHeaders(i + 1).Width = 600
    Next
    LockWindowUpdate 0
    
    If lvwMain.ListItems.Count > 0 Then
        Dim item As ListItem
        On Error Resume Next
        Set item = lvwMain.ListItems(strKey)
        If Err <> 0 Then
            Set item = lvwMain.ListItems(1)
            item.Selected = True
            item.EnsureVisible
        Else
            Err.Clear
            item.Selected = True
            item.EnsureVisible
        End If
    End If
    Call SetMenu
End Sub

Public Sub SetMenu()
    Dim i As Integer
    i = InStr(picTable.Tag, "'")
    lblTable.Caption = " " & picSplit.Tag & IIf(i > 0, "——" & Mid(picTable.Tag, 1, i - 1), "")
    stbThis.Panels(2) = "该字典共有" & lvwMain.ListItems.Count & "条编码。"
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
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"): objControl.BeginGroup = True
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    objMenu.ID = conMenu_EditPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Add, "增加(&A)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改(&M)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除(&D)")
        
        Set objControl = .Add(xtpControlButton, conMenu_Edit_AddGroup, "增加分类(&I)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ModifyGroup, "修改分类(&U)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_DeleteGroup, "删除分类(&E)")
        
        Set objControl = .Add(xtpControlButton, conMenu_Edit_setDefault, "设为缺省项(&F)"): objControl.BeginGroup = True
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    objMenu.ID = conMenu_ViewPopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_ToolBar, "工具栏(&T)")
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False
            .Add xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False
        End With
        
        
        'Set objControl = .Add(xtpControlButton, conMenu_View_XP, "OutLook 2003 样式")
        'objControl.BeginGroup = True
        'Set objControl = .Add(xtpControlButton, conMenu_View_OutLook, "OutLook 2000 样式")
        
        Set objControl = .Add(xtpControlButton, conMenu_View_BigIcon, "大图标(&G)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_View_MiniIcon, "小图标(&M)")
        Set objControl = .Add(xtpControlButton, conMenu_View_List, "列表(&L)")
        Set objControl = .Add(xtpControlButton, conMenu_view_Report, "详细资料(&D)")
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)"): objControl.BeginGroup = True
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
        
        Set objControl = .Add(xtpControlButton, conMenu_Edit_AddGroup, "分类"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Add, "增加"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除")
        
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
        
        Dim ControlComboSys As CommandBarControlCustom
        Set ControlComboSys = .Add(xtpControlCustom, conTool_System, "系统")
        ControlComboSys.Handle = cmbSys.hWnd
        ControlComboSys.BeginGroup = True
        cmbSys.Width = 3000
    End With
    For Each objControl In objBar.Controls
        objControl.Style = xtpButtonIconAndCaption
    Next
    
    '命令的快键绑定:公共部份主界面已处理
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyP, conMenu_File_Print '打印
        .Add FCONTROL, vbKeyA, conMenu_Edit_Add '新增
        .Add FCONTROL, vbKeyO, conMenu_Edit_Modify '修改
        .Add 0, vbKeyDelete, conMenu_Edit_Delete '删除
        .Add 0, vbKeyF5, conMenu_View_Refresh '刷新
        .Add 0, vbKeyF1, conMenu_Help_Help '帮助
    End With
    
    '设置一些公共的不常用命令
    With cbsMain.Options
        .AddHiddenCommand conMenu_File_PrintSet '打印设置
        .AddHiddenCommand conMenu_File_Excel '输出到Excel
    End With
End Sub


Private Sub picSplit2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Button = 1 Then
        If tvwMain.Width + X < 220 Or lvwMain.Width - X < 200 Then
            Exit Sub
        End If
        tvwMain.Width = tvwMain.Width + X
        picSplit2.Left = picSplit2.Left + X
        lvwMain.Left = lvwMain.Left + X
        lvwMain.Width = lvwMain.Width - X
    End If

End Sub

Private Sub tplTable_GroupExpanding(ByVal Group As XtremeSuiteControls.ITaskPanelGroup, ByVal Expanding As Boolean, Cancel As Boolean)
    Dim tplItem As TaskPanelGroupItem
    If Expanding Then
        For Each tplItem In Group.Items
            If Trim(tplItem.Caption) = Trim(picSplit.Tag) Then
                tplItem.Selected = True
            Else
                tplItem.Selected = False
            End If
        Next
    End If
End Sub

Private Sub tplTable_ItemClick(ByVal item As XtremeSuiteControls.ITaskPanelGroupItem)
    Dim strIcon As String, i As Integer
    If item.IconIndex = 111 Then
        strIcon = "W"
    Else
        strIcon = "R"
    End If
    If lblTable.Tag = strIcon & item.Caption Then Exit Sub
    picTable.Tag = item.Tag
    picSplit.Tag = item.Caption
    
    lblTable.Tag = strIcon & item.Caption
    FillList
End Sub

Private Sub Edit_Modify()
    frmDictEdit.编辑编码 mcolSys("C" & cmbSys.ItemData(cmbSys.ListIndex)), Mid(lblTable.Tag, 2), Mid(lvwMain.SelectedItem.Key, 2), 1
End Sub

Private Sub FillTree(ByVal strTable As String)
    
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim nodTmp As Node
    strSQL = " Select * From " & strTable & " Where nvl(末级,0)=0 Start with Nvl(上级,0)=0 connect by prior 编码 =上级"
                 
    Set rsTmp = gcnOracle.Execute(strSQL, adOpenStatic, adLockReadOnly)
    With tvwMain
        .Nodes.Clear
        .Nodes.Add , , "Root", "全部", "Root", "Root"
        Do Until rsTmp.EOF
            If IsNull(rsTmp!上级) Then
                tvwMain.Nodes.Add "Root", tvwChild, "B" & rsTmp!编码, "[" & rsTmp!编码 & "]" & rsTmp!名称, "Group", "GroupOpen"
            Else
                If nodTmp Is Nothing Then
                    Set nodTmp = tvwMain.Nodes.Add("B" & rsTmp!上级, tvwChild, "B" & rsTmp!编码, "[" & rsTmp!编码 & "]" & rsTmp!名称, "Group", "Group")
                Else
                    tvwMain.Nodes.Add "B" & rsTmp!上级, tvwChild, "B" & rsTmp!编码, "[" & rsTmp!编码 & "]" & rsTmp!名称, "Group", "GroupOpen"
                End If
            End If
            rsTmp.MoveNext
        Loop
        .Nodes.item("Root").Expanded = True
        .Nodes.item("Root").Selected = True
        Call tvwMain_NodeClick(.Nodes.item("Root"))
    End With
    
    
End Sub

Private Sub tvwMain_DblClick()
    If Not tvwMain.SelectedItem Is Nothing And mblnModifyGroup Then Edit_ModifyGroup
End Sub

Private Sub tvwMain_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Not tvwMain.SelectedItem Is Nothing And mblnModifyGroup Then Edit_ModifyGroup
    End If
End Sub

Private Sub tvwMain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Dim Control As CommandBarControl, objControl As CommandBarControl
        Dim Popup As CommandBar
        
        Set Popup = cbsMain.Add("Popup", xtpBarPopup)
        
        With Popup.Controls
            Set objControl = .Add(xtpControlButton, conMenu_Edit_AddGroup, "增加分类(&I)")
            Set objControl = .Add(xtpControlButton, conMenu_Edit_ModifyGroup, "修改分类(&U)")
            Set objControl = .Add(xtpControlButton, conMenu_Edit_DeleteGroup, "删除分类(&E)")
            
            Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)"): objControl.BeginGroup = True
            
        End With
            
        Popup.ShowPopup
    End If
End Sub

Private Sub tvwMain_NodeClick(ByVal Node As MSComctlLib.Node)
    If Not mLastNode Is Nothing Then
        If mLastNode = Node Then Exit Sub
    End If
    Node.Expanded = True
    If tvwMain.Tag <> "" Then
        Call ShowList(tvwMain.Tag, Mid(Node.Key, 2))
    End If
    Set mLastNode = Node
End Sub

Public Sub frmRefresh()
    Set mLastNode = Nothing
    Call FillList
End Sub
Private Sub Edit_Add()
    '增加
    
    If tvwMain.Visible Then
        If Not tvwMain.SelectedItem Is Nothing Then
            Set mLastNode = tvwMain.SelectedItem
        End If
        If Not mLastNode Is Nothing Then
            frmDictEdit.编辑编码 mcolSys("C" & cmbSys.ItemData(cmbSys.ListIndex)), Mid(lblTable.Tag, 2), , 1, Mid(mLastNode.Key, 2)
        Else
            frmDictEdit.编辑编码 mcolSys("C" & cmbSys.ItemData(cmbSys.ListIndex)), Mid(lblTable.Tag, 2), , 1
        End If
    Else
        frmDictEdit.编辑编码 mcolSys("C" & cmbSys.ItemData(cmbSys.ListIndex)), Mid(lblTable.Tag, 2), , 1
    End If
End Sub

Private Sub Edit_AddGroup()
    '增加分类
    If Not tvwMain.SelectedItem Is Nothing Then
        Set mLastNode = tvwMain.SelectedItem
    End If
    If Not mLastNode Is Nothing Then
        frmDictEdit.编辑编码 mcolSys("C" & cmbSys.ItemData(cmbSys.ListIndex)), Mid(lblTable.Tag, 2), , 0, Mid(mLastNode.Key, 2)
    Else
        frmDictEdit.编辑编码 mcolSys("C" & cmbSys.ItemData(cmbSys.ListIndex)), Mid(lblTable.Tag, 2), , 0
    End If
End Sub

Private Sub Edit_ModifyGroup()
    If Not mLastNode Is Nothing Then
        frmDictEdit.编辑编码 mcolSys("C" & cmbSys.ItemData(cmbSys.ListIndex)), Mid(lblTable.Tag, 2), Mid(mLastNode.Key, 2), 0
    End If
End Sub

Private Sub Edit_DeleteGroup()
    On Error GoTo errHandle
    If Not tvwMain.SelectedItem Is Nothing Then
        Set mLastNode = tvwMain.SelectedItem
    End If
    If Not mLastNode Is Nothing Then
        If MsgBox("你确认要删除《" & Mid(lblTable.Tag, 2) & "》中名称为“" & mLastNode.Text & "”的分类以及它的所有下级项目吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
            '用过程进行封装
            gstrSQL = "Delete " & mstrOwner & "." & Mid(lblTable.Tag, 2) & _
                    " Where 编码 In (Select 编码 From " & mstrOwner & "." & Mid(lblTable.Tag, 2) & _
                    " Start With Nvl(上级, '0') = '" & Mid(mLastNode.Key, 2) & "'" & _
                    " Connect By Prior 编码 = 上级)"

            gstrSQL = "ZL_字典管理_execute('" & Replace(gstrSQL, "'", "''") & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            
            gstrSQL = "delete from " & mstrOwner & "." & Mid(lblTable.Tag, 2) & _
                " where 编码='" & Mid(mLastNode.Key, 2) & "'"
            '用过程进行封装
            gstrSQL = "ZL_字典管理_execute('" & Replace(gstrSQL, "'", "''") & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            Call frmRefresh
            Call SetMenu
            
        End If
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub
