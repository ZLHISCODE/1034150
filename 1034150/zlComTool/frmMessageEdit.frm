VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Begin VB.Form frmMessageEdit 
   AutoRedraw      =   -1  'True
   Caption         =   "消息"
   ClientHeight    =   6435
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   9510
   Icon            =   "frmMessageEdit.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6435
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Tag             =   "可变化的"
   Begin zlComTool.ColorPicker ColorForeColor 
      Height          =   2190
      Left            =   5670
      TabIndex        =   9
      Top             =   4365
      Width           =   2190
      _ExtentX        =   3863
      _ExtentY        =   3863
   End
   Begin zlComTool.ColorPicker ColorFillColor 
      Height          =   2190
      Left            =   4260
      TabIndex        =   8
      Top             =   1545
      Width           =   2190
      _ExtentX        =   3863
      _ExtentY        =   3863
   End
   Begin VB.TextBox txtSubject 
      Height          =   300
      Left            =   1800
      MaxLength       =   200
      TabIndex        =   3
      Top             =   2190
      Width           =   4665
   End
   Begin VB.CommandButton cmdReceive 
      Caption         =   "收件人(&R)"
      Height          =   315
      Left            =   300
      TabIndex        =   0
      Top             =   1800
      Width           =   1100
   End
   Begin VB.TextBox txtReceive 
      Height          =   300
      Left            =   1410
      Locked          =   -1  'True
      MaxLength       =   200
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1830
      Width           =   4665
   End
   Begin RichTextLib.RichTextBox rtfContent 
      Height          =   3585
      Left            =   240
      TabIndex        =   4
      Top             =   2550
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   6324
      _Version        =   393217
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      MaxLength       =   4000
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmMessageEdit.frx":6852
   End
   Begin MSComDlg.CommonDialog cdg 
      Left            =   7575
      Top             =   1275
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   3435
      Top             =   270
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessageEdit.frx":68EF
            Key             =   "FILLCOLOR"
            Object.Tag             =   "562"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessageEdit.frx":6A59
            Key             =   "LINECOLOR"
            Object.Tag             =   "563"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessageEdit.frx":6BB2
            Key             =   "FORECOLOR"
            Object.Tag             =   "564"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   480
      Top             =   165
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeCommandBars.ImageManager imgPublic 
      Bindings        =   "frmMessageEdit.frx":6CFF
      Left            =   1305
      Top             =   180
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmMessageEdit.frx":6D13
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "主题(&S)："
      Height          =   180
      Index           =   3
      Left            =   600
      TabIndex        =   2
      Top             =   2250
      Width           =   810
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "收件人："
      Height          =   180
      Index           =   2
      Left            =   165
      TabIndex        =   7
      Top             =   1890
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "发件时间："
      Height          =   180
      Index           =   1
      Left            =   3000
      TabIndex        =   6
      Top             =   1530
      Width           =   900
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "发件人："
      Height          =   180
      Index           =   0
      Left            =   165
      TabIndex        =   5
      Top             =   1530
      Width           =   720
   End
End
Attribute VB_Name = "frmMessageEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'保存当前所编编辑的消息ID
Public mstrID As String     'strID 消息ID。如果为空，表示是新消息
Public mstrByID As String   'strByID 要进行转发、答复的原始ID
Public mlng类型 As Long     '当前用户对应的消息类型
Public mlngMode As Long     '打开方式。1-答复；2-全部答复；3-转发

Dim mstr会话ID As String
Dim mblnSend As Boolean     '准备发送
Dim mblnDelete As Boolean   '是否已经删除

Private mrsUser As ADODB.Recordset '存放收件人姓名,用户名,人员性质的记录集

Dim mblnChange As Boolean
'保存用于查找的
Public mblnCase As Boolean
Public mblnBegin As Boolean
Public mstrFind As String
Dim mblnHigh As Boolean, mblnLow As Boolean  '重要性
Dim mblnSetState As Boolean '是否初始化调用


Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)

    On Error Resume Next
    Dim objControl As CommandBarControl
    Dim i As Integer
    Dim lngID As Long, lng会话ID As Long
    Dim lst As ListItem
    Dim objMenu As CommandBarPopup
     
    Select Case Control.ID
    Case conMenu_File_Send
        '发送消息
       
        If mblnSend = False Then Exit Sub
        
        If SaveMessage(True) = True Then
            '更新主界面
            '首先判断是否合适的显示位置
            With frmMessageManager
                On Error Resume Next
                If .mlngIndex = 1 Then
                    Unload Me
                    Exit Sub
                End If
                If .mlngIndex = 0 Or .mlngIndex = 3 Then
                    '删除草稿消息
                    Unload Me
                    Exit Sub
                End If
                '创建已发送消息
              End With
              Unload Me
              '出现图标
              Call frmMessageRead.UpdateNotify
        End If
    Case conMenu_File_Save
        '保存
        If SaveMessage(False) = True Then
            On Error Resume Next
            frmMessageManager.FillList
        End If
    Case conMenu_View_ToolBar_Text
            '按钮文字
            For i = 2 To cbsMain.Count
                For Each objControl In Me.cbsMain(i).Controls
                    objControl.Style = IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
                Next
            Next
            Me.cbsMain.RecalcLayout
            Call Form_Resize
    Case conMenu_File_SaveAs
        '另存
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
        
    Case conMenu_Edit_Cut
        '进行剪切
        If rtfContent.SelLength = 0 Then Exit Sub
        
        Clipboard.SetText rtfContent.SelRTF, vbCFRTF
        rtfContent.SelText = ""
    Case conMenu_Edit_Copy
        '复制文本
        If rtfContent.SelLength = 0 Then Exit Sub
        Clipboard.SetText rtfContent.SelRTF, vbCFRTF
    Case conMenu_Edit_plaster
        '粘贴文本
        If Clipboard.GetFormat(vbCFText) = True Then
            If Clipboard.GetText(vbCFRTF) <> "" Then
                rtfContent.SelRTF = Clipboard.GetText(vbCFRTF)
            Else
                rtfContent.SelRTF = Clipboard.GetText(vbCFText)
            End If
        End If
    Case conMenu_Edit_Clear
        '清除所选的文本
            rtfContent.SelText = ""
    Case conMenu_Edit_CheckAll
        '全部选择
        rtfContent.SelStart = 0
        rtfContent.SelLength = Len(rtfContent.Text)
    Case conMenu_View_Find
        '查找第一个
        Set frmMessageFind.frmMain = Me
        frmMessageFind.Show vbModal, Me
    Case conMenu_View_FindNext
        '继续查找
        Call FindText
    Case conMenu_Format_ForeColor
        '选择字体颜色
        Call ColorForeColor_pOK
    Case conMenu_Format_FillColor
        '选择背景颜色
        Call ColorFillColor_pOK
    Case conMenu_Format_Sig
        '对文本加上项目符号
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_FormatPopup)
        objMenu.CommandBar.FindControl(, conMenu_Format_Sig).Checked = Not objMenu.CommandBar.FindControl(, conMenu_Format_Sig).Checked
        rtfContent.SelBullet = objMenu.CommandBar.FindControl(, conMenu_Format_Sig).Checked
    Case conMenu_Format_Left
        '文本靠左
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_FormatPopup)
        objMenu.CommandBar.FindControl(, conMenu_Format_Left).Checked = True
        objMenu.CommandBar.FindControl(, conMenu_Format_Center).Checked = False
        objMenu.CommandBar.FindControl(, conMenu_Format_Right).Checked = False
        rtfContent.SelAlignment = 0
    Case conMenu_Format_Center
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_FormatPopup)
        objMenu.CommandBar.FindControl(, conMenu_Format_Left).Checked = False
        objMenu.CommandBar.FindControl(, conMenu_Format_Center).Checked = True
        objMenu.CommandBar.FindControl(, conMenu_Format_Right).Checked = False
        rtfContent.SelAlignment = 2
    Case conMenu_Format_Right
        '靠右
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_FormatPopup)
        objMenu.CommandBar.FindControl(, conMenu_Format_Left).Checked = False
        objMenu.CommandBar.FindControl(, conMenu_Format_Center).Checked = False
        objMenu.CommandBar.FindControl(, conMenu_Format_Right).Checked = True
        rtfContent.SelAlignment = 1
    Case conMenu_Format_Decrease
        '减少缩进量
        i = IIf(rtfContent.SelIndent > 360, rtfContent.SelIndent, 360)
        rtfContent.SelIndent = i - 360
    Case conMenu_Format_Increase
        ' 增加缩进量
        i = IIf(rtfContent.SelIndent < rtfContent.Width - 1000, rtfContent.SelIndent, rtfContent.Width - 1000)
        rtfContent.SelIndent = i + 360
    Case conMenu_Edit_Reply
        '答复
        frmMessageEdit.OpenWindow "", mstrID, mlng类型, 1
        Unload Me
    Case conMenu_Edit_AllReply
        '全部答复
        frmMessageEdit.OpenWindow "", mstrID, mlng类型, 2
        Unload Me
    Case conMenu_Edit_Transmit
        '转发消息
        frmMessageEdit.OpenWindow "", mstrID, mlng类型, 3
        Unload Me
    Case conMenu_Action_Hight
        '设置消息的重要性：高
        mblnHigh = Not mblnHigh
        
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ActionPopup)
        objMenu.CommandBar.FindControl(, conMenu_Action_Hight).Checked = mblnHigh
        cbsMain.FindControl(, conMenu_Action_Hight).Checked = mblnHigh
        
        objMenu.CommandBar.FindControl(, conMenu_Action_Low).Checked = False  '菜单
        cbsMain.FindControl(, conMenu_Action_Low).Checked = False '工具栏
        mblnLow = False
    Case conMenu_Action_Low
        '设置消息的重要性：低
        mblnLow = Not mblnLow
        
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ActionPopup)
        objMenu.CommandBar.FindControl(, conMenu_Action_Low).Checked = mblnLow
        cbsMain.FindControl(, conMenu_Action_Low).Checked = mblnLow
        
        objMenu.CommandBar.FindControl(, conMenu_Action_Hight).Checked = False '菜单
        cbsMain.FindControl(, conMenu_Action_Hight).Checked = False '工具栏
        mblnHigh = False
    Case conMenu_Format_Font
        '字体
        If Control.Type = xtpControlComboBox Then
            rtfContent.SelFontName = Control.Text
        End If
        
    Case conMenu_Format_SIZE
        '字号
        If Control.Type = xtpControlComboBox Then
            rtfContent.SelFontSize = Control.Text
        End If
    Case conMenu_FORMAT_BOLD
        '粗体
        rtfContent.SelBold = Not rtfContent.SelBold
    Case conMenu_FORMAT_ITALIC
        '斜体
        rtfContent.SelItalic = Not rtfContent.SelItalic
    Case conMenu_FORMAT_UNDERLINE
        '下划线
        rtfContent.SelUnderline = Not rtfContent.SelUnderline
    Case conMenu_Help_Help
        Call ShowHelp(App.ProductName, Me.hWnd, "ZL9AppTool\" & Me.Name, 0)
    Case conMenu_Help_Web_Mail
        '发送反馈
         Call zlMailTo(Me.hWnd)
    Case conMenu_Help_Web_Home
        '主页
        Call zlHomePage(Me.hWnd)
    Case conMenu_Help_Web_Forum
        '论坛
        Call zlWebForum(Me.hWnd)
    Case conMenu_Help_About
        '关于
        ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
    Case conMenu_File_Exit
    
        Unload Me
    End Select
    cbsMain.RecalcLayout
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long

    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)

    On Error Resume Next

    '窗体其它控件Resize处理
    
    '调整第一行各控件的位置
    
    lbl(0).Top = lngTop + 30
    lbl(0).Left = 60
    lbl(1).Top = lbl(0).Top
    lbl(1).Left = ScaleWidth - lbl(1).Width - 60
    '调整第二行各控件的位置
    cmdReceive.Left = lbl(0).Left
    If mblnSend = False Then
        cmdReceive.Top = lbl(0).Top + lbl(0).Height + 30
    Else
        cmdReceive.Top = lbl(0).Top + 90
    End If
    
    txtReceive.Top = cmdReceive.Top + 25
    txtReceive.Left = cmdReceive.Left + cmdReceive.Width + 60
    txtReceive.Width = ScaleWidth - txtReceive.Left - 60
    
    lbl(2).Left = lbl(0).Left
    '调整第三行各控件的位置
    txtSubject.Top = cmdReceive.Top + cmdReceive.Height + 60
    txtSubject.Left = txtReceive.Left
    txtSubject.Width = txtReceive.Width
    
    lbl(3).Left = lbl(0).Left
    If mblnSend = True Then
        lbl(2).Top = txtReceive.Top + 60
        lbl(3).Top = txtSubject.Top + 60
    Else
        lbl(2).Top = txtReceive.Top
        lbl(3).Top = txtSubject.Top
    End If
        
    '调整编辑控件的位置
    rtfContent.Left = lbl(0).Left
    rtfContent.Top = txtSubject.Top + txtSubject.Height + 60
    rtfContent.Width = ScaleWidth - rtfContent.Left - 60
    rtfContent.Height = lngBottom - rtfContent.Top - 30
    
       Me.Refresh
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)

    '首先设置可见性
    If mblnSetState Then
        If InStr(frmMessageManager.mstrPrivs, "发送消息") = 0 Then
            Select Case Control.ID
                Case conMenu_File_Send, conMenu_ActionPopup
                    Control.Visible = False
                    mblnSetState = False
            End Select
        End If
        
    End If
    
    Select Case Control.ID
        Case conMenu_File_Send, conMenu_File_Save, conMenu_Action_Hight, conMenu_Action_Low, _
             conMenu_Edit_Cut, conMenu_Edit_plaster, conMenu_Edit_Clear
            '新消息
            Control.Enabled = mblnSend
        Case conMenu_FormatPopup
            Control.Visible = mblnSend
        Case conMenu_Edit_Reply, conMenu_Edit_AllReply, conMenu_Edit_Transmit
            '打开消息
            Control.Enabled = Not mblnSend
        Case conMenu_FORMAT_BOLD: Control.Checked = IIf(IsNull(rtfContent.SelBold), False, rtfContent.SelBold)
        Case conMenu_FORMAT_ITALIC: Control.Checked = IIf(IsNull(rtfContent.SelItalic), False, rtfContent.SelItalic)
        Case conMenu_FORMAT_UNDERLINE: Control.Checked = IIf(IsNull(rtfContent.SelUnderline), False, rtfContent.SelUnderline)
        Case conMenu_Format_Left: Control.Checked = rtfContent.SelAlignment = rtfLeft
        Case conMenu_Format_Center: Control.Checked = rtfContent.SelAlignment = rtfCenter
        Case conMenu_Format_Right: Control.Checked = rtfContent.SelAlignment = rtfRight
    End Select
End Sub

Private Sub ColorFillColor_GotFocus()
    ColorFillColor.Tag = "Focused"
End Sub

Private Sub ColorFillColor_pOK()
    Dim lngSelFillColor As Long
    lngSelFillColor = IIf(ColorFillColor.Color = tomAutoColor, ColorFillColor.AutoColor, ColorFillColor.Color)
    rtfContent.BackColor = lngSelFillColor
    SetColorIcon "FILLCOLOR", conMenu_Format_FillColor, lngSelFillColor
    SendKeys "{ESCAPE}"
End Sub

Private Sub ColorForeColor_pOK()
    Dim lngSelForeColor As Long
    lngSelForeColor = IIf(ColorForeColor.Color = tomAutoColor, ColorForeColor.AutoColor, ColorForeColor.Color)
    rtfContent.SelColor = lngSelForeColor
    SetColorIcon "FORECOLOR", conMenu_Format_ForeColor, lngSelForeColor
    SendKeys "{ESCAPE}"
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    Call cbsMain_Resize
End Sub

Private Sub Form_Load()
    RestoreWinState Me, App.ProductName
    Call LoadMessage
    Call InitCommandBar
    Call SetState
    mblnChange = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = True Then
        If MsgBox("如果你就这样退出的话，所有的修改都不会生效。" & vbCrLf & "是否确认退出？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = 1
            Exit Sub
        End If
    End If
    SaveWinState Me, App.ProductName
    mstr会话ID = ""
    mblnSend = False
    mblnDelete = False

    mblnChange = False
    mblnHigh = False
    mblnLow = False
    mblnSetState = False
    Set mrsUser = Nothing
End Sub

Private Sub cmdReceive_Click()
    Dim str收件人  As String
    Dim rsUser As ADODB.Recordset
    
    Set rsUser = mrsUser
    str收件人 = txtReceive.Text
    If frmSelectReceiver.Get收件人(str收件人, rsUser) = True Then
        
        Set mrsUser = rsUser
        
        txtReceive.Text = str收件人
        txtSubject.SetFocus
    End If
End Sub

Private Sub CoolBar1_HeightChanged(ByVal NewHeight As Single)
    Call Form_Resize
End Sub

Private Sub rtfContent_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If (Shift And vbCtrlMask) > 0 Then
            'Call mnuFileSend_Click
        End If
    End If
End Sub

Private Sub txtReceive_Change()
    mblnChange = True
End Sub

Private Sub txtReceive_GotFocus()
    zlCommFun.OpenIme False
End Sub

Private Sub txtReceive_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If (Shift And vbCtrlMask) > 0 Then
'            Call mnuFileSend_Click
        Else
            txtSubject.SetFocus
        End If
    End If
End Sub

Private Sub txtSubject_Change()
    mblnChange = True
End Sub

Private Sub txtSubject_GotFocus()
    zlCommFun.OpenIme True
End Sub

Private Sub rtfContent_Change()
    mblnChange = True
End Sub

Private Sub rtfContent_GotFocus()
    zlCommFun.OpenIme True
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(Me.hWnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(Me.hWnd)
End Sub

Public Sub OpenWindow(ByVal strID As String, ByVal strByID As String, Optional ByVal lng类型 As Long, Optional ByVal lngMode As Long)
'功能：根据参数来显示消息编辑窗口
'strID 消息ID。如果为空，表示是新消息
'strByID 要进行转发、答复的原始ID
'lngMode  处理方式。1-答复；2-全部答复；3-转发
    Dim frmMessage As frmMessageEdit
    Dim frmTemp As Form
    
    '查找该消息是否已经打开了编辑窗口
    If strID <> "" Then
        For Each frmTemp In Forms
            If frmTemp.Name = "frmMessageEdit" Then
                If frmTemp.mstrID = strID And frmTemp.mlng类型 = lng类型 Then
                    Set frmMessage = frmTemp
                    Exit For
                End If
            End If
        Next
    End If
    If frmMessage Is Nothing Then
        Set frmMessage = New frmMessageEdit
        frmMessage.mstrID = strID
        frmMessage.mstrByID = strByID
        frmMessage.mlng类型 = lng类型
        frmMessage.mlngMode = lngMode
    End If
    frmMessage.Show , gfrmMain
    
End Sub

Private Sub SetState()
    '首先设置可见性
    cmdReceive.Visible = mblnSend
    lbl(0).Visible = Not mblnSend
    lbl(1).Visible = Not mblnSend
    lbl(2).Visible = Not mblnSend
    If Not mblnSend Then
        txtReceive.Appearance = 0
        txtReceive.BorderStyle = 0
        txtReceive.BackColor = BackColor
        txtReceive.Enabled = False
        txtSubject.Appearance = 0
        txtSubject.BorderStyle = 0
        txtSubject.BackColor = BackColor
        txtSubject.Enabled = False
    End If
   
    Dim cmdBar As CommandBar, i As Integer
    For i = 1 To cbsMain.Count
        Set cmdBar = cbsMain.item(i)
        If cmdBar.Title = "格式工具栏" Then
            cmdBar.Visible = mblnSend
        End If
    Next
    'CoolBar1.Bands("three").Visible = mblnSend
    '接着设置可用性
    rtfContent.Locked = Not mblnSend
    '调用工具栏的初始化过程
    mblnSetState = True
End Sub

Private Sub LoadMessage()
'装入数据
    Dim lngID As Long
    Dim rsTemp As New ADODB.Recordset
    Dim str发件人 As String
    Dim str收件人 As String
    Dim str姓名 As String
    
    lngID = Val(IIf(mstrID = "", mstrByID, mstrID)) '
    If lngID = 0 Then
        mblnDelete = False
        mblnSend = True
        Exit Sub '是全新的，用不着装入
    End If
    
    '得到邮件正文
    rsTemp.CursorLocation = adUseClient
    Set rsTemp = zlDatabase.OpenCursor(Me.Caption, "zlTools", "b_ComFunc.Get_Zlmessage", lngID, mlng类型, gstrDbUser)
    If rsTemp.RecordCount <= 0 Then Exit Sub
    mstr会话ID = rsTemp("会话ID")
    mblnDelete = IIf(rsTemp("删除") = 1, True, False)
    mblnSend = (mstrID = "") Or (mlng类型 <> 2 And mlng类型 <> 1) '新邮件或未发送邮件
    
    lbl(1).Caption = "发送时间：" & IIf(IsNull(rsTemp("时间")), "", Format(rsTemp("时间"), "yyyy-MM-dd HH:mm:ss"))
    str发件人 = IIf(IsNull(rsTemp("发件人")), "", rsTemp("发件人"))
    str收件人 = IIf(IsNull(rsTemp("收件人")), "", rsTemp("收件人"))
    txtSubject.Text = IIf(IsNull(rsTemp("主题")), "", rsTemp("主题"))
    rtfContent.TextRTF = IIf(IsNull(rsTemp("内容")), "", rsTemp("内容"))
    rtfContent.BackColor = IIf(IsNull(rsTemp("背景色")), RGB(255, 255, 255), rsTemp("背景色"))
    
    '得到邮递地址
    Set rsTemp = zlDatabase.OpenCursor(Me.Caption, "zlTools", "b_ComFunc.Get_Zlmsgstate", lngID)
    
    '创建空的收件人记录集
    Dim strFild As String
    strFild = "用户名,Varchar2,30;姓名,varchar2,30;收件人,varchar2,30"
    
    '更新界面内容
    
    
    Select Case mlngMode
        Case 1 '答复
            Set mrsUser = NewClientRecord(strFild)
            rsTemp.Filter = "类型=0 or 类型=1"
            mrsUser.AddNew
            mrsUser.Fields("姓名") = rsTemp("身份")
            mrsUser.Fields("用户名") = rsTemp("用户")
            
            txtReceive.Text = mrsUser.Fields("姓名")
            If Left(txtSubject.Text, 3) <> "答复：" Then
                txtSubject.Text = "答复：" & txtSubject.Text
            End If
        Case 2 '全部答复
            Set mrsUser = NewClientRecord(strFild)
            str姓名 = ""
            If str收件人 = "所有人员" Or str收件人 = "本部门人员" Or str收件人 = "本科室人员" Then
                txtReceive.Text = str收件人
            ElseIf InStr(str收件人, "]") > 0 And InStr(str收件人, "[") > 0 Then
                txtReceive.Text = str收件人
            Else
                If str发件人 <> str收件人 And InStr(str收件人, str发件人 & ",") = 0 And InStr(str收件人, "," & str发件人) = 0 Then
                    txtReceive.Text = str发件人 & "," & str收件人
                Else
                    txtReceive.Text = str收件人
                End If
            End If
            If Left(txtSubject.Text, 3) <> "答复：" Then
                txtSubject.Text = "答复：" & txtSubject.Text
            End If
            rsTemp.Filter = "类型=3 or 类型=2"
            Do Until rsTemp.EOF
                mrsUser.AddNew
                mrsUser.Fields("姓名") = rsTemp("身份")
                mrsUser.Fields("用户名") = rsTemp("用户")
                str姓名 = str姓名 & rsTemp("身份") & ","
                rsTemp.MoveNext
            Loop
            If str发件人 <> str姓名 And InStr(str姓名, str发件人 & ",") = 0 And InStr(str姓名, "," & str发件人) = 0 Then
            
                mrsUser.AddNew
                mrsUser.Fields("姓名") = str发件人
                mrsUser.Fields("用户名") = gstrDbUser
                
            End If
        Case 3 '转发
            Set mrsUser = NewClientRecord(strFild)
            txtReceive.Text = ""
            If Left(txtSubject.Text, 3) <> "转发：" Then
                txtSubject.Text = "转发：" & txtSubject.Text
            End If
        Case Else
            Set mrsUser = NewClientRecord(strFild)
            txtReceive = str收件人
            rsTemp.Filter = "类型=3 or 类型=2"
            
            Do Until rsTemp.EOF
                mrsUser.AddNew
                mrsUser.Fields("姓名") = str发件人
                mrsUser.Fields("用户名") = gstrDbUser
            
                rsTemp.MoveNext
            Loop
    End Select
    
    
    lbl(0).Caption = "发件人：     " & str发件人
    Me.Caption = "消息  " & IIf(txtSubject.Text = "", "", "-  " & txtSubject.Text)
    If mlngMode <> 0 Then
        '把原件加上区别
        With rtfContent
            .SelStart = 0
            .SelText = vbCrLf & "----------原始消息------------" & vbCrLf
            .SelStart = 2
            .SelLength = Len("----------原始消息------------")
            .SelFontName = "宋体"
            .SelFontSize = 9
            .SelBold = False
            .SelItalic = False
            .SelColor = 0
            
            .SelLength = Len(.Text)
            .SelIndent = 720
            
            .SelStart = 0
            .SelLength = 2
            .SelColor = RGB(0, 0, 255)
            .SelFontName = "宋体"
            .SelFontSize = 9
            .SelBold = False
            .SelItalic = False
            .SelLength = 0
        End With
    End If
    'Update_Zlmsgstate_Idtntify(身份,消息ID,类型,用户)
    Call zlDatabase.OpenCursor(Me.Caption, "zltools", "b_ComFunc.Update_Zlmsgstate_Idtntify", _
                                gstrUserName, _
                                lngID, _
                                mlng类型, _
                                gstrDbUser)
    
    Call frmMessageRead.UpdateNotify
    Dim lst As ListItem
    On Error Resume Next
    '改成已读
'    Set lst = frmMessageManager.lvwMain.ListItems("C" & mlng类型 & lngID)
    If Not lst Is Nothing Then
        If lst.Icon <> "Script" Then '不是草稿的消息才可能更改
            If InStr(lst.Icon, "Reply") > 0 Then
                lst.Icon = "ReadReply"
                lst.SmallIcon = "ReadReply"
            Else
                lst.Icon = "Read"
                lst.SmallIcon = "Read"
            End If
        End If
        frmMessageManager.Refresh
    End If
    Err.Clear
End Sub

Private Function SaveMessage(ByVal blnSend As Boolean) As Boolean
    Dim lngID As Long
    Dim lng会话ID As Long
    Dim lngCount As Long
    
    Dim strUsers As String '存储已经保存了的用户名
    Dim rsTemp As ADODB.Recordset
    '没修改，且只是保存
    
    If blnSend = True And txtReceive.Text = "" And mrsUser Is Nothing Then
        MsgBox "请选择收件人。", vbExclamation, gstrSysName
        cmdReceive.SetFocus
        Exit Function
    End If
    If zlCommFun.StrIsValid(txtReceive.Text, txtReceive.MaxLength, cmdReceive.hWnd, "收件人") = False Then
        Exit Function
    End If
    
    If zlCommFun.StrIsValid(txtSubject.Text, txtSubject.MaxLength, txtSubject.hWnd, "主题") = False Then
        Exit Function
    End If
    If LenB(StrConv(rtfContent.TextRTF, vbFromUnicode)) > 4000 Then
        MsgBox "正文的字符太多，或者格式太复杂了。", vbExclamation, gstrSysName
        rtfContent.SetFocus
        Exit Function
    End If
    
    On Error GoTo errHandle
    gcnOracle.BeginTrans
    '处理主表
    'Save_Zlmessage(ID,会话ID,发件人,收件人,主题,内容,背景色)
    Set rsTemp = zlDatabase.OpenCursor(Me.Caption, "zltools", "b_ComFunc.Save_Zlmessage", _
                                Val(mstrID), _
                                Val(mstr会话ID), _
                                gstrUserName, _
                                txtReceive.Text, _
                                txtSubject.Text, _
                                rtfContent.TextRTF, _
                                rtfContent.BackColor)
    
    If rsTemp.RecordCount > 0 Then
        lngID = zlCommFun.NVL(rsTemp.Fields(0), 0)
        lng会话ID = zlCommFun.NVL(rsTemp.Fields(1), 0)
        '处理从表
    
        'Insert_Zlmsgstate (消息ID,类型,用户,身份,删除,状态)
         Call zlDatabase.OpenCursor(Me.Caption, "zltools", "b_ComFunc.Insert_Zlmsgstate", _
                                     lngID, _
                                     Val(IIf(blnSend, "1", "0")), _
                                     gstrDbUser, _
                                     gstrUserName, _
                                     Val(IIf(mblnDelete = True And blnSend = False, "1", "0")), _
                                     CStr(IIf(blnSend, "1", "0") & "00" & IIf(mblnHigh = True, "1", IIf(mblnLow = True, "2", "0"))))
        ''''' 增加收件人的记录
         If Not mrsUser Is Nothing Then
             If mrsUser.State = adStateOpen Then
                 If mrsUser.RecordCount > 0 Then mrsUser.MoveFirst
                 Do Until mrsUser.EOF
                    Call zlDatabase.OpenCursor(Me.Caption, "zltools", "b_ComFunc.Insert_Zlmsgstate", _
                                               lngID, _
                                               Val(IIf(blnSend, "2", "3")), _
                                               IIf(IsNull(mrsUser.Fields("用户名")), "", mrsUser.Fields("用户名")), _
                                               IIf(IsNull(mrsUser.Fields("姓名")), "", mrsUser.Fields("姓名")), _
                                               0, _
                                               CStr("000" & IIf(mblnHigh = True, "1", IIf(mblnLow = True, "2", "0"))))
                    
                    mrsUser.MoveNext
                Loop
            End If
        End If
        
        '最后为原件加上答复或转发标志

        'Update_Zlmsgstate_State(模式,消息ID,类型,用户)
        Call zlDatabase.OpenCursor(Me.Caption, "zltools", "b_ComFunc.Update_Zlmsgstate_State", _
                                    mlngMode, _
                                    Val(mstrByID), _
                                    mlng类型, _
                                    gstrDbUser)
    End If

    gcnOracle.CommitTrans
    mstrID = lngID
    mstr会话ID = lng会话ID
    mblnChange = False
    SaveMessage = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    gcnOracle.RollbackTrans
End Function

Public Sub FindText()
    '如果查找内容为空，直接退出
    Dim lngPos As Long, lngStart As Long
    Dim strText As String
    
    If mstrFind = "" Then Exit Sub
    
    strText = rtfContent.Text
    If mblnBegin = False Then
        lngStart = rtfContent.SelStart + rtfContent.SelLength + 1
    Else
        lngStart = 1
    End If
    
    lngPos = InStr(lngStart, IIf(mblnCase = True, strText, UCase(strText)), IIf(mblnCase = True, mstrFind, UCase(mstrFind)))
    
    If lngPos = 0 Then
        MsgBox "查找结束，没找到“" & mstrFind & "”", vbInformation, gstrSysName
    Else
        rtfContent.SelStart = lngPos - 1
        rtfContent.SelLength = Len(mstrFind)
    End If
    
End Sub

Private Sub txtSubject_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If (Shift And vbCtrlMask) > 0 Then
'            Call mnuFileSend_Click
        Else
            rtfContent.SetFocus
        End If
    End If
End Sub


Private Sub InitCommandBar()
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    
    Dim objCustControl As CommandBarControlCustom       '自定义控件

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
        .LargeIcons = False
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
        Set objControl = .Add(xtpControlButton, conMenu_File_Send, "发送(&E)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Save, "保存(&S)")
        Set objControl = .Add(xtpControlButton, conMenu_File_SaveAs, "另存为(&A)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"): objControl.BeginGroup = True
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    objMenu.ID = conMenu_EditPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Cut, "剪切(&T)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Copy, "复制(&C)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_plaster, "粘贴(&P)")
        
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Clear, "清除(&L)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_CheckAll, "全选(&A)")
        
        Set objControl = .Add(xtpControlButton, conMenu_View_Find, "查找(&F)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_View_FindNext, "查找下一处(&N)")
       
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FormatPopup, "格式(&F)", -1, False)
    objMenu.ID = conMenu_FormatPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Format_Sig, "项目符号(&S)")
        
        Set objControl = .Add(xtpControlButton, conMenu_Format_Left, "靠左对齐(&L)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Format_Center, "居中对齐(&C)")
        Set objControl = .Add(xtpControlButton, conMenu_Format_Right, "靠右对齐(&R)")
        Set objControl = .Add(xtpControlButton, conMenu_Format_Decrease, "减少缩进量(&D)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Format_Increase, "增加缩进量(&I)")
       
    End With
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    objMenu.ID = conMenu_ViewPopup
    With objMenu.CommandBar.Controls
'        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_ToolBar, "工具栏(&T)")
'        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False
'        End With
    End With
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ActionPopup, "动作(&A)", -1, False)
    objMenu.ID = conMenu_ActionPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Reply, "答复(&R)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_AllReply, "全部答复(&L)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Transmit, "转发(&W)")
        
        Set objControl = .Add(xtpControlButton, conMenu_Action_Hight, "重要性高(&H)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Action_Low, "重要性低(&O)")
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
    Set objBar = cbsMain.Add("新建工具栏", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagHideWrap
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_Send, "发送")
        Set objControl = .Add(xtpControlButton, conMenu_File_Save, "保存")
        
        Set objControl = .Add(xtpControlButton, conMenu_Action_Hight, "重要"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Action_Low, "次要")
        '----
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Reply, "答复"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_AllReply, "全部答复")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Transmit, "转发")
        
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
        
    End With
    
    For Each objControl In objBar.Controls
        objControl.Style = xtpButtonIconAndCaption
    Next
    
    '
    Set objBar = cbsMain.Add("格式工具栏", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagHideWrap
    
    With objBar.Controls
        .Add xtpControlButton, conMenu_Edit_Cut, "剪切", -1, False
        .Add xtpControlButton, conMenu_Edit_Copy, "复制", -1, False
        .Add xtpControlButton, conMenu_Edit_plaster, "粘贴", -1, False
        
        Set objControl = .Add(xtpControlComboBox, conMenu_Format_Font, "字体")
        objControl.BeginGroup = True
        objControl.Width = 140
        objControl.DropDownListStyle = True
        '装入字体和字号
        Dim lngCount As Long, lngDefault As Long
        
        For lngCount = 0 To Screen.FontCount - 1
            objControl.AddItem Screen.Fonts(lngCount)
            If objControl.List(lngCount) = "宋体" Then
                lngDefault = lngCount
            End If
        Next
        If lngDefault > 0 Then objControl.ListIndex = lngDefault
        
        Dim ControlComboSize As CommandBarComboBox
        Set ControlComboSize = .Add(xtpControlComboBox, conMenu_Format_SIZE, "字号")
        ControlComboSize.Width = 60
        ControlComboSize.AddItem "8"
        ControlComboSize.AddItem "9"
        ControlComboSize.AddItem "10"
        ControlComboSize.AddItem "11"
        ControlComboSize.AddItem "12"
        ControlComboSize.AddItem "14"
        ControlComboSize.AddItem "16"
        ControlComboSize.AddItem "18"
        ControlComboSize.AddItem "20"
        ControlComboSize.AddItem "22"
        ControlComboSize.AddItem "24"
        ControlComboSize.AddItem "26"
        ControlComboSize.AddItem "28"
        ControlComboSize.AddItem "36"
        ControlComboSize.AddItem "48"
        ControlComboSize.AddItem "72"
        ControlComboSize.DropDownListStyle = True
        ControlComboSize.ListIndex = 3
        
        Set objControl = .Add(xtpControlButton, conMenu_FORMAT_BOLD, "加粗", -1, False)
        objControl.BeginGroup = True
        .Add xtpControlButton, conMenu_FORMAT_ITALIC, "斜体", -1, False
        .Add xtpControlButton, conMenu_FORMAT_UNDERLINE, "下划线", -1, False
        
        Set objControl = .Add(xtpControlButton, conMenu_Format_Left, "左对齐", -1, False): objControl.BeginGroup = True
        .Add xtpControlButton, conMenu_Format_Center, "居中", -1, False
        .Add xtpControlButton, conMenu_Format_Right, "右对齐", -1, False
        
        Set objPopup = .Add(xtpControlSplitButtonPopup, conMenu_Format_ForeColor, "字体颜色")
        Set objCustControl = objPopup.CommandBar.Controls.Add(xtpControlCustom, 0, "")
        objCustControl.Handle = ColorForeColor.hWnd
        objControl.BeginGroup = True
        
        Set objPopup = .Add(xtpControlSplitButtonPopup, conMenu_Format_FillColor, "背景色")
        Set objCustControl = objPopup.CommandBar.Controls.Add(xtpControlCustom, 0, "")
        objCustControl.Handle = ColorFillColor.hWnd
        
        Set objControl = .Add(xtpControlButton, conMenu_View_Find, "&查找"): objControl.BeginGroup = True
        .Add xtpControlButton, conMenu_Format_Sig, "&项目符号"
        .Add xtpControlButton, conMenu_Format_Decrease, "&减少缩进量"
        .Add xtpControlButton, conMenu_Format_Increase, "&增加缩进量"
        

    End With

    For Each objControl In objBar.Controls
        objControl.Style = xtpButtonIcon
    Next
    '命令的快键绑定:公共部份主界面已处理
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyF, conMenu_View_Find '查找
        .Add 0, vbKeyF1, conMenu_Help_Help '帮助
    End With
    
    '设置一些公共的不常用命令
    With cbsMain.Options
        .AddHiddenCommand conMenu_File_PrintSet '打印设置
        .AddHiddenCommand conMenu_File_Excel '输出到Excel
    End With
    cbsMain.RecalcLayout
    
    '初始化颜色
     
    ColorForeColor.Color = rtfContent.SelColor
    ColorFillColor.Color = rtfContent.BackColor
    
    SetColorIcon "FORECOLOR", conMenu_Format_ForeColor, ColorForeColor.Color
    SetColorIcon "FILLCOLOR", conMenu_Format_FillColor, ColorFillColor.Color
    
    ColorFillColor.Visible = False
    ColorForeColor.Visible = False
    
    Me.BackColor = cbsMain.GetSpecialColor(XPCOLOR_SPLITTER_FACE)
    cmdReceive.BackColor = cbsMain.GetSpecialColor(XPCOLOR_TOOLBAR_FACE)
    
    
End Sub

Private Sub SetColorIcon(Key As String, ID As Long, Color As OLE_COLOR)
    Dim ctlPictureBox As VB.PictureBox
    Set ctlPictureBox = Controls.Add("VB.PictureBox", "ctlPictureBox1")
    Dim ListImage As ListImage
    Set ListImage = imgColor.ListImages(Key)
    
    ctlPictureBox.AutoRedraw = True
    ctlPictureBox.AutoSize = True
    ctlPictureBox.BackColor = imgColor.MaskColor
    
    ctlPictureBox.Picture = ListImage.ExtractIcon
    
   ' If Color = vbWhite Then Color = RGB(254, 254, 254)
    ctlPictureBox.Line (1, ctlPictureBox.Height * 0.6)-(ctlPictureBox.Width, ctlPictureBox.Height), Color, BF
    ctlPictureBox.Refresh

    '更新图标
    imgColor.ListImages.Remove imgColor.ListImages(Key).Index
    imgColor.ListImages.Add 1, Key, ctlPictureBox.Image

    '更新 Tag 属性
    imgColor.ListImages(1).Tag = ID
    cbsMain.AddImageList imgColor
    cbsMain.RecalcLayout
    
    Me.Controls.Remove ctlPictureBox
    Set ctlPictureBox = Nothing
End Sub

Private Sub ColorForeColor_GotFocus()
    ColorForeColor.Tag = "Focused"
End Sub
