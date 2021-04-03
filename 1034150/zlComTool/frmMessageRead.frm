VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Begin VB.Form frmMessageRead 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "读消息"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1890
   Icon            =   "frmMessageRead.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   1890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer timMessage 
      Interval        =   1000
      Left            =   630
      Top             =   1140
   End
   Begin VB.Image imgClose 
      Height          =   240
      Left            =   75
      Picture         =   "frmMessageRead.frx":6852
      Top             =   750
      Width           =   240
   End
   Begin VB.Image imgMail 
      Height          =   480
      Left            =   30
      Picture         =   "frmMessageRead.frx":699C
      Top             =   225
      Width           =   480
   End
   Begin VB.Image imgTipClose 
      Height          =   270
      Left            =   585
      Picture         =   "frmMessageRead.frx":7666
      Top             =   375
      Width           =   810
   End
   Begin XtremeSuiteControls.PopupControl popMsg 
      Left            =   615
      Top             =   1665
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   4
   End
End
Attribute VB_Name = "frmMessageRead"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const WM_MOUSEMOVE = &H200
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Dim mblnIcon As Boolean '图标已经显示
Dim mdatLast As Date     '最后一次显示通知图标的时间
'------------------------------------------------------------------------------------------
'Popupcontrol 控件显示用到的API
Const IDOK = 1
Const IDCLOSE = 2
Const IDSITE = 3

Private Sub Form_Load()
    Dim lngTime As Long
    
    mdatLast = zlDatabase.CurrentDate()
    
    If Val(zlDatabase.GetPara("登录检查邮件消息")) = 1 Then
        '只要有未读消息就提醒
        mdatLast = CDate("1900-01-01")
        
        lngTime = Val(zlDatabase.GetPara("邮件消息检查周期"))
        If lngTime < 10 Or lngTime > 60 Then lngTime = 30
        timMessage.Interval = lngTime * 1000
    End If
    
    mblnIcon = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '只需对X进行处理就行了。这是MouseMove是一个通用事件
    '如果事件是由鼠标移动引起的,X的值为"1E00";如果事件是按鼠标右键引起的,X的值为"1E3C"......
    If Hex(X) <> "1E0F" Then Exit Sub '鼠标左键按下
    
    If gblnMessageShow = False Then
        '打开消息收发窗口
        With frmMessageManager
            On Error Resume Next
            .Show , gfrmMain
            If Err.Number <> 0 Then
                ShowWindow .hWnd, SW_SHOWNORMAL
                BringWindowToTop .hWnd
                SetActiveWindow .hWnd
            End If
        End With
    Else
        With frmMessageManager
            .mlngIndexPre = -1
            If .mlngIndex = 1 Then
               .FillList
            End If
            On Error Resume Next
            .Show , gfrmMain
            If Err.Number <> 0 Then
                ShowWindow .hWnd, SW_SHOWNORMAL
                BringWindowToTop .hWnd
                SetActiveWindow .hWnd
            End If
        End With
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call RemoveIcon
End Sub

Private Sub timMessage_Timer()
    Call UpdateNotify
End Sub

Public Sub UpdateNotify()
'更新通知消息
    Dim rsTemp As New ADODB.Recordset
    Dim datTemp As Date
    Dim str主题 As String

    '读出未读消息的最大值

    gstrSQL = "select A.时间,A.主题 From zlmessages A, (select max(A.时间) as 时间 " & _
              "  from zlmessages A, zlmsgstate B " & _
              "  where A.ID=B.消息ID and B.类型=2 and B.删除=0 and substr(B.状态,1,1)='0' and B.用户=[1]) B where A.时间=B.时间 "
              
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, gstrDbUser)
    If IsNull(rsTemp("时间")) Then
        Call RemoveIcon
    Else
        Do Until rsTemp.EOF
            datTemp = rsTemp("时间")
            str主题 = "" & rsTemp!主题
            rsTemp.MoveNext
        Loop
        If datTemp > mdatLast Then
            '表示有新邮件到了
            If gblnMessageShow = True Then
                If frmMessageManager.mlngIndex = 1 Then
                    frmMessageManager.mlngIndexPre = -1
                    frmMessageManager.FillList    '直接刷新列表
                End If
            End If
            If Len(str主题) > 20 Then str主题 = Mid(str主题, 1, 17) & "..."
            Call PopShow(0, str主题)
            Call AddIcon
            mdatLast = datTemp        '把这个时间做为最后一次的
        ElseIf datTemp < mdatLast Then
            '表示该邮件已经读了
            Call RemoveIcon
        End If
    End If
End Sub

Private Sub AddIcon()
    Dim t As NOTIFYICONDATA
    
    If mblnIcon = False Then
        t.cbSize = Len(t)
        t.hWnd = Me.hWnd   '事件发生的载体
        t.uId = 1&
        t.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        t.ucallbackMessage = WM_MOUSEMOVE
        t.hIcon = Me.Icon
        t.szTip = "您有新的、未打开过的消息到了" & Chr$(0)
        Shell_NotifyIcon NIM_ADD, t
        Beep
    End If
    mblnIcon = True
End Sub

Private Sub RemoveIcon()
    Dim t As NOTIFYICONDATA
    t.cbSize = Len(t)
    t.hWnd = Me.hWnd   '事件发生的载体
    t.uId = 1&
    Shell_NotifyIcon NIM_DELETE, t
    mblnIcon = False
End Sub

Private Sub popMsg_ItemClick(ByVal item As XtremeSuiteControls.IPopupControlItem)
    
    Dim frmTmp As Form, lngStyle As Long
    On Error Resume Next
    If item.ID = IDCLOSE Or item.ID = IDOK Then
        popMsg.Close
    End If

    If item.ID = IDSITE Then
        popMsg.Close
        frmMessageManager.Show , gfrmMain
        If Err.Number = 401 Then
            Unload frmMessageManager
            Call PopShow(1, "现在不能打开消息阅读窗体，请先关闭当前窗体！")
        ElseIf Err.Number <> 0 Then
            If ErrCenter() = 1 Then
                Resume
            End If
        End If
    End If

End Sub

Private Sub PopShow(ByVal lngType As Long, ByVal strMsg As String)
    '显示消息提示
    On Error Resume Next
    
    popMsg.Animation = 2 'pop窗体出现动画 0-None(无) 1-Fade(渐现)　2-Slide(滑动) 3-Unfold(展现)
    popMsg.AnimateDelay = 500 '动画延时 ms
    popMsg.ShowDelay = 5000 '显示延时 ms
    popMsg.Transparency = 200 '透明度
    
    If lngType = 0 Then
        SetOffice2003Theme (strMsg)
    Else
        SetToolTipTheme (strMsg)
    End If
    popMsg.Show
End Sub

Private Sub SetOffice2003Theme(ByVal str主题 As String)
    '显示　新消息提示
    Dim item As PopupControlItem
    popMsg.RemoveAllItems
    
    Set item = popMsg.AddItem(50, 27, 200, 45, "您有新的消息！")
    item.Bold = True
    item.Hyperlink = False
    
    Set item = popMsg.AddItem(12, 20, 12, 27, "")
    item.SetIcon Me.imgMail.Picture.Handle, xtpPopupItemIconNormal
    item.IconIndex = 1
    
    Set item = popMsg.AddItem(50, 45, 270, 95, str主题)
    item.TextColor = RGB(0, 61, 178)
    item.ID = IDSITE
    item.TextAlignment = DT_LEFT Or DT_WORDBREAK
    
    Set item = popMsg.AddItem(250, 10, 266, 26, "")
    item.SetIcon Me.imgClose.Picture.Handle, xtpPopupItemIconNormal
    item.ID = IDCLOSE
    item.Button = True
    
    popMsg.VisualTheme = xtpPopupThemeOffice2003
    popMsg.SetSize 270, 100

End Sub

Private Sub SetToolTipTheme(ByVal strTip As String)
    '显示提示
    Dim item As PopupControlItem
    
    popMsg.RemoveAllItems
    
    Set item = popMsg.AddItem(0, 0, 220, 90, "", RGB(255, 255, 225), 0)
    
    Set item = popMsg.AddItem(20, 30, 200, 100, strTip)
    item.TextAlignment = DT_CENTER Or DT_WORDBREAK
    item.Hyperlink = False
    
    Set item = popMsg.AddItem(5, 0, 170, 25, "提示")
    item.TextAlignment = DT_SINGLELINE Or DT_LEFT
    item.Bold = True
    item.Hyperlink = False
    
    Set item = popMsg.AddItem(220 - 20, 2, 220 - 2, 2 + 18, "")
    item.SetIcons Me.imgTipClose.Picture.Handle, 0, xtpPopupItemIconNormal Or xtpPopupItemIconSelected Or xtpPopupItemIconPressed
    item.IconIndex = 0
    item.ID = IDCLOSE
   
    popMsg.VisualTheme = xtpPopupThemeCustom
    popMsg.SetSize 220, 90

End Sub
