VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.Unicode.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Begin VB.Form frmMessageManager 
   Caption         =   "��Ϣ�շ�����"
   ClientHeight    =   7245
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   11265
   Icon            =   "frmMessageManager.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7245
   ScaleWidth      =   11265
   ShowInTaskbar   =   0   'False
   Tag             =   "�ɱ仯��"
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
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
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
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
         Caption         =   "�ռ���"
         BeginProperty Font 
            Name            =   "����"
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
Dim mblnLoad As Boolean   '���ڻ�δ��ʱΪ��

Dim mstrKey As String     'δ���µ��ʼ�ID
Dim sngStartY As Single   '�ƶ�ǰ����λ��
Dim mblnItem As Boolean   'Ϊ���ʾ������ListViewĳһ����
Dim mintColumn As Integer '����ListView������

Public mlngIndexPre As Long       '��ʾ֮ǰ���ĸ�Ŀ¼
Public mlngIndex As Long          '��ʾ��ǰ���ĸ�Ŀ¼
Public mstrPrivs As String        'ֻ����Ϣ�շ���ģ���Ȩ��
Public mblnShowAll As Boolean     '��ʾ�Ѷ�
Public mblnLogin As Boolean       '��¼ʱ��ʾ����δ���ʼ�
Const con�ݸ� = 0
Const con�ռ��� = 1
Const con�ѷ�����Ϣ = 2
Const con��ɾ����Ϣ = 3


Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)

    On Error Resume Next
    Dim objControl As CommandBarControl
    Dim i As Integer
    Dim lngID As Long, lng�ỰID As Long
    
    Select Case Control.ID
    Case conMenu_File_PrintSet
        '��ӡ����
        Call zlPrintSet
    Case conMenu_File_Preview
        'Ԥ��
        Call subPrint(2)
    Case conMenu_File_Print
        '��ӡ
        Call subPrint(1)
    Case conMenu_File_Excel
        '�����Excel
        Call subPrint(3)
    Case conMenu_File_SaveAs
        '���Ϊ�ļ�
        On Error Resume Next
        If rtfContent.Text = "" Then Exit Sub
        cdg.CancelError = True
        cdg.Filter = "RTF�ļ�(*.RTF)|*.rtf"
        '����ʱ����ʾ���Ҳ�����ֻ����
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
        '����
        frmMessageEdit.OpenWindow "", ""
        Call FillList
    Case conMenu_Edit_Modify
        '��,�޸�
        Call Edit_Modify
    Case conMenu_Edit_Delete
        'ɾ��
        Call Edit_Delete
    Case conMenu_Edit_Reuse
        '��ԭ
        Call Edit_Restore
    Case conMenu_Edit_Reply
        '��
        If rpcMain.FocusedRow Is Nothing Then Exit Sub
        lngID = Val(rpcMain.Rows(rpcMain.FocusedRow.Index).Record(0).Caption)
        lng�ỰID = Val(rpcMain.Rows(rpcMain.FocusedRow.Index).Record(1).Caption)
        frmMessageEdit.OpenWindow "", lngID, lng�ỰID, 1
        Call FillList
    Case conMenu_Edit_AllReply
        'ȫ����
        If rpcMain.FocusedRow Is Nothing Then Exit Sub
        lngID = Val(rpcMain.Rows(rpcMain.FocusedRow.Index).Record(0).Caption)
        lng�ỰID = Val(rpcMain.Rows(rpcMain.FocusedRow.Index).Record(1).Caption)
        frmMessageEdit.OpenWindow "", lngID, lng�ỰID, 2
        Call FillList
    Case conMenu_Edit_Transmit
        'ת��
        If rpcMain.FocusedRow Is Nothing Then Exit Sub
        lngID = Val(rpcMain.Rows(rpcMain.FocusedRow.Index).Record(0).Caption)
        lng�ỰID = Val(rpcMain.Rows(rpcMain.FocusedRow.Index).Record(1).Caption)
        frmMessageEdit.OpenWindow "", lngID, lng�ỰID, 3
        Call FillList
    Case conMenu_View_ToolBar_Button
        '������
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
        '��ť����
        For i = 2 To cbsMain.Count
            For Each objControl In Me.cbsMain(i).Controls
                objControl.Style = IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
        Next
        Me.cbsMain.RecalcLayout
        Call Form_Resize
    Case conMenu_View_ToolBar_Size
        '��ͼ��
        Me.cbsMain.Options.LargeIcons = Not Me.cbsMain.Options.LargeIcons
        
        If Me.cbsMain.Options.LargeIcons = True Then
            cbarTool.Bands.item(2).MinHeight = 520
        Else
            cbarTool.Bands.item(2).MinHeight = 425
        End If
        Me.cbsMain.RecalcLayout
        Call Form_Resize
    Case conMenu_View_StatusBar
        '״̬��
        Me.stbThis.Visible = Not Me.stbThis.Visible
        Me.cbsMain.RecalcLayout
        Call Form_Resize
    Case conMenu_View_PreviewWindow
        'Ԥ������
        Me.rtfContent.Visible = Not Me.rtfContent.Visible
        picSplitH.Visible = Me.rtfContent.Visible
        Me.cbsMain.RecalcLayout
        Call Form_Resize
    Case conMenu_View_ShowAll
        '��ʾ�Ѷ�
        mblnShowAll = Not mblnShowAll
        Me.cbsMain.RecalcLayout
        mlngIndexPre = -1 'ǿ��ˢ��
        Call FillList
    Case conMenu_View_Login
       '��¼ʱ����
        mblnLogin = Not mblnLogin
        Call zlDatabase.SetPara("��¼����ʼ���Ϣ", IIf(mblnLogin, "1", "0"))
        Me.cbsMain.RecalcLayout
    Case conMenu_View_Find
        '���������Ϣ
        lng�ỰID = Val(rpcMain.Rows(rpcMain.FocusedRow.Index).Record(1).Caption)
        frmMessageRelate.FillList lng�ỰID
    Case conMenu_View_Refresh
        'ˢ��
        mlngIndexPre = -1 'ǿ��ˢ��
        Call FillList
    Case conMenu_Help_Help
        '����
        Call ShowHelp(App.ProductName, Me.hWnd, "ZL9AppTool\" & Me.Name, 0)
    Case conMenu_Help_Web_Home
        'Web�ϵ�����
        Call zlHomePage(Me.hWnd)
    Case conMenu_Help_Web_Forum
        '������̳
        Call zlWebForum(Me.hWnd)
    Case conMenu_Help_Web_Mail
        '���ͷ���
         Call zlMailTo(Me.hWnd)
    Case conMenu_Help_About
        '����
        ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
    Case conMenu_File_Exit
        '�˳�
        Unload Me
    End Select
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnEnabled As Boolean
    
    blnEnabled = Not (rpcMain.FocusedRow Is Nothing)
    
    'Ȩ�޿���
    If InStr(mstrPrivs, "������Ϣ") = 0 Then
        Select Case Control.ID
        Case conMenu_Edit_Add, conMenu_Edit_Reply, conMenu_Edit_AllReply, conMenu_Edit_Transmit
            '���ӣ��𸴣�ȫ���𸴣�ת��
            Control.Enabled = False
            
        End Select
    End If
    
    '�˵�����
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
    stbThis.Panels(2).Text = "����" & lngMessage & "����Ϣ" & IIf(lngSum = 0, "��", "��������" & lngSum & "��δ����")

    Select Case Control.ID
    Case conMenu_Edit_Modify, conMenu_Edit_Delete, conMenu_Edit_Reply, conMenu_Edit_AllReply, conMenu_Edit_Transmit, conMenu_View_Find
        '�޸ģ�ɾ�����𸴣�ȫ���𸴣�ת��,���������Ϣ
        Control.Enabled = blnEnabled
    Case conMenu_Edit_Reuse
        '��ԭ
        Control.Enabled = (mlngIndex = 3 And Not (rpcMain.FocusedRow Is Nothing))
    Case conMenu_File_Print, conMenu_File_Preview, conMenu_File_Excel
        '��ӡ,Ԥ��,�����Excel
        Control.Enabled = lngMessage > 0
    Case conMenu_File_SaveAs
        '���Ϊ
        Control.Enabled = rtfContent.Text <> ""
    Case conMenu_View_ToolBar_Button '������
        If cbsMain.Count >= 2 Then
            Control.Checked = Me.cbsMain(2).Visible
        End If
    Case conMenu_View_ToolBar_Text 'ͼ������
        If cbsMain.Count >= 2 Then
            Control.Checked = Not (Me.cbsMain(2).Controls(1).Style = xtpButtonIcon)
        End If
    Case conMenu_View_ToolBar_Size '��ͼ��
        Control.Checked = Me.cbsMain.Options.LargeIcons
    Case conMenu_View_StatusBar '״̬��
        Control.Checked = Me.stbThis.Visible
    Case conMenu_View_PreviewWindow 'Ԥ������
        Control.Checked = Me.rtfContent.Visible
    Case conMenu_View_ShowAll '��ʾ�Ѷ�
        Control.Checked = mblnShowAll
    Case conMenu_View_Login '��¼����
        Control.Checked = mblnLogin
    End Select
    
End Sub

Private Sub Form_Activate()
    If mblnLoad = True Then
        Call Form_Resize 'Ϊ��ʹCoolBar����Ӧ�߶�
        mlngIndexPre = -1 'ǿ��ˢ��
        Call FillList
    End If
    mblnLoad = False
End Sub

Private Sub Form_Load()
    gblnMessageShow = True
    If gblnMessageGet = False Then
        '����̨��û�д���Ϣ֪ͨ���ڣ�ֻ���Լ�������
        Load frmMessageRead
    End If
    Call DeleteMessage

    mblnLoad = True
    '-----------
    RestoreWinState Me, App.ProductName
    mblnShowAll = Val(zlDatabase.GetPara("��ʾ�Ѷ��ʼ�")) <> 0
    mblnLogin = Val(zlDatabase.GetPara("��¼����ʼ���Ϣ")) <> 0
    
    mstrPrivs = GetPrivFunc(0, 12) 'ȡȨ��
    '-----------------------------------------------------------------------------------------------------------
        '�½���
        Call InitCommandBar
        Dim tpGroup As TaskPanelGroup
    
        Set tpGroup = tplCon.Groups.Add(101, "����")
        
        tpGroup.Items.Add(con�ݸ�, "�ݸ�", xtpTaskItemTypeLink, con�ݸ� + 2).Selected = False
        tpGroup.Items.Add(con�ռ���, "�ռ���", xtpTaskItemTypeLink, con�ռ��� + 2).Selected = False
        tpGroup.Items.Add(con�ѷ�����Ϣ, "�ѷ�����Ϣ", xtpTaskItemTypeLink, con�ѷ�����Ϣ + 2).Selected = False
        tpGroup.Items.Add(con��ɾ����Ϣ, "��ɾ����Ϣ", xtpTaskItemTypeLink, con��ɾ����Ϣ + 2).Selected = False
        
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
    '���ó�ʼ��ѡ��
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
        '����̨��û�д���Ϣ֪ͨ���ڣ�����˳�ʱ����һ������
        Unload frmMessageRead
    End If
    
    mstrKey = ""
    mlngIndexPre = 0
    Call zlDatabase.SetPara("��ʾ�Ѷ��ʼ�", IIf(mblnShowAll, 1, 0))
    SaveWinState Me, App.ProductName
End Sub


Private Sub Edit_Delete()
    On Error GoTo errHandle
    Dim strKey As String
    Dim intIndex As Long
    Dim rsTemp As New ADODB.Recordset
    
    Dim lngID As Long, lng�ỰID As Long
    If rpcMain.FocusedRow Is Nothing Then Exit Sub
    lngID = Val(rpcMain.Rows(rpcMain.FocusedRow.Index).Record(0).Caption)
    lng�ỰID = Val(rpcMain.Rows(rpcMain.FocusedRow.Index).Record(1).Caption)
    
    gcnOracle.BeginTrans
    If mlngIndex <> 3 Then
        'Delete_Zlmsgstate(ɾ��,��ϢID,����,�û�)
        Call zlDatabase.OpenCursor(Me.Caption, "zltools", "b_ComFunc.Delete_Zlmsgstate", _
                                    1, _
                                    lngID, _
                                    lng�ỰID, _
                                    gstrDbUser)
    Else
        If MsgBox("��ȷ��Ҫɾ������Ϊ��" & rpcMain.Rows(rpcMain.FocusedRow.Index).Record(3).Caption & "������Ϣ��", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
            gcnOracle.RollbackTrans
            Exit Sub
        End If
        Me.MousePointer = 11
        Call zlDatabase.OpenCursor(Me.Caption, "zltools", "b_ComFunc.Delete_Zlmsgstate", _
                                            2, _
                                            lngID, _
                                            lng�ỰID, _
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
'��ԭ��ɾ����Ϣ
    On Error GoTo errHandle
'    Dim intIndex As Long
    
    Dim lngID As Long, lng�ỰID As Long
    lngID = Val(rpcMain.Rows(rpcMain.FocusedRow.Index).Record(0).Caption)
    lng�ỰID = Val(rpcMain.Rows(rpcMain.FocusedRow.Index).Record(1).Caption)
    
    'b_ComFunc.Restore_Zlmsgstate(��ϢID,����,�û�)
    Call zlDatabase.OpenCursor(Me.Caption, "zltools", "b_ComFunc.Restore_Zlmsgstate", _
                                lngID, _
                                lng�ỰID, _
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
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Add, "����(&A)")
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, "��(&O)")
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��(&D)")
            Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)"): objControl.BeginGroup = True
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
'����:���д�ӡ,Ԥ���������EXCEL
'����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    If gstrUserName = "" Then Call GetUserInfo
    Dim objPrint As Object
    
    Set objPrint = New zlPrintLvw
    objPrint.Title.Text = IIf(InStr(lblTitle.Caption, "��Ϣ") > 0, lblTitle.Caption, lblTitle.Caption & "�����Ϣ")
   ' Set objPrint.Body.objData = lvwMain
    objPrint.BelowAppItems.Add "��ӡ�ˣ�" & gstrUserName
    objPrint.BelowAppItems.Add "��ӡʱ�䣺" & Format(zlDatabase.CurrentDate, "yyyy��MM��dd��")
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
'����:װ����Ϣ��lvwMain

    Dim rsTemp As New ADODB.Recordset
    Dim lst As ListItem
    'Dim strKey As String
    Dim strTemp As String
    Dim strICO As String
    
    '�������ͬһ��Ŀ¼�����˳�
    'If mlngIndexPre = mlngIndex Then Exit Sub
    mlngIndexPre = mlngIndex
    mstrKey = ""
    '���浱ǰ��ѡ����
    Select Case mlngIndex
        Case 0
            lblTitle.Caption = "�ݸ�"
        Case 1
            lblTitle.Caption = "�ռ���"
        Case 2
            lblTitle.Caption = "�ѷ�����Ϣ"
        Case 3
            lblTitle.Caption = "��ɾ����Ϣ"
    End Select
    rsTemp.CursorLocation = adUseClient
    
    'Get_Mail_List(��Ϣ����_In,�û�_In,��ʾ�Ѷ�_In,�ỰID)
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
            
            Set rptItem = rptRecord.AddItem(IIf(IsNull(rsTemp!����), "", rsTemp!����))
            rptItem.Caption = IIf(IsNull(rsTemp!����), "", rsTemp!����)
            'Item.Icon = 2
            
            strTemp = IIf(IsNull(rsTemp("״̬")), "0000", rsTemp("״̬"))
            If Mid(strTemp, 4, 1) <> "0" Then
                Set rptItem = rptRecord.AddItem(IIf(Mid(strTemp, 4, 1) = 1, "��", "��"))
                rptItem.Caption = IIf(Mid(strTemp, 4, 1) = 1, "��", "��")
                rptItem.Icon = IIf(Mid(strTemp, 4, 1) = 1, 5, 6)
            Else
                Set rptItem = rptRecord.AddItem("")
                rptItem.Caption = ""
            End If
            
            
            Set rptItem = rptRecord.AddItem(IIf(IsNull(rsTemp!����), "", rsTemp!����))
            rptItem.Caption = IIf(IsNull(rsTemp!����), "", rsTemp!����)
            
            If rsTemp("����") = 0 Then
                strICO = "Script"
            Else
                strICO = IIf(Mid(strTemp, 1, 1) = "1", "Read", "New") & IIf(Mid(strTemp, 2, 2) <> "00", "Reply", "")   '�Ѷ�+�Ѵ���
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
            
            Set rptItem = rptRecord.AddItem(IIf(IsNull(rsTemp!������), "", rsTemp!������))
            rptItem.Caption = IIf(IsNull(rsTemp!������), "", rsTemp!������)
            
            Set rptItem = rptRecord.AddItem(IIf(IsNull(rsTemp!�ռ���), "", Trim(rsTemp!�ռ���)))
            rptItem.Caption = IIf(IsNull(rsTemp!�ռ���), "", Trim(rsTemp!�ռ���))
            
            Set rptItem = rptRecord.AddItem(IIf(IsNull(rsTemp!ʱ��), "", Trim(rsTemp!ʱ��)))
            rptItem.Caption = IIf(IsNull(rsTemp!ʱ��), "", Trim(rsTemp!ʱ��))
            
            Set rptItem = rptRecord.AddItem(IIf(IsNull(rsTemp!�ỰID), "0", rsTemp!�ỰID))
            rptItem.Caption = IIf(IsNull(rsTemp!�ỰID), "0", rsTemp!�ỰID)

            
            rsTemp.MoveNext
        Loop
        .Populate
    End With
    'ͳһ������ʾ�ı�
    Call FillText
End Sub

Public Sub FillText()
'����:����Ϣ������װ�뵽RichText��

    Dim rsTemp As New ADODB.Recordset
    
    If rpcMain.FocusedRow Is Nothing Then
        '����ԭ�м�ֵ
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
        
        rtfContent.BackColor = IIf(IsNull(rsTemp("����ɫ")), RGB(255, 255, 255), rsTemp("����ɫ"))
        rtfContent.TextRTF = IIf(IsNull(rsTemp("����")), "", rsTemp("����"))
    Else
        rtfContent.Text = ""
        rtfContent.BackColor = RGB(255, 255, 255)
    End If
End Sub



Private Sub DeleteMessage()
'���ܣ�ɾ����ʱ����Ϣ
    'ɾ��������ǰ����Ϣ ,������ϵͳ��������ȡ��
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
        .Columns.Add 1, "����", 100, True
        .Columns.Add 2, " ", 18, False
        .Columns.Add 3, "����", 1200, True
        .Columns.Add 4, "������", 500, True
        .Columns.Add 5, "�ռ���", 800, True
        .Columns.Add 6, "ʱ��", 900, True
        .Columns.Add 7, "�ỰID", 800, True
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
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û�п���ʾ����Ϣ..."
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
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False
    
    Set cbsMain.Icons = imgPublic.Icons
    
    '�˵�����:������������
    '    ���xtpControlPopup���͵�����ID���¸�ֵ
    '-----------------------------------------------------
    cbsMain.ActiveMenuBar.Title = "�˵�"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    objMenu.ID = conMenu_FilePopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)��")
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "��ӡԤ��(&V)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Excel, "�����&Excel��")
        Set objControl = .Add(xtpControlButton, conMenu_File_SaveAs, "���Ϊ(&A)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"): objControl.BeginGroup = True
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", -1, False)
    objMenu.ID = conMenu_EditPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Add, "����(&A)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, "��(&O)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��(&D)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "��ԭ(&S)"): objControl.BeginGroup = True
        
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Reply, "��(&R)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_AllReply, "ȫ����(&L)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Transmit, "ת��(&W)")
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    objMenu.ID = conMenu_ViewPopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_ToolBar, "������(&T)")
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False
            .Add xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False
            .Add xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False
        End With
        Set objControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)")
        Set objControl = .Add(xtpControlButton, conMenu_View_PreviewWindow, "Ԥ������(&P)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_View_ShowAll, "��ʾ�Ѷ�(&E)")
        
        Set objControl = .Add(xtpControlButton, conMenu_View_Login, "��¼ʱ��δ���ʼ�����(&W)"): objControl.BeginGroup = True
        
        Set objControl = .Add(xtpControlButton, conMenu_View_Find, "���������Ϣ(&F)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)")
    End With
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    objMenu.ID = conMenu_HelpPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)")
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Help_Web, "&WEB�ϵ�" & gstr֧���̼���)
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Help_Web_Home, gstr֧���̼��� & "��ҳ(&H)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Forum, gstr֧���̼��� & "��̳(&F)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        End With
        Set objControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��"): objControl.BeginGroup = True
    End With

    '����������:������������
    '-----------------------------------------------------
    Set objBar = cbsMain.Add("������", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagHideWrap
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ")
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��")
        
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Add, "����"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, "��")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "��ԭ")
        
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Reply, "��"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Transmit, "ת��")
        
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
    End With
    For Each objControl In objBar.Controls
        objControl.Style = xtpButtonIconAndCaption
    Next
    
    '����Ŀ����:���������������Ѵ���
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyP, conMenu_File_Print '��ӡ
        
        .Add FCONTROL, vbKeyA, conMenu_Edit_Add '����
        .Add FCONTROL, vbKeyO, conMenu_Edit_Modify '��
        .Add 0, vbKeyDelete, conMenu_Edit_Delete 'ɾ��
        
        .Add FCONTROL, vbKeyAdd, conMenu_View_Expend_AllExpend 'չ��������
        .Add FCONTROL, vbKeySubtract, conMenu_View_Expend_AllCollapse '�۵�������
        
        .Add FCONTROL, vbKeyF, conMenu_View_Find '����
        
        .Add 0, vbKeyF5, conMenu_View_Refresh 'ˢ��
        
        .Add 0, vbKeyF1, conMenu_Help_Help '����
    End With
    
    '����һЩ�����Ĳ���������
    With cbsMain.Options
        .AddHiddenCommand conMenu_File_PrintSet '��ӡ����
        .AddHiddenCommand conMenu_File_Excel '�����Excel
    End With
End Sub

Private Sub Edit_Modify()
    Dim lngID As Long, lng�ỰID As Long
    
    If rpcMain.FocusedRow Is Nothing Then Exit Sub
    lngID = Val(rpcMain.Rows(rpcMain.FocusedRow.Index).Record(0).Caption)
    lng�ỰID = Val(rpcMain.Rows(rpcMain.FocusedRow.Index).Record(1).Caption)
    frmMessageEdit.OpenWindow lngID, "", lng�ỰID
    Call FillList
End Sub
