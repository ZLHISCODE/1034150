VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Begin VB.Form frmMessageEdit 
   AutoRedraw      =   -1  'True
   Caption         =   "��Ϣ"
   ClientHeight    =   6435
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   9510
   Icon            =   "frmMessageEdit.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6435
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Tag             =   "�ɱ仯��"
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
      Caption         =   "�ռ���(&R)"
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
      Caption         =   "����(&S)��"
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
      Caption         =   "�ռ��ˣ�"
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
      Caption         =   "����ʱ�䣺"
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
      Caption         =   "�����ˣ�"
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

'���浱ǰ����༭����ϢID
Public mstrID As String     'strID ��ϢID�����Ϊ�գ���ʾ������Ϣ
Public mstrByID As String   'strByID Ҫ����ת�����𸴵�ԭʼID
Public mlng���� As Long     '��ǰ�û���Ӧ����Ϣ����
Public mlngMode As Long     '�򿪷�ʽ��1-�𸴣�2-ȫ���𸴣�3-ת��

Dim mstr�ỰID As String
Dim mblnSend As Boolean     '׼������
Dim mblnDelete As Boolean   '�Ƿ��Ѿ�ɾ��

Private mrsUser As ADODB.Recordset '����ռ�������,�û���,��Ա���ʵļ�¼��

Dim mblnChange As Boolean
'�������ڲ��ҵ�
Public mblnCase As Boolean
Public mblnBegin As Boolean
Public mstrFind As String
Dim mblnHigh As Boolean, mblnLow As Boolean  '��Ҫ��
Dim mblnSetState As Boolean '�Ƿ��ʼ������


Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)

    On Error Resume Next
    Dim objControl As CommandBarControl
    Dim i As Integer
    Dim lngID As Long, lng�ỰID As Long
    Dim lst As ListItem
    Dim objMenu As CommandBarPopup
     
    Select Case Control.ID
    Case conMenu_File_Send
        '������Ϣ
       
        If mblnSend = False Then Exit Sub
        
        If SaveMessage(True) = True Then
            '����������
            '�����ж��Ƿ���ʵ���ʾλ��
            With frmMessageManager
                On Error Resume Next
                If .mlngIndex = 1 Then
                    Unload Me
                    Exit Sub
                End If
                If .mlngIndex = 0 Or .mlngIndex = 3 Then
                    'ɾ���ݸ���Ϣ
                    Unload Me
                    Exit Sub
                End If
                '�����ѷ�����Ϣ
              End With
              Unload Me
              '����ͼ��
              Call frmMessageRead.UpdateNotify
        End If
    Case conMenu_File_Save
        '����
        If SaveMessage(False) = True Then
            On Error Resume Next
            frmMessageManager.FillList
        End If
    Case conMenu_View_ToolBar_Text
            '��ť����
            For i = 2 To cbsMain.Count
                For Each objControl In Me.cbsMain(i).Controls
                    objControl.Style = IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
                Next
            Next
            Me.cbsMain.RecalcLayout
            Call Form_Resize
    Case conMenu_File_SaveAs
        '���
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
        
    Case conMenu_Edit_Cut
        '���м���
        If rtfContent.SelLength = 0 Then Exit Sub
        
        Clipboard.SetText rtfContent.SelRTF, vbCFRTF
        rtfContent.SelText = ""
    Case conMenu_Edit_Copy
        '�����ı�
        If rtfContent.SelLength = 0 Then Exit Sub
        Clipboard.SetText rtfContent.SelRTF, vbCFRTF
    Case conMenu_Edit_plaster
        'ճ���ı�
        If Clipboard.GetFormat(vbCFText) = True Then
            If Clipboard.GetText(vbCFRTF) <> "" Then
                rtfContent.SelRTF = Clipboard.GetText(vbCFRTF)
            Else
                rtfContent.SelRTF = Clipboard.GetText(vbCFText)
            End If
        End If
    Case conMenu_Edit_Clear
        '�����ѡ���ı�
            rtfContent.SelText = ""
    Case conMenu_Edit_CheckAll
        'ȫ��ѡ��
        rtfContent.SelStart = 0
        rtfContent.SelLength = Len(rtfContent.Text)
    Case conMenu_View_Find
        '���ҵ�һ��
        Set frmMessageFind.frmMain = Me
        frmMessageFind.Show vbModal, Me
    Case conMenu_View_FindNext
        '��������
        Call FindText
    Case conMenu_Format_ForeColor
        'ѡ��������ɫ
        Call ColorForeColor_pOK
    Case conMenu_Format_FillColor
        'ѡ�񱳾���ɫ
        Call ColorFillColor_pOK
    Case conMenu_Format_Sig
        '���ı�������Ŀ����
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_FormatPopup)
        objMenu.CommandBar.FindControl(, conMenu_Format_Sig).Checked = Not objMenu.CommandBar.FindControl(, conMenu_Format_Sig).Checked
        rtfContent.SelBullet = objMenu.CommandBar.FindControl(, conMenu_Format_Sig).Checked
    Case conMenu_Format_Left
        '�ı�����
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
        '����
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_FormatPopup)
        objMenu.CommandBar.FindControl(, conMenu_Format_Left).Checked = False
        objMenu.CommandBar.FindControl(, conMenu_Format_Center).Checked = False
        objMenu.CommandBar.FindControl(, conMenu_Format_Right).Checked = True
        rtfContent.SelAlignment = 1
    Case conMenu_Format_Decrease
        '����������
        i = IIf(rtfContent.SelIndent > 360, rtfContent.SelIndent, 360)
        rtfContent.SelIndent = i - 360
    Case conMenu_Format_Increase
        ' ����������
        i = IIf(rtfContent.SelIndent < rtfContent.Width - 1000, rtfContent.SelIndent, rtfContent.Width - 1000)
        rtfContent.SelIndent = i + 360
    Case conMenu_Edit_Reply
        '��
        frmMessageEdit.OpenWindow "", mstrID, mlng����, 1
        Unload Me
    Case conMenu_Edit_AllReply
        'ȫ����
        frmMessageEdit.OpenWindow "", mstrID, mlng����, 2
        Unload Me
    Case conMenu_Edit_Transmit
        'ת����Ϣ
        frmMessageEdit.OpenWindow "", mstrID, mlng����, 3
        Unload Me
    Case conMenu_Action_Hight
        '������Ϣ����Ҫ�ԣ���
        mblnHigh = Not mblnHigh
        
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ActionPopup)
        objMenu.CommandBar.FindControl(, conMenu_Action_Hight).Checked = mblnHigh
        cbsMain.FindControl(, conMenu_Action_Hight).Checked = mblnHigh
        
        objMenu.CommandBar.FindControl(, conMenu_Action_Low).Checked = False  '�˵�
        cbsMain.FindControl(, conMenu_Action_Low).Checked = False '������
        mblnLow = False
    Case conMenu_Action_Low
        '������Ϣ����Ҫ�ԣ���
        mblnLow = Not mblnLow
        
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ActionPopup)
        objMenu.CommandBar.FindControl(, conMenu_Action_Low).Checked = mblnLow
        cbsMain.FindControl(, conMenu_Action_Low).Checked = mblnLow
        
        objMenu.CommandBar.FindControl(, conMenu_Action_Hight).Checked = False '�˵�
        cbsMain.FindControl(, conMenu_Action_Hight).Checked = False '������
        mblnHigh = False
    Case conMenu_Format_Font
        '����
        If Control.Type = xtpControlComboBox Then
            rtfContent.SelFontName = Control.Text
        End If
        
    Case conMenu_Format_SIZE
        '�ֺ�
        If Control.Type = xtpControlComboBox Then
            rtfContent.SelFontSize = Control.Text
        End If
    Case conMenu_FORMAT_BOLD
        '����
        rtfContent.SelBold = Not rtfContent.SelBold
    Case conMenu_FORMAT_ITALIC
        'б��
        rtfContent.SelItalic = Not rtfContent.SelItalic
    Case conMenu_FORMAT_UNDERLINE
        '�»���
        rtfContent.SelUnderline = Not rtfContent.SelUnderline
    Case conMenu_Help_Help
        Call ShowHelp(App.ProductName, Me.hWnd, "ZL9AppTool\" & Me.Name, 0)
    Case conMenu_Help_Web_Mail
        '���ͷ���
         Call zlMailTo(Me.hWnd)
    Case conMenu_Help_Web_Home
        '��ҳ
        Call zlHomePage(Me.hWnd)
    Case conMenu_Help_Web_Forum
        '��̳
        Call zlWebForum(Me.hWnd)
    Case conMenu_Help_About
        '����
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

    '���������ؼ�Resize����
    
    '������һ�и��ؼ���λ��
    
    lbl(0).Top = lngTop + 30
    lbl(0).Left = 60
    lbl(1).Top = lbl(0).Top
    lbl(1).Left = ScaleWidth - lbl(1).Width - 60
    '�����ڶ��и��ؼ���λ��
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
    '���������и��ؼ���λ��
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
        
    '�����༭�ؼ���λ��
    rtfContent.Left = lbl(0).Left
    rtfContent.Top = txtSubject.Top + txtSubject.Height + 60
    rtfContent.Width = ScaleWidth - rtfContent.Left - 60
    rtfContent.Height = lngBottom - rtfContent.Top - 30
    
       Me.Refresh
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)

    '�������ÿɼ���
    If mblnSetState Then
        If InStr(frmMessageManager.mstrPrivs, "������Ϣ") = 0 Then
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
            '����Ϣ
            Control.Enabled = mblnSend
        Case conMenu_FormatPopup
            Control.Visible = mblnSend
        Case conMenu_Edit_Reply, conMenu_Edit_AllReply, conMenu_Edit_Transmit
            '����Ϣ
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
        If MsgBox("�����������˳��Ļ������е��޸Ķ�������Ч��" & vbCrLf & "�Ƿ�ȷ���˳���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = 1
            Exit Sub
        End If
    End If
    SaveWinState Me, App.ProductName
    mstr�ỰID = ""
    mblnSend = False
    mblnDelete = False

    mblnChange = False
    mblnHigh = False
    mblnLow = False
    mblnSetState = False
    Set mrsUser = Nothing
End Sub

Private Sub cmdReceive_Click()
    Dim str�ռ���  As String
    Dim rsUser As ADODB.Recordset
    
    Set rsUser = mrsUser
    str�ռ��� = txtReceive.Text
    If frmSelectReceiver.Get�ռ���(str�ռ���, rsUser) = True Then
        
        Set mrsUser = rsUser
        
        txtReceive.Text = str�ռ���
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

Public Sub OpenWindow(ByVal strID As String, ByVal strByID As String, Optional ByVal lng���� As Long, Optional ByVal lngMode As Long)
'���ܣ����ݲ�������ʾ��Ϣ�༭����
'strID ��ϢID�����Ϊ�գ���ʾ������Ϣ
'strByID Ҫ����ת�����𸴵�ԭʼID
'lngMode  ����ʽ��1-�𸴣�2-ȫ���𸴣�3-ת��
    Dim frmMessage As frmMessageEdit
    Dim frmTemp As Form
    
    '���Ҹ���Ϣ�Ƿ��Ѿ����˱༭����
    If strID <> "" Then
        For Each frmTemp In Forms
            If frmTemp.Name = "frmMessageEdit" Then
                If frmTemp.mstrID = strID And frmTemp.mlng���� = lng���� Then
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
        frmMessage.mlng���� = lng����
        frmMessage.mlngMode = lngMode
    End If
    frmMessage.Show , gfrmMain
    
End Sub

Private Sub SetState()
    '�������ÿɼ���
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
        If cmdBar.Title = "��ʽ������" Then
            cmdBar.Visible = mblnSend
        End If
    Next
    'CoolBar1.Bands("three").Visible = mblnSend
    '�������ÿ�����
    rtfContent.Locked = Not mblnSend
    '���ù������ĳ�ʼ������
    mblnSetState = True
End Sub

Private Sub LoadMessage()
'װ������
    Dim lngID As Long
    Dim rsTemp As New ADODB.Recordset
    Dim str������ As String
    Dim str�ռ��� As String
    Dim str���� As String
    
    lngID = Val(IIf(mstrID = "", mstrByID, mstrID)) '
    If lngID = 0 Then
        mblnDelete = False
        mblnSend = True
        Exit Sub '��ȫ�µģ��ò���װ��
    End If
    
    '�õ��ʼ�����
    rsTemp.CursorLocation = adUseClient
    Set rsTemp = zlDatabase.OpenCursor(Me.Caption, "zlTools", "b_ComFunc.Get_Zlmessage", lngID, mlng����, gstrDbUser)
    If rsTemp.RecordCount <= 0 Then Exit Sub
    mstr�ỰID = rsTemp("�ỰID")
    mblnDelete = IIf(rsTemp("ɾ��") = 1, True, False)
    mblnSend = (mstrID = "") Or (mlng���� <> 2 And mlng���� <> 1) '���ʼ���δ�����ʼ�
    
    lbl(1).Caption = "����ʱ�䣺" & IIf(IsNull(rsTemp("ʱ��")), "", Format(rsTemp("ʱ��"), "yyyy-MM-dd HH:mm:ss"))
    str������ = IIf(IsNull(rsTemp("������")), "", rsTemp("������"))
    str�ռ��� = IIf(IsNull(rsTemp("�ռ���")), "", rsTemp("�ռ���"))
    txtSubject.Text = IIf(IsNull(rsTemp("����")), "", rsTemp("����"))
    rtfContent.TextRTF = IIf(IsNull(rsTemp("����")), "", rsTemp("����"))
    rtfContent.BackColor = IIf(IsNull(rsTemp("����ɫ")), RGB(255, 255, 255), rsTemp("����ɫ"))
    
    '�õ��ʵݵ�ַ
    Set rsTemp = zlDatabase.OpenCursor(Me.Caption, "zlTools", "b_ComFunc.Get_Zlmsgstate", lngID)
    
    '�����յ��ռ��˼�¼��
    Dim strFild As String
    strFild = "�û���,Varchar2,30;����,varchar2,30;�ռ���,varchar2,30"
    
    '���½�������
    
    
    Select Case mlngMode
        Case 1 '��
            Set mrsUser = NewClientRecord(strFild)
            rsTemp.Filter = "����=0 or ����=1"
            mrsUser.AddNew
            mrsUser.Fields("����") = rsTemp("���")
            mrsUser.Fields("�û���") = rsTemp("�û�")
            
            txtReceive.Text = mrsUser.Fields("����")
            If Left(txtSubject.Text, 3) <> "�𸴣�" Then
                txtSubject.Text = "�𸴣�" & txtSubject.Text
            End If
        Case 2 'ȫ����
            Set mrsUser = NewClientRecord(strFild)
            str���� = ""
            If str�ռ��� = "������Ա" Or str�ռ��� = "��������Ա" Or str�ռ��� = "��������Ա" Then
                txtReceive.Text = str�ռ���
            ElseIf InStr(str�ռ���, "]") > 0 And InStr(str�ռ���, "[") > 0 Then
                txtReceive.Text = str�ռ���
            Else
                If str������ <> str�ռ��� And InStr(str�ռ���, str������ & ",") = 0 And InStr(str�ռ���, "," & str������) = 0 Then
                    txtReceive.Text = str������ & "," & str�ռ���
                Else
                    txtReceive.Text = str�ռ���
                End If
            End If
            If Left(txtSubject.Text, 3) <> "�𸴣�" Then
                txtSubject.Text = "�𸴣�" & txtSubject.Text
            End If
            rsTemp.Filter = "����=3 or ����=2"
            Do Until rsTemp.EOF
                mrsUser.AddNew
                mrsUser.Fields("����") = rsTemp("���")
                mrsUser.Fields("�û���") = rsTemp("�û�")
                str���� = str���� & rsTemp("���") & ","
                rsTemp.MoveNext
            Loop
            If str������ <> str���� And InStr(str����, str������ & ",") = 0 And InStr(str����, "," & str������) = 0 Then
            
                mrsUser.AddNew
                mrsUser.Fields("����") = str������
                mrsUser.Fields("�û���") = gstrDbUser
                
            End If
        Case 3 'ת��
            Set mrsUser = NewClientRecord(strFild)
            txtReceive.Text = ""
            If Left(txtSubject.Text, 3) <> "ת����" Then
                txtSubject.Text = "ת����" & txtSubject.Text
            End If
        Case Else
            Set mrsUser = NewClientRecord(strFild)
            txtReceive = str�ռ���
            rsTemp.Filter = "����=3 or ����=2"
            
            Do Until rsTemp.EOF
                mrsUser.AddNew
                mrsUser.Fields("����") = str������
                mrsUser.Fields("�û���") = gstrDbUser
            
                rsTemp.MoveNext
            Loop
    End Select
    
    
    lbl(0).Caption = "�����ˣ�     " & str������
    Me.Caption = "��Ϣ  " & IIf(txtSubject.Text = "", "", "-  " & txtSubject.Text)
    If mlngMode <> 0 Then
        '��ԭ����������
        With rtfContent
            .SelStart = 0
            .SelText = vbCrLf & "----------ԭʼ��Ϣ------------" & vbCrLf
            .SelStart = 2
            .SelLength = Len("----------ԭʼ��Ϣ------------")
            .SelFontName = "����"
            .SelFontSize = 9
            .SelBold = False
            .SelItalic = False
            .SelColor = 0
            
            .SelLength = Len(.Text)
            .SelIndent = 720
            
            .SelStart = 0
            .SelLength = 2
            .SelColor = RGB(0, 0, 255)
            .SelFontName = "����"
            .SelFontSize = 9
            .SelBold = False
            .SelItalic = False
            .SelLength = 0
        End With
    End If
    'Update_Zlmsgstate_Idtntify(���,��ϢID,����,�û�)
    Call zlDatabase.OpenCursor(Me.Caption, "zltools", "b_ComFunc.Update_Zlmsgstate_Idtntify", _
                                gstrUserName, _
                                lngID, _
                                mlng����, _
                                gstrDbUser)
    
    Call frmMessageRead.UpdateNotify
    Dim lst As ListItem
    On Error Resume Next
    '�ĳ��Ѷ�
'    Set lst = frmMessageManager.lvwMain.ListItems("C" & mlng���� & lngID)
    If Not lst Is Nothing Then
        If lst.Icon <> "Script" Then '���ǲݸ����Ϣ�ſ��ܸ���
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
    Dim lng�ỰID As Long
    Dim lngCount As Long
    
    Dim strUsers As String '�洢�Ѿ������˵��û���
    Dim rsTemp As ADODB.Recordset
    'û�޸ģ���ֻ�Ǳ���
    
    If blnSend = True And txtReceive.Text = "" And mrsUser Is Nothing Then
        MsgBox "��ѡ���ռ��ˡ�", vbExclamation, gstrSysName
        cmdReceive.SetFocus
        Exit Function
    End If
    If zlCommFun.StrIsValid(txtReceive.Text, txtReceive.MaxLength, cmdReceive.hWnd, "�ռ���") = False Then
        Exit Function
    End If
    
    If zlCommFun.StrIsValid(txtSubject.Text, txtSubject.MaxLength, txtSubject.hWnd, "����") = False Then
        Exit Function
    End If
    If LenB(StrConv(rtfContent.TextRTF, vbFromUnicode)) > 4000 Then
        MsgBox "���ĵ��ַ�̫�࣬���߸�ʽ̫�����ˡ�", vbExclamation, gstrSysName
        rtfContent.SetFocus
        Exit Function
    End If
    
    On Error GoTo errHandle
    gcnOracle.BeginTrans
    '��������
    'Save_Zlmessage(ID,�ỰID,������,�ռ���,����,����,����ɫ)
    Set rsTemp = zlDatabase.OpenCursor(Me.Caption, "zltools", "b_ComFunc.Save_Zlmessage", _
                                Val(mstrID), _
                                Val(mstr�ỰID), _
                                gstrUserName, _
                                txtReceive.Text, _
                                txtSubject.Text, _
                                rtfContent.TextRTF, _
                                rtfContent.BackColor)
    
    If rsTemp.RecordCount > 0 Then
        lngID = zlCommFun.NVL(rsTemp.Fields(0), 0)
        lng�ỰID = zlCommFun.NVL(rsTemp.Fields(1), 0)
        '����ӱ�
    
        'Insert_Zlmsgstate (��ϢID,����,�û�,���,ɾ��,״̬)
         Call zlDatabase.OpenCursor(Me.Caption, "zltools", "b_ComFunc.Insert_Zlmsgstate", _
                                     lngID, _
                                     Val(IIf(blnSend, "1", "0")), _
                                     gstrDbUser, _
                                     gstrUserName, _
                                     Val(IIf(mblnDelete = True And blnSend = False, "1", "0")), _
                                     CStr(IIf(blnSend, "1", "0") & "00" & IIf(mblnHigh = True, "1", IIf(mblnLow = True, "2", "0"))))
        ''''' �����ռ��˵ļ�¼
         If Not mrsUser Is Nothing Then
             If mrsUser.State = adStateOpen Then
                 If mrsUser.RecordCount > 0 Then mrsUser.MoveFirst
                 Do Until mrsUser.EOF
                    Call zlDatabase.OpenCursor(Me.Caption, "zltools", "b_ComFunc.Insert_Zlmsgstate", _
                                               lngID, _
                                               Val(IIf(blnSend, "2", "3")), _
                                               IIf(IsNull(mrsUser.Fields("�û���")), "", mrsUser.Fields("�û���")), _
                                               IIf(IsNull(mrsUser.Fields("����")), "", mrsUser.Fields("����")), _
                                               0, _
                                               CStr("000" & IIf(mblnHigh = True, "1", IIf(mblnLow = True, "2", "0"))))
                    
                    mrsUser.MoveNext
                Loop
            End If
        End If
        
        '���Ϊԭ�����ϴ𸴻�ת����־

        'Update_Zlmsgstate_State(ģʽ,��ϢID,����,�û�)
        Call zlDatabase.OpenCursor(Me.Caption, "zltools", "b_ComFunc.Update_Zlmsgstate_State", _
                                    mlngMode, _
                                    Val(mstrByID), _
                                    mlng����, _
                                    gstrDbUser)
    End If

    gcnOracle.CommitTrans
    mstrID = lngID
    mstr�ỰID = lng�ỰID
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
    '�����������Ϊ�գ�ֱ���˳�
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
        MsgBox "���ҽ�����û�ҵ���" & mstrFind & "��", vbInformation, gstrSysName
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
    
    Dim objCustControl As CommandBarControlCustom       '�Զ���ؼ�

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
        .LargeIcons = False
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
        Set objControl = .Add(xtpControlButton, conMenu_File_Send, "����(&E)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Save, "����(&S)")
        Set objControl = .Add(xtpControlButton, conMenu_File_SaveAs, "���Ϊ(&A)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"): objControl.BeginGroup = True
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", -1, False)
    objMenu.ID = conMenu_EditPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Cut, "����(&T)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Copy, "����(&C)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_plaster, "ճ��(&P)")
        
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Clear, "���(&L)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_CheckAll, "ȫѡ(&A)")
        
        Set objControl = .Add(xtpControlButton, conMenu_View_Find, "����(&F)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_View_FindNext, "������һ��(&N)")
       
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FormatPopup, "��ʽ(&F)", -1, False)
    objMenu.ID = conMenu_FormatPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Format_Sig, "��Ŀ����(&S)")
        
        Set objControl = .Add(xtpControlButton, conMenu_Format_Left, "�������(&L)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Format_Center, "���ж���(&C)")
        Set objControl = .Add(xtpControlButton, conMenu_Format_Right, "���Ҷ���(&R)")
        Set objControl = .Add(xtpControlButton, conMenu_Format_Decrease, "����������(&D)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Format_Increase, "����������(&I)")
       
    End With
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    objMenu.ID = conMenu_ViewPopup
    With objMenu.CommandBar.Controls
'        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_ToolBar, "������(&T)")
'        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False
'        End With
    End With
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ActionPopup, "����(&A)", -1, False)
    objMenu.ID = conMenu_ActionPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Reply, "��(&R)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_AllReply, "ȫ����(&L)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Transmit, "ת��(&W)")
        
        Set objControl = .Add(xtpControlButton, conMenu_Action_Hight, "��Ҫ�Ը�(&H)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Action_Low, "��Ҫ�Ե�(&O)")
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
    Set objBar = cbsMain.Add("�½�������", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagHideWrap
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_Send, "����")
        Set objControl = .Add(xtpControlButton, conMenu_File_Save, "����")
        
        Set objControl = .Add(xtpControlButton, conMenu_Action_Hight, "��Ҫ"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Action_Low, "��Ҫ")
        '----
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Reply, "��"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_AllReply, "ȫ����")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Transmit, "ת��")
        
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
        
    End With
    
    For Each objControl In objBar.Controls
        objControl.Style = xtpButtonIconAndCaption
    Next
    
    '
    Set objBar = cbsMain.Add("��ʽ������", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagHideWrap
    
    With objBar.Controls
        .Add xtpControlButton, conMenu_Edit_Cut, "����", -1, False
        .Add xtpControlButton, conMenu_Edit_Copy, "����", -1, False
        .Add xtpControlButton, conMenu_Edit_plaster, "ճ��", -1, False
        
        Set objControl = .Add(xtpControlComboBox, conMenu_Format_Font, "����")
        objControl.BeginGroup = True
        objControl.Width = 140
        objControl.DropDownListStyle = True
        'װ��������ֺ�
        Dim lngCount As Long, lngDefault As Long
        
        For lngCount = 0 To Screen.FontCount - 1
            objControl.AddItem Screen.Fonts(lngCount)
            If objControl.List(lngCount) = "����" Then
                lngDefault = lngCount
            End If
        Next
        If lngDefault > 0 Then objControl.ListIndex = lngDefault
        
        Dim ControlComboSize As CommandBarComboBox
        Set ControlComboSize = .Add(xtpControlComboBox, conMenu_Format_SIZE, "�ֺ�")
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
        
        Set objControl = .Add(xtpControlButton, conMenu_FORMAT_BOLD, "�Ӵ�", -1, False)
        objControl.BeginGroup = True
        .Add xtpControlButton, conMenu_FORMAT_ITALIC, "б��", -1, False
        .Add xtpControlButton, conMenu_FORMAT_UNDERLINE, "�»���", -1, False
        
        Set objControl = .Add(xtpControlButton, conMenu_Format_Left, "�����", -1, False): objControl.BeginGroup = True
        .Add xtpControlButton, conMenu_Format_Center, "����", -1, False
        .Add xtpControlButton, conMenu_Format_Right, "�Ҷ���", -1, False
        
        Set objPopup = .Add(xtpControlSplitButtonPopup, conMenu_Format_ForeColor, "������ɫ")
        Set objCustControl = objPopup.CommandBar.Controls.Add(xtpControlCustom, 0, "")
        objCustControl.Handle = ColorForeColor.hWnd
        objControl.BeginGroup = True
        
        Set objPopup = .Add(xtpControlSplitButtonPopup, conMenu_Format_FillColor, "����ɫ")
        Set objCustControl = objPopup.CommandBar.Controls.Add(xtpControlCustom, 0, "")
        objCustControl.Handle = ColorFillColor.hWnd
        
        Set objControl = .Add(xtpControlButton, conMenu_View_Find, "&����"): objControl.BeginGroup = True
        .Add xtpControlButton, conMenu_Format_Sig, "&��Ŀ����"
        .Add xtpControlButton, conMenu_Format_Decrease, "&����������"
        .Add xtpControlButton, conMenu_Format_Increase, "&����������"
        

    End With

    For Each objControl In objBar.Controls
        objControl.Style = xtpButtonIcon
    Next
    '����Ŀ����:���������������Ѵ���
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyF, conMenu_View_Find '����
        .Add 0, vbKeyF1, conMenu_Help_Help '����
    End With
    
    '����һЩ�����Ĳ���������
    With cbsMain.Options
        .AddHiddenCommand conMenu_File_PrintSet '��ӡ����
        .AddHiddenCommand conMenu_File_Excel '�����Excel
    End With
    cbsMain.RecalcLayout
    
    '��ʼ����ɫ
     
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

    '����ͼ��
    imgColor.ListImages.Remove imgColor.ListImages(Key).Index
    imgColor.ListImages.Add 1, Key, ctlPictureBox.Image

    '���� Tag ����
    imgColor.ListImages(1).Tag = ID
    cbsMain.AddImageList imgColor
    cbsMain.RecalcLayout
    
    Me.Controls.Remove ctlPictureBox
    Set ctlPictureBox = Nothing
End Sub

Private Sub ColorForeColor_GotFocus()
    ColorForeColor.Tag = "Focused"
End Sub
