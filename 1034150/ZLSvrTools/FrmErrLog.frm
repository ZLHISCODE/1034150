VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmErrLog 
   BackColor       =   &H80000005&
   Caption         =   "������־����"
   ClientHeight    =   5790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8010
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "FrmErrLog.frx":0000
   ScaleHeight     =   5790
   ScaleWidth      =   8010
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PicFind 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3165
      Left            =   330
      ScaleHeight     =   3135
      ScaleWidth      =   3405
      TabIndex        =   6
      Top             =   1125
      Visible         =   0   'False
      Width           =   3435
      Begin VB.Frame Fra���� 
         BackColor       =   &H80000005&
         Height          =   3270
         Left            =   -30
         TabIndex        =   7
         Top             =   -120
         Width           =   3465
         Begin VB.ComboBox Cbo�û��� 
            Height          =   300
            Left            =   960
            TabIndex        =   17
            Top             =   840
            Width           =   2385
         End
         Begin VB.ComboBox Cbo�������� 
            Height          =   300
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   1200
            Width           =   2385
         End
         Begin VB.ComboBox Cbo����վ 
            Height          =   300
            Left            =   960
            TabIndex        =   15
            Top             =   480
            Width           =   2385
         End
         Begin VB.Frame FraHead 
            BackColor       =   &H80000005&
            Height          =   405
            Left            =   60
            TabIndex        =   12
            Top             =   0
            Width           =   3375
            Begin VB.PictureBox PicClose 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   200
               Left            =   3105
               Picture         =   "FrmErrLog.frx":04F9
               ScaleHeight     =   195
               ScaleWidth      =   210
               TabIndex        =   13
               TabStop         =   0   'False
               Top             =   150
               Width           =   215
            End
            Begin VB.Label LblHead 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "��������"
               ForeColor       =   &H80000008&
               Height          =   180
               Left            =   90
               TabIndex        =   14
               Top             =   160
               Width           =   720
            End
         End
         Begin VB.CommandButton cmdReset 
            Cancel          =   -1  'True
            Caption         =   "��������"
            Height          =   350
            Left            =   210
            TabIndex        =   11
            Top             =   2685
            Width           =   915
         End
         Begin VB.CommandButton Cmdȷ�� 
            Caption         =   "ȷ��(&O)"
            Height          =   350
            Left            =   1515
            TabIndex        =   10
            Top             =   2685
            Width           =   915
         End
         Begin VB.CommandButton Cmdȡ�� 
            Caption         =   "ȡ��(&C)"
            Height          =   350
            Left            =   2430
            TabIndex        =   9
            Top             =   2685
            Width           =   915
         End
         Begin MSComCtl2.DTPicker dtpDateEnd 
            Height          =   315
            Left            =   960
            TabIndex        =   8
            Top             =   2235
            Width           =   2385
            _ExtentX        =   4207
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy��MM��dd��"
            Format          =   116457475
            CurrentDate     =   37029
         End
         Begin MSComCtl2.DTPicker dtpDateStart 
            Height          =   315
            Left            =   960
            TabIndex        =   18
            Top             =   1582
            Width           =   2385
            _ExtentX        =   4207
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy��MM��dd��"
            Format          =   116457475
            CurrentDate     =   37029
         End
         Begin VB.Label Lbl�������� 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "����ʱ��"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   210
            TabIndex        =   23
            Top             =   1620
            Width           =   720
         End
         Begin VB.Label Lbl�û��� 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "�û���"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   375
            TabIndex        =   22
            Top             =   900
            Width           =   540
         End
         Begin VB.Label Lbl�������� 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��������"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   195
            TabIndex        =   21
            Top             =   1260
            Width           =   720
         End
         Begin VB.Label Lbl����վ 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "����վ"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   375
            TabIndex        =   20
            Top             =   540
            Width           =   540
         End
         Begin VB.Label lblTo 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "��"
            Height          =   180
            Left            =   960
            TabIndex        =   19
            Top             =   1965
            Width           =   180
         End
      End
   End
   Begin MSComctlLib.ImageList ImgLvw 
      Left            =   30
      Top             =   1140
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
            Picture         =   "FrmErrLog.frx":0A47
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmErrLog.frx":0BA1
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmErrLog.frx":19F3
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton CmdFind 
      Caption         =   "����(&F)"
      Height          =   350
      Left            =   1020
      TabIndex        =   3
      Top             =   630
      Width           =   1100
   End
   Begin VB.CommandButton CmdView 
      Caption         =   "�鿴(&V)"
      Height          =   350
      Left            =   4380
      TabIndex        =   1
      Top             =   630
      Width           =   1100
   End
   Begin VB.PictureBox PicMain 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   255
      ScaleHeight     =   465
      ScaleWidth      =   495
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   570
      Width           =   495
      Begin VB.Image imgMain 
         Height          =   480
         Left            =   0
         Picture         =   "FrmErrLog.frx":2845
         Top             =   0
         Width           =   480
      End
   End
   Begin MSComctlLib.ListView LvwList 
      Height          =   4155
      Left            =   315
      TabIndex        =   0
      Top             =   1125
      Width           =   7305
      _ExtentX        =   12885
      _ExtentY        =   7329
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImgLvw"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "����"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "����վ"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "�û���"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "ʱ��"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "�������"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "������Ϣ"
         Object.Width           =   5292
      EndProperty
   End
   Begin VB.CommandButton CmdDelete 
      Caption         =   "ɾ��(&D)"
      Height          =   350
      Left            =   5670
      TabIndex        =   2
      Top             =   630
      Width           =   1100
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "������־����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   195
      TabIndex        =   4
      Top             =   120
      Width           =   1440
   End
End
Attribute VB_Name = "FrmErrLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'���ڵ���listview�и�
Private Declare Function ImageList_Create Lib "COMCTL32" (ByVal MinCx As Long, ByVal MinCy As Long, ByVal Flags As Long, ByVal cInitial As Long, ByVal cGrow As Long) As Long
Private Declare Function ImageList_Destroy Lib "comctl32.dll" (ByVal himl As Long) As Long
Private Const LVM_FIRST = &H1000
Private Const LVM_SETIMAGELIST = (LVM_FIRST + 3)
Private Const LVSIL_SMALL = 1
Private Const LVM_UPDATE = (LVM_FIRST + 42)
Private hImageList As Long

Private RecLog As New ADODB.Recordset                       '��־��¼��
Private strSQL As String                                    'SQL���
Private StrDefaultSQL As String                             'ȱʡ���Ҵ�
Private StrFindSQL As String                                '���Ҵ�

Private Type MousePoint
    X As Single
    Y As Single
End Type
Private Type WindowRect
    Left As Single
    Top As Single
End Type
Private CurMousePoint As MousePoint
Private CurWindowRect As WindowRect

Private Sub CmdDelete_Click()
    Dim ItemThis As ListItem
    '��ʾ������"ɾ��ѡ��˵�"
    If LvwList.ListItems.Count = 0 Then Exit Sub
    If LvwList.SelectedItem Is Nothing Then Exit Sub
    
    For Each ItemThis In LvwList.ListItems
        If ItemThis.Selected Then Exit For
    Next
    
    If ItemThis.Selected = False Then Exit Sub
    PopupMenu frmRegMenus.TrackMenu, 2, CmdDelete.Left, CmdDelete.Top + CmdDelete.Height
End Sub

Private Sub cmdReset_Click()
    Cbo����վ.Text = ""
    Cbo�û���.Text = ""
    
    dtpDateStart.value = date
    dtpDateEnd.value = date
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Unload frmRegMenus
    SetListViewRowHeight_Destroy
End Sub

Private Sub FraHead_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    PicFind_MouseDown Button, Shift, X, Y
End Sub

Private Sub FraHead_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    PicFind_MouseMove Button, Shift, X, Y
End Sub

Private Sub Fra����_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    PicFind_MouseDown Button, Shift, X, Y
End Sub

Private Sub Fra����_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    PicFind_MouseMove Button, Shift, X, Y
End Sub

Private Sub LvwList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With LvwList
        .Sorted = False
        
        .SortKey = ColumnHeader.Index - 1
        .SortOrder = IIf(.SortOrder = 0, 1, 0)
        .Sorted = True
    End With
End Sub

Private Sub LvwList_DblClick()
    If LvwList.ListItems.Count = 0 Then Exit Sub
    If LvwList.SelectedItem Is Nothing Then Exit Sub
    
    CmdView_Click
End Sub

Private Sub LvwList_KeyDown(KeyCode As Integer, Shift As Integer)
    If LvwList.ListItems.Count = 0 Then Exit Sub
    If LvwList.SelectedItem Is Nothing Then Exit Sub
    
    If KeyCode = vbKeyDelete Then Call DeleteCurLog(Me, False): Exit Sub
    If KeyCode = vbKeyReturn Then CmdView_Click
End Sub

Private Sub LvwList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ItemThis As ListItem
    '��ʾ������"ɾ��ѡ��˵�"
    
    If Button <> 2 Then Exit Sub
    If LvwList.ListItems.Count = 0 Then Exit Sub
    If LvwList.SelectedItem Is Nothing Then Exit Sub
    
    For Each ItemThis In LvwList.ListItems
        If ItemThis.Selected Then Exit For
    Next
    
    If ItemThis.Selected = False Then Exit Sub
    PopupMenu frmRegMenus.TrackMenu, 2
End Sub

Private Sub CmdView_Click()
    Dim ItemThis As ListItem
    If LvwList.ListItems.Count = 0 Then Exit Sub
    If LvwList.SelectedItem Is Nothing Then Exit Sub
    
    Set ItemThis = LvwList.SelectedItem
    With FrmErrLogProperty
        .Txt�Ự�� = ItemThis.Tag
        .Txt����վ = ItemThis.SubItems(1)
        .Txt�û��� = ItemThis.SubItems(2)
        .Txt�������� = ItemThis
        .Txt������� = ItemThis.SubItems(4)
        .Txt����ʱ�� = ItemThis.SubItems(3)
        .Txt������Ϣ = Space(4) & ItemThis.SubItems(5)
        .Show 1
    End With
End Sub

Private Sub PicClose_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then RaisEffect PicClose, -2
End Sub

Private Sub PicClose_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then RaisEffect PicClose, 2
    
    If X > 0 And X < PicClose.Width And Y > 0 And Y < PicClose.Height Then Cmdȡ��_Click
End Sub

Private Sub PicFind_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        With CurMousePoint
            .X = X
            .Y = Y
        End With
    End If
End Sub

Private Sub PicFind_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        With CurWindowRect
            .Left = PicFind.Left + X - CurMousePoint.X
            .Top = PicFind.Top + Y - CurMousePoint.Y
            
            If .Left < ScaleLeft Then .Left = ScaleLeft
            If .Left + PicFind.Width > ScaleWidth Then .Left = ScaleWidth - PicFind.Width
            If .Top < ScaleTop Then .Top = ScaleTop
            If .Top + PicFind.Height > ScaleHeight Then .Top = ScaleHeight - PicFind.Height
        End With
        
        With PicFind
            .Move CurWindowRect.Left, CurWindowRect.Top
        End With
    End If
End Sub

Private Sub cmdFind_Click()
    With PicFind
        .Visible = True
        
        CmdFind.Enabled = .Visible Xor True
        CmdDelete.Enabled = CmdFind.Enabled
        CmdView.Enabled = CmdFind.Enabled
        LvwList.Enabled = CmdFind.Enabled
        
        Cbo����վ.SetFocus
    End With
End Sub

Private Sub Cmdȡ��_Click()
    CmdFind.Enabled = True
    CmdDelete.Enabled = (LvwList.ListItems.Count <> 0)
    CmdView.Enabled = (LvwList.ListItems.Count <> 0)
    LvwList.Enabled = CmdFind.Enabled
    LvwList.SetFocus
    PicFind.Visible = False
End Sub

Private Sub Cmdȷ��_Click()
    If GetFindSQL = False Then Exit Sub
    
    cmdDelete.Enabled = True
    CmdView.Enabled = True
    LvwList.Enabled = True
    LvwList.SetFocus
    PicFind.Visible = False
    frmMDIMain.stbThis.Panels(2).Text = "���ڲ��ң�"
    Call RefreshData
    
    CmdFind.Enabled = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{Tab}", 1
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim StrDate As String
    
    With frmRegMenus
        .Bln��־ = False
        Set .FrmObj = Me
    End With
    
    RaisEffect PicClose, 2
    
    '��ȡ���û�ѡ�������
    Call InitCons
    
    '����ȱʡ���Ҵ�(���ҵ����������־)
    StrDate = Format(CurrentDate(), "yyyy-MM-dd")
    StrDefaultSQL = " ʱ�� Between To_Date('" & StrDate & " 00:00:00','yyyy-MM-dd hh24:mi:ss') And To_date('" & StrDate & " 23:59:59','yyyy-MM-dd hh24:mi:ss')"
    
    Call RefreshData
    SetListViewRowHeight LvwList.hwnd, 15
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    With LvwList
        .Width = ScaleWidth - .Left
        .Height = ScaleHeight - .Top
    End With
    With CmdDelete
        .Left = LvwList.Width - 300 - .Width
    End With
    With CmdView
        .Left = CmdDelete.Left - 150 - .Width
    End With
    With CmdFind
        .Left = PicMain.Left + PicMain.Width + 150
    End With
End Sub

Private Function GetFindSQL() As Boolean
    Dim strDateStart As String, strDateEnd As String
    
    '--�������������Ӧ�Ĳ��Ҵ�--
    GetFindSQL = False
    StrFindSQL = ""
    'Substr(����վ, Instr(����վ, '\') + 1):���˹���վ�������������Ϊ�����ϼ��ݣ���Ϊԭ���İ汾��¼�Ĺ���վ��Ϣ��ʽΪ"������\����վ"������Ϊ"����վ"
    If Cbo����վ.Text <> "" Then StrFindSQL = StrFindSQL & IIf(StrFindSQL = "", " ", " And ") & " Substr(����վ, Instr(����վ, '\') + 1) = '" & Cbo����վ.Text & "'"
    If Cbo�û���.Text <> "" Then StrFindSQL = StrFindSQL & IIf(StrFindSQL = "", " ", " And ") & " �û��� = '" & Cbo�û���.Text & "'"
    StrFindSQL = StrFindSQL & IIf(StrFindSQL = "", " ", " And ") & " ����=" & Cbo��������.ListIndex + 1
    strDateStart = Format(dtpDateStart, "yyyy-MM-dd")
    strDateEnd = Format(dtpDateEnd, "yyyy-MM-dd")
    StrFindSQL = StrFindSQL & IIf(StrFindSQL = "", " ", " And ") & " ʱ�� Between To_Date('" & strDateStart & " 00:00:00','yyyy-MM-dd hh24:mi:ss') And To_date('" & strDateEnd & " 23:59:59','yyyy-MM-dd hh24:mi:ss')"
    
    GetFindSQL = True
End Function

Private Function InitCons()
    Call ReadInitData(Cbo����վ, Right(Cbo����վ.name, 3))
    Call ReadInitData(Cbo�û���, Right(Cbo�û���.name, 3))
    
    With Cbo��������
        .Clear
        .AddItem "�洢���̴���"
        .AddItem "������������"
        .AddItem "Ӧ�ó�������"
        .AddItem "�ͻ�����������"
        .ListIndex = 0
    End With
    
    dtpDateStart.value = CurrentDate()
    dtpDateEnd.value = CurrentDate()
End Function

Private Function ReadInitData(ByVal ConObj As Object, ByVal StrColumnName As String)
    Dim RecInit As ADODB.Recordset
    Dim strSQL As String
    '--��ȡ��ʼֵ--
On Error GoTo errHandle
    
    With ConObj
        .Clear
    End With
    
    If StrColumnName = "����վ" Then
        strSQL = "Select Distinct " & StrColumnName & " As ColumnName From Zlclients"
    Else
        strSQL = "Select Distinct " & StrColumnName & " As ColumnName From �ϻ���Ա��"
    End If
    Set RecInit = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)
    With RecInit
        Do While Not .EOF
            If Not IsNull(!ColumnName) Then
                ConObj.AddItem !ColumnName
            End If
            .MoveNext
        Loop
    End With
    Exit Function
errHandle:
    MsgBox Err.Description, vbCritical, Me.Caption
End Function

Private Function RefreshData()
    '--���ݲ��Ҵ�,���»�ȡ����--
On Error GoTo errHandle
    Set RecLog = OpenCursor(gcnOracle, "ZLTOOLS.B_Runmana.Get_Log", "������־", IIf(StrFindSQL = "", StrDefaultSQL, StrFindSQL))
   
    Call LoadData
    Exit Function
errHandle:
    MsgBox Err.Description, vbCritical, Me.Caption
End Function

Private Function LoadData()
    Dim lngCount As Long
    Dim ItemThis As ListItem
    '--װ��--
On Error GoTo errHandle
    LvwList.ListItems.Clear
    With RecLog
        Do While Not .EOF
            Set ItemThis = LvwList.ListItems.Add(, "K_" & .AbsolutePosition, !��������, , 3)
            With ItemThis
                .SubItems(1) = IIf(IsNull(RecLog!����վ), "", Mid(RecLog!����վ, InStr(RecLog!����վ, "\") + 1))
                .SubItems(2) = IIf(IsNull(RecLog!�û���), "", RecLog!�û���)
                .SubItems(3) = IIf(IsNull(RecLog!ʱ��), "", RecLog!ʱ��)
                .SubItems(4) = IIf(IsNull(RecLog!�������), "", RecLog!�������)
                .SubItems(5) = IIf(IsNull(RecLog!������Ϣ), "", RecLog!������Ϣ)
                .Tag = RecLog!�Ự��
            End With
            .MoveNext
        Loop
    End With
    With LvwList
        If .ListItems.Count <> 0 Then
            .ListItems(1).Selected = True
            .SelectedItem.Selected = True
        End If
        
        CmdView.Enabled = (.ListItems.Count <> 0)
        CmdDelete.Enabled = (.ListItems.Count <> 0)
    End With
    If CmdFind.Enabled = False Then
        frmMDIMain.stbThis.Panels(2).Text = "������ϣ������ҵ�" & RecLog.RecordCount & "�����ݣ�"
    End If
    Exit Function
errHandle:
    MsgBox Err.Description, vbCritical, Me.Caption
End Function

Public Function SupportPrint() As Boolean
'���ر������Ƿ�֧�ִ�ӡ���������ڵ���
    SupportPrint = True
End Function

Public Sub SubPrint(ByVal bytMode As Byte)
'�������ڵ��ã�ʵ�־���Ĵ�ӡ����
'���û�пɴ�ӡ�ģ�������һ���յĽӿ�

'����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    Dim objPrint As zlPrintLvw
    
    Set objPrint = New zlPrintLvw
    objPrint.Title.Text = "������־"
    Set objPrint.Body.objData = LvwList
    objPrint.BelowAppItems.Add "��ӡʱ�䣺" & Format(CurrentDate, "yyyy��MM��dd��")
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

'����listview�и�
Private Sub SetListViewRowHeight(ByVal listViewHwnd As Long, ByVal rowHeight As Long)
    Call SetListViewRowHeight_Destroy
    hImageList = ImageList_Create(1, rowHeight, 1, 0, 0)
    SendMessage listViewHwnd, LVM_SETIMAGELIST, LVSIL_SMALL, ByVal hImageList
    SendMessage listViewHwnd, LVM_UPDATE, 0, ByVal 0
End Sub

Private Sub SetListViewRowHeight_Destroy()
    If hImageList <> 0 Then ImageList_Destroy hImageList
End Sub
