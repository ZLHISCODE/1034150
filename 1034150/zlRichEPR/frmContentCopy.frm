VERSION 5.00
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmContentCopy 
   Caption         =   "ר�ø���"
   ClientHeight    =   9825
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14190
   Icon            =   "frmContentCopy.frx":0000
   LinkTopic       =   "ר�ø���"
   ScaleHeight     =   9825
   ScaleWidth      =   14190
   StartUpPosition =   1  '����������
   Begin XtremeReportControl.ReportControl RptThis 
      Height          =   4440
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   2895
      _Version        =   589884
      _ExtentX        =   5106
      _ExtentY        =   7832
      _StockProps     =   0
      BorderStyle     =   2
      MultipleSelection=   0   'False
   End
   Begin VB.Frame fraThis 
      Height          =   700
      Left            =   5280
      TabIndex        =   1
      Top             =   4200
      Width           =   3135
      Begin VB.CommandButton cmdCancle 
         Cancel          =   -1  'True
         Caption         =   "ȡ��"
         Height          =   360
         Left            =   2040
         TabIndex        =   3
         Top             =   240
         Width           =   990
      End
      Begin VB.CommandButton cmdInsert 
         Caption         =   "����"
         Height          =   360
         Left            =   960
         TabIndex        =   2
         Top             =   240
         Width           =   990
      End
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   5760
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContentCopy.frx":6852
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContentCopy.frx":6DEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContentCopy.frx":7386
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContentCopy.frx":7920
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContentCopy.frx":7EBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContentCopy.frx":8454
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmContentCopy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mfrmContent As frmDockEPRContent
Attribute mfrmContent.VB_VarHelpID = -1
Private mblnOk As Boolean
Private Enum mCol
    ID = 0: ��ҳID: ����ID: ��������: ���ʱ��: ��������: �༭��ʽ: ������Դ: ��Ժ����: ����ʱ��:
End Enum
Public Function ShowMe(ByVal frmParent As Object, ByVal patiantID As String, ByVal patiantPageId As String, ByVal lngPatiFrom As Long) As Boolean
    mblnOk = False
    Call Me.zlRefresh(patiantID, patiantPageId, lngPatiFrom)
    Me.Show 1, frmParent
    ShowMe = mblnOk
End Function

Private Sub cmdCancle_Click()
    Unload Me
End Sub

Private Sub cmdInsert_Click()
    If Not mfrmContent Is Nothing Then
        If mfrmContent.edtThis.SelText <> "" Then
            mfrmContent.edtThis.Copy    '�������ı���ʽ�������������򣨷ŵ������壩
            mblnOk = True
            Unload Me
        Else
            mblnOk = False
            MsgBox "����ѡ�����ݣ�", vbOKOnly + vbInformation, gstrSysName
        End If
    End If
        
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
Select Case Item.ID
        Case 1
        Item.Handle = Me.RptThis.hWnd
        Case 2
        Item.Handle = mfrmContent.hWnd
        Case 3
        Item.Handle = Me.fraThis.hWnd
End Select
End Sub

Private Sub dkpMan_Resize()
    Me.cmdInsert.Move Me.fraThis.Width - Me.cmdInsert.Width - Me.cmdCancle.Width - 200, 160
    Me.cmdCancle.Move Me.fraThis.Width - Me.cmdCancle.Width - 200, 160
End Sub


Private Sub Form_Load()
    Dim rptCol As ReportColumn
    Dim panList As Pane, panContent As Pane, panNew As Pane
    Me.dkpMan.Options.ThemedFloatingFrames = True
    Me.dkpMan.Options.HideClient = True
    Set panList = dkpMan.CreatePane(1, 200, 100, DockLeftOf, Nothing)
    panList.MaxTrackSize.Width = 270
    panList.Title = "�����б�"
    panList.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    Set mfrmContent = New frmDockEPRContent
    mfrmContent.mIsShowAnnex = True
    Set panContent = dkpMan.CreatePane(2, 200, 300, DockRightOf, Nothing)
    panContent.Title = "��������"
    panContent.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    Set panNew = dkpMan.CreatePane(3, 100, 40, DockBottomOf, panContent)
    panNew.MaxTrackSize.Height = 40
    panNew.Options = PaneNoFloatable Or PaneNoHideable
    panNew.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    With Me.RptThis
        Set rptCol = .Columns.Add(mCol.ID, "ID", 110, True): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.��ҳID, "��ҳID", 110, True): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.����ID, "����ID", 110, True): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.��������, "��������", 20, True): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.���ʱ��, "���ʱ��", 120, True): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.��������, "��������", 120, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.�༭��ʽ, "�༭��ʽ", 120, True): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.������Դ, "��Դ", 120, True): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.��Ժ����, "��Դ", 120, True): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.����ʱ��, "����ʱ��", 120, True): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
'        '.SetImageList Me.imgList
        .AllowColumnRemove = False
        .MultipleSelection = False
        .ShowItemsInGroups = False
        .ShowHeader = False
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û�п���ʾ����Ŀ..."
            .VerticalGridStyle = xtpGridSolid
        End With
    End With
        If Me.RptThis.Rows.Count > 0 Then
            'Me.RptThis.Rows(1).Selected = True
            'Call mfrmContent.zlRefresh(Me.RptThis.Rows(1).Record(mCol.ID).Value, "NOUSE")
        End If
End Sub
Public Function zlRefresh(ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngPatiFrom As Long) As Boolean
    '���ܣ�ˢ��װ��ָ������Ĳ����ļ��嵥������λ��ָ�����ļ���
Dim strGroups As String
Dim rptRcd As ReportRecord
Dim rptItem As ReportRecordItem
Dim rptRow As ReportRow
Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    Me.RptThis.Tag = ""
    Me.RptThis.SetImageList Me.imgList
    If lngPatiFrom = 2 Then
        gstrSQL = "Select ID, ���, ����id, ��ҳid, ������Դ, ��������, ���ʱ��, ����ʱ��, ��������, �༭��ʽ, ��Ժ����" & vbNewLine & _
                   "From (Select r.Id, r.���, r.����id, r.��ҳid, r.������Դ, r.��������, To_Char(r.���ʱ��, 'yyyy-mm-dd hh24:mi') As ���ʱ��, r.����ʱ��, r.��������," & vbNewLine & _
                   "              r.�༭��ʽ, '��' || r.��ҳid || '��סԺ����' || '(' || To_Char(m.��Ժ����, 'yyyy-mm-dd hh24:mi:ss') || ')' As ��Ժ����" & vbNewLine & _
                   "       From ���Ӳ�����¼ R, ������ҳ M" & vbNewLine & _
                   "       Where r.�������� In (2, 5, 6) And nvl(r.�༭��ʽ,0)=0 And m.����id = r.����id And m.��ҳid = r.��ҳid And r.����id = [1] And r.������Դ = [2]"
        If InStr(gstrPrivsEpr, "��ʷ�ļ�") = 0 Then 'ûȨ��ֻ�ܿ����ξ���
            gstrSQL = gstrSQL & " And r.��ҳid=[3] "
        End If
        gstrSQL = gstrSQL & "       Union" & vbNewLine & _
                   "       Select r.Id, r.���, r.����id, r.��ҳid, r.������Դ, r.��������, To_Char(r.���ʱ��, 'yyyy-mm-dd hh24:mi') As ���ʱ��, r.����ʱ��, r.��������," & vbNewLine & _
                   "              r.�༭��ʽ, '��' || r.��ҳid || '��סԺ����' || '(' || To_Char(m.��Ժ����, 'yyyy-mm-dd hh24:mi:ss') || ')' As ��Ժ����" & vbNewLine & _
                   "       From ���Ӳ�����¼ R, ������ҳ M, ����ҽ������ L, ����ҽ����¼ A, Ӱ�����¼ B" & vbNewLine & _
                   "       Where r.�������� = 7 And nvl(r.�༭��ʽ,0)=0 And r.Id = l.����id And l.ҽ��id = a.Id And a.Id = b.ҽ��id And a.������� <> 'C' And m.����id = r.����id And" & vbNewLine & _
                   "             m.��ҳid = r.��ҳid And r.����id = [1] And r.������Դ = [2]"
        If InStr(gstrPrivsEpr, "��ʷ�ļ�") = 0 Then 'ûȨ��ֻ�ܿ����ξ���
            gstrSQL = gstrSQL & " And r.��ҳid=[3] "
        End If
        gstrSQL = gstrSQL & ") Order By ��Ժ����, ID, ���, ����ʱ��"
    Else
         gstrSQL = "Select r.id,r.����id,r.��ҳid,r.������Դ,r.��������, To_Char(r.���ʱ��, 'yyyy-mm-dd hh24:mi') As ���ʱ��,r.����ʱ��, r.Id, r.��������, r.�༭��ʽ," & vbNewLine & _
            "'���ﲡ��'||'('||to_char(nvl(m.ִ��ʱ��,m.�Ǽ�ʱ��),'yyyy-mm-dd hh24:mi:ss')||')' as ��Ժ���� " & vbNewLine & _
            "From ���Ӳ�����¼ r,���˹Һż�¼  m " & vbNewLine & _
            "Where r.�������� in (1,5,6,7) And nvl(r.�༭��ʽ,0)=0 and M.����ID = r.����ID and m.ID=r.��ҳid And r.����ID = [1] And r.������Դ = [2] " & vbNewLine & _
            "Order By m.ִ��ʱ��,r.Id,r.���,r.����ʱ�� asc "
    End If
   
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngPatiID, lngPatiFrom, lngPageId)
    Me.RptThis.Records.DeleteAll
    With rsTemp
        strGroups = ""
        Do While Not .EOF
            Set rptRcd = Me.RptThis.Records.Add()
            rptRcd.AddItem CStr(!ID)
            rptRcd.AddItem CStr(!��ҳID)
            rptRcd.AddItem CStr(!����ID)
            rptRcd.AddItem (CStr(!��������))
            rptRcd.AddItem CStr(NVL(!���ʱ��, ""))
            Set rptItem = rptRcd.AddItem(CStr(!��������)): rptItem.Icon = NVL(!��������, 0) - 1
            rptRcd.AddItem CStr(!�༭��ʽ)
            rptRcd.AddItem CStr(!������Դ)
            rptRcd.AddItem CStr(NVL(!��Ժ����, ""))
            rptRcd.AddItem CStr(NVL(!����ʱ��))
            .MoveNext
        Loop
        With Me.RptThis
            .SortOrder.Add .Columns.Find(mCol.ID)
            .SortOrder.Add .Columns.Find(mCol.����ʱ��)
            .SortOrder.Column(0).SortAscending = True
            .SortOrder.Column(1).SortAscending = True
            .GroupsOrder.Add .Columns.Find(mCol.��Ժ����)
            .GroupsOrder(0).SortAscending = False
            .Populate
        End With
    End With
    Exit Function
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_Unload(Cancel As Integer)
        Unload mfrmContent
        Set mfrmContent = Nothing
End Sub

Private Sub RptThis_SelectionChanged()
    Dim lngRecordId As Long
    On Error GoTo errHand
    If Me.RptThis.FocusedRow Is Nothing Then
        Exit Sub
    End If
    If Me.RptThis.FocusedRow.Record Is Nothing Then
        Exit Sub
    End If
    lngRecordId = Me.RptThis.FocusedRow.Record.Item(mCol.ID).Value
    If Val(Me.RptThis.Tag) <> Me.RptThis.FocusedRow.Index Then
        mfrmContent.mIsShowAnnex = False
        Call mfrmContent.zlRefresh(lngRecordId, "NOUSE")
        RptThis.Tag = Me.RptThis.FocusedRow.Index
    End If
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


