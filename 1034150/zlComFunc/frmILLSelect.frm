VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmILLSelect 
   AutoRedraw      =   -1  'True
   Caption         =   "����ѡ����"
   ClientHeight    =   5505
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8385
   Icon            =   "frmILLSelect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   8385
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraBottom 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   585
      Left            =   45
      TabIndex        =   13
      Top             =   4890
      Width           =   8280
      Begin VB.CommandButton cmdUnUse 
         Caption         =   "ȡ������(&U)"
         Height          =   350
         Left            =   3405
         TabIndex        =   9
         Top             =   120
         Width           =   1230
      End
      Begin VB.ComboBox cbo���� 
         Height          =   300
         Left            =   1530
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   150
         Width           =   1590
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   5655
         TabIndex        =   5
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   6750
         TabIndex        =   6
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdCommon 
         Caption         =   "��Ϊ����(&M)"
         Height          =   350
         Left            =   255
         TabIndex        =   7
         Top             =   120
         Width           =   1230
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsList 
      Height          =   4245
      Left            =   3315
      TabIndex        =   4
      Top             =   615
      Width           =   5025
      _cx             =   8864
      _cy             =   7488
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmILLSelect.frx":058A
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   -1  'True
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   5
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   2
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin MSComctlLib.ImageList iimg16 
      Left            =   1125
      Top             =   3405
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmILLSelect.frx":0618
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmILLSelect.frx":0BB2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraTop 
      Height          =   645
      Left            =   0
      TabIndex        =   11
      Top             =   -75
      Width           =   8385
      Begin VB.ComboBox cbo��� 
         Height          =   300
         Left            =   6045
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   225
         Width           =   2160
      End
      Begin VB.ComboBox cbo���� 
         Height          =   300
         Left            =   1005
         TabIndex        =   1
         Top             =   225
         Width           =   2160
      End
      Begin VB.Label lbl��� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������"
         Height          =   180
         Left            =   5250
         TabIndex        =   12
         Top             =   285
         Width           =   720
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ӧ����"
         Height          =   180
         Left            =   210
         TabIndex        =   0
         Top             =   285
         Width           =   720
      End
   End
   Begin VB.Frame fraLR 
      BorderStyle     =   0  'None
      Height          =   4245
      Left            =   3225
      MousePointer    =   9  'Size W E
      TabIndex        =   10
      Top             =   615
      Width           =   45
   End
   Begin MSComctlLib.TreeView tvwTree_s 
      Height          =   4245
      Left            =   15
      TabIndex        =   3
      Top             =   630
      Width           =   3150
      _ExtentX        =   5556
      _ExtentY        =   7488
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   441
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "iimg16"
      Appearance      =   1
   End
End
Attribute VB_Name = "frmILLSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mfrmParent As Object
Private mstr��� As String
Private mlng���˿���ID As Long
Private mstr�Ա� As String
Private mblnMultiSel As Boolean
Private mblnICD10 As Boolean

Private mrsList As ADODB.Recordset

Private mblnOK As Boolean
Private mstrLike As String
Private mlngPreDept As Long
Private mintPreClass As Integer
Private mstrPreNode As String

Public Function ShowMe(frmParent As Object, ByVal str��� As String, ByVal lng���˿���ID As Long, _
    Optional ByVal str�Ա� As String, Optional ByVal blnMultiSel As Boolean, Optional ByVal blnICD10 As Boolean = True) As ADODB.Recordset
    mstr��� = str���
    mlng���˿���ID = lng���˿���ID
    mstr�Ա� = str�Ա�
    mblnMultiSel = blnMultiSel
    mblnICD10 = blnICD10
    
    Set mfrmParent = frmParent
    Me.Show 1, frmParent
    
    If mblnOK Then Set ShowMe = mrsList
End Function

Private Sub cbo����_Click()
    Call SetControlEnabled
End Sub

Private Sub cbo����_Click()
    Dim rsTmp As ADODB.Recordset
    Dim lngRow As Long, strSQL As String
    Dim intIdx As Integer, blnDo As Boolean, i As Long
    Dim vRect As RECT, blnCancel As Boolean
        
    If cbo����.ListIndex = -1 Then Exit Sub
    If cbo����.ItemData(cbo����.ListIndex) = mlngPreDept And cbo����.ItemData(cbo����.ListIndex) <> -1 Then Exit Sub
    
    blnDo = True
    If cbo����.ItemData(cbo����.ListIndex) = -1 Then
        'ѡ����������
        strSQL = "Select Distinct A.ID,A.����,A.����,A.����" & _
            " From ���ű� A,��������˵�� B" & _
            " Where A.ID=B.����ID And B.������� IN(2,3)" & _
            " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
            " Order by A.����"
        vRect = GetControlRect(cbo����.hwnd)
        Set rsTmp = gobjComLib.zlDatabase.ShowSelect(Me, strSQL, 0, IIf(mblnICD10, "ѡ�񼲲�", "ѡ�����"), , , , , , True, vRect.Left, vRect.Top, cbo����.Height, blnCancel, , True)
        If Not rsTmp Is Nothing Then
            intIdx = SeekCboIndex(cbo����, rsTmp!id)
            '������Click�¼�,�ڱ��¼�����ʱһ������
            If intIdx <> -1 Then
                Call gobjComLib.zlControl.CboSetIndex(cbo����.hwnd, intIdx)
            Else
                cbo����.AddItem rsTmp!���� & "-" & rsTmp!����, cbo����.ListCount - 1
                cbo����.ItemData(cbo����.NewIndex) = rsTmp!id
                Call gobjComLib.zlControl.CboSetIndex(cbo����.hwnd, cbo����.NewIndex)
            End If
        Else
            If Not blnCancel Then
                MsgBox "û�п������ݣ����ȵ����Ź��������á�", vbInformation, gstrSysName
            End If
            '�ָ������еĿ���(������Click)
            intIdx = SeekCboIndex(cbo����, mlngPreDept)
            Call gobjComLib.zlControl.CboSetIndex(cbo����.hwnd, intIdx)
            blnDo = False
        End If
    End If
    mlngPreDept = cbo����.ItemData(cbo����.ListIndex)
    
    '��ȡ����
    If blnDo Then
        Call SetControlEnabled
        Call FillTreeData
    End If
End Sub

Private Sub cbo����_GotFocus()
    Call gobjComLib.zlControl.TxtSelAll(cbo����)
End Sub

Private Sub cbo����_KeyPress(KeyAscii As Integer)
    Dim blnCancel As Boolean
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If cbo����.ListIndex = -1 Then
            Call cbo����_Validate(blnCancel)
        End If
        If Not blnCancel Then
            Call cbo����_Validate(False)
            Call gobjComLib.zlCommFun.PressKey(vbKeyTab)
        End If
    End If
End Sub

Private Sub cbo����_Validate(Cancel As Boolean)
'���ܣ��������������,�Զ�ƥ��ִ�п���
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, intIdx As Long
    Dim vRect As RECT, blnCancel As Boolean
    Dim strInput As String, i As Long
    
    If cbo����.ListIndex <> -1 Then Exit Sub '��ѡ��,���ô���
    If cbo����.Text = "" Then Cancel = True: Exit Sub '������
    
    On Error GoTo errH
    
    strInput = UCase(gobjComLib.zlCommFun.GetNeedName(cbo����.Text))
    strSQL = "Select Distinct A.ID,A.����,A.����,A.����" & _
        " From ���ű� A,��������˵�� B" & _
        " Where A.ID=B.����ID And B.������� IN(2,3)" & _
        " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
        " And (Upper(A.����) Like [1] Or Upper(A.����) Like [2] Or Upper(A.����) Like [2])" & _
        " Order by A.����"
    
    vRect = GetControlRect(cbo����.hwnd)
    Set rsTmp = gobjComLib.zlDatabase.ShowSQLSelect(Me, strSQL, 0, IIf(mblnICD10, "����ѡ��", "���ѡ��"), False, "", "", False, False, _
        True, vRect.Left, vRect.Top, cbo����.Height, blnCancel, False, True, strInput & "%", mstrLike & strInput & "%")
    If Not rsTmp Is Nothing Then
        intIdx = SeekCboIndex(cbo����, rsTmp!id)
        If intIdx <> -1 Then
            cbo����.ListIndex = intIdx
        Else
            cbo����.AddItem rsTmp!���� & "-" & rsTmp!����, cbo����.ListCount - 1
            cbo����.ItemData(cbo����.NewIndex) = rsTmp!id
            cbo����.ListIndex = cbo����.NewIndex
        End If
    Else
        If Not blnCancel Then
            MsgBox "δ�ҵ���Ӧ�Ŀ��ҡ�", vbInformation, gstrSysName
        End If
        Cancel = True: Exit Sub
    End If
    Exit Sub
errH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Sub

Private Sub cbo���_Click()
    If mintPreClass = cbo���.ListIndex Then Exit Sub
    mintPreClass = cbo���.ListIndex
    
    Call FillTreeData
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCommon_Click()
    Dim arrSQL As Variant, i As Integer
    
    If cbo����.ListIndex = -1 Then
        MsgBox "��ָ����ǰ" & IIf(mblnICD10, "����", "���") & "�ĳ��ÿ��ҡ�", vbInformation, gstrSysName
        cbo����.SetFocus: Exit Sub
    End If
    If cbo����.ItemData(cbo����.ListIndex) = cbo����.ItemData(cbo����.ListIndex) Then
        MsgBox "��" & IIf(mblnICD10, "����", "���") & "�Ѿ�����Ϊ""" & cbo����.Text & """�ĳ���" & IIf(mblnICD10, "����", "���") & "��", vbInformation, gstrSysName
        cbo����.SetFocus: Exit Sub
    End If
    
    arrSQL = Array()
    With vsList
        If mblnMultiSel Then
            For i = .FixedRows To .Rows - 1
                If Val(.TextMatrix(i, 0)) <> 0 And .RowData(i) <> 0 Then
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    If mblnICD10 Then
                        arrSQL(UBound(arrSQL)) = "zl_�����������_Insert(" & .RowData(i) & "," & cbo����.ItemData(cbo����.ListIndex) & ")"
                    Else
                        arrSQL(UBound(arrSQL)) = "zl_������Ͽ���_Insert(" & .RowData(i) & "," & cbo����.ItemData(cbo����.ListIndex) & ")"
                    End If
                End If
            Next
        End If
        If UBound(arrSQL) = -1 Then
            If .RowData(.Row) = 0 Then
                MsgBox "û��ѡ��" & IIf(mblnICD10, "����", "���") & "��", vbInformation, gstrSysName
                Exit Sub
            End If
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            If mblnICD10 Then
                arrSQL(UBound(arrSQL)) = "zl_�����������_Insert(" & .RowData(.Row) & "," & cbo����.ItemData(cbo����.ListIndex) & ")"
            Else
                arrSQL(UBound(arrSQL)) = "zl_������Ͽ���_Insert(" & .RowData(.Row) & "," & cbo����.ItemData(cbo����.ListIndex) & ")"
            End If
        End If
    End With
    
    On Error GoTo errH
    gcnOracle.BeginTrans
    For i = 0 To UBound(arrSQL)
        Call gobjComLib.zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    gcnOracle.CommitTrans
        
    MsgBox "�����á�", vbInformation, gstrSysName
    vsList.SetFocus
    Exit Sub
errH:
    gcnOracle.RollbackTrans
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Sub

Private Sub cmdOK_Click()
    Dim strFilter As String
    Dim i As Long
    
    With vsList
        If mblnMultiSel Then
            For i = .FixedRows To .Rows - 1
                If Val(.TextMatrix(i, 0)) <> 0 Then
                    strFilter = strFilter & " Or ��ĿID=" & .RowData(i)
                End If
            Next
            strFilter = Mid(strFilter, 5)
        End If
        If strFilter = "" Then
            If .RowData(.Row) = 0 Then
                MsgBox "û��ѡ��" & IIf(mblnICD10, "����", "���") & "��", vbInformation, gstrSysName
                Exit Sub
            End If
            strFilter = "��ĿID=" & .RowData(.Row)
        End If
        
        mrsList.Filter = strFilter
        If mrsList.EOF Then
            MsgBox "û��ѡ��" & IIf(mblnICD10, "����", "���") & "��", vbInformation, gstrSysName
            Exit Sub
        End If
    End With
    
    mblnOK = True
    Unload Me
End Sub

Private Sub cmdUnUse_Click()
    Dim arrSQL As Variant, i As Integer
    
    If MsgBox("ȷʵҪ��ѡ���" & IIf(mblnICD10, "����", "���") & "��" & gobjComLib.zlCommFun.GetNeedName(cbo����.Text) & "��ȡ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    arrSQL = Array()
    With vsList
        If mblnMultiSel Then
            For i = .FixedRows To .Rows - 1
                If Val(.TextMatrix(i, 0)) <> 0 And .RowData(i) <> 0 Then
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    If mblnICD10 Then
                        arrSQL(UBound(arrSQL)) = "Zl_�����������_Delete(" & .RowData(i) & "," & cbo����.ItemData(cbo����.ListIndex) & ")"
                    Else
                        arrSQL(UBound(arrSQL)) = "Zl_������Ͽ���_Delete(" & .RowData(i) & "," & cbo����.ItemData(cbo����.ListIndex) & ")"
                    End If
                End If
            Next
        End If
        If UBound(arrSQL) = -1 Then
            If .RowData(.Row) = 0 Then
                MsgBox "û��ѡ��" & IIf(mblnICD10, "����", "���") & "��", vbInformation, gstrSysName
                Exit Sub
            End If
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            If mblnICD10 Then
                arrSQL(UBound(arrSQL)) = "Zl_�����������_Delete(" & .RowData(.Row) & "," & cbo����.ItemData(cbo����.ListIndex) & ")"
            Else
                arrSQL(UBound(arrSQL)) = "Zl_������Ͽ���_Delete(" & .RowData(.Row) & "," & cbo����.ItemData(cbo����.ListIndex) & ")"
            End If
        End If
    End With
    
    On Error GoTo errH
    gcnOracle.BeginTrans
    For i = 0 To UBound(arrSQL)
        Call gobjComLib.zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    gcnOracle.CommitTrans
    
    mstrPreNode = ""
    Call tvwTree_s_NodeClick(tvwTree_s.SelectedItem)
    Exit Sub
errH:
    gcnOracle.RollbackTrans
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Sub

Private Sub InitListTable()
'���ܣ���ʼ��ҽ���嵥��ʽ
    Dim arrHead As Variant, strHead As String, i As Long
    
    If mblnICD10 Then
        strHead = ",255,4;����,1000,1;����,550,1;����,2500,1;˵��,1500,1"
    Else
        strHead = ",255,4;����,1000,1;����,2500,1;˵��,1500,1;����,850,1"
    End If
    arrHead = Split(strHead, ";")
    With vsList
        .Clear
        .FixedRows = 1: .FixedCols = 0
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(.FixedCols + i) = False
                .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
            End If
        Next
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
    End With
End Sub

Private Sub Form_Load()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim blnDept As Boolean, blnHave As Boolean
        
    Call InitListTable
    Call gobjComLib.RestoreWinState(Me, App.ProductName, mfrmParent.Name & IIf(mblnICD10, 1, 0))
    
    mblnOK = False
    mlngPreDept = -1
    mintPreClass = -1
    mstrPreNode = ""
    Set mrsList = Nothing
    mstrLike = IIf(Val(gobjComLib.zlDatabase.GetPara("����ƥ��")) = 0, "%", "") '����ƥ�䷽ʽ
    
    If Not mblnICD10 Then Me.Caption = "���ѡ����"
    
    On Error GoTo errH
    
    '����Ƿ��ж�Ӧ����
    If mblnICD10 Then
        If mstr��� = "" Then
            strSQL = "Select A.* From ����������� A,������Ա B,�ϻ���Ա�� C" & _
                " Where A.����ID=B.����ID And B.��ԱID=C.��ԱID And C.�û���=User And Rownum=1"
        Else
            strSQL = "Select A.* From ��������Ŀ¼ I,����������� A,������Ա B,�ϻ���Ա�� C" & _
                " Where I.ID=A.����ID And A.����ID=B.����ID And B.��ԱID=C.��ԱID" & _
                " And (I.����ʱ�� is Null Or I.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " And C.�û���=User And Instr([1],I.���)>0 And Rownum=1"
        End If
    Else
        If mstr��� = "" Then mstr��� = "1,2"
        strSQL = "Select A.* From �������Ŀ¼ I,������Ͽ��� A,������Ա B,�ϻ���Ա�� C" & _
            " Where I.ID=A.���ID And A.����ID=B.����ID And B.��ԱID=C.��ԱID" & _
            " And C.�û���=User And Instr([1],I.���)>0 And Rownum=1"
    End If
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr���)
    If Not rsTmp.EOF Then blnDept = True
    
    '��ʾ��ǰ��Ա����
    strSQL = "Select A.ID,A.����,A.����,A.����,Max(Nvl(C.ȱʡ,0)) as ȱʡ" & _
        " From ���ű� A,��������˵�� B,������Ա C,�ϻ���Ա�� D" & _
        " Where A.ID=B.����ID And B.�������� IN('�ٴ�','���','����','����','����','Ӫ��')" & _
        " And A.ID=C.����ID And C.��ԱID=D.��ԱID And D.�û���=User" & _
        " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� Is Null)" & _
        " Group by A.ID,A.����,A.����,A.����" & _
        " Order by A.����"
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    cbo����.AddItem IIf(mblnICD10, "���м���", "�������")
    Do While Not rsTmp.EOF
        blnHave = True
        cbo����.AddItem rsTmp!���� & "-" & rsTmp!����
        cbo����.ItemData(cbo����.NewIndex) = rsTmp!id
        If blnDept Then
            If rsTmp!id = mlng���˿���ID Then
                Call gobjComLib.zlControl.CboSetIndex(cbo����.hwnd, cbo����.NewIndex)
            ElseIf cbo����.ListIndex = -1 And rsTmp!ȱʡ = 1 Then
                Call gobjComLib.zlControl.CboSetIndex(cbo����.hwnd, cbo����.NewIndex)
            End If
        End If
        
        cbo����.AddItem rsTmp!����
        cbo����.ItemData(cbo����.NewIndex) = rsTmp!id
        If rsTmp!id = mlng���˿���ID Then
            cbo����.ListIndex = cbo����.NewIndex
        ElseIf cbo����.ListIndex = -1 And rsTmp!ȱʡ = 1 Then
            cbo����.ListIndex = cbo����.NewIndex
        End If
        
        rsTmp.MoveNext
    Loop
    cbo����.AddItem "<��������...>"
    cbo����.ItemData(cbo����.NewIndex) = -1
    
    If cbo����.ListIndex = -1 Then
        If Not blnDept Or Not blnHave Then
            '���κμ�����Ӧ��������ʱ,������Ա�޶�Ӧ����ʱ��ȱʡ��ʾ���м���
            Call gobjComLib.zlControl.CboSetIndex(cbo����.hwnd, 0)
        Else
            Call gobjComLib.zlControl.CboSetIndex(cbo����.hwnd, 1)
        End If
    End If
    If cbo����.ListCount > 0 And cbo����.ListIndex = -1 Then
        cbo����.ListIndex = 0
    End If
    
    '��ʾ�����������
    If mblnICD10 Then
        If mstr��� = "" Then
            strSQL = "Select ����,���,�Ƿ���� From ����������� Order by ���ȼ�"
        Else
            strSQL = "Select ����,���,�Ƿ���� From ����������� Where Instr([1],����)>0 Order by ���ȼ�"
        End If
        Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr���)
        Do While Not rsTmp.EOF
            cbo���.AddItem rsTmp!���� & ". " & rsTmp!���
            cbo���.ItemData(cbo���.NewIndex) = NVL(rsTmp!�Ƿ����, 0)
            rsTmp.MoveNext
        Loop
        Call gobjComLib.zlControl.CboSetIndex(cbo���.hwnd, 0)
        If cbo���.ListCount = 1 Then cbo���.Locked = True
    Else
        lbl���.Visible = False
        cbo���.Visible = False
    End If
    
    'ȱʡ��ȡ����
    Call FillTreeData
    Exit Sub
errH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    fraTop.Left = 0
    fraTop.Top = -75
    fraTop.Width = Me.ScaleWidth
    
    If fraTop.Width - cbo���.Width - 200 > 4135 Then
        cbo���.Left = fraTop.Width - cbo���.Width - 200
        lbl���.Left = cbo���.Left - lbl���.Width - 75
    End If
    
    fraBottom.Left = 0
    fraBottom.Top = Me.ScaleHeight - fraBottom.Height
    fraBottom.Width = Me.ScaleWidth
    
    If fraBottom.Width - cmdCancel.Width - 550 > 7000 Then
        cmdCancel.Left = fraBottom.Width - cmdCancel.Width - 800
        cmdOK.Left = cmdCancel.Left - cmdOK.Width
    End If
    
    tvwTree_s.Left = 0
    tvwTree_s.Top = fraTop.Top + fraTop.Height + 15
    tvwTree_s.Height = Me.ScaleHeight - tvwTree_s.Top - fraBottom.Height
    
    fraLR.Top = tvwTree_s.Top
    fraLR.Left = tvwTree_s.Left + tvwTree_s.Width
    fraLR.Height = tvwTree_s.Height
    
    vsList.Top = tvwTree_s.Top
    vsList.Left = IIf(tvwTree_s.Visible, fraLR.Left + fraLR.Width, 0)
    vsList.Width = Me.ScaleWidth - vsList.Left
    vsList.Height = tvwTree_s.Height
    
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call gobjComLib.SaveWinState(Me, App.ProductName, mfrmParent.Name & IIf(mblnICD10, 1, 0))
End Sub

Private Sub fraLR_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If tvwTree_s.Width + X < 1000 Or vsList.Width - X < 1000 Then Exit Sub
        fraLR.Left = fraLR.Left + X
        tvwTree_s.Width = tvwTree_s.Width + X
        vsList.Left = vsList.Left + X
        vsList.Width = vsList.Width - X
    End If
End Sub

Private Sub FillTreeData()
'���ܣ���ȡ�����������ݣ������ǿ��Ҷ�Ӧ����ֻ��Ӧ�ķ���
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim objNode As Node
    
    '�������
    Set mrsList = Nothing
    tvwTree_s.Nodes.Clear
    vsList.Rows = vsList.FixedRows
    vsList.Rows = vsList.FixedRows + 1
    
    'ICD-10����Ƿ��з���
    If mblnICD10 Then
        If cbo���.ItemData(cbo���.ListIndex) = 0 Then
            tvwTree_s.Visible = False
            fraLR.Visible = False
        Else
            tvwTree_s.Visible = True
            fraLR.Visible = True
        End If
        Call Form_Resize
    End If
    
    Screen.MousePointer = 11
    Me.Refresh
    
    On Error GoTo errH
    
    If mblnICD10 Then
        If cbo���.ItemData(cbo���.ListIndex) <> 0 Then 'Ϊ0��ʾ���ּ���û�з���
            If cbo����.ItemData(cbo����.ListIndex) = 0 Then
                strSQL = "Select ID,�ϼ�ID,���,���� From ����������� Where ���=[1]" & _
                    " Start With �ϼ�ID is Null Connect by Prior ID=�ϼ�ID Order by Level,���"
            Else
                strSQL = _
                    " Select Distinct B.����id From ����������� A, ��������Ŀ¼ B" & _
                    " Where A.����id = B.ID And A.����id = [2]" & _
                    " And (B.����ʱ�� is Null Or B.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))"
                strSQL = _
                    "Select Max(Level) as ��ID, ID, �ϼ�id, ���, ����" & vbNewLine & _
                    "From ����������� Where ���=[1]" & vbNewLine & _
                    "Start With ID In (" & strSQL & ")" & vbNewLine & _
                    "Connect By Prior �ϼ�id = ID" & vbNewLine & _
                    "Group By ID, �ϼ�ID, ���, ����" & vbNewLine & _
                    "Order By ��ID Desc"
                strSQL = "Select ID, �ϼ�id, ���, ���� From (" & strSQL & ")"
            End If
            Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Left(cbo���.Text, 1), cbo����.ItemData(cbo����.ListIndex))
            Do Until rsTmp.EOF
                If IsNull(rsTmp!�ϼ�ID) Then
                    Set objNode = tvwTree_s.Nodes.Add(, , "_" & rsTmp!id, "��" & rsTmp!��� & "��" & Trim(rsTmp!����), 1, 2)
                Else
                    Set objNode = tvwTree_s.Nodes.Add("_" & rsTmp!�ϼ�ID, 4, "_" & rsTmp!id, "��" & rsTmp!��� & "��" & Trim(rsTmp!����), 1, 2)
                End If
                rsTmp.MoveNext
            Loop
        End If
    Else
        If cbo����.ItemData(cbo����.ListIndex) = 0 Then
            strSQL = "Select ID,�ϼ�ID,����,���� From ������Ϸ��� Where Instr([1],���)>0" & _
                " Start With �ϼ�ID is Null Connect by Prior ID=�ϼ�ID Order by Level,����"
        Else
            strSQL = _
                " Select Distinct C.����ID From ������Ͽ��� A, �������Ŀ¼ B,����������� C" & _
                " Where A.���ID = B.ID And B.ID=C.���ID And A.����ID = [2]"
            strSQL = _
                "Select Max(Level) as ��ID, ID, �ϼ�id, ����, ����" & vbNewLine & _
                "From ������Ϸ��� Where Instr([1],���)>0" & vbNewLine & _
                "Start With ID In (" & strSQL & ")" & vbNewLine & _
                "Connect By Prior �ϼ�id = ID" & vbNewLine & _
                "Group By ID, �ϼ�ID, ����, ����" & vbNewLine & _
                "Order By ��ID Desc"
            strSQL = "Select ID, �ϼ�id, ����, ���� From (" & strSQL & ")"
        End If
        Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr���, cbo����.ItemData(cbo����.ListIndex))
        Do Until rsTmp.EOF
            If IsNull(rsTmp!�ϼ�ID) Then
                Set objNode = tvwTree_s.Nodes.Add(, , "_" & rsTmp!id, "[" & rsTmp!���� & "]" & Trim(rsTmp!����), 1, 2)
            Else
                Set objNode = tvwTree_s.Nodes.Add("_" & rsTmp!�ϼ�ID, 4, "_" & rsTmp!id, "[" & rsTmp!���� & "]" & Trim(rsTmp!����), 1, 2)
            End If
            rsTmp.MoveNext
        Loop
    End If
    
    If tvwTree_s.Nodes.count > 0 Then
        tvwTree_s.Nodes(1).Selected = True
        tvwTree_s.Nodes(1).Expanded = True
        tvwTree_s.Nodes(1).EnsureVisible
    End If
    
    Screen.MousePointer = 0
    Call FillListData
    Exit Sub
errH:
    Screen.MousePointer = 0
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Sub

Private Sub FillListData()
    Dim strSQL As String
    Dim str�Ա� As String
    Dim lng����ID As Long
    Dim i As Long
    
    On Error GoTo errH
    
    Screen.MousePointer = 11
    
    vsList.Rows = vsList.FixedRows
    vsList.Rows = vsList.FixedRows + 1
    vsList.ColHidden(0) = Not mblnMultiSel
    
    If mblnICD10 Then
        If mstr�Ա� Like "*��*" Then
            str�Ա� = "��"
        ElseIf mstr�Ա� Like "*Ů*" Then
            str�Ա� = "Ů"
        End If
        
        If cbo����.ItemData(cbo����.ListIndex) <> 0 Then
            strSQL = _
                " Select A.ID as ��ĿID,A.����,A.���,A.����,A.����,A.˵��" & _
                " From ��������Ŀ¼ A,����������� B" & _
                " Where A.ID=B.����ID And B.����ID=[1] And A.���=[2]" & _
                " And (A.����ʱ�� is Null Or A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))"
        Else
            strSQL = "Select A.ID as ��ĿID,A.����,A.���,A.����,A.����,A.˵�� From ��������Ŀ¼ A" & _
                " Where A.���=[2] And (A.����ʱ�� is Null Or A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))"
        End If
        If str�Ա� <> "" Then strSQL = strSQL & " And (A.�Ա�����=[4] Or A.�Ա����� is Null)"
        
        If cbo���.ItemData(cbo���.ListIndex) <> 0 Then 'Ϊ0��ʾ���ּ���û�з���
            If tvwTree_s.SelectedItem Is Nothing Then
                vsList.Row = 1: Screen.MousePointer = 0: Exit Sub
            End If
            
            lng����ID = Val(Mid(tvwTree_s.SelectedItem.Key, 2))
            strSQL = strSQL & " And A.����ID=[3]"
        End If
        strSQL = strSQL & " Order by A.����,A.���"
        
        Set mrsList = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, cbo����.ItemData(cbo����.ListIndex), Left(cbo���.Text, 1), lng����ID, str�Ա�)
        If Not mrsList.EOF Then
            With vsList
                .Redraw = flexRDNone
                .Rows = .FixedRows + mrsList.RecordCount
                For i = 1 To mrsList.RecordCount
                    .RowData(i) = Val(mrsList!��ĿID)
                    .TextMatrix(i, 0) = 0
                    .TextMatrix(i, 1) = NVL(mrsList!����)
                    .Cell(flexcpData, i, 1) = CStr(NVL(mrsList!����))
                    If NVL(mrsList!����) = .Cell(flexcpData, i - 1, 1) Then
                        If Not IsNull(mrsList!���) Then
                            .TextMatrix(i, 1) = .TextMatrix(i, 1) & "." & mrsList!���
                            If .TextMatrix(i - 1, 1) = .Cell(flexcpData, i - 1, 1) And mrsList!��� = 2 Then
                                .TextMatrix(i - 1, 1) = .TextMatrix(i - 1, 1) & ".1"
                            End If
                        End If
                    End If
                    
                    .TextMatrix(i, 2) = NVL(mrsList!����)
                    .TextMatrix(i, 3) = NVL(mrsList!����)
                    .TextMatrix(i, 4) = NVL(mrsList!˵��)
                    mrsList.MoveNext
                Next
                .Redraw = flexRDDirect
            End With
        End If
    Else
        If tvwTree_s.SelectedItem Is Nothing Then
            vsList.Row = 1: Screen.MousePointer = 0: Exit Sub
        End If
        lng����ID = Val(Mid(tvwTree_s.SelectedItem.Key, 2))
        
        If cbo����.ItemData(cbo����.ListIndex) <> 0 Then
            strSQL = _
                " Select A.ID as ��ĿID,A.����,A.����,A.˵��,A.����" & _
                " From �������Ŀ¼ A,������Ͽ��� B,����������� C" & _
                " Where A.ID=B.���ID And A.ID=C.���ID And B.����ID=[1] And Instr([2],A.���)>0 And C.����ID=[3]" & _
                " Order by A.����"
        Else
            strSQL = "Select A.ID as ��ĿID,A.����,A.����,A.˵��,A.����" & _
                " From �������Ŀ¼ A,����������� B" & _
                " Where Instr([2],A.���)>0 And A.ID=B.���ID And B.����ID=[3]" & _
                " Order by A.����"
        End If
        Set mrsList = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, cbo����.ItemData(cbo����.ListIndex), mstr���, lng����ID)
        If Not mrsList.EOF Then
            With vsList
                .Redraw = flexRDNone
                .Rows = .FixedRows + mrsList.RecordCount
                For i = 1 To mrsList.RecordCount
                    .RowData(i) = Val(mrsList!��ĿID)
                    .TextMatrix(i, 0) = 0
                    .TextMatrix(i, 1) = NVL(mrsList!����)
                    .TextMatrix(i, 2) = NVL(mrsList!����)
                    .TextMatrix(i, 3) = NVL(mrsList!˵��)
                    .TextMatrix(i, 4) = NVL(mrsList!����)
                    mrsList.MoveNext
                Next
                .Redraw = flexRDDirect
            End With
        End If
    End If
    
    vsList.Row = 1: vsList.Col = 1
    Screen.MousePointer = 0
    
    Call vsList_AfterRowColChange(-1, -1, vsList.Row, vsList.Col)
    Exit Sub
errH:
    Screen.MousePointer = 0
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Sub

Private Sub tvwTree_s_NodeClick(ByVal Node As MSComctlLib.Node)
    If mstrPreNode = Node.Key Then Exit Sub
    mstrPreNode = Node.Key
    
    Call FillListData
End Sub

Private Function NVL(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'���ܣ��൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
    NVL = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Private Sub vsList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim blnEnabled As Boolean, i As Integer
    
    Call SetControlEnabled
    
    '�������ݵ�����£�ֻ��ȡ�������������ҵĳ��ü���
    If vsList.RowData(vsList.Row) <> 0 And cbo����.ListIndex > 0 Then
        For i = 0 To cbo����.ListCount - 1
            If cbo����.ItemData(cbo����.ListIndex) = cbo����.ItemData(i) Then
                blnEnabled = True: Exit For
            End If
        Next
    End If
    cmdUnUse.Enabled = blnEnabled
End Sub

Private Sub vsList_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Col = 0 Then
        If Val(vsList.TextMatrix(Row, 0)) <> 0 Then
            vsList.Cell(flexcpBackColor, Row, 0, Row, vsList.Cols - 1) = &HC0FFFF
        Else
            vsList.Cell(flexcpBackColor, Row, 0, Row, vsList.Cols - 1) = vsList.BackColor
        End If
    End If
End Sub

Private Sub vsList_DblClick()
    If vsList.MouseRow >= vsList.FixedRows Then
        Call cmdOK_Click
    End If
End Sub

Private Sub vsList_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call cmdOK_Click
    ElseIf KeyAscii = 32 Then
        If mblnMultiSel And vsList.Col > 0 And vsList.RowData(vsList.Row) <> 0 Then
            vsList.TextMatrix(vsList.Row, 0) = IIf(Val(vsList.TextMatrix(vsList.Row, 0)) = 0, 1, 0)
        End If
    End If
End Sub

Private Sub SetControlEnabled()
    Dim blnEnabled As Boolean
    
    '��Ϊ���õĿ�����
    blnEnabled = True
    If cbo����.ListIndex = -1 Then
        blnEnabled = False
    ElseIf cbo����.ListIndex <> -1 And cbo����.ListIndex <> -1 Then
        If cbo����.ItemData(cbo����.ListIndex) = cbo����.ItemData(cbo����.ListIndex) Then
            blnEnabled = False
        End If
    End If
    If blnEnabled Then
        If vsList.Row >= vsList.FixedRows Then
            blnEnabled = vsList.RowData(vsList.Row) <> 0
        End If
    End If
    cmdCommon.Enabled = blnEnabled
    
    'ȷ����ť�Ŀ�����
    blnEnabled = True
    If vsList.Row >= vsList.FixedRows Then
        blnEnabled = vsList.RowData(vsList.Row) <> 0
    Else
        blnEnabled = False
    End If
    cmdOK.Enabled = blnEnabled
End Sub

Private Sub vsList_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If vsList.RowData(Row) = 0 Then
        Cancel = True
    ElseIf Col <> 0 Then
        Cancel = True
    End If
End Sub
