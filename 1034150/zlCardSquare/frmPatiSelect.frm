VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmPatiSelect 
   Caption         =   "����ѡ��"
   ClientHeight    =   7020
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10710
   Icon            =   "frmPatiSelect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7020
   ScaleWidth      =   10710
   Begin VB.PictureBox picCmd 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   0
      ScaleHeight     =   525
      ScaleWidth      =   10710
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   6495
      Width           =   10710
      Begin VB.CommandButton cmdPatiMerge 
         Caption         =   "���˺ϲ�(&M)"
         Enabled         =   0   'False
         Height          =   350
         Left            =   225
         TabIndex        =   6
         Top             =   120
         Width           =   1650
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   8505
         TabIndex        =   5
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   9600
         TabIndex        =   4
         Top             =   120
         Width           =   1100
      End
   End
   Begin VB.PictureBox picInfo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   0
      ScaleHeight     =   390
      ScaleWidth      =   10710
      TabIndex        =   1
      Top             =   0
      Width           =   10710
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ѡ��һ����Ŀ,Ȼ����ȷ��"
         Height          =   180
         Left            =   180
         TabIndex        =   2
         Top             =   120
         Width           =   2430
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsSelect 
      Height          =   5655
      Left            =   1665
      TabIndex        =   0
      Top             =   630
      Width           =   6825
      _cx             =   12039
      _cy             =   9975
      Appearance      =   2
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
      GridLines       =   8
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   10
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   300
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
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
      ExplorerBar     =   7
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
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
   Begin VB.Shape shpLine 
      Height          =   1815
      Left            =   255
      Top             =   1125
      Width           =   675
   End
End
Attribute VB_Name = "frmPatiSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngSys As Long, mlngModule As Long
Private mfrmMain As Form, mobjControl As Object
Private mrsBindings As ADODB.Recordset, mrsOutSel As ADODB.Recordset
Private mblnShowHead As Boolean, mstr������ As String, mstrHideCols As String '��1,��2,...
Private mblnOK As Boolean
Private mstrWinTittle As String, mstrNotes As String, mblnShowPatiMerge As Boolean, mblnNotShowWin As Boolean
Private marrInput() As Variant
Private mstrSQL As String
'-------------------------------------------------------------------------------------------------------------------
'�ؼ���λ
Private Type ty_ctlObject_Locale
    '�ؼ���λ��
    Left As Single
    Top As Single
    Width As Single
    Height As Single
    '�����б�����С�߶ȺͿ���
    minWidth As Single
    minHeight As Single
    
    '�½��б���ʵ��λ��
    DownLeft As Single
    DownTop As Single
    DownWidth As Single
    DownHeight As Single
    '��ģ���
    ScreenWidth As Single
    ScreenHeight As Single
End Type
Private mTyCtl_Locale As ty_ctlObject_Locale
'-------------------------------------------------------------------------------------------------------------------
'--API����
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private mcnOracle As ADODB.Connection

Private Sub SetWindowsProperty()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ô�������
    '����:���˺�
    '����:2012-08-21 11:33:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    cmdPatiMerge.Visible = mblnShowPatiMerge
    
    If mstrWinTittle <> "" Then Me.Caption = mstrWinTittle & IIf(InStr(mstrWinTittle, "ѡ��") > 0, "", "ѡ��")
    picInfo.Visible = False
    If mstrNotes <> "" Then
        lblInfo.Caption = mstrNotes
        picInfo.Visible = True
    End If
    If mblnNotShowWin Then
        Call zlControl.FormSetCaption(Me, False, False)
    End If
    '���ڲ��˺ϲ�����ڴ���ʱ,��ʾ����Ĺ��ܰ�ť
   picCmd.Visible = mblnNotShowWin = False Or mblnShowPatiMerge
   If mblnShowPatiMerge Then
        vsSelect.AllowSelection = True
        vsSelect.AllowBigSelection = True
        vsSelect.SelectionMode = flexSelectionListBox
   End If
End Sub

Private Function GetBoundData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������
    '����: �ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2012-08-21 14:54:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsBindings As ADODB.Recordset, strLog As String
    Dim objDatabase As Object, objTemp As clsDataBase
    
    Screen.MousePointer = 11
    Set objDatabase = zlDatabase
    If Not mcnOracle Is Nothing Then
        Set objTemp = New clsDataBase
        Call objTemp.InitCommon(mcnOracle)
        Set objDatabase = objTemp
    End If
    
    Screen.MousePointer = 11
    mblnOK = False
    '��SQL���
    If UBound(marrInput) >= 0 Then
        Set mrsBindings = OpenSQLRecord(mstrSQL, strLog)
    Else
        Set mrsBindings = objDatabase.OpenSQLRecord(mstrSQL, mstrWinTittle & "ѡ��", mstrSQL)
    End If
    Set objDatabase = Nothing: Set objTemp = Nothing

    'û�������򷵻�
    If mrsBindings.EOF Then
        Screen.MousePointer = 0
        Set mrsBindings = Nothing: Set mrsOutSel = Nothing
        mblnOK = False: Unload Me: Exit Function
    End If
    If mrsBindings.RecordCount = 1 Then
        Set mrsOutSel = mrsBindings
        Screen.MousePointer = 0
        mblnOK = True: Unload Me: Exit Function
    End If
    Screen.MousePointer = 0
    GetBoundData = True
End Function

Public Function ShowSelect(ByVal frmMain As Form, ByVal cnOracle As ADODB.Connection, ByVal lngSys As Long, _
    ByVal lngModule As Long, ByVal objControl As Object, _
    ByVal strSql As String, _
    ByVal strWinTittle As String, _
    ByVal strNotes As String, _
    ByVal blnShowHead As Boolean, _
    ByVal blnShowPatiMerge As Boolean, _
    ByVal blnNotShowWin As Boolean, _
    ByVal str������ As String, _
    ByVal strHideCols As String, _
    ByRef rsOutSel As ADODB.Recordset, _
     ParamArray arrInput() As Variant) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ѡ�������
    '���:frmMain-���õ�������
    '     lngSys-ϵͳ��
    '     lngModule-ģ���
    '     objControl-�ؼ�����(Ŀǰֻ֧:textBox,Combox)
    '     strSQL-����Ϊ��,��Ҫ�ֶ�,ID,......
    '     str����-���Ի�����Ĳ�����.
    '     blnShowHead-�Ƿ���ʾ����ͷ
    '
    '����:rsOutSel-ѡ���ļ�¼��
    '����:ѡ�з���True, ���򷵻�False(���԰�Esc���з���)
    '����:���˺�
    '����:2009-01-01 15:35:30
    '����:52913
    '---------------------------------------------------------------------------------------------------------------------------------------------
   marrInput = arrInput
   Set rsOutSel = Nothing: Set mcnOracle = cnOracle
   mstrWinTittle = strWinTittle: mstrNotes = strNotes: mblnShowHead = blnShowHead
   mblnShowPatiMerge = blnShowPatiMerge: mblnNotShowWin = blnNotShowWin
   mstrSQL = strSql
   'ֻ��һ������,��ֱ�ӷ���
    Set mfrmMain = frmMain: mlngSys = lngSys: mlngModule = lngModule
    mblnOK = False: Set mobjControl = objControl
    mblnShowHead = blnShowHead: mstr������ = str������: mstrHideCols = strHideCols
    On Error Resume Next
    If Not frmMain Is Nothing Then
        Me.Show 1, frmMain
    Else
        Me.Show 1
    End If
    On Error GoTo 0
    Set rsOutSel = mrsOutSel
    ShowSelect = mblnOK
End Function
Private Function SelectedItem() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ѡ��ָ��������
    '���:
    '����:
    '����:
    '����:���˺�
    '����:2009-07-20 12:21:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngID As Long
    With vsSelect
        lngID = Val(.TextMatrix(.Row, .ColIndex("ID")))
    End With
    If lngID = 0 Then Exit Function
    
    Set mrsOutSel = mrsBindings
    mrsOutSel.Filter = "ID=" & lngID
    If mrsOutSel.RecordCount = 0 Then Exit Function
    mblnOK = True
    Unload Me
    SelectedItem = True
End Function


Private Function zlBindingData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������
    '���:
    '����:
    '����:
    '����:���˺�
    '����:2009-07-20 11:37:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    If GetBoundData = False Then Exit Function
    
    '��ʼ������
    With vsSelect
        .Redraw = flexRDNone
        Err = 0: On Error Resume Next
        Set vsSelect.Font = mobjControl.Font
        Set Me.Font = mobjControl.Font
        Err = 0: On Error GoTo 0
        Set .DataSource = mrsBindings
        If mrsBindings.EOF Then .Rows = 2: .Clear 1
        For i = 0 To .Cols - 1
            .ColKey(i) = Trim(.TextMatrix(0, i))
            If UCase(.ColKey(i)) Like "*ID" Then .ColHidden(i) = True
            If InStr(1, "," & mstrHideCols & ",", "," & UCase(.ColKey(i)) & ",") > 0 Then .ColHidden(i) = True
            If .ColKey(i) Like "*���" Then
                .ColAlignment(i) = flexAlignRightCenter
            End If
            .AutoSizeMode = flexAutoSizeColWidth
            .AutoSize 0, .Cols - 1
        Next
        .RowHidden(0) = False
        If mblnShowHead = False Then .RowHidden(0) = True
        '�ָ���˳��
        .Redraw = flexRDBuffered
        '�ָ�����ؼ�˳��
        If mstr������ <> "" Then Call zl_vsGrid_Para_Restore(mlngSys, mlngModule, vsSelect, mstr������)
    End With
    zlBindingData = True
End Function
Private Sub InitCtrlLocal()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ָ���ؼ���ʼ���ؼ�λ��
    '���:
    '����:
    '����:
    '����:���˺�
    '����:2009-07-20 10:04:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngH As Long, lngW As Long, vRect As RECT, sngX As Single, sngY As Single
    
    '�޶�Ӧ����ؼ�
    If mobjControl Is Nothing Then
        With mTyCtl_Locale
            .Top = 0
            .Left = 0
            .Width = Me.Width
            .Height = Me.Height
            .minHeight = vsSelect.RowHeight(0) * 5  'һ������
            .minHeight = .minHeight + IIf(mblnNotShowWin = False Or mblnShowPatiMerge, picCmd.Height, 0)
            .minHeight = .minHeight + IIf(mstrNotes <> "", picInfo.Height, 0)
            .minWidth = .Width
            .ScreenHeight = GetSystemMetrics(SM_CYFULLSCREEN) * 15   '��Ļ���ø߶�
            .ScreenWidth = Screen.Width  ' GetSystemMetrics(SM_CXVSCROLL) * 15 + 75  '��Ļ���ÿ���
        End With
        Exit Sub
    End If
   
   'ͨ��Api������ؼ������������Ϣ
    Select Case UCase(TypeName(mobjControl))
    Case UCase("VSFlexGrid")
        Call CalcPosition(sngX, sngY, mobjControl)
        lngH = mobjControl.CellHeight
        lngW = mobjControl.CellWidth
        sngY = sngY - lngH
    Case UCase("BILLEDIT")
        Call CalcPosition(sngX, sngY, mobjControl.MsfObj)
        lngH = mobjControl.MsfObj.CellHeight
        lngW = mobjControl.MsfObj.CellWidth
    Case Else
        vRect = GetControlRect(mobjControl.hWnd)
        sngX = vRect.Left - 15
        sngY = vRect.Top
        lngH = mobjControl.Height
        lngW = mobjControl.Width
    End Select

    With mTyCtl_Locale
        .Top = sngY
        .Left = sngX
        .Width = lngW
        .Height = lngH
        .minHeight = vsSelect.RowHeight(0) * 5 'һ������
        .minWidth = .Width
        .ScreenHeight = GetSystemMetrics(SM_CYFULLSCREEN) * 15   '��Ļ���ø߶�
        .ScreenWidth = Screen.Width  ' GetSystemMetrics(SM_CXVSCROLL) * 15 + 75  '��Ļ���ÿ���
    End With
End Sub

Public Sub ReSetWindowsFormLocal()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������ô��ڵĴ�С��λ��
    '���:
    '����:
    '����:
    '����:���˺�
    '����:2009-07-20 10:30:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblColsWidth As Double, dblRowsHeight As Double
    Dim dblTemp As Double, lngPicHeight As Long
    Dim i As Long
    
    '��λ
    With mTyCtl_Locale
        .DownTop = .Top + .Height
        .DownLeft = .Left
        .DownWidth = .Width
    End With
    lngPicHeight = IIf(mstrNotes <> "", picInfo.Height, 0)
    lngPicHeight = lngPicHeight + IIf(mblnNotShowWin = False Or mblnShowPatiMerge, picCmd.Height, 0)
    
    
    '�����������Ŀ���
    dblColsWidth = 0
    For i = 0 To vsSelect.Cols - 1
        If Not vsSelect.ColHidden(i) Then
            dblColsWidth = dblColsWidth + vsSelect.ColWidth(i) + Screen.TwipsPerPixelX
        End If
    Next
    dblColsWidth = dblColsWidth + 300
    
    '�����������ĸ߶�
    dblRowsHeight = vsSelect.Cell(flexcpHeight, 0, 0, 0, 0)
    
    dblRowsHeight = (dblRowsHeight) * (vsSelect.Rows) + 100 + lngPicHeight
    If dblRowsHeight < mTyCtl_Locale.minHeight Then dblRowsHeight = mTyCtl_Locale.minHeight
    
    dblColsWidth = IIf(dblColsWidth < mTyCtl_Locale.minWidth, mTyCtl_Locale.minWidth, dblColsWidth)
    If vsSelect.Width > dblColsWidth - 300 Then
        vsSelect.ExtendLastCol = True
    Else
        vsSelect.ExtendLastCol = False
    End If
    
    If mblnNotShowWin = False Then
        If (mTyCtl_Locale.ScreenWidth - Me.Width) \ 2 > 0 Then
            Me.Left = (mTyCtl_Locale.ScreenWidth - Me.Width) \ 2
        End If
        If (mTyCtl_Locale.ScreenHeight - Me.Width) \ 2 > 0 Then
            Me.Top = (mTyCtl_Locale.ScreenHeight - Me.Width) \ 2
        End If
        Exit Sub
    End If
    
    With mTyCtl_Locale
        '���㴰���Y������½Ӹ߶�
        If .ScreenHeight - (.Top + .Height + dblRowsHeight) < 0 Then
            '֤���ؼ�������Ҫ�߶ȱȿؼ����µ�λ��Ҫ��
            If dblRowsHeight < .Top Then
                '֤���ϲ�����װ������,��˿ؼ�����,�����ϲ���
               .DownHeight = dblRowsHeight
               .DownTop = .Top - dblRowsHeight
            ElseIf .Top > .ScreenHeight - (.Top + .Height) Then
                '�����������������,�˷�֧��ʾ������
                .DownTop = 0
                .DownHeight = .Top
            Else '�˷�֧��ʾ������
                .DownHeight = .ScreenHeight - (.Top + .Height)
            End If
        Else
            '֤�������б������ڿؼ��·���ʾ
            .DownHeight = dblRowsHeight
        End If
        
        '���㴰���Y�������������
        If .ScreenWidth - .Left >= dblColsWidth Then
            '������װ�������п�
            .DownWidth = dblColsWidth
            .DownLeft = .Left
        Else
           If .Left + .Width >= dblColsWidth Then
                '������װ�������п�
                .DownLeft = .Left + .Width - dblColsWidth
                .DownWidth = dblColsWidth
           ElseIf .Left + .Width > .ScreenWidth - .Left Then
                '֤����ߴ����ұ�
                .DownWidth = .Left + .Width
                .DownLeft = 0
           Else
                '֤���ұߴ������
                .DownWidth = .ScreenWidth - .Left
                .DownLeft = .Left
           End If
        End If
        
        '���Խ��ж�λ��
        Me.Left = .DownLeft
        Me.Top = .DownTop
        Me.Width = .DownWidth
        Me.Height = .DownHeight
    End With
    
End Sub

Public Sub CalcPosition(ByRef X As Single, ByRef Y As Single, ByVal objBill As Object, Optional blnNoBill As Boolean = False)
    '----------------------------------------------------------------------
    '���ܣ� ����X,Y��ʵ�����꣬��������Ļ���������
    '������ X---���غ��������
    '       Y---�������������
    '----------------------------------------------------------------------
    Dim objPoint As POINTAPI
    Call ClientToScreen(objBill.hWnd, objPoint)
    If blnNoBill Then
        X = objPoint.X * 15
        Y = objPoint.Y * 15 + objBill.Height
    Else
        X = objPoint.X * 15 + objBill.CellLeft
        Y = objPoint.Y * 15 + objBill.CellTop + objBill.CellHeight
    End If
End Sub

Public Function zl_vsGrid_Para_Save(ByVal lngSys As Long, ByVal lngModule As Long, ByVal vsGrid As VSFlexGrid, ByVal strKey As String, _
    Optional blnǿ�Ʊ��� As Boolean = False) As Boolean
    '------------------------------------------------------------------------------
    '����:����vsFlex�Ŀ��ȵ�ע���
    '����:vsGrid-��Ӧ������ؼ�
    '     strCaption-������
    '     strKey-����
    '����:����ɹ�,����True,���򷵻�False
    '����:���˺�
    '����:2008/03/03
    '------------------------------------------------------------------------------
    Dim intCol As Integer, strCol As String, strColCaption As String, intRow As Integer
    Dim zlDatabase As New clsDataBase
    
    If blnǿ�Ʊ��� = False Then
        zl_vsGrid_Para_Save = True
        If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 0 Then Exit Function
    End If
    
    zl_vsGrid_Para_Save = False
    With vsGrid
        strCol = ""
        For intCol = 0 To .Cols - 1
            strCol = strCol & "|" & .ColKey(intCol) & "," & .ColWidth(intCol) & "," & IIf(.ColHidden(intCol), 1, 0)
        Next
    End With
    If strCol <> "" Then strCol = Mid(strCol, 2)
    '�����ʽ:������,�п�,������|������,�п�,������|...
    zlDatabase.SetPara strKey, strCol, lngSys, lngModule
    zl_vsGrid_Para_Save = True
End Function

Public Function zl_vsGrid_Para_Restore(ByVal lngSys As Long, ByVal lngModule As Long, ByVal vsGrid As VSFlexGrid, ByVal strKey As String, _
    Optional blnǿ�ƻָ����� As Boolean = False) As Boolean
    '------------------------------------------------------------------------------
    '����:�����ݿ��лָ�����Ŀ��ȵ���Ϣ
    '����:vsGrid-��Ӧ������ؼ�
    '     strCaption-������
    '     strKey-����
    '     blnǿ�ƻָ�����-�����Ƿ񽫱���Ĳ���ֵ,����ǿ�ƻָ�
    '����:�ָ��ɹ�,����True,���򷵻�False
    '����:���˺�
    '����:2008/03/03
    '------------------------------------------------------------------------------
    Dim strParaValue As String, intCols As Integer, arrReg As Variant, ArrTemp As Variant, intCol As Integer, intRow As Integer
    Dim intTemp As Integer, strColName As String
    Dim zlDatabase As New clsDataBase
    
    If blnǿ�ƻָ����� = False Then
        zl_vsGrid_Para_Restore = True
        If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 0 Then Exit Function
    End If
    strParaValue = zlDatabase.GetPara(strKey, lngSys, lngModule)
    
    
    zl_vsGrid_Para_Restore = False
    If strParaValue = "" Then Exit Function
    'strParaValue:�����ʽ:������,�п�,������|������,�п�,������|...
    Err = 0: On Error GoTo Errhand:
    arrReg = Split(strParaValue, "|")
    intCols = UBound(arrReg) + 1
    With vsGrid
        For intCol = 0 To intCols - 1
            ArrTemp = Split(arrReg(intCol) & ",,", ",")
            strColName = ArrTemp(0)
            intTemp = .ColIndex(strColName)
            If intTemp <> -1 Then
                .ColWidth(intTemp) = Val(ArrTemp(1))
                If Val(ArrTemp(2)) = 1 Then
                    .ColHidden(intTemp) = True
                Else
                    .ColHidden(intTemp) = False
                End If
                If .ColWidth(intTemp) = 0 Then .ColHidden(intTemp) = True
                .ColPosition(.ColIndex(strColName)) = intCol
            End If
        Next
    End With
    zl_vsGrid_Para_Restore = True
    Exit Function
Errhand:
End Function

Private Sub cmdCancel_Click()
    Set mrsOutSel = Nothing
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Call vsSelect_DblClick
End Sub

Private Sub cmdPatiMerge_Click()
    Dim lng����ID As Long
    Dim lng����ID1 As Long
    Dim i As Long
    
    With vsSelect
        For i = 1 To .Rows - 1
            If .IsSelected(i) Then
                If lng����ID <> 0 Then
                    If lng����ID1 <> 0 Then Exit For
                    lng����ID1 = Val(.TextMatrix(i, .ColIndex("����ID")))
                    If lng����ID1 <> 0 Then Exit For
                Else
                    lng����ID = Val(.TextMatrix(i, .ColIndex("����ID")))
                End If
            End If
        Next
    End With
    If frmMergePatient.zlShowPatiMerge(Me, lng����ID1, lng����ID) = False Then Exit Sub
    
    '������
    If zlBindingData = False Then Exit Sub
    '��ʼ���ؼ�λ��
    Call InitCtrlLocal
    '��������λ��
    Call ReSetWindowsFormLocal
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyReturn
        Call SelectedItem
    Case vbKeyEscape
        Unload Me: Exit Sub
    End Select
End Sub

Private Sub Form_Load()
    Call SetWindowsProperty
    '������
    If zlBindingData = False Then Exit Sub
    '��ʼ���ؼ�λ��
    Call InitCtrlLocal
    '��������λ��
    Call ReSetWindowsFormLocal
End Sub

Private Sub Form_Resize()
    Dim lngLineWidth As Long
    Err = 0: On Error Resume Next
    lngLineWidth = IIf(shpLine.Visible, 20, 0)
    With shpLine
        .Left = ScaleLeft
        .Top = ScaleTop
        .Width = ScaleWidth
        .Height = ScaleHeight
    End With
    With picInfo
        .Left = ScaleLeft + lngLineWidth
        .Top = ScaleTop + lngLineWidth
        .Width = ScaleWidth - lngLineWidth * 2
    End With
    With picCmd
        .Left = ScaleLeft + lngLineWidth
        .Top = ScaleHeight - .Height - lngLineWidth
        .Width = ScaleWidth - lngLineWidth * 2
    End With
    
    With vsSelect
        .Left = ScaleLeft + lngLineWidth
        .Width = ScaleWidth - lngLineWidth * 2
        .Top = ScaleTop + IIf(picInfo.Visible, picInfo.Top + picInfo.Height, 0)
        .Height = IIf(picCmd.Visible, picCmd.Top, ScaleHeight - lngLineWidth) - .Top
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mstr������ <> "" Then zl_vsGrid_Para_Save mlngSys, mlngModule, vsSelect, mstr������
End Sub

Private Sub picCmd_Resize()
    Err = 0: On Error Resume Next
    With picCmd
        cmdCancel.Left = .ScaleWidth - cmdCancel.Width - 100
        cmdOK.Left = cmdCancel.Left - cmdOK.Width - 50
    End With
End Sub

Private Sub vsSelect_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Call ReSetWindowsFormLocal
End Sub
Private Sub vsSelect_DblClick()
    Call SelectedItem
End Sub

Private Sub vsSelect_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    Call SelectedItem
End Sub
Private Function OpenSQLRecord(ByVal strSql As String, strLog As String) As ADODB.Recordset
    Static cmdData As New ADODB.Command
    Dim strPar As String, arrPar As Variant
    Dim lngLeft As Long, lngRight As Long
    Dim intMax As Integer, i As Integer
    Dim strSeq As String, varValue As Variant
    
    '�����Զ���[x]����
    lngLeft = InStr(1, strSql, "[")
    Do While lngLeft > 0
        lngRight = InStr(lngLeft + 1, strSql, "]")
        
        '������������"[����]����"
        strSeq = Mid(strSql, lngLeft + 1, lngRight - lngLeft - 1)
        If IsNumeric(strSeq) Then
            i = CInt(strSeq)
            strPar = strPar & "," & i
            If i > intMax Then intMax = i
        End If
        
        lngLeft = InStr(lngRight + 1, strSql, "[")
    Loop

    '�滻Ϊ"?"����
    strLog = strSql
    For i = 1 To intMax
        strSql = Replace(strSql, "[" & i & "]", "?")
        
        '��������SQL���ٵ����
        varValue = marrInput(i - 1)
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '����
            strLog = Replace(strLog, "[" & i & "]", varValue)
        Case "String" '�ַ�
            strLog = Replace(strLog, "[" & i & "]", "'" & Replace(varValue, "'", "''") & "'")
        Case "Date" '����
            strLog = Replace(strLog, "[" & i & "]", "To_Date('" & Format(varValue, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')")
        End Select
    Next

    '���ԭ�в���:��Ȼ�����ظ�ִ��
    cmdData.CommandText = "" '��Ϊ����ʱ�����������
    Do While cmdData.Parameters.count > 0
        cmdData.Parameters.Delete 0
    Loop

    '�����µĲ���
    arrPar = Split(Mid(strPar, 2), ",")
    For i = 0 To UBound(arrPar)
        varValue = marrInput((arrPar(i) - 1))
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '����
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adVarNumeric, adParamInput, 30, varValue)
        Case "String" '�ַ�
            intMax = zlCommFun.ActualLen(varValue)
            If intMax <= 2000 Then
                intMax = IIf(intMax <= 200, 200, 2000)
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adVarChar, adParamInput, intMax, varValue)
            Else
                If intMax < 4000 Then intMax = 4000
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adLongVarChar, adParamInput, intMax, varValue)
            End If
        Case "Date" '����
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adDBTimeStamp, adParamInput, , varValue)
        End Select
    Next

    'ִ�з��ؼ�¼��
    If cmdData.ActiveConnection Is Nothing Then
        Set cmdData.ActiveConnection = IIf(mcnOracle Is Nothing, gcnOracle, mcnOracle) '���Ƚ���
    End If
    cmdData.CommandText = strSql
    
    Call SQLTest(App.ProductName, mstrWinTittle & "ѡ��", strLog)
    Set OpenSQLRecord = cmdData.Execute
    Call SQLTest
End Function

Private Sub vsSelect_SelChange()
    '------------------------------------------------------------------------------
    '����:���úϲ���ť�Ƿ����
    '����:����
    '����:20012/09/28
    '�����:54215
    '------------------------------------------------------------------------------
     If cmdPatiMerge.Visible = True Then cmdPatiMerge.Enabled = vsSelect.SelectedRows >= 2

End Sub