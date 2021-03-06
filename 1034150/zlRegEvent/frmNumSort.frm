VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmNumSortSel 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "号别选择"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   7575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6300
      TabIndex        =   2
      Top             =   615
      Width           =   1100
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   6300
      TabIndex        =   1
      Top             =   135
      Width           =   1100
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshPlan 
      Height          =   5715
      Left            =   60
      TabIndex        =   0
      Top             =   45
      Width           =   6045
      _ExtentX        =   10663
      _ExtentY        =   10081
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      RowHeightMin    =   250
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLinesFixed  =   1
      AllowUserResizing=   1
      MousePointer    =   1
      FormatString    =   "^  号别|^    科室|^      项目|^  医生|时间段|限号|已挂"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmNumSort.frx":0000
      _NumberOfBands  =   1
      _Band(0).Cols   =   8
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
End
Attribute VB_Name = "frmNumSortSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const STR_COMP = "|',~" '分隔字符串
Private mrsPlan As New ADODB.Recordset
Private mlngSect As Long
Private mlngID As Long
Private strSQL As String
Private mstrReturn As String
Private mblnOK As Boolean
Private i As Long
Private mstr号别 As String

Public Function ShowMe(ByVal lng挂号ID As String, strReturn As String, frmParent As Form) As Boolean
'显示本窗体并返回选择的是否正确
On Error GoTo errHandle

    mblnOK = False
    '先找到执行科室和号别
    strSQL = "Select B.号别,B.执行部门ID,A.收费细目ID " & _
        " From 门诊费用记录 A,病人挂号记录 B" & _
        " Where A.记录性质=4 and A.记录状态=1 And A.序号=1 And b.记录性质=1 and b.记录状态=1 and A.NO=B.NO And B.ID=[1]"
    Set mrsPlan = zlDatabase.OpenSQLRecord(strSQL, "号别选择器", lng挂号ID)
    
    If mrsPlan.RecordCount > 0 Then
        mrsPlan.MoveFirst
        mlngSect = mrsPlan!执行部门id
        mlngID = mrsPlan!收费细目ID
        mstr号别 = mrsPlan!号别
    Else
        Exit Function
    End If
    
    Me.Show 1, frmParent
    '号表ID,项目ID,医生ID,医生,科室ID,科室,号类,号别
    If Not mblnOK Then
        strReturn = ",,,,,,,"
    Else
        strReturn = mstrReturn
        ShowMe = True
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SetPlanGrid()
    Dim i As Integer
    
    '初始安排表
    With mshPlan
        .Redraw = False
        .Clear: .Rows = 2: .Cols = 17
        .TextMatrix(0, 0) = "IDS" '号表ID_项目ID_医生ID
        .TextMatrix(0, 1) = "号类"
        .TextMatrix(0, 2) = "号别"
        .TextMatrix(0, 3) = "科室"
        .TextMatrix(0, 4) = "项目"
        .TextMatrix(0, 5) = "医生"
        .TextMatrix(0, 6) = "限号"
        .TextMatrix(0, 7) = "已挂"
        .TextMatrix(0, 8) = "日"
        .TextMatrix(0, 9) = "一"
        .TextMatrix(0, 10) = "二"
        .TextMatrix(0, 11) = "三"
        .TextMatrix(0, 12) = "四"
        .TextMatrix(0, 13) = "五"
        .TextMatrix(0, 14) = "六"
        .TextMatrix(0, 15) = "病案"
        .TextMatrix(0, 16) = "分诊"
        
        If Not Visible Then
            .ColWidth(0) = 0
            .ColWidth(1) = 500
            .ColWidth(2) = 550
            .ColWidth(3) = 1150
            .ColWidth(4) = 1250
            .ColWidth(5) = 700
            .ColWidth(6) = 500
            .ColWidth(7) = 500
            .ColWidth(8) = 280
            .ColWidth(9) = 280
            .ColWidth(10) = 280
            .ColWidth(11) = 280
            .ColWidth(12) = 280
            .ColWidth(13) = 280
            .ColWidth(14) = 280
            .ColWidth(15) = 500
            .ColWidth(16) = 500
        End If
        
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
        .ColAlignment(2) = 1
        .ColAlignment(3) = 1
        .ColAlignment(4) = 1
        .ColAlignment(5) = 1
        .ColAlignment(6) = 1
        .ColAlignment(7) = 1
        .ColAlignment(8) = 4
        .ColAlignment(9) = 4
        .ColAlignment(10) = 4
        .ColAlignment(11) = 4
        .ColAlignment(12) = 4
        .ColAlignment(13) = 4
        .ColAlignment(14) = 4
        .ColAlignment(15) = 4
        .ColAlignment(16) = 4
        
        If Not Visible Then Call RestoreFlexState(mshPlan, App.ProductName & "\" & Me.Name)
        
        For i = 0 To .Cols - 1
            .ColAlignmentFixed(i) = flexAlignCenterCenter
        Next
        
        .RowHeight(0) = 300
        
        .Redraw = True
    End With
End Sub

Private Function ShowPlans(Optional strSort As String = "号别", Optional blnDesc As Boolean) As Boolean
'功能：读取当日安排内容
    Dim i As Integer
    Dim strTime As String, strState As String
    
    On Error GoTo errH
    '该部分语句当时读取各种安排的挂号情况
    strState = _
        "Select A.ID as 安排ID,B.日期,B.已挂数" & vbCrLf & _
        " From 挂号安排 A,病人挂号汇总 B" & vbCrLf & _
        " Where a.科室ID = B.科室ID And a.项目ID = B.项目ID" & vbCrLf & _
        " And Nvl(A.医生ID,0)=Nvl(B.医生ID,0)" & vbCrLf & _
        " And Nvl(A.医生姓名,'医生')=Nvl(B.医生姓名,'医生') And (a.号码=b.号码 or b.号码 is null )" & vbCrLf & _
        " And B.日期=Trunc(Sysdate)"
    '该部分语句取当时所对应的时间段
    strTime = _
        "Select 时间段 From 时间段 Where" & vbCrLf & _
        " ('3000-01-10 '||To_Char(SysDate,'HH24:MI:SS')" & vbCrLf & _
        " Between" & vbCrLf & _
        " Decode(Sign(开始时间 - 终止时间),1,'3000-01-09 '||To_Char(开始时间,'HH24:MI:SS'),'3000-01-10 '||To_Char(开始时间,'HH24:MI:SS'))" & vbCrLf & _
        " And" & vbCrLf & _
        " '3000-01-10 '||To_Char(终止时间,'HH24:MI:SS'))" & vbCrLf & _
        " Or" & vbCrLf & _
        " ('3000-01-10 '||To_Char(SysDate,'HH24:MI:SS')" & vbCrLf & _
        " Between" & vbCrLf & _
        " '3000-01-10 '||To_Char(开始时间,'HH24:MI:SS')" & vbCrLf & _
        " And" & vbCrLf & _
        " Decode(Sign(开始时间 - 终止时间),1,'3000-01-11 '||To_Char(终止时间,'HH24:MI:SS'),'3000-01-10 '||To_Char(终止时间,'HH24:MI:SS')))"
    '该部分语句取时间内的安排及状态
    strSQL = _
        "Select Distinct P.ID,A.日期,P.号码 as 号别,P.号类," & vbCrLf & _
        " Decode(To_Char(SysDate,'D'),'1',P.周日,'2',P.周一,'3',P.周二,'4',P.周三,'5',P.周四,'6',P.周五,'7',P.周六) as 时间," & vbCrLf & _
        " P.科室ID,B.名称 As 科室,P.项目ID,C.名称 As 项目,P.医生ID,P.医生姓名 as 医生,Nvl(A.已挂数,0) as 已挂," & vbCrLf & _
        " F.限号数 as 限号,F.限约数 as 限约,Nvl(P.病案必须,0) as 病案,Nvl(C.项目特性,0) as 急诊," & vbCrLf & _
        " P.周日 as 日,P.周一 as 一,P.周二 as 二,P.周三 as 三,P.周四 as 四,P.周五 as 五,P.周六 as 六," & vbCrLf & _
        " Decode(P.分诊方式,1,'指定',2,'动态',3,'平均',NULL) as 分诊" & vbCrLf & _
        " From 挂号安排 P,(" & strState & ") A,部门表 B,收费项目目录 C,时间段 X,挂号安排限制 F " & vbCrLf & _
        " Where P.ID=A.安排ID(+) And P.科室ID=B.ID And P.项目ID=C.ID AND P.科室id <>[1] and P.项目ID=[2]" & vbCrLf & _
        " And P.Id = f.安排id(+) And" & _
        " Decode(To_Char(sysdate, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7', '周六', Null) =f.限制项目(+)" & vbNewLine & _
        " And SysDate Between C.建档时间 And Nvl(C.撤档时间,To_Date('3000-01-01','YYYY-MM-DD')) " & vbCrLf & _
        " And Decode(To_Char(SysDate,'D'),'1',P.周日,'2',P.周一,'3',P.周二,'4',P.周三,'5',P.周四,'6',P.周五,'7',P.周六,NULL) IN" & vbCrLf & _
        " (" & strTime & ") Order by " & strSort & IIf(blnDesc, " Desc", "") & IIf(strSort <> "号别", ",号别", "")
    
    Set mrsPlan = zlDatabase.OpenSQLRecord(strSQL, "号别选择器", mlngSect, mlngID)
    If mrsPlan.RecordCount > 0 Then
        mrsPlan.MoveFirst
        mshPlan.Rows = mrsPlan.RecordCount + 1
        For i = 1 To mrsPlan.RecordCount
            mshPlan.RowData(i) = mrsPlan!科室ID
            mshPlan.TextMatrix(i, 0) = mrsPlan!ID & "," & mrsPlan!项目ID & "," & IIf(IsNull(mrsPlan!医生ID), 0, mrsPlan!医生ID)
            mshPlan.TextMatrix(i, 1) = IIf(IsNull(mrsPlan!号类), "", mrsPlan!号类)
            mshPlan.TextMatrix(i, 2) = mrsPlan!号别
            mshPlan.TextMatrix(i, 3) = mrsPlan!科室
            mshPlan.TextMatrix(i, 4) = mrsPlan!项目
            mshPlan.TextMatrix(i, 5) = IIf(IsNull(mrsPlan!医生), "", mrsPlan!医生)
            mshPlan.TextMatrix(i, 6) = IIf(IsNull(mrsPlan!限号), "", mrsPlan!限号)
            mshPlan.TextMatrix(i, 7) = IIf(mrsPlan!已挂 = 0, "", mrsPlan!已挂)
            mshPlan.TextMatrix(i, 8) = Left(IIf(IsNull(mrsPlan!日), "", mrsPlan!日), 1)
            mshPlan.TextMatrix(i, 9) = Left(IIf(IsNull(mrsPlan!一), "", mrsPlan!一), 1)
            mshPlan.TextMatrix(i, 10) = Left(IIf(IsNull(mrsPlan!二), "", mrsPlan!二), 1)
            mshPlan.TextMatrix(i, 11) = Left(IIf(IsNull(mrsPlan!三), "", mrsPlan!三), 1)
            mshPlan.TextMatrix(i, 12) = Left(IIf(IsNull(mrsPlan!四), "", mrsPlan!四), 1)
            mshPlan.TextMatrix(i, 13) = Left(IIf(IsNull(mrsPlan!五), "", mrsPlan!五), 1)
            mshPlan.TextMatrix(i, 14) = Left(IIf(IsNull(mrsPlan!六), "", mrsPlan!六), 1)
            mshPlan.TextMatrix(i, 15) = IIf(mrsPlan!病案 = 1, "√", "")
            mshPlan.TextMatrix(i, 16) = IIf(IsNull(mrsPlan!分诊), "", mrsPlan!分诊)
            mrsPlan.MoveNext
        Next
    Else
        Set mrsPlan = Nothing
        Call SetPlanGrid
    End If
    
    mshPlan.Col = 0: mshPlan.ColSel = mshPlan.Cols - 1
    Call mshPlan_EnterCell
    
    ShowPlans = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Set mrsPlan = Nothing
End Function

Private Sub cmdOK_Click()
    mshPlan_DblClick
End Sub

Private Sub mshPlan_DblClick()
    If mshPlan.Row > 0 Then
        If mshPlan.TextMatrix(mshPlan.Row, 0) = "" Then
            MsgBox "没有适合换号的号别。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '号表ID,项目ID,医生ID,医生,科室ID,科室,号类,号别
        mstrReturn = mshPlan.TextMatrix(mshPlan.Row, 0) & "," & mshPlan.TextMatrix(mshPlan.Row, 5) & "," & mshPlan.RowData(mshPlan.Row) & "," & mshPlan.TextMatrix(mshPlan.Row, 3) & "," & mshPlan.TextMatrix(mshPlan.Row, 1) & "," & mshPlan.TextMatrix(mshPlan.Row, 2)
        mblnOK = True
        Unload Me
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    SetPlanGrid
    ShowPlans
End Sub

Private Sub mshPlan_EnterCell()
    Dim i As Integer, blnPre As Boolean
    Dim intRow As Integer, intCol As Integer
    
    blnPre = mshPlan.Redraw
    intRow = mshPlan.Row: intCol = mshPlan.Col
    mshPlan.Redraw = False
    
    For i = 0 To mshPlan.Cols - 1
        mshPlan.Col = i
        mshPlan.CellBackColor = mshPlan.BackColorSel
        mshPlan.CellForeColor = mshPlan.ForeColorSel
    Next
    
    mshPlan.Row = intRow:  mshPlan.Col = intCol
    mshPlan.Redraw = blnPre
End Sub

Private Sub mshPlan_LeaveCell()
    Dim i As Integer, blnPre As Boolean
    
    blnPre = mshPlan.Redraw
    mshPlan.Redraw = False
    
    For i = 0 To mshPlan.Cols - 1
        mshPlan.Col = i
        mshPlan.CellBackColor = mshPlan.BackColor
        mshPlan.CellForeColor = mshPlan.ForeColor
    Next
    
    mshPlan.Redraw = blnPre
End Sub

Private Sub mshPlan_SelChange()
    If mshPlan.Rows = 2 Then Exit Sub
    mshPlan.RowSel = mshPlan.Row
End Sub

