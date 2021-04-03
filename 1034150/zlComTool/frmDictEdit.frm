VERSION 5.00
Begin VB.Form frmDictEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4320
   Icon            =   "frmDictEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   4320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmd上级 
      Caption         =   "…"
      Height          =   270
      Left            =   2430
      TabIndex        =   8
      Top             =   1875
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.CheckBox Chk是否 
      Caption         =   "Check1"
      Height          =   195
      Index           =   0
      Left            =   315
      TabIndex        =   7
      Top             =   2445
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.CheckBox chk末级 
      Caption         =   "Check1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   285
      TabIndex        =   6
      Top             =   3105
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.Frame fraSplit 
      Height          =   4485
      Left            =   2700
      TabIndex        =   5
      Top             =   -510
      Width           =   30
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   0
      Left            =   420
      TabIndex        =   4
      Top             =   1860
      Visible         =   0   'False
      Width           =   2025
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   2970
      TabIndex        =   2
      Top             =   630
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2970
      TabIndex        =   1
      Top             =   180
      Width           =   1100
   End
   Begin VB.CheckBox chkLog 
      Caption         =   "Check1"
      Height          =   315
      Left            =   300
      TabIndex        =   0
      Top             =   2760
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "简码"
      Height          =   180
      Index           =   0
      Left            =   450
      TabIndex        =   3
      Top             =   960
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "frmDictEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mstrOwner As String       '当前编辑表的所有者名字
Dim mstrTable As String       '当前编辑的表名
Dim mstr编码 As String        '当前编辑的记录标识
Dim mint编码  As Integer      '编码字段的序号
Dim mint名称  As Integer      '名称字段的序号
Dim mint简码  As Integer      '简码字段的序号
Dim mint编码长度  As Integer  '调用的源

Dim mlng类型() As Long        '字段类型,为1表示数字型,2表示日期
Dim mblnChange As Boolean

Private Sub cmd上级_Click()
    Dim vRect As RECT
    Dim rsTmp As ADODB.Recordset
    Dim blnCancel As Boolean
    
    vRect = GetControlRect(txtEdit(cmd上级.Tag).hWnd)
    
     gstrSQL = "Select * From (select '0' as ID,null as 上级ID,'' as 编码,'全部' as 名称,0 as 末级 From dual " & _
              "union all Select to_char(编码) as ID,nvl(上级,0) As 上级ID, to_char(编码) as 编码, 名称, 末级 " & _
             " From " & mstrOwner & "." & mstrTable & " Where nvl(末级,0)=0 ) Order by nvl(上级ID,0),Id "
             
     Set rsTmp = zlDatabase.ShowSelect(Me, gstrSQL, 1, "项目", , , , , , False, vRect.Left, vRect.Top, txtEdit(cmd上级.Tag).Height, blnCancel, , True)
            
    If Not blnCancel Then
        If Not rsTmp Is Nothing Then
            txtEdit(cmd上级.Tag).Tag = IIf(txtEdit(cmd上级.Tag).Text = "", "全部", txtEdit(cmd上级.Tag).Text)
            txtEdit(cmd上级.Tag).Text = IIf(IsNull(rsTmp("编码")), "", rsTmp("编码"))
        End If
    End If

End Sub

Private Sub Form_Activate()
    txtEdit(mint名称).SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = False Then Exit Sub
    If MsgBox("如果你就这样退出的话，所有的修改都不会生效。" & vbCrLf & "是否确认退出？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer
    If IsValid() = False Then Exit Sub
    If Save编码() = False Then Exit Sub
    If mstr编码 <> "" Then
        mblnChange = False
        Unload Me
        Exit Sub
    End If
    mstr编码 = ""
    chkLog.Value = 0
    For i = 1 To lblEdit.Count - 1
        txtEdit(i).Text = ""
    Next
    If mstr编码 = "" Then txtEdit(mint编码).Text = zlDatabase.GetMax(mstrOwner & "." & mstrTable, "编码", mint编码长度)
    mblnChange = False
    txtEdit(mint名称).SetFocus
'    Unload Me
End Sub

Private Function IsValid() As Boolean
'功能:分析所输入的内容是否有效
'参数:
'返回值:有效返回True,否则为False
    Dim i As Integer
    Dim strTemp As String
    For i = 1 To lblEdit.Count - 1
        strTemp = Trim(txtEdit(i).Text)
        If zlCommFun.StrIsValid(strTemp, txtEdit(i).MaxLength, txtEdit(i).hWnd) = False Then
            zlControl.TxtSelAll txtEdit(i)
            Exit Function
        End If
        If i = mint编码 Or i = mint名称 Then
            If Len(strTemp) = 0 Then
                MsgBox lblEdit(i).Tag & "不能为空。", vbExclamation, gstrSysName
                txtEdit(i).Text = ""
                txtEdit(i).SetFocus
                Exit Function
            End If
        End If
        If mlng类型(i) = 1 Then
            '数字型字段
            If strTemp <> "" And Not IsNumeric(strTemp) Then
                MsgBox lblEdit(i).Tag & "应该输入数字。", vbExclamation, gstrSysName
                zlControl.TxtSelAll txtEdit(i)
                txtEdit(i).SetFocus
                Exit Function
            End If
        
        End If
        If mlng类型(i) = 2 Then
            '数字型字段
            If strTemp <> "" Then
                If Not IsDate(strTemp) Then
                    MsgBox lblEdit(i).Tag & "不是日期格式(yyyy-mm-dd)。", vbExclamation, gstrSysName
                    zlControl.TxtSelAll txtEdit(i)
                    txtEdit(i).SetFocus
                    Exit Function
                End If
                If zlCommFun.ActualLen(strTemp) <> 10 Then
                    MsgBox lblEdit(i).Tag & "长度不对,应该为10位(yyyy-mm-dd)。", vbExclamation, gstrSysName
                    zlControl.TxtSelAll txtEdit(i)
                    txtEdit(i).SetFocus
                    Exit Function
                End If
                Err = 0
                On Error Resume Next
                strTemp = Format(strTemp, "yyyy-mm-dd")
                If Err <> 0 Or Not IsDate(strTemp) Then
                        MsgBox lblEdit(i).Tag & "不是日期格式(yyyy-mm-dd)。", vbExclamation, gstrSysName
                        zlControl.TxtSelAll txtEdit(i)
                        txtEdit(i).SetFocus
                        Exit Function
                End If
            End If
            
        End If
    Next
    
    If chk末级.Visible = True Then
        If chk末级.Value <> 1 And chkLog.Value = 1 Then
            MsgBox "只有末级项目，才能设为缺省值。", vbInformation, gstrSysName
            chkLog.Value = 0
            Exit Function
        End If
    End If

    IsValid = True
End Function

Private Function Save编码() As Boolean
'功能:对输入的内容进行保存
'参数:
'返回值:成功返回True,否则为False
    Dim strSQL As String
    Dim strTemp As String
    Dim i As Integer
    Dim lngSystem As Long
    
    With frmDictManager.cmbSys
        lngSystem = .ItemData(.ListIndex) \ 100
    End With
    
    On Error GoTo errHandle
    If mstr编码 = "" Then       '新增一条记录
        strSQL = "insert into " & mstrOwner & "." & mstrTable & " ("
        For i = 1 To lblEdit.Count - 1
            strSQL = strSQL & lblEdit(i).Tag & ","
            If mlng类型(i) = 2 Then
                strTemp = strTemp & "to_Date('" & Format(Trim(txtEdit(i).Text), "yyyy-mm-dd") & "','yyyy-mm-dd'),"
            Else
                strTemp = strTemp & "'" & Trim(txtEdit(i).Text) & "',"
            End If
        Next
        
        For i = 1 To Chk是否.Count - 1
            strSQL = strSQL & Chk是否(i).Tag & ","
            strTemp = strTemp & IIf(Chk是否(i).Value = 1, "1,", "0,")
        Next
        
        If chk末级.Tag <> "" Then
            strSQL = strSQL & chk末级.Tag & ","
            strTemp = strTemp & IIf(chk末级.Value = 1, "1,", "0,")
        End If
        
        If chkLog.Visible = False Then
            strSQL = Left(strSQL, Len(strSQL) - 1)
            strTemp = Left(strTemp, Len(strTemp) - 1)
        Else
            strSQL = strSQL & chkLog.Tag
            strTemp = strTemp & IIf(chkLog.Value = 1, "1", "0")
        End If
        
        
        strSQL = strSQL & ") values ( " & strTemp & ")"
    Else    '修改
        strSQL = "update " & mstrOwner & "." & mstrTable & " set "
        For i = 1 To lblEdit.Count - 1
            If mlng类型(i) = 2 Then
                strSQL = strSQL & lblEdit(i).Tag & "=" & "to_Date('" & Format(Trim(txtEdit(i).Text), "yyyy-mm-dd") & "','yyyy-mm-dd'),"
            Else
                strSQL = strSQL & lblEdit(i).Tag & "=" & "'" & Trim(txtEdit(i).Text) & "',"
            End If
            
        Next
        
        For i = 1 To Chk是否.Count - 1
            strSQL = strSQL & Chk是否(i).Tag & "=" & IIf(Chk是否(i).Value = 1, "1,", "0,")
        Next
        
        If chk末级.Tag <> "" Then
            strSQL = strSQL & chk末级.Tag & "=" & IIf(chk末级.Value = 1, "1,", "0,")
        End If
        
        If chkLog.Visible = False Then
            strSQL = Left(strSQL, Len(strSQL) - 1)
        Else
            strSQL = strSQL & chkLog.Tag & "=" & IIf(chkLog.Value = 1, "1", "0")
        End If
        strSQL = strSQL & " where 编码 = '" & mstr编码 & "'"
    End If
    gcnOracle.BeginTrans
    If chkLog.Tag = "缺省标志" And chkLog.Value = 1 Then
        strTemp = "update " & mstrOwner & "." & mstrTable & " set 缺省标志=0"
        '用过程进行封装
        gstrSQL = "ZL_字典管理_execute('" & Replace(strTemp, "'", "''") & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    End If
    '用过程进行封装
    gstrSQL = "ZL_字典管理_execute('" & Replace(strSQL, "'", "''") & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    If chk末级.Tag <> "" Then
        If txtEdit(cmd上级.Tag).Tag <> "" Then
            '改了上级
            Call UpdateMain(0)
        Else
            Call UpdateMain(IIf(chk末级.Value = 1, "1", "0"))
        End If
    Else
        Call UpdateMain(1)
    End If
    gcnOracle.CommitTrans
    Save编码 = True
    Exit Function

errHandle:
    If ErrCenter() = 1 Then Resume
    gcnOracle.RollbackTrans
End Function

Private Sub UpdateMain(ByVal str末级 As String)
'功能：更新主界面
    Dim lst As ListItem
    Dim ch As ColumnHeader
    Dim lngCount As Long
    Dim strTemp As String
    
    If str末级 = 0 Then
        Call frmDictManager.frmRefresh
        Exit Sub
    End If
    With frmDictManager.lvwMain
        If mstr编码 = "" Then
'            If str末级 = 1 Then
                Set lst = .ListItems.Add(, "C" & txtEdit(mint编码).Text, txtEdit(mint名称).Text, "Item", "Item")
                If .ListItems.Count = 1 Then
                    lst.Selected = True
                End If
'            Else
'                '树加一个结点
'            End If
        Else
            If mstr编码 <> txtEdit(mint编码).Text Then
                
                '编码改变，就要修改其Key值
                .ListItems.Remove .SelectedItem.Key
                Set lst = .ListItems.Add(, "C" & txtEdit(mint编码).Text, txtEdit(mint名称).Text, "Item", "Item")
                lst.Selected = True
                lst.EnsureVisible
            Else
                Set lst = .SelectedItem
                lst.Text = txtEdit(mint名称).Text
            End If
        End If
        
        For Each ch In .ColumnHeaders
            strTemp = ch.Text
            If strTemp <> "名称" Then
                For lngCount = 1 To lblEdit.Count - 1
                    If strTemp = lblEdit(lngCount).Tag Then '表示相同字段
                        Exit For
                    End If
                Next
                
                If lngCount < lblEdit.Count Then
                    '在编辑框中找到
                    If mlng类型(lngCount) = 2 Then
                        lst.SubItems(ch.SubItemIndex) = Format(Trim(txtEdit(lngCount).Text), "yyyy-mm-dd")
                    Else
                        If lblEdit(lngCount).Tag = "上级" Then
                            lst.SubItems(ch.SubItemIndex) = txtEdit(lngCount).Tag
                        Else
                            lst.SubItems(ch.SubItemIndex) = txtEdit(lngCount).Text
                        End If
                    End If
                Else
                    If strTemp = "缺省标志" Then
                        If chkLog.Value = 1 Then
                            '把ListView中该列的值全清空
                            For lngCount = 1 To .ListItems.Count
                                .ListItems(lngCount).SubItems(ch.SubItemIndex) = ""
                            Next
                        End If
                        lst.SubItems(ch.SubItemIndex) = IIf(chkLog.Value = 1, "√", "")
                    End If
 
                End If
                Dim intChk As Integer
                If strTemp Like "是否*" Then
                    For intChk = 1 To Chk是否.Count - 1
                        If strTemp = Chk是否(intChk).Tag Then
                            lst.SubItems(ch.SubItemIndex) = IIf(Chk是否(intChk).Value = 1, "√", "")
                        End If
                    Next
                End If
            End If
        Next
    End With
    Call frmDictManager.SetMenu
End Sub

Public Function 编辑编码(ByVal strOwner As String, ByVal strTable As String, Optional str编码 As String = "", Optional int末级 As Integer = -1, Optional str上级 As String) As Boolean
'功能:用来与调用窗口进行通讯的程序
'参数:strTable  要编辑的表名
'     str编码     要编辑的表的主关键字
'返回值:成功返回True,否则为False
    Dim rs编码表 As New ADODB.Recordset, rsTmp As New ADODB.Recordset
    Dim fld As Field
    Dim lst As ListItem
    Dim sngY As Single     '当前编辑框的高度
    Dim sngMaxW As Single  '编辑框的最大宽度
    Dim intTemp As Integer
    Dim intChkTmp As Integer
    
    '初始化变量
    sngY = 200
    sngMaxW = 0
    mstrOwner = strOwner
    mstrTable = strTable
    mstr编码 = str编码
    
    mint编码长度 = 0
    mint名称 = 0
    mint简码 = 0
    chkLog.Tag = ""
    chk末级.Tag = ""
    
    On Error Resume Next
    rs编码表.CursorLocation = adUseClient
    
    gstrSQL = "select * from " & strOwner & "." & strTable & " where 编码 = [1]"
    Set rs编码表 = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, str编码)
    
    ReDim mlng类型(0 To rs编码表.Fields.Count)
    For Each fld In rs编码表.Fields
        If fld.Name = "缺省标志" Then
            '是逻辑类型
            chkLog.Caption = fld.Name
            chkLog.Tag = fld.Name
            chkLog.Caption = fld.Name & IIf(fld.Name = "缺省标志", "（注意：这个标志具有排它性）", "")
            chkLog.Left = 200
            chkLog.Width = 300 + Me.TextWidth(chkLog.Caption)
            If chkLog.Width + 200 > sngMaxW Then sngMaxW = chkLog.Width + 200
            chkLog.Value = IIf(IIf(IsNull(fld.Value), 0, fld.Value), 1, 0)
            chkLog.Visible = True
            
        ElseIf fld.Name Like "是否*" Then
            intChkTmp = Chk是否.Count
            Load Chk是否(intChkTmp)
            Chk是否(intChkTmp).Caption = fld.Name
            Chk是否(intChkTmp).Tag = fld.Name
            Chk是否(intChkTmp).Left = 200
            Chk是否(intChkTmp).Width = 300 + Me.TextWidth(Chk是否(intChkTmp).Caption)
            If Chk是否(intChkTmp).Width + 200 > sngMaxW Then sngMaxW = Chk是否(intChkTmp).Width + 200
            Chk是否(intChkTmp).Value = IIf(IIf(IsNull(fld.Value), 0, fld.Value), 1, 0)
            
            Chk是否(intChkTmp).Top = sngY
            sngY = sngY + Chk是否(intChkTmp).Height + 100
            If Chk是否(intChkTmp).Width + Chk是否(intChkTmp).Left > sngMaxW Then sngMaxW = Chk是否(intChkTmp).Width + Chk是否(intChkTmp).Left
            
            Chk是否(intChkTmp).Visible = True
            
        ElseIf fld.Name = "末级" Then
            chk末级.Caption = fld.Name
            chk末级.Tag = fld.Name
            chk末级.Left = 200
            chk末级.Width = 300 + Me.TextWidth(chk末级.Caption)
            If chk末级.Width + 200 > sngMaxW Then sngMaxW = chk末级.Width + 200
            If int末级 <> -1 Then
                chk末级.Value = IIf(IIf(IsNull(int末级), 0, int末级), 1, 0)
            Else
                chk末级.Value = IIf(IIf(IsNull(fld.Value), 0, fld.Value), 1, 0)
            End If
            
        Else
            intTemp = lblEdit.Count
            Load lblEdit(intTemp)
            Load txtEdit(intTemp)
            
            If fld.Type = adNumeric Then
                '数字型
                mlng类型(intTemp) = 1
            ElseIf fld.Type = adDate Or fld.Type = adDBTimeStamp Or fld.Type = adDBDate Or fld.Type = adDBTime Then
                mlng类型(intTemp) = 2
            End If
            '快捷字母不能超过9
            lblEdit(intTemp).Caption = fld.Name & "(&" & intTemp & ")"
            
            '记录下一些特殊字段的序号
            If fld.Name = "名称" Then mint名称 = intTemp
            If fld.Name = "简码" Then mint简码 = intTemp
            If fld.Name = "编码" Then
                mint编码 = intTemp
                mint编码长度 = fld.DefinedSize
            End If
            lblEdit(intTemp).Tag = fld.Name
            lblEdit(intTemp).Left = 200
            txtEdit(intTemp).Left = lblEdit(intTemp).Left + lblEdit(intTemp).Width + 100
            
            If fld.Type = adVarChar Then
                txtEdit(intTemp).MaxLength = fld.DefinedSize
                txtEdit(intTemp).Width = 300 + fld.DefinedSize * 100
            ElseIf fld.Type = adDate Or fld.Type = adDBTimeStamp Or fld.Type = adDBDate Or fld.Type = adDBTime Then
                txtEdit(intTemp).MaxLength = 10
                txtEdit(intTemp).Width = 300 + fld.Precision * 100
            Else
                txtEdit(intTemp).MaxLength = fld.Precision
                txtEdit(intTemp).Width = 300 + fld.Precision * 100
            End If
            If txtEdit(intTemp).Width > 3000 Then txtEdit(intTemp).Width = 3000
            If chkLog.Width + 200 > sngMaxW Then sngMaxW = chkLog.Width + 200
            If fld.Type = adDate Or fld.Type = adDBTimeStamp Or fld.Type = adDBDate Or fld.Type = adDBTime Then
                txtEdit(intTemp).Text = Format(fld.Value, "yyyy-mm-dd")
            Else
                txtEdit(intTemp).Text = IIf(IsNull(fld.Value), "", fld.Value)
            End If
            txtEdit(intTemp).Top = sngY
            lblEdit(intTemp).Top = txtEdit(intTemp).Top + 75
            sngY = sngY + txtEdit(intTemp).Height + 100
            If txtEdit(intTemp).Width + txtEdit(intTemp).Left > sngMaxW Then sngMaxW = txtEdit(intTemp).Width + txtEdit(intTemp).Left
            lblEdit(intTemp).Visible = True
            txtEdit(intTemp).Visible = True
            
            '设置Tab顺序
            lblEdit(intTemp).TabIndex = (intTemp - 1) * 2
            txtEdit(intTemp).TabIndex = (intTemp - 1) * 2 + 1
            If fld.Name = "上级" Then
                txtEdit(intTemp).Enabled = False
                If txtEdit(intTemp).Text = "" And str上级 <> "" Then
                    If str上级 <> "oot" Then
                        txtEdit(intTemp).Text = str上级
                    End If
                End If
                cmd上级.Left = txtEdit(intTemp).Left + txtEdit(intTemp).Width
                cmd上级.Top = txtEdit(intTemp).Top + 10
                If cmd上级.Width + txtEdit(intTemp).Width + txtEdit(intTemp).Left > sngMaxW Then sngMaxW = cmd上级.Left + cmd上级.Width
                cmd上级.Visible = True
                cmd上级.TabIndex = (intTemp - 1) * 2 + 2
                cmd上级.Tag = intTemp
            End If
            
        End If
    Next
    
    If chkLog.Tag <> "" Then
        chkLog.Top = sngY
        sngY = sngY + chkLog.Height + 100 '把可选
        chkLog.TabIndex = intTemp * 2
    End If
    
    If mstr编码 = "" Then txtEdit(mint编码).Text = zlDatabase.GetMax(mstrOwner & "." & strTable, "编码", mint编码长度)
    fraSplit.Top = -500
    fraSplit.Left = sngMaxW + 250
    cmdOK.Left = sngMaxW + 500
    cmdCancel.Left = cmdOK.Left
    
    frmDictEdit.Width = cmdOK.Left + cmdOK.Width + 250
    frmDictEdit.Height = sngY + 500
    '为了显示完几个按钮，且使窗口美观。窗口的高度保证在一定的值之上
    If frmDictEdit.Height < 2300 Then frmDictEdit.Height = 2300
    fraSplit.Height = frmDictEdit.Height + 1000
    
    frmDictEdit.Caption = mstrTable & IIf(int末级 = 0, "[分类]", "[项目]")
    frmDictEdit.txtEdit(1).SetFocus
    
    mblnChange = False
    frmDictEdit.Show vbModal
End Function

Private Sub chkLog_Click()
    mblnChange = True
End Sub

Private Sub chkLog_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtEdit_Change(Index As Integer)
    mblnChange = True
    On Error Resume Next
    If Index = mint名称 Then
        txtEdit(mint简码).Text = zlCommFun.SpellCode(txtEdit(Index).Text)
    End If
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
    If lblEdit(Index).Tag = "名称" Then
        zlCommFun.OpenIme True
    ElseIf lblEdit(Index).Tag = "编码" Or lblEdit(Index).Tag = "简码" Or mlng类型(Index) = 1 Then
        zlCommFun.OpenIme False
    End If
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    ElseIf lblEdit(Index).Tag = "编码" Then
        If InStr("0123456789" & Chr(vbKeyBack) & Chr(vbKeyDelete), Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
