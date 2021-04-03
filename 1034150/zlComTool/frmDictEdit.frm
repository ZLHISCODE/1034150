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
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmd�ϼ� 
      Caption         =   "��"
      Height          =   270
      Left            =   2430
      TabIndex        =   8
      Top             =   1875
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.CheckBox Chk�Ƿ� 
      Caption         =   "Check1"
      Height          =   195
      Index           =   0
      Left            =   315
      TabIndex        =   7
      Top             =   2445
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.CheckBox chkĩ�� 
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
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   2970
      TabIndex        =   2
      Top             =   630
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
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
      Caption         =   "����"
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
Dim mstrOwner As String       '��ǰ�༭�������������
Dim mstrTable As String       '��ǰ�༭�ı���
Dim mstr���� As String        '��ǰ�༭�ļ�¼��ʶ
Dim mint����  As Integer      '�����ֶε����
Dim mint����  As Integer      '�����ֶε����
Dim mint����  As Integer      '�����ֶε����
Dim mint���볤��  As Integer  '���õ�Դ

Dim mlng����() As Long        '�ֶ�����,Ϊ1��ʾ������,2��ʾ����
Dim mblnChange As Boolean

Private Sub cmd�ϼ�_Click()
    Dim vRect As RECT
    Dim rsTmp As ADODB.Recordset
    Dim blnCancel As Boolean
    
    vRect = GetControlRect(txtEdit(cmd�ϼ�.Tag).hWnd)
    
     gstrSQL = "Select * From (select '0' as ID,null as �ϼ�ID,'' as ����,'ȫ��' as ����,0 as ĩ�� From dual " & _
              "union all Select to_char(����) as ID,nvl(�ϼ�,0) As �ϼ�ID, to_char(����) as ����, ����, ĩ�� " & _
             " From " & mstrOwner & "." & mstrTable & " Where nvl(ĩ��,0)=0 ) Order by nvl(�ϼ�ID,0),Id "
             
     Set rsTmp = zlDatabase.ShowSelect(Me, gstrSQL, 1, "��Ŀ", , , , , , False, vRect.Left, vRect.Top, txtEdit(cmd�ϼ�.Tag).Height, blnCancel, , True)
            
    If Not blnCancel Then
        If Not rsTmp Is Nothing Then
            txtEdit(cmd�ϼ�.Tag).Tag = IIf(txtEdit(cmd�ϼ�.Tag).Text = "", "ȫ��", txtEdit(cmd�ϼ�.Tag).Text)
            txtEdit(cmd�ϼ�.Tag).Text = IIf(IsNull(rsTmp("����")), "", rsTmp("����"))
        End If
    End If

End Sub

Private Sub Form_Activate()
    txtEdit(mint����).SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = False Then Exit Sub
    If MsgBox("�����������˳��Ļ������е��޸Ķ�������Ч��" & vbCrLf & "�Ƿ�ȷ���˳���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer
    If IsValid() = False Then Exit Sub
    If Save����() = False Then Exit Sub
    If mstr���� <> "" Then
        mblnChange = False
        Unload Me
        Exit Sub
    End If
    mstr���� = ""
    chkLog.Value = 0
    For i = 1 To lblEdit.Count - 1
        txtEdit(i).Text = ""
    Next
    If mstr���� = "" Then txtEdit(mint����).Text = zlDatabase.GetMax(mstrOwner & "." & mstrTable, "����", mint���볤��)
    mblnChange = False
    txtEdit(mint����).SetFocus
'    Unload Me
End Sub

Private Function IsValid() As Boolean
'����:����������������Ƿ���Ч
'����:
'����ֵ:��Ч����True,����ΪFalse
    Dim i As Integer
    Dim strTemp As String
    For i = 1 To lblEdit.Count - 1
        strTemp = Trim(txtEdit(i).Text)
        If zlCommFun.StrIsValid(strTemp, txtEdit(i).MaxLength, txtEdit(i).hWnd) = False Then
            zlControl.TxtSelAll txtEdit(i)
            Exit Function
        End If
        If i = mint���� Or i = mint���� Then
            If Len(strTemp) = 0 Then
                MsgBox lblEdit(i).Tag & "����Ϊ�ա�", vbExclamation, gstrSysName
                txtEdit(i).Text = ""
                txtEdit(i).SetFocus
                Exit Function
            End If
        End If
        If mlng����(i) = 1 Then
            '�������ֶ�
            If strTemp <> "" And Not IsNumeric(strTemp) Then
                MsgBox lblEdit(i).Tag & "Ӧ���������֡�", vbExclamation, gstrSysName
                zlControl.TxtSelAll txtEdit(i)
                txtEdit(i).SetFocus
                Exit Function
            End If
        
        End If
        If mlng����(i) = 2 Then
            '�������ֶ�
            If strTemp <> "" Then
                If Not IsDate(strTemp) Then
                    MsgBox lblEdit(i).Tag & "�������ڸ�ʽ(yyyy-mm-dd)��", vbExclamation, gstrSysName
                    zlControl.TxtSelAll txtEdit(i)
                    txtEdit(i).SetFocus
                    Exit Function
                End If
                If zlCommFun.ActualLen(strTemp) <> 10 Then
                    MsgBox lblEdit(i).Tag & "���Ȳ���,Ӧ��Ϊ10λ(yyyy-mm-dd)��", vbExclamation, gstrSysName
                    zlControl.TxtSelAll txtEdit(i)
                    txtEdit(i).SetFocus
                    Exit Function
                End If
                Err = 0
                On Error Resume Next
                strTemp = Format(strTemp, "yyyy-mm-dd")
                If Err <> 0 Or Not IsDate(strTemp) Then
                        MsgBox lblEdit(i).Tag & "�������ڸ�ʽ(yyyy-mm-dd)��", vbExclamation, gstrSysName
                        zlControl.TxtSelAll txtEdit(i)
                        txtEdit(i).SetFocus
                        Exit Function
                End If
            End If
            
        End If
    Next
    
    If chkĩ��.Visible = True Then
        If chkĩ��.Value <> 1 And chkLog.Value = 1 Then
            MsgBox "ֻ��ĩ����Ŀ��������Ϊȱʡֵ��", vbInformation, gstrSysName
            chkLog.Value = 0
            Exit Function
        End If
    End If

    IsValid = True
End Function

Private Function Save����() As Boolean
'����:����������ݽ��б���
'����:
'����ֵ:�ɹ�����True,����ΪFalse
    Dim strSQL As String
    Dim strTemp As String
    Dim i As Integer
    Dim lngSystem As Long
    
    With frmDictManager.cmbSys
        lngSystem = .ItemData(.ListIndex) \ 100
    End With
    
    On Error GoTo errHandle
    If mstr���� = "" Then       '����һ����¼
        strSQL = "insert into " & mstrOwner & "." & mstrTable & " ("
        For i = 1 To lblEdit.Count - 1
            strSQL = strSQL & lblEdit(i).Tag & ","
            If mlng����(i) = 2 Then
                strTemp = strTemp & "to_Date('" & Format(Trim(txtEdit(i).Text), "yyyy-mm-dd") & "','yyyy-mm-dd'),"
            Else
                strTemp = strTemp & "'" & Trim(txtEdit(i).Text) & "',"
            End If
        Next
        
        For i = 1 To Chk�Ƿ�.Count - 1
            strSQL = strSQL & Chk�Ƿ�(i).Tag & ","
            strTemp = strTemp & IIf(Chk�Ƿ�(i).Value = 1, "1,", "0,")
        Next
        
        If chkĩ��.Tag <> "" Then
            strSQL = strSQL & chkĩ��.Tag & ","
            strTemp = strTemp & IIf(chkĩ��.Value = 1, "1,", "0,")
        End If
        
        If chkLog.Visible = False Then
            strSQL = Left(strSQL, Len(strSQL) - 1)
            strTemp = Left(strTemp, Len(strTemp) - 1)
        Else
            strSQL = strSQL & chkLog.Tag
            strTemp = strTemp & IIf(chkLog.Value = 1, "1", "0")
        End If
        
        
        strSQL = strSQL & ") values ( " & strTemp & ")"
    Else    '�޸�
        strSQL = "update " & mstrOwner & "." & mstrTable & " set "
        For i = 1 To lblEdit.Count - 1
            If mlng����(i) = 2 Then
                strSQL = strSQL & lblEdit(i).Tag & "=" & "to_Date('" & Format(Trim(txtEdit(i).Text), "yyyy-mm-dd") & "','yyyy-mm-dd'),"
            Else
                strSQL = strSQL & lblEdit(i).Tag & "=" & "'" & Trim(txtEdit(i).Text) & "',"
            End If
            
        Next
        
        For i = 1 To Chk�Ƿ�.Count - 1
            strSQL = strSQL & Chk�Ƿ�(i).Tag & "=" & IIf(Chk�Ƿ�(i).Value = 1, "1,", "0,")
        Next
        
        If chkĩ��.Tag <> "" Then
            strSQL = strSQL & chkĩ��.Tag & "=" & IIf(chkĩ��.Value = 1, "1,", "0,")
        End If
        
        If chkLog.Visible = False Then
            strSQL = Left(strSQL, Len(strSQL) - 1)
        Else
            strSQL = strSQL & chkLog.Tag & "=" & IIf(chkLog.Value = 1, "1", "0")
        End If
        strSQL = strSQL & " where ���� = '" & mstr���� & "'"
    End If
    gcnOracle.BeginTrans
    If chkLog.Tag = "ȱʡ��־" And chkLog.Value = 1 Then
        strTemp = "update " & mstrOwner & "." & mstrTable & " set ȱʡ��־=0"
        '�ù��̽��з�װ
        gstrSQL = "ZL_�ֵ����_execute('" & Replace(strTemp, "'", "''") & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    End If
    '�ù��̽��з�װ
    gstrSQL = "ZL_�ֵ����_execute('" & Replace(strSQL, "'", "''") & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    If chkĩ��.Tag <> "" Then
        If txtEdit(cmd�ϼ�.Tag).Tag <> "" Then
            '�����ϼ�
            Call UpdateMain(0)
        Else
            Call UpdateMain(IIf(chkĩ��.Value = 1, "1", "0"))
        End If
    Else
        Call UpdateMain(1)
    End If
    gcnOracle.CommitTrans
    Save���� = True
    Exit Function

errHandle:
    If ErrCenter() = 1 Then Resume
    gcnOracle.RollbackTrans
End Function

Private Sub UpdateMain(ByVal strĩ�� As String)
'���ܣ�����������
    Dim lst As ListItem
    Dim ch As ColumnHeader
    Dim lngCount As Long
    Dim strTemp As String
    
    If strĩ�� = 0 Then
        Call frmDictManager.frmRefresh
        Exit Sub
    End If
    With frmDictManager.lvwMain
        If mstr���� = "" Then
'            If strĩ�� = 1 Then
                Set lst = .ListItems.Add(, "C" & txtEdit(mint����).Text, txtEdit(mint����).Text, "Item", "Item")
                If .ListItems.Count = 1 Then
                    lst.Selected = True
                End If
'            Else
'                '����һ�����
'            End If
        Else
            If mstr���� <> txtEdit(mint����).Text Then
                
                '����ı䣬��Ҫ�޸���Keyֵ
                .ListItems.Remove .SelectedItem.Key
                Set lst = .ListItems.Add(, "C" & txtEdit(mint����).Text, txtEdit(mint����).Text, "Item", "Item")
                lst.Selected = True
                lst.EnsureVisible
            Else
                Set lst = .SelectedItem
                lst.Text = txtEdit(mint����).Text
            End If
        End If
        
        For Each ch In .ColumnHeaders
            strTemp = ch.Text
            If strTemp <> "����" Then
                For lngCount = 1 To lblEdit.Count - 1
                    If strTemp = lblEdit(lngCount).Tag Then '��ʾ��ͬ�ֶ�
                        Exit For
                    End If
                Next
                
                If lngCount < lblEdit.Count Then
                    '�ڱ༭�����ҵ�
                    If mlng����(lngCount) = 2 Then
                        lst.SubItems(ch.SubItemIndex) = Format(Trim(txtEdit(lngCount).Text), "yyyy-mm-dd")
                    Else
                        If lblEdit(lngCount).Tag = "�ϼ�" Then
                            lst.SubItems(ch.SubItemIndex) = txtEdit(lngCount).Tag
                        Else
                            lst.SubItems(ch.SubItemIndex) = txtEdit(lngCount).Text
                        End If
                    End If
                Else
                    If strTemp = "ȱʡ��־" Then
                        If chkLog.Value = 1 Then
                            '��ListView�и��е�ֵȫ���
                            For lngCount = 1 To .ListItems.Count
                                .ListItems(lngCount).SubItems(ch.SubItemIndex) = ""
                            Next
                        End If
                        lst.SubItems(ch.SubItemIndex) = IIf(chkLog.Value = 1, "��", "")
                    End If
 
                End If
                Dim intChk As Integer
                If strTemp Like "�Ƿ�*" Then
                    For intChk = 1 To Chk�Ƿ�.Count - 1
                        If strTemp = Chk�Ƿ�(intChk).Tag Then
                            lst.SubItems(ch.SubItemIndex) = IIf(Chk�Ƿ�(intChk).Value = 1, "��", "")
                        End If
                    Next
                End If
            End If
        Next
    End With
    Call frmDictManager.SetMenu
End Sub

Public Function �༭����(ByVal strOwner As String, ByVal strTable As String, Optional str���� As String = "", Optional intĩ�� As Integer = -1, Optional str�ϼ� As String) As Boolean
'����:��������ô��ڽ���ͨѶ�ĳ���
'����:strTable  Ҫ�༭�ı���
'     str����     Ҫ�༭�ı�����ؼ���
'����ֵ:�ɹ�����True,����ΪFalse
    Dim rs����� As New ADODB.Recordset, rsTmp As New ADODB.Recordset
    Dim fld As Field
    Dim lst As ListItem
    Dim sngY As Single     '��ǰ�༭��ĸ߶�
    Dim sngMaxW As Single  '�༭��������
    Dim intTemp As Integer
    Dim intChkTmp As Integer
    
    '��ʼ������
    sngY = 200
    sngMaxW = 0
    mstrOwner = strOwner
    mstrTable = strTable
    mstr���� = str����
    
    mint���볤�� = 0
    mint���� = 0
    mint���� = 0
    chkLog.Tag = ""
    chkĩ��.Tag = ""
    
    On Error Resume Next
    rs�����.CursorLocation = adUseClient
    
    gstrSQL = "select * from " & strOwner & "." & strTable & " where ���� = [1]"
    Set rs����� = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, str����)
    
    ReDim mlng����(0 To rs�����.Fields.Count)
    For Each fld In rs�����.Fields
        If fld.Name = "ȱʡ��־" Then
            '���߼�����
            chkLog.Caption = fld.Name
            chkLog.Tag = fld.Name
            chkLog.Caption = fld.Name & IIf(fld.Name = "ȱʡ��־", "��ע�⣺�����־���������ԣ�", "")
            chkLog.Left = 200
            chkLog.Width = 300 + Me.TextWidth(chkLog.Caption)
            If chkLog.Width + 200 > sngMaxW Then sngMaxW = chkLog.Width + 200
            chkLog.Value = IIf(IIf(IsNull(fld.Value), 0, fld.Value), 1, 0)
            chkLog.Visible = True
            
        ElseIf fld.Name Like "�Ƿ�*" Then
            intChkTmp = Chk�Ƿ�.Count
            Load Chk�Ƿ�(intChkTmp)
            Chk�Ƿ�(intChkTmp).Caption = fld.Name
            Chk�Ƿ�(intChkTmp).Tag = fld.Name
            Chk�Ƿ�(intChkTmp).Left = 200
            Chk�Ƿ�(intChkTmp).Width = 300 + Me.TextWidth(Chk�Ƿ�(intChkTmp).Caption)
            If Chk�Ƿ�(intChkTmp).Width + 200 > sngMaxW Then sngMaxW = Chk�Ƿ�(intChkTmp).Width + 200
            Chk�Ƿ�(intChkTmp).Value = IIf(IIf(IsNull(fld.Value), 0, fld.Value), 1, 0)
            
            Chk�Ƿ�(intChkTmp).Top = sngY
            sngY = sngY + Chk�Ƿ�(intChkTmp).Height + 100
            If Chk�Ƿ�(intChkTmp).Width + Chk�Ƿ�(intChkTmp).Left > sngMaxW Then sngMaxW = Chk�Ƿ�(intChkTmp).Width + Chk�Ƿ�(intChkTmp).Left
            
            Chk�Ƿ�(intChkTmp).Visible = True
            
        ElseIf fld.Name = "ĩ��" Then
            chkĩ��.Caption = fld.Name
            chkĩ��.Tag = fld.Name
            chkĩ��.Left = 200
            chkĩ��.Width = 300 + Me.TextWidth(chkĩ��.Caption)
            If chkĩ��.Width + 200 > sngMaxW Then sngMaxW = chkĩ��.Width + 200
            If intĩ�� <> -1 Then
                chkĩ��.Value = IIf(IIf(IsNull(intĩ��), 0, intĩ��), 1, 0)
            Else
                chkĩ��.Value = IIf(IIf(IsNull(fld.Value), 0, fld.Value), 1, 0)
            End If
            
        Else
            intTemp = lblEdit.Count
            Load lblEdit(intTemp)
            Load txtEdit(intTemp)
            
            If fld.Type = adNumeric Then
                '������
                mlng����(intTemp) = 1
            ElseIf fld.Type = adDate Or fld.Type = adDBTimeStamp Or fld.Type = adDBDate Or fld.Type = adDBTime Then
                mlng����(intTemp) = 2
            End If
            '�����ĸ���ܳ���9
            lblEdit(intTemp).Caption = fld.Name & "(&" & intTemp & ")"
            
            '��¼��һЩ�����ֶε����
            If fld.Name = "����" Then mint���� = intTemp
            If fld.Name = "����" Then mint���� = intTemp
            If fld.Name = "����" Then
                mint���� = intTemp
                mint���볤�� = fld.DefinedSize
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
            
            '����Tab˳��
            lblEdit(intTemp).TabIndex = (intTemp - 1) * 2
            txtEdit(intTemp).TabIndex = (intTemp - 1) * 2 + 1
            If fld.Name = "�ϼ�" Then
                txtEdit(intTemp).Enabled = False
                If txtEdit(intTemp).Text = "" And str�ϼ� <> "" Then
                    If str�ϼ� <> "oot" Then
                        txtEdit(intTemp).Text = str�ϼ�
                    End If
                End If
                cmd�ϼ�.Left = txtEdit(intTemp).Left + txtEdit(intTemp).Width
                cmd�ϼ�.Top = txtEdit(intTemp).Top + 10
                If cmd�ϼ�.Width + txtEdit(intTemp).Width + txtEdit(intTemp).Left > sngMaxW Then sngMaxW = cmd�ϼ�.Left + cmd�ϼ�.Width
                cmd�ϼ�.Visible = True
                cmd�ϼ�.TabIndex = (intTemp - 1) * 2 + 2
                cmd�ϼ�.Tag = intTemp
            End If
            
        End If
    Next
    
    If chkLog.Tag <> "" Then
        chkLog.Top = sngY
        sngY = sngY + chkLog.Height + 100 '�ѿ�ѡ
        chkLog.TabIndex = intTemp * 2
    End If
    
    If mstr���� = "" Then txtEdit(mint����).Text = zlDatabase.GetMax(mstrOwner & "." & strTable, "����", mint���볤��)
    fraSplit.Top = -500
    fraSplit.Left = sngMaxW + 250
    cmdOK.Left = sngMaxW + 500
    cmdCancel.Left = cmdOK.Left
    
    frmDictEdit.Width = cmdOK.Left + cmdOK.Width + 250
    frmDictEdit.Height = sngY + 500
    'Ϊ����ʾ�꼸����ť����ʹ�������ۡ����ڵĸ߶ȱ�֤��һ����ֵ֮��
    If frmDictEdit.Height < 2300 Then frmDictEdit.Height = 2300
    fraSplit.Height = frmDictEdit.Height + 1000
    
    frmDictEdit.Caption = mstrTable & IIf(intĩ�� = 0, "[����]", "[��Ŀ]")
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
    If Index = mint���� Then
        txtEdit(mint����).Text = zlCommFun.SpellCode(txtEdit(Index).Text)
    End If
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
    If lblEdit(Index).Tag = "����" Then
        zlCommFun.OpenIme True
    ElseIf lblEdit(Index).Tag = "����" Or lblEdit(Index).Tag = "����" Or mlng����(Index) = 1 Then
        zlCommFun.OpenIme False
    End If
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    ElseIf lblEdit(Index).Tag = "����" Then
        If InStr("0123456789" & Chr(vbKeyBack) & Chr(vbKeyDelete), Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
