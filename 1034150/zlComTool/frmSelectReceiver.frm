VERSION 5.00
Begin VB.Form frmSelectReceiver 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�ռ���ѡ��"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6375
   Icon            =   "frmSelectReceiver.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame fra 
      Height          =   75
      Index           =   2
      Left            =   -30
      TabIndex        =   23
      Top             =   1305
      Width           =   8145
   End
   Begin VB.Frame fra 
      Height          =   75
      Index           =   0
      Left            =   -30
      TabIndex        =   21
      Top             =   510
      Width           =   8145
   End
   Begin VB.ComboBox cmbSystem 
      Height          =   300
      Left            =   1875
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   180
      Width           =   4365
   End
   Begin VB.OptionButton optPick 
      Caption         =   "������Ա(&A)"
      Height          =   195
      Index           =   0
      Left            =   450
      TabIndex        =   19
      Top             =   720
      Width           =   1365
   End
   Begin VB.OptionButton optPick 
      Caption         =   "��������Ա(&D)"
      Height          =   195
      Index           =   1
      Left            =   2265
      TabIndex        =   18
      Top             =   720
      Width           =   1485
   End
   Begin VB.OptionButton optPick 
      Caption         =   "��������Ա(&F)"
      Height          =   195
      Index           =   2
      Left            =   4260
      TabIndex        =   17
      Top             =   720
      Width           =   1485
   End
   Begin VB.Frame fra 
      Height          =   3945
      Index           =   1
      Left            =   165
      TabIndex        =   10
      Top             =   1950
      Width           =   6045
      Begin VB.ListBox lst 
         Height          =   3480
         Index           =   0
         ItemData        =   "frmSelectReceiver.frx":000C
         Left            =   240
         List            =   "frmSelectReceiver.frx":000E
         TabIndex        =   16
         Top             =   300
         Width           =   2385
      End
      Begin VB.ListBox lst 
         Height          =   3480
         Index           =   1
         ItemData        =   "frmSelectReceiver.frx":0010
         Left            =   3450
         List            =   "frmSelectReceiver.frx":0012
         TabIndex        =   15
         Top             =   270
         Width           =   2385
      End
      Begin VB.CommandButton cmdFunc 
         Caption         =   "<<"
         Height          =   350
         Index           =   0
         Left            =   2760
         MousePointer    =   1  'Arrow
         TabIndex        =   14
         ToolTipText     =   "ȫ���Ƴ�"
         Top             =   2580
         Width           =   540
      End
      Begin VB.CommandButton cmdFunc 
         Caption         =   "&<"
         Height          =   350
         Index           =   1
         Left            =   2760
         MousePointer    =   1  'Arrow
         TabIndex        =   13
         ToolTipText     =   "�����Ƴ�"
         Top             =   2160
         Width           =   540
      End
      Begin VB.CommandButton cmdFunc 
         Caption         =   "&>"
         Height          =   350
         Index           =   2
         Left            =   2760
         MousePointer    =   1  'Arrow
         TabIndex        =   12
         ToolTipText     =   "��������"
         Top             =   915
         Width           =   540
      End
      Begin VB.CommandButton cmdFunc 
         Caption         =   ">>"
         Height          =   350
         Index           =   3
         Left            =   2760
         MousePointer    =   1  'Arrow
         TabIndex        =   11
         ToolTipText     =   "ȫ������"
         Top             =   540
         Width           =   540
      End
   End
   Begin VB.OptionButton optPick 
      Caption         =   "ָ����Ա(&I)"
      Height          =   195
      Index           =   3
      Left            =   450
      TabIndex        =   9
      Top             =   1065
      Value           =   -1  'True
      Width           =   1365
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4560
      TabIndex        =   8
      Top             =   5985
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   3240
      TabIndex        =   7
      Top             =   5985
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   330
      TabIndex        =   6
      Top             =   5985
      Width           =   1100
   End
   Begin VB.OptionButton optPick 
      Caption         =   "������Ա(&N)"
      Height          =   195
      Index           =   4
      Left            =   2265
      TabIndex        =   5
      Top             =   1065
      Width           =   1365
   End
   Begin VB.OptionButton optPick 
      Caption         =   "��Ա����(&X)"
      Height          =   195
      Index           =   5
      Left            =   4260
      TabIndex        =   4
      Top             =   1065
      Width           =   1590
   End
   Begin VB.OptionButton optPick 
      Caption         =   "���Ҽ���(&S)"
      Height          =   195
      Index           =   7
      Left            =   2265
      TabIndex        =   3
      Top             =   1545
      Width           =   1305
   End
   Begin VB.OptionButton optPick 
      Caption         =   "��Ա����(&S)"
      Height          =   195
      Index           =   6
      Left            =   450
      TabIndex        =   2
      Top             =   1545
      Width           =   1305
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      Left            =   3720
      TabIndex        =   1
      Top             =   1500
      Width           =   1965
   End
   Begin VB.CommandButton cmdFind 
      Height          =   315
      Left            =   5700
      Picture         =   "frmSelectReceiver.frx":0014
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1485
      Width           =   390
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "�û�����ϵͳ(&S)"
      Height          =   180
      Left            =   375
      TabIndex        =   22
      Top             =   240
      Width           =   1350
   End
End
Attribute VB_Name = "frmSelectReceiver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnOK As Boolean

Dim mstr�ռ��� As String           '�ռ��˵�����

Dim mrs��Ա As New ADODB.Recordset '������Ա�嵥
Dim mrsϵͳ As New ADODB.Recordset '������ϵͳ

Private mrsUser As New ADODB.Recordset

Private Sub cmbSystem_Click()
    Dim strOwner As String
    
    mrsϵͳ.Filter = "���=" & cmbSystem.ItemData(cmbSystem.ListIndex)
    strOwner = mrsϵͳ("������")
    
    If mrs��Ա.State = 1 Then mrs��Ա.Close
    gstrSQL = "Select A.���� As ���ű��, B.����, D.�û���" & vbNewLine & _
            "From " & strOwner & ".���ű� A, " & strOwner & ".������Ա C, " & strOwner & ".�ϻ���Ա�� D, " & strOwner & ".��Ա�� B" & vbNewLine & _
            "Where A.ID = C.����id And B.ID = C.��Աid And C.��Աid = D.��Աid And C.ȱʡ = 1 And (B.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or B.����ʱ�� Is Null) " & vbNewLine & _
            "Order By B.����"


    Call zlDatabase.OpenRecordset(mrs��Ա, gstrSQL, Me.Caption)
    
    Call optPick_Click(0)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    Dim strOwner As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHandle
    If Trim(txt����.Text) = "" Then Exit Sub
    lst(0).Clear
    mrsϵͳ.Filter = "���=" & cmbSystem.ItemData(cmbSystem.ListIndex)
    strOwner = mrsϵͳ("������")
    
    If optPick(6).Value = True Then
    
        gstrSQL = "select DISTINCT B.����,D.�û��� " & _
                  " from " & strOwner & ".���ű� A," & strOwner & ".��Ա�� B," & _
                  strOwner & ".������Ա C," & strOwner & ".�ϻ���Ա�� D " & _
                  "  where A.ID=C.����ID and B.ID=C.��ԱID and C.��ԱID=D.��ԱID And (B.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or B.����ʱ�� Is Null) and C.ȱʡ=1 " & _
                  " And Upper(B.����) Like '%" & UCase(Trim(txt����.Text)) & "%' order by B.����"
                  
        Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
        Do Until rsTemp.EOF
            lst(0).AddItem rsTemp("����") & "(" & rsTemp("�û���") & ")"
            rsTemp.MoveNext
        Loop
    ElseIf optPick(7).Value = True Then
        gstrSQL = "Select Distinct A.����,A.���� From " & strOwner & ".���ű� A Where (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null)" & _
                  " And Upper(A.����) Like '%" & UCase(Trim(txt����.Text)) & "%' order by A.����,A.����"
        Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
        Do Until rsTemp.EOF
            lst(0).AddItem rsTemp("����") & "-" & rsTemp("����")
            rsTemp.MoveNext
        Loop
        
    End If
    If lst(0).ListCount > 0 Then lst(0).ListIndex = 0
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Err = 0
End Sub

Private Sub cmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, "ZL9AppTool\" & Me.Name, 0)
End Sub

Private Sub cmdOK_Click()
    Dim i As Long
    Dim intPos  As Long
    Dim strTemp As String
    Dim strOwner As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    mrsϵͳ.Filter = "���=" & cmbSystem.ItemData(cmbSystem.ListIndex)
    strOwner = mrsϵͳ("������")
    
'    mstr�û� = ""
'    mstr���� = ""
    mstr�ռ��� = ""
    
    Dim strFild As String
    strFild = "�û���,Varchar2,30;����,varchar2,30;�ռ���,varchar2,30"
    Set mrsUser = NewClientRecord(strFild)

    
    If optPick(3).Value = True Or optPick(4).Value = True Or optPick(6).Value = True Then
        
        '�����б��õ���Ա����
        For i = 0 To lst(1).ListCount - 1
            If lst(1).List(i) <> "" Then
                'ȥ�����ߵ�����
                mrsUser.AddNew
                intPos = InStr(lst(1).List(i), "(")
                strTemp = Mid(lst(1).List(i), intPos + 1)
                strTemp = Mid(strTemp, 1, Len(strTemp) - 1)
                mrsUser.Fields("�û���") = strTemp
                '����ǰΪ�û�����
                strTemp = Mid(lst(1).List(i), 1, intPos - 1)
                mstr�ռ��� = mstr�ռ��� & strTemp & ","
                mrsUser.Fields("����") = strTemp
                mrsUser.Fields("�ռ���") = strTemp
            End If
        Next
        If mstr�ռ��� <> "" Then
            mstr�ռ��� = Mid(mstr�ռ���, 1, Len(mstr�ռ���) - 1)
        End If
    ElseIf optPick(5).Value = True Then
        '��Ա����:�Էֺŷָ�
        For i = 0 To lst(1).ListCount - 1
            mstr�ռ��� = mstr�ռ��� & lst(1).List(i) & ";"
        Next
        If mstr�ռ��� <> "" Then
           
            gstrSQL = "Select Distinct B.����, D.�û���" & vbNewLine & _
                    "From " & strOwner & ".��Ա����˵�� E, " & strOwner & ".�ϻ���Ա�� D, " & strOwner & ".��Ա�� B" & vbNewLine & _
                    "Where B.ID = E.��Աid And B.ID = D.��Աid And (B.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or B.����ʱ�� Is Null) And Instr('" & mstr�ռ��� & "', E.��Ա����) > 0"
            Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
            Do Until rsTemp.EOF
                mrsUser.AddNew
                mrsUser.Fields("�û���") = rsTemp.Fields("�û���")
                mrsUser.Fields("����") = rsTemp.Fields("����")
                rsTemp.MoveNext
            Loop
            mstr�ռ��� = "[" & Mid(mstr�ռ���, 1, Len(mstr�ռ���) - 1) & "]"
            
        End If
    ElseIf optPick(7).Value = True Then
        For i = 0 To lst(1).ListCount - 1
            mstr�ռ��� = mstr�ռ��� & lst(1).List(i) & ";"
        Next
        If mstr�ռ��� <> "" Then
            
            gstrSQL = "select DISTINCT B.����,D.�û��� " & _
                      " from " & strOwner & ".���ű� A," & strOwner & ".��Ա�� B," & _
                      strOwner & ".������Ա C," & strOwner & ".�ϻ���Ա�� D " & _
                      "  where A.ID=C.����ID and B.ID=C.��ԱID and C.��ԱID=D.��ԱID And (B.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or B.����ʱ�� Is Null) " & _
                      "  And Instr('" & mstr�ռ��� & "', A.����||'-'||A.���� ) > 0" & _
                      " order by B.����"
            Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
            Do Until rsTemp.EOF
                mrsUser.AddNew
                mrsUser.Fields("�û���") = rsTemp.Fields("�û���")
                mrsUser.Fields("����") = rsTemp.Fields("����")
                rsTemp.MoveNext
            Loop
            mstr�ռ��� = "{" & Mid(mstr�ռ���, 1, Len(mstr�ռ���) - 1) & "}"
            
        End If
    Else
        If optPick(2).Value = True Then
        '�����ݿ��еõ���Ա����
            mstr�ռ��� = "��������Ա"
            mrs��Ա.Filter = "���ű��='" & gstrDeptCode & "'"
        ElseIf optPick(1).Value = True Then
            mstr�ռ��� = "��������Ա"
            If gstrDeptCode = "" Then
                mrs��Ա.Filter = "���ű��='��'"
            Else
                mrs��Ա.Filter = "���ű�� like '" & gstrDeptCode & "%'"
            End If
        Else
            mstr�ռ��� = "������Ա"
            mrs��Ա.Filter = 0
        End If
        Do Until mrs��Ա.EOF
            mrsUser.AddNew
            mrsUser.Fields("�ռ���") = mstr�ռ���
            mrsUser.Fields("�û���") = mrs��Ա("�û���")
            mrsUser.Fields("����") = mrs��Ա("����")
            
            mrs��Ա.MoveNext
        Loop
    End If
        
    mblnOK = True
    Unload Me
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmdFunc_Click(Index As Integer)
    '����ָ����Ա��ѡ��
    Dim strTemp As String
    Dim i As Long
    
    Select Case Index
        Case 0
            lst(1).Clear
        Case 1
            i = lst(1).ListIndex
            If i >= 0 Then
                lst(1).RemoveItem i
                If i > lst(1).ListCount - 1 Then
                    lst(1).ListIndex = lst(1).ListCount - 1
                Else
                    lst(1).ListIndex = i
                End If
            End If
        Case 2
            If lst(0).ListIndex >= 0 Then
                strTemp = lst(0).List(lst(0).ListIndex)
                For i = 0 To lst(1).ListCount - 1
                    If lst(1).List(i) = strTemp Then Exit For
                Next
                If i > lst(1).ListCount - 1 Then lst(1).AddItem strTemp
                If lst(1).ListIndex < 0 Then lst(1).ListIndex = 0
            End If
        Case 3
            lst(1).Clear
            For i = 0 To lst(0).ListCount - 1
                lst(1).AddItem lst(0).List(i)
            Next
            If lst(1).ListIndex < 0 And lst(1).ListCount > 0 Then lst(1).ListIndex = 0
    End Select
End Sub

Private Sub Form_Load()
    cmdFind.Enabled = False
    txt����.Enabled = False
End Sub

Private Sub lst_DblClick(Index As Integer)
    If Index = 0 Then
        cmdFunc_Click 2
    Else
        cmdFunc_Click 1
    End If
End Sub

Private Sub optPick_Click(Index As Integer)
    If mrs��Ա.State = 0 Then Exit Sub
    Dim strOwner As String
    Dim var�ռ��� As Variant, strTmp As String, i As Integer

    Dim blnList As Boolean, rsTemp As New ADODB.Recordset
    On Error GoTo errHandle
    mrsϵͳ.Filter = "���=" & cmbSystem.ItemData(cmbSystem.ListIndex)
    strOwner = mrsϵͳ("������")
    
    blnList = optPick(3).Value Or optPick(4).Value
    fra(1).Enabled = blnList
    lst(0).Enabled = blnList
    lst(1).Enabled = blnList
    cmdFunc(0).Enabled = blnList
    cmdFunc(1).Enabled = blnList
    cmdFunc(2).Enabled = blnList
    cmdFunc(3).Enabled = blnList
    
    cmdFind.Enabled = False
    txt����.Enabled = False
    
    '����Ҫ�б�
    lst(0).Clear

    
    If blnList = True Then
        If optPick(3).Value = True Then
            '��������Ա��ѡȡ
            gstrSQL = "select DISTINCT B.����,D.�û��� " & _
                      " from " & strOwner & ".���ű� A," & strOwner & ".��Ա�� B," & _
                      strOwner & ".������Ա C," & strOwner & ".�ϻ���Ա�� D " & _
                      "  where A.ID=C.����ID and B.ID=C.��ԱID and C.��ԱID=D.��ԱID And (B.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or B.����ʱ�� Is Null) and C.ȱʡ=1 order by B.����"
        Else
            '��������Ա��ѡȡ
            gstrSQL = "select DISTINCT B.����,D.�û��� " & _
                      " from " & strOwner & ".���ű� A," & strOwner & ".��Ա�� B," & _
                      strOwner & ".������Ա C," & strOwner & ".�ϻ���Ա�� D,V$session S " & _
                      "  where A.ID=C.����ID and B.ID=C.��ԱID and C.��ԱID=D.��ԱID And (B.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or B.����ʱ�� Is Null) and C.ȱʡ=1 AND D.�û���=S.USERNAME order by B.����"
        End If
        Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
        Do Until rsTemp.EOF
            lst(0).AddItem rsTemp("����") & "(" & rsTemp("�û���") & ")"
            rsTemp.MoveNext
        Loop
        If lst(0).ListCount > 0 Then lst(0).ListIndex = 0
        
        lst(1).Clear
        If InStr(mstr�ռ���, "]") <= 0 And InStr(mstr�ռ���, "[") <= 0 Then
            If Not mrsUser Is Nothing Then
                If mrsUser.State = adStateOpen Then
                    If mrsUser.RecordCount > 0 Then mrsUser.MoveFirst
                    Do Until mrsUser.EOF
                        lst(1).AddItem mrsUser.Fields("����") & "(" & mrsUser.Fields("�û���") & ")"
                        mrsUser.MoveNext
                    Loop
                End If
            End If
        End If
    End If
    
    
    If optPick(5).Value = True Then
        lst(0).Clear
        fra(1).Enabled = True
        lst(0).Enabled = True
        lst(1).Enabled = True
        cmdFunc(0).Enabled = True
        cmdFunc(1).Enabled = True
        cmdFunc(2).Enabled = True
        cmdFunc(3).Enabled = True
        
        lst(1).Clear
        If InStr(mstr�ռ���, "]") > 0 And InStr(mstr�ռ���, "[") > 0 Then
            strTmp = Mid(mstr�ռ���, 2, Len(mstr�ռ���) - 2)
            If InStr(strTmp, ";") > 0 Then
                var�ռ��� = Split(strTmp, ";")
                For i = LBound(var�ռ���) To UBound(var�ռ���)
                    lst(1).AddItem var�ռ���(i)
                Next
            Else
                lst(1).AddItem strTmp
            End If

        End If
        gstrSQL = "Select ����,���� From " & strOwner & ".��Ա���ʷ���"
        Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
        Do Until rsTemp.EOF
            lst(0).AddItem rsTemp("����")
            rsTemp.MoveNext
        Loop
        If lst(0).ListCount > 0 Then lst(0).ListIndex = 0
    End If
    
    If optPick(6).Value = True Then

        lst(0).Clear
        fra(1).Enabled = True
        lst(0).Enabled = True
        lst(1).Enabled = True
        cmdFunc(0).Enabled = True
        cmdFunc(1).Enabled = True
        cmdFunc(2).Enabled = True
        cmdFunc(3).Enabled = True
        
        cmdFind.Enabled = True
        txt����.Enabled = True
        
        lst(1).Clear
        If InStr(mstr�ռ���, "]") <= 0 And InStr(mstr�ռ���, "[") <= 0 Then
            If Not mrsUser Is Nothing Then
                If mrsUser.State = adStateOpen Then
                    If mrsUser.RecordCount > 0 Then mrsUser.MoveFirst
                    Do Until mrsUser.EOF
                        lst(1).AddItem mrsUser.Fields("����") & "(" & mrsUser.Fields("�û���") & ")"
                        mrsUser.MoveNext
                    Loop
                End If
            End If
        End If
        
    End If
    
    If optPick(7).Value = True Then
        lst(0).Clear
        fra(1).Enabled = True
        lst(0).Enabled = True
        lst(1).Enabled = True
        cmdFunc(0).Enabled = True
        cmdFunc(1).Enabled = True
        cmdFunc(2).Enabled = True
        cmdFunc(3).Enabled = True
        
        cmdFind.Enabled = True
        txt����.Enabled = True
        
        lst(1).Clear
        If InStr(mstr�ռ���, "}") > 0 And InStr(mstr�ռ���, "{") > 0 Then
            strTmp = Mid(mstr�ռ���, 2, Len(mstr�ռ���) - 2)
            If InStr(strTmp, ";") > 0 Then
                var�ռ��� = Split(strTmp, ";")
                For i = LBound(var�ռ���) To UBound(var�ռ���)
                    lst(1).AddItem var�ռ���(i)
                Next
            Else
                lst(1).AddItem strTmp
            End If

        End If

        If lst(0).ListCount > 0 Then lst(0).ListIndex = 0
    End If
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Function Get�ռ���(str�ռ��� As String, rsUser As ADODB.Recordset) As Boolean
    
    Dim var�ռ��� As Variant, strTmp As String, i As Integer
    On Error GoTo errHandle
    mblnOK = False
    mstr�ռ��� = str�ռ���
    
    Set mrsUser = rsUser
    '-----------------------------------
    '���ݴ����Ĳ���������ʾ
    lst(1).Clear
    Select Case str�ռ���
        Case "������Ա"
            optPick(0).Value = True
        Case "��������Ա"
            optPick(1).Value = True
        Case "��������Ա"
            optPick(2).Value = True
        Case Else
            If InStr(str�ռ���, "[") > 0 And InStr(str�ռ���, "]") > 0 Then
                '��������
                optPick(5).Value = True
                lst(1).Clear
                strTmp = Mid(str�ռ���, 2, Len(str�ռ���) - 2)
                If InStr(strTmp, ";") > 0 Then
                    var�ռ��� = Split(strTmp, ";")
                    For i = 0 To UBound(var�ռ���)
                        lst(1).AddItem var�ռ���(i)
                    Next
                Else
                    lst(1).AddItem strTmp
                End If
            Else
                optPick(3).Value = True
                If Not rsUser Is Nothing Then
                    If rsUser.State = adStateOpen Then
                        If rsUser.RecordCount > 0 Then rsUser.MoveFirst
                        Do Until rsUser.EOF
                            lst(1).AddItem rsUser.Fields("����") & "(" & rsUser.Fields("�û���") & ")"
                            rsUser.MoveNext
                        Loop
                    End If
                End If
            End If
            If lst(1).ListCount > 0 Then lst(1).ListIndex = 0
    End Select
    
    '�õ�ϵͳ
    gstrSQL = "select A.���,A.���� ||'��'||A.���||'��' as ����,A.������ from zlsystems A, (select owner from all_tables where " & _
               " table_name in ('���ű�','��Ա��','������Ա','�ϻ���Ա��') " & _
               " group by owner " & _
               " having count(table_name)=4) B " & _
               " Where A.������ = B.owner"
    Call zlDatabase.OpenRecordset(mrsϵͳ, gstrSQL, Me.Caption)
    
    If mrsϵͳ.EOF Then
        MsgBox "�㲻����ѡ���ռ��˵�Ȩ�ޣ�����ʹ�ñ����ܡ�", vbInformation, gstrSysName
        Exit Function
    End If
    cmbSystem.Clear
    Do Until mrsϵͳ.EOF
        cmbSystem.AddItem mrsϵͳ("����")
        cmbSystem.ItemData(cmbSystem.NewIndex) = mrsϵͳ("���")
        mrsϵͳ.MoveNext
    Loop
    If cmbSystem.ListCount > 0 Then cmbSystem.ListIndex = 0
    If cmbSystem.ListCount = 1 Then cmbSystem.Enabled = False
    
    
    'ͨ��cmbSystem��ѡ���Ѿ��õ���Ա�嵥
    
    frmSelectReceiver.Show vbModal
    Get�ռ��� = mblnOK
    If mblnOK = True Then
        str�ռ��� = mstr�ռ���
        Set rsUser = mrsUser
    End If
    If mrs��Ա.State = 1 Then mrs��Ա.Close
    Set mrs��Ա = Nothing
    If mrsϵͳ.State = 1 Then mrsϵͳ.Close
    Set mrsϵͳ = Nothing
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function



