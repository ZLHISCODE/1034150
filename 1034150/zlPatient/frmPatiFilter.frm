VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPatiFilter 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7740
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   7740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdDef 
      Caption         =   "ȱʡ(&D)"
      Height          =   350
      Left            =   6480
      TabIndex        =   7
      Top             =   2475
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6480
      TabIndex        =   6
      Top             =   735
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6480
      TabIndex        =   5
      Top             =   300
      Width           =   1100
   End
   Begin VB.Frame fraBdr 
      Height          =   3060
      Left            =   120
      TabIndex        =   8
      Top             =   30
      Width           =   6225
      Begin VB.CheckBox chk�Ǽ� 
         Caption         =   "�Ǽ�ʱ��"
         Height          =   195
         Left            =   180
         TabIndex        =   27
         Top             =   338
         Value           =   1  'Checked
         Width           =   1020
      End
      Begin VB.CheckBox chk���� 
         Caption         =   "��������"
         Height          =   195
         Left            =   180
         TabIndex        =   24
         Top             =   713
         Value           =   1  'Checked
         Width           =   1020
      End
      Begin VB.CheckBox chk��Ժ 
         Caption         =   "��Ժʱ��"
         Height          =   195
         Left            =   180
         TabIndex        =   20
         Top             =   2250
         Value           =   1  'Checked
         Width           =   1020
      End
      Begin VB.CommandButton cmd���� 
         Caption         =   "��"
         Height          =   255
         Left            =   5730
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "�ȼ���F3"
         Top             =   1450
         Width           =   285
      End
      Begin VB.TextBox txtIdentity 
         Height          =   300
         IMEMode         =   1  'ON
         Left            =   2520
         MaxLength       =   50
         TabIndex        =   2
         Top             =   1800
         Width           =   3480
      End
      Begin VB.ComboBox cboIDKind 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1230
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1800
         Width           =   1290
      End
      Begin VB.CheckBox chk��Ժ 
         Caption         =   "��Ժʱ��"
         Height          =   195
         Left            =   180
         TabIndex        =   3
         Top             =   2625
         Value           =   1  'Checked
         Width           =   1020
      End
      Begin VB.ComboBox cbo�Ա� 
         Height          =   300
         Left            =   3945
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   1050
         Width           =   2085
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         Left            =   3945
         MaxLength       =   30
         TabIndex        =   16
         Top             =   1425
         Width           =   2085
      End
      Begin MSComCtl2.DTPicker dtp��ԺE 
         Height          =   300
         Left            =   3945
         TabIndex        =   17
         Top             =   2565
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   144310275
         CurrentDate     =   40544
      End
      Begin MSComCtl2.DTPicker dtp��ԺB 
         Height          =   300
         Left            =   1230
         TabIndex        =   18
         Top             =   2565
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   144310275
         CurrentDate     =   40544
      End
      Begin MSComCtl2.DTPicker dtp��ԺE 
         Height          =   300
         Left            =   3945
         TabIndex        =   21
         Top             =   2190
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   144310275
         CurrentDate     =   40544
      End
      Begin MSComCtl2.DTPicker dtp��ԺB 
         Height          =   300
         Left            =   1230
         TabIndex        =   22
         Top             =   2190
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   144310275
         CurrentDate     =   40544
      End
      Begin MSComCtl2.DTPicker dtp����E 
         Height          =   300
         Left            =   3945
         TabIndex        =   25
         Top             =   660
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   144310275
         CurrentDate     =   40544
      End
      Begin MSComCtl2.DTPicker dtp�Ǽ�E 
         Height          =   300
         Left            =   3945
         TabIndex        =   28
         Top             =   285
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   144310275
         CurrentDate     =   40544
      End
      Begin MSComCtl2.DTPicker dtp����B 
         Height          =   300
         Left            =   1230
         TabIndex        =   32
         Top             =   660
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   144310275
         CurrentDate     =   40544
      End
      Begin MSComCtl2.DTPicker dtp�Ǽ�B 
         Height          =   300
         Left            =   1230
         TabIndex        =   33
         Top             =   285
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   144310275
         CurrentDate     =   40544
      End
      Begin VB.TextBox txt��� 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1230
         TabIndex        =   31
         Top             =   1425
         Visible         =   0   'False
         Width           =   2085
      End
      Begin VB.ComboBox cbo�ѱ� 
         Height          =   300
         Left            =   1230
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1425
         Width           =   2085
      End
      Begin VB.TextBox txtסԺ�� 
         Height          =   300
         IMEMode         =   1  'ON
         Left            =   1230
         MaxLength       =   18
         TabIndex        =   30
         Top             =   1050
         Width           =   2085
      End
      Begin VB.Label lbl�Ǽ� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   180
         Left            =   3540
         TabIndex        =   29
         Top             =   345
         Width           =   180
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   180
         Left            =   3540
         TabIndex        =   26
         Top             =   720
         Width           =   180
      End
      Begin VB.Label lbl��Ժ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   180
         Left            =   3540
         TabIndex        =   23
         Top             =   2250
         Width           =   180
      End
      Begin VB.Label lbl��Ժ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   180
         Left            =   3540
         TabIndex        =   19
         Top             =   2625
         Width           =   180
      End
      Begin VB.Label lblIDKind 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������"
         Height          =   180
         Left            =   480
         TabIndex        =   14
         Top             =   1860
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ��"
         Height          =   180
         Left            =   630
         TabIndex        =   13
         Top             =   1110
         Width           =   540
      End
      Begin VB.Label lbl���� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   3450
         TabIndex        =   11
         Top             =   1485
         Width           =   360
      End
      Begin VB.Label lbl�ѱ� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ѱ�"
         Height          =   180
         Left            =   810
         TabIndex        =   10
         Top             =   1485
         Width           =   360
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ա�"
         Height          =   180
         Left            =   3450
         TabIndex        =   9
         Top             =   1110
         Width           =   360
      End
      Begin VB.Label lbl��� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   180
         Left            =   480
         TabIndex        =   12
         Top             =   1485
         Visible         =   0   'False
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmPatiFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������

Public mbytType As Byte '��:�����嵥����0-����,1-��Ժ,2-��Ժ,3-����,4-����
Public mstrFilter As String '��:����
Public mstrFilterInfo As String '������Ϣ ר�ù�������
Public mbytInFun As Byte '0-��ͨ����,1-���ⲡ�˹��˵���

Private Const mstrIDKind = "1-����;2-���￨;3-�����;4-ҽ����;5-����֤��;6-IC����;7-�ֻ���"
Private WithEvents mobjIDCard As clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1


Private Sub cmd����_Click()
    Dim rsTmp As ADODB.Recordset
    Set rsTmp = GetArea(Me, txt����, True)
    If Not rsTmp Is Nothing Then
        txt����.Text = rsTmp!����
        txt����.SelStart = Len(txt����.Text)
        txt����.SetFocus
    Else
        SelAll txt����
        txt����.SetFocus
    End If
End Sub

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, _
                            ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
    If txtIdentity.Text = "" And Not txtIdentity.Locked And Me.ActiveControl Is txtIdentity Then
        cboIDKind.ListIndex = 4
        txtIdentity.Text = strID
    End If
End Sub


Private Sub chk�Ǽ�_Click()
    If chk�Ǽ�.Tag <> "" Then chk�Ǽ�.Value = 0: Exit Sub
    dtp�Ǽ�B.Enabled = (chk�Ǽ�.Value = 1)
    dtp�Ǽ�E.Enabled = dtp�Ǽ�B.Enabled
    If dtp�Ǽ�B.Enabled Then dtp�Ǽ�B.SetFocus
End Sub

Private Sub chk����_Click()
    If chk����.Tag <> "" Then chk����.Value = 0: Exit Sub
    dtp����B.Enabled = (chk����.Value = 1)
    dtp����E.Enabled = dtp����B.Enabled
    If dtp����B.Enabled Then dtp����B.SetFocus
End Sub

Private Sub chk��Ժ_Click()
    If chk��Ժ.Tag <> "" Then chk��Ժ.Value = 0: Exit Sub
    dtp��ԺB.Enabled = (chk��Ժ.Value = 1)
    dtp��ԺE.Enabled = dtp��ԺB.Enabled
    If dtp��ԺB.Enabled Then dtp��ԺB.SetFocus
End Sub

Private Sub chk��Ժ_Click()
    If chk��Ժ.Tag <> "" Then chk��Ժ.Value = 0: Exit Sub
    dtp��ԺB.Enabled = (chk��Ժ.Value = 1)
    dtp��ԺE.Enabled = dtp��ԺB.Enabled
    If dtp��ԺB.Enabled Then dtp��ԺB.SetFocus
End Sub

Private Sub cmdCancel_Click()
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (False): Set mobjIDCard = Nothing
    gblnOK = False
    Hide
End Sub

Private Sub cmdDef_Click()
    Form_Load
End Sub



Private Sub cmdOK_Click()
    txtסԺ��.Text = Trim(txtסԺ��.Text)
    txtIdentity.Text = Trim(txtIdentity.Text)
    
    If txtסԺ��.Text = "" And txtIdentity.Text = "" Then
        If chk�Ǽ�.Value = 0 And chk��Ժ.Value = 0 And chk��Ժ.Value = 0 And mbytType <> 1 Then
            MsgBox "������ѡ��һ���Ǽ�ʱ�䷶Χ.", vbInformation, gstrSysName
            chk�Ǽ�.Value = 1
            Exit Sub
        End If
        
        If mbytType = 0 Then
            If chk�Ǽ�.Value = 0 Then
                MsgBox "������ѡ��һ���Ǽ�ʱ�䷶Χ.", vbInformation, gstrSysName
                chk�Ǽ�.Value = 1
                Exit Sub
            End If
        End If
    End If
        
    Call MakeFilter
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (False): Set mobjIDCard = Nothing
    gblnOK = True
    Hide
End Sub

Private Sub dtp����E_Change()
    dtp����B.MaxDate = dtp����E.Value
End Sub

Private Sub dtp��ԺE_Change()
    dtp��ԺB.MaxDate = dtp��ԺE.Value
End Sub

Private Sub dtp�Ǽ�E_Change()
    dtp�Ǽ�B.MaxDate = dtp�Ǽ�E.Value
End Sub

Private Sub dtp��ԺE_Change()
    dtp��ԺB.MaxDate = dtp��ԺE.Value
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    Select Case mbytType
        Case 0
            dtp�Ǽ�B.SetFocus
        Case 1
            chk��Ժ.SetFocus
        Case 2
            dtp��ԺB.SetFocus
        Case 3, 4
            dtp�Ǽ�B.SetFocus
    End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim rsTmp As ADODB.Recordset
    Dim curDate As Date, datTmp As Date, i As Integer
    
    txtIdentity.Text = ""
    '�������
    cboIDKind.Clear
    For i = 0 To UBound(Split(mstrIDKind, ";"))
        cboIDKind.AddItem Split(mstrIDKind, ";")(i)
    Next
    cboIDKind.ListIndex = 0
    
    lbl�ѱ�.Visible = mbytInFun = 0
    cbo�ѱ�.Visible = mbytInFun = 0
    lbl���.Visible = mbytInFun = 1
    txt���.Visible = mbytInFun = 1
    
    If mbytInFun = 0 Then
        '�ѱ�
        If glngSys Like "8??" Then
            lbl�ѱ�.Caption = "��Ա�ȼ�"
        Else
            If mbytType = 0 Or mbytType = 3 Or mbytType = 4 Then
                lbl�ѱ�.Caption = "����ѱ�"
            Else
                lbl�ѱ�.Caption = "סԺ�ѱ�"
            End If
        End If
        
        Set rsTmp = Nothing
        Set rsTmp = GetDictData("�ѱ�")
        cbo�ѱ�.Clear
        cbo�ѱ�.AddItem "���зѱ�"
        cbo�ѱ�.ListIndex = 0
        If Not rsTmp Is Nothing Then
            For i = 1 To rsTmp.RecordCount
                cbo�ѱ�.AddItem rsTmp!���� & "-" & rsTmp!����
                rsTmp.MoveNext
            Next
        End If
    ElseIf mbytInFun = 1 Then
        chk�Ǽ�.Caption = "����ʱ��"
    End If
    
    '�Ա�
    Set rsTmp = Nothing
    Set rsTmp = GetDictData("�Ա�")
    cbo�Ա�.Clear
    cbo�Ա�.AddItem "�����Ա�"
    cbo�Ա�.ListIndex = 0
    If Not rsTmp Is Nothing Then
        For i = 1 To rsTmp.RecordCount
            cbo�Ա�.AddItem rsTmp!���� & "-" & rsTmp!����
            rsTmp.MoveNext
        Next
    End If
    
    
    '���ó�ʼ����
    On Error Resume Next    '����ע����洢��Чʱ��ʱ����
    curDate = zlDatabase.Currentdate
    dtp�Ǽ�B.MaxDate = Format(DateAdd("d", 1, curDate), dtp�Ǽ�E.CustomFormat)
    dtp����B.MaxDate = Format(curDate, dtp����E.CustomFormat)
    dtp��ԺB.MaxDate = Format(DateAdd("d", 1, curDate), dtp��ԺE.CustomFormat)
    dtp��ԺB.MaxDate = Format(DateAdd("d", 1, curDate), dtp��ԺE.CustomFormat)
        
    datTmp = Format(curDate, "yyyy-MM-dd 00:00:00")
    dtp�Ǽ�B.Value = datTmp
    datTmp = Format(curDate, "yyyy-MM-dd 23:59:59")
    dtp�Ǽ�E.Value = datTmp
    
    datTmp = CDate(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName, "������ʼʱ��", Format(DateAdd("yyyy", -100, curDate), "yyyy-MM-dd")))
    dtp����B.Value = datTmp
    datTmp = CDate(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName, "��������ʱ��", Format(dtp����B.MaxDate, dtp����E.CustomFormat)))
    dtp����E.Value = datTmp
    
    datTmp = CDate(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName, "��Ժ��ʼʱ��", Format(curDate, "YYYY-MM-DD")))
    dtp��ԺB.Value = datTmp
    datTmp = CDate(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName, "��Ժ����ʱ��", Format(dtp��ԺB.MaxDate, dtp��ԺE.CustomFormat)))
    dtp��ԺE.Value = datTmp
    
    datTmp = CDate(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName, "��Ժ��ʼʱ��", Format(curDate, "YYYY-MM-DD")))
    dtp��ԺB.Value = datTmp
    datTmp = CDate(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName, "��Ժ����ʱ��", Format(dtp��ԺB.MaxDate, dtp��ԺE.CustomFormat)))
    dtp��ԺE.Value = datTmp
    
    On Error GoTo 0
    
    
    Select Case mbytType
        Case 0 '���в���
            chk�Ǽ�.Value = 1
            chk����.Value = 0
            chk��Ժ.Value = 0
            chk��Ժ.Value = 0
        Case 1 '��Ժ����
            chk�Ǽ�.Value = 0
            chk����.Value = 0
            chk��Ժ.Value = 0
            chk��Ժ.Value = 0: chk��Ժ.Tag = 1
        Case 2 '��Ժ����
            chk�Ǽ�.Value = 0
            chk����.Value = 0
            chk��Ժ.Value = 0
            chk��Ժ.Value = 1
        Case 3, 4 '���ﲡ��
            chk�Ǽ�.Value = 1
            chk����.Value = 0
            chk��Ժ.Value = 0: chk��Ժ.Tag = 1
            chk��Ժ.Value = 0: chk��Ժ.Tag = 1
    End Select
    
    If glngSys Like "8??" And Not Visible Then
        chk��Ժ.Visible = False
        dtp��ԺB.Visible = False
        dtp��ԺE.Visible = False
        lbl��Ժ.Visible = False
        chk��Ժ.Visible = False
        dtp��ԺB.Visible = False
        dtp��ԺE.Visible = False
        lbl��Ժ.Visible = False
        fraBdr.Height = fraBdr.Height - 900
        Me.Height = Me.Height - 900
        cmdOK.Top = cmdOK.Top - 100
        cmdCancel.Top = cmdCancel.Top - 100
        cmdDef.Top = cmdDef.Top - 800
    End If
End Sub

Public Sub MakeFilter()
    mstrFilter = ""
    mstrFilterInfo = "" 'ֻ�����������е�����
    If chk�Ǽ�.Value = 1 Then
        mstrFilter = mstrFilter & " And A.�Ǽ�ʱ�� Between [3] And [4]"
        mstrFilterInfo = mstrFilterInfo & " And A.�Ǽ�ʱ�� Between [3] And [4]"
    End If
    If chk����.Value = 1 Then mstrFilter = mstrFilter & " And A.�������� Between [5] And [6]"
    If chk��Ժ.Value = 1 Then mstrFilter = mstrFilter & " And P.��Ժ���� Between [7] And [8]"
    If chk��Ժ.Value = 1 Then mstrFilter = mstrFilter & " And P.��Ժ���� Between [9] And [10]"
    
    If txtסԺ��.Text <> "" Then
        mstrFilter = mstrFilter & " And A.סԺ��=[11]"
        mstrFilterInfo = mstrFilterInfo & " And A.סԺ��=[11]"
    End If
    If cbo�Ա�.ListIndex <> 0 Then mstrFilter = mstrFilter & " And A.�Ա�=[12]"
    If Trim(txt����.Text) <> "" Then mstrFilter = mstrFilter & " And A.����=[13]"
    
    '���������������ⲡ�˹���
    If txt���.Visible Then
        If txt���.Text <> "" Then mstrFilter = mstrFilter & " And C.���=[14]"
    Else
        '��ͬ�Ĳ鿴��Χʱ������ͬ
        If cbo�ѱ�.ListIndex <> 0 Then
            If mbytType = 0 Or mbytType = 3 Or mbytType = 4 Then
                mstrFilter = mstrFilter & " And A.�ѱ�=[14]"
            Else
                mstrFilter = mstrFilter & " And P.�ѱ�=[14]"
            End If
        End If
    End If
    
    If Trim(txtIdentity.Text) <> "" Then
        Select Case Val(cboIDKind.Text) '"1-����;2-���￨;3-�����;4-ҽ����;5-����֤��;6-IC����;7-�ֻ���"
            Case 1
                If chk�Ǽ�.Value = 1 Or chk��Ժ.Value = 1 Or chk��Ժ.Value = 1 Then
                    mstrFilter = Replace(mstrFilter, "�Ǽ�ʱ��", "�Ǽ�ʱ��+0") & " And A.���� like [15]"
                    mstrFilterInfo = Replace(mstrFilterInfo, "�Ǽ�ʱ��", "�Ǽ�ʱ��+0") & " And A.���� like [15]"
                Else
                    mstrFilter = Replace(mstrFilter, "�Ǽ�ʱ��", "�Ǽ�ʱ��+0") & " And A.����=[15]"
                    mstrFilterInfo = Replace(mstrFilterInfo, "�Ǽ�ʱ��", "�Ǽ�ʱ��+0") & " And A.����=[15]"
                End If
            Case 2
                mstrFilter = Replace(mstrFilter, "�Ǽ�ʱ��", "�Ǽ�ʱ��+0") & " And A.���￨��=[15]"
                mstrFilterInfo = Replace(mstrFilterInfo, "�Ǽ�ʱ��", "�Ǽ�ʱ��+0") & " And A.���￨��=[15]"
            Case 3
                mstrFilter = Replace(mstrFilter, "�Ǽ�ʱ��", "�Ǽ�ʱ��+0") & " And A.�����=[15]"
                mstrFilterInfo = Replace(mstrFilterInfo, "�Ǽ�ʱ��", "�Ǽ�ʱ��+0") & " And A.�����=[15]"
            Case 4
                mstrFilter = Replace(mstrFilter, "�Ǽ�ʱ��", "�Ǽ�ʱ��+0") & " And A.ҽ����=[15]"
                mstrFilterInfo = Replace(mstrFilterInfo, "�Ǽ�ʱ��", "�Ǽ�ʱ��+0") & " And A.ҽ����=[15]"
            Case 5
                mstrFilter = Replace(mstrFilter, "�Ǽ�ʱ��", "�Ǽ�ʱ��+0") & " And A.����֤��=[15]"
                mstrFilterInfo = Replace(mstrFilterInfo, "�Ǽ�ʱ��", "�Ǽ�ʱ��+0") & " And A.����֤��=[15]"
            Case 6
                mstrFilter = Replace(mstrFilter, "�Ǽ�ʱ��", "�Ǽ�ʱ��+0") & " And A.IC����=[15]"
                mstrFilterInfo = Replace(mstrFilterInfo, "�Ǽ�ʱ��", "�Ǽ�ʱ��+0") & " And A.IC����=[15]"
            Case 7
                mstrFilter = Replace(mstrFilter, "�Ǽ�ʱ��", "�Ǽ�ʱ��+0") & " And A.�ֻ���=[15]"
                mstrFilterInfo = Replace(mstrFilterInfo, "�Ǽ�ʱ��", "�Ǽ�ʱ��+0") & " And A.�ֻ���=[15]"
        End Select
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mbytInFun = 0
    
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName, "������ʼʱ��", Format(Me.dtp����B.Value, "YYYY-MM-DD")
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName, "��������ʱ��", Format(Me.dtp����E.Value, "yyyy-MM-dd 23:59:59")
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName, "��Ժ��ʼʱ��", Format(Me.dtp��ԺB.Value, "YYYY-MM-DD")
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName, "��Ժ����ʱ��", Format(Me.dtp��ԺE.Value, "yyyy-MM-dd 23:59:59")
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName, "��Ժ��ʼʱ��", Format(Me.dtp��ԺB.Value, "YYYY-MM-DD")
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName, "��Ժ����ʱ��", Format(Me.dtp��ԺE.Value, "yyyy-MM-dd 23:59:59")
End Sub

Private Sub txtIdentity_Change()
    If mobjIDCard Is Nothing Then
        Set mobjIDCard = New clsIDCard
        Call mobjIDCard.SetParent(Me.hWnd)
    End If
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (txtIdentity.Text = "" And Not txtIdentity.Locked)
End Sub

Private Sub txtIdentity_GotFocus()
    Call zlControl.TxtSelAll(txtIdentity)
    
    If mobjIDCard Is Nothing Then
        Set mobjIDCard = New clsIDCard
        Call mobjIDCard.SetParent(Me.hWnd)
    End If
    If Not mobjIDCard Is Nothing And txtIdentity.Text = "" And Not txtIdentity.Locked Then mobjIDCard.SetEnabled (True)
End Sub
'����27819 by lesfeng 2010-02-02
Private Sub txtIdentity_KeyPress(KeyAscii As Integer)
    '59340:������,2013-04-23,ȡ����дת��(��Ϊ��������ΪСд�ַ��������޷���ѯ����,�����Upper�ᵼ���޷�ʹ���������²�ѯЧ��)
    'KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If InStr(":��;��?��'��||", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub txtIdentity_LostFocus()
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (False)
End Sub

Private Sub txt���_GotFocus()
    Call zlControl.TxtSelAll(txt���)
End Sub

Private Sub txt���_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt����_GotFocus()
    SelAll txt����
    Call OpenIme(gstrIme)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt����.Text <> "" Then
            Set rsTmp = GetArea(Me, txt����)
            If Not rsTmp Is Nothing Then
                txt����.Text = rsTmp!����
                Call zlCommFun.PressKey(vbKeyTab)
            Else
                SelAll txt����
                txt����.SetFocus
            End If
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        CheckInputLen txt����, KeyAscii
    End If
End Sub

Private Sub txt����_LostFocus()
    If gstrIme <> "���Զ�����" Then Call OpenIme
End Sub

Private Sub txtסԺ��_GotFocus()
    Call zlControl.TxtSelAll(txtסԺ��)
End Sub

Private Sub txtסԺ��_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
