VERSION 5.00
Begin VB.Form frmSetPara 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "参数设置"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5370
   Icon            =   "frmSetPara.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   5370
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fra 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   3615
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   5055
      Begin VB.Frame fraOpt 
         BackColor       =   &H80000005&
         Caption         =   "药品说明书"
         Height          =   1455
         Left            =   120
         TabIndex        =   14
         Top             =   1920
         Width           =   4815
         Begin VB.OptionButton optType 
            BackColor       =   &H80000005&
            Caption         =   "产品公用"
            Height          =   255
            Index           =   0
            Left            =   360
            MaskColor       =   &H00C0C0FF&
            TabIndex        =   17
            Top             =   360
            Width           =   1695
         End
         Begin VB.OptionButton optType 
            BackColor       =   &H80000005&
            Caption         =   "产品非公用"
            Height          =   255
            Index           =   1
            Left            =   3120
            TabIndex        =   16
            Top             =   360
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.TextBox txtPara 
            Height          =   360
            Index           =   1
            Left            =   1440
            MaxLength       =   8
            TabIndex        =   15
            Top             =   840
            Width           =   2895
         End
         Begin VB.Label lblName 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            Caption         =   "说明书端口"
            Height          =   180
            Index           =   1
            Left            =   360
            TabIndex        =   18
            ToolTipText     =   "药品说明书端口号"
            Top             =   930
            Width           =   900
         End
      End
      Begin VB.TextBox txtPara 
         Height          =   360
         Index           =   3
         Left            =   1560
         MaxLength       =   8
         TabIndex        =   12
         Top             =   1320
         Width           =   2895
      End
      Begin VB.TextBox txtPara 
         Height          =   360
         Index           =   2
         Left            =   1560
         MaxLength       =   8
         TabIndex        =   11
         Top             =   150
         Width           =   2895
      End
      Begin VB.TextBox txtPara 
         Height          =   360
         Index           =   0
         Left            =   1560
         MaxLength       =   16
         TabIndex        =   8
         Top             =   720
         Width           =   2895
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "审查端口"
         Height          =   180
         Index           =   3
         Left            =   240
         TabIndex        =   13
         ToolTipText     =   "用药审查端口号"
         Top             =   1410
         Width           =   720
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "医院编码"
         Height          =   180
         Index           =   2
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   720
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "IP"
         Height          =   180
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   810
         Width           =   180
      End
   End
   Begin VB.TextBox txtIn 
      Height          =   375
      Left            =   120
      MaxLength       =   50
      TabIndex        =   4
      Top             =   600
      Width           =   3855
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00EFF0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   5370
      TabIndex        =   0
      Top             =   3750
      Width           =   5370
      Begin VB.CommandButton cmdPara 
         Caption         =   "取消(&C)"
         Height          =   360
         Index           =   1
         Left            =   2880
         TabIndex        =   2
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdPara 
         BackColor       =   &H8000000E&
         Caption         =   "确定(&O)"
         Height          =   360
         Index           =   0
         Left            =   1680
         TabIndex        =   1
         Top             =   120
         Width           =   1100
      End
   End
   Begin VB.Frame fra 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   1335
      Index           =   0
      Left            =   360
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   3855
      Begin VB.CheckBox chk 
         BackColor       =   &H80000005&
         Caption         =   "启用药师审方干预系统"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Label lblInfo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "医院编码"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   120
      TabIndex        =   3
      Top             =   210
      Width           =   720
   End
End
Attribute VB_Name = "frmSetPara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum CMD_ENUM
    CMD_OK = 0
    CMD_CANCEL = 1
End Enum

Private Sub cmdPara_Click(Index As Integer)
    Dim blnOK As Boolean
    Dim strPara As String
    Dim strSQL As String
    Dim lngID As Long
    Dim rsTmp As ADODB.Recordset
    
    If Index = CMD_OK Then
        If gbytPass = DT And gstrVersion = "4.0" Then
            strPara = Trim(txtIn.Text)
        ElseIf gbytPass = MK And gstrVersion = "4.0" Then
            strPara = IIf(chk(0).Value = vbChecked, 1, 0)
        ElseIf gbytPass = HZYY Then
            strPara = HZYY_SetPara
        End If
        On Error GoTo errH
        strSQL = "Select count(1) as RowCount  From zlParameters Where 系统 = [1] And Nvl(模块, 0) = 0 And 参数号 = 90001"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "合理用药监测配置", glngSys)
        If Not rsTmp.EOF Then
            If rsTmp!RowCount = 0 Then
                lngID = zlDatabase.GetNextId("zlParameters")
                strSQL = "Insert Into zlParameters(ID, 系统, 模块, 参数号, 参数名, 参数值) Values (" & lngID & ", " & glngSys & ", Null, 90001, '合理用药监测配置','" & strPara & "')"
                Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
                blnOK = True
            End If
        End If
        If Not blnOK Then
            Call zlDatabase.SetPara(90001, strPara, glngSys)
        End If
    End If
    Unload Me
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    Dim strPara As String
    
    Me.Width = 4215
    Me.Height = 2535
    fra(0).Visible = False: fra(1).Visible = False
    txtIn.Visible = False
    lblInfo.Visible = False
    If gbytPass = DT And gstrVersion = "4.0" Then
        txtIn.Visible = True
        lblInfo.Visible = True
        strPara = zlDatabase.GetPara(90001, glngSys, , "1513")
        txtIn.Text = strPara
    ElseIf gbytPass = MK And gstrVersion = "4.0" Then
        fra(0).Visible = True
        strPara = zlDatabase.GetPara(90001, glngSys, , "0")
        chk(0).Value = IIf(Val(strPara) = 1, vbChecked, vbUnchecked)
    ElseIf gbytPass = HZYY Then
        Me.Height = 4785
        Me.Width = 5460
        fra(1).Visible = True
        strPara = HZYY_GetPara
        txtPara(0).Text = gstrIP
        txtPara(1).Text = gstrPort
        txtPara(2).Text = gstrHOSCODE
        txtPara(3).Text = gstrPortPlus
        optType(0).Value = (gbytType = 0)
        optType(1).Value = (gbytType = 1)
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If gbytPass = MK And gstrVersion = "4.0" Then
        fra(0).Move 120, 120, Me.ScaleWidth - 240, Me.ScaleHeight - picBottom.Height - 120
    ElseIf gbytPass = HZYY Then
        fra(1).Move 120, 120, Me.ScaleWidth - 240, Me.ScaleHeight - picBottom.Height - 120
    End If
    cmdPara(CMD_CANCEL).Left = picBottom.Width - 1100 - 120
    cmdPara(CMD_OK).Left = cmdPara(CMD_CANCEL).Left - 1100 - 60
End Sub

Private Sub optType_Click(Index As Integer)
    gbytType = Index
End Sub

Private Sub txtPara_Change(Index As Integer)
    If gbytPass = HZYY Then
        If Index = 0 Then
            gstrIP = txtPara(Index)
        ElseIf Index = 1 Then
            gstrPort = txtPara(Index)
        ElseIf Index = 2 Then
            gstrHOSCODE = txtPara(Index)
        ElseIf Index = 3 Then
            gstrPortPlus = txtPara(Index)
        End If
    End If
End Sub
