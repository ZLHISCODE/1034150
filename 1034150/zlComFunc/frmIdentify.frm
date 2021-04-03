VERSION 5.00
Begin VB.Form frmIdentify 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "病人身份验证"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5670
   Icon            =   "frmIdentify.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdReadIC 
      Caption         =   "读卡"
      Height          =   405
      Left            =   4500
      TabIndex        =   1
      Top             =   1230
      Width           =   585
   End
   Begin VB.TextBox txtPass 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   1455
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1815
      Width           =   3015
   End
   Begin VB.TextBox txtCard 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   1455
      TabIndex        =   0
      Top             =   1230
      Width           =   3015
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2505
      TabIndex        =   3
      Top             =   2865
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3720
      TabIndex        =   4
      Top             =   2865
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   0
      TabIndex        =   7
      Top             =   2700
      Width           =   6900
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   840
      Left            =   0
      ScaleHeight     =   840
      ScaleWidth      =   5670
      TabIndex        =   5
      Top             =   0
      Width           =   5670
      Begin VB.Label lblPati 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "病人：张永康，男，30岁"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   255
         TabIndex        =   10
         Top             =   105
         Width           =   2640
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   4845
         Picture         =   "frmIdentify.frx":058A
         Top             =   45
         Width           =   720
      End
      Begin VB.Label lblMoney 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "剩余款额：1000.00，本次金额：1000.00"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   240
         TabIndex        =   6
         Top             =   465
         Width           =   4320
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000010&
         X1              =   -45
         X2              =   6000
         Y1              =   810
         Y2              =   810
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "卡  号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   600
      TabIndex        =   9
      Top             =   1320
      Width           =   720
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "密  码"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   615
      TabIndex        =   8
      Top             =   1890
      Width           =   720
   End
End
Attribute VB_Name = "frmIdentify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOK As Boolean
Private mblnCardHide As Boolean
Private mbytCardLen As Byte
Private mintCount As Integer

Private mobjICCard As Object 'IC卡对象

Public Function ShowMe(frmParent As Object, ByVal lngSys As Long, ByVal lng病人ID As Long, ByVal cur金额 As Currency) As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim intMouse As Integer
    
    mblnOK = False
    mintCount = 3
    
    intMouse = Screen.MousePointer
    Screen.MousePointer = 0
    
    '读取就诊卡信息
    On Error GoTo errH
    strSQL = "Select A.姓名,A.性别,A.年龄,A.就诊卡号,A.卡验证码,Nvl(B.预交余额,0)-Nvl(B.费用余额,0) as 余额" & _
        " From 病人信息 A,病人余额 B Where A.病人ID=[1] And A.病人ID=B.病人ID(+) And B.性质(+)=1"
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Name, lng病人ID)
    Me.lblPati.Caption = "病人：" & gobjComLib.zlCommFun.NVL(rsTmp!姓名) & _
        IIf(Not IsNull(rsTmp!性别), "，" & rsTmp!性别, "") & _
        IIf(Not IsNull(rsTmp!年龄), "，" & rsTmp!年龄, "")
    Me.lblMoney.Caption = "剩余款额：" & Format(rsTmp!余额, "0.00") & "，本次金额：" & Format(cur金额, "0.00")
    Me.txtCard.Tag = gobjComLib.zlCommFun.NVL(rsTmp!就诊卡号)
    Me.txtPass.Tag = gobjComLib.zlCommFun.NVL(rsTmp!卡验证码)
    On Error GoTo 0
    '如果没有就诊卡，则不作验证
    If Me.txtCard.Tag = "" Then
        Screen.MousePointer = intMouse
        ShowMe = True: Exit Function
    End If
    
    '读取系统参数：
    '就诊卡号密文显示
    mblnCardHide = Val(gobjComLib.zlDatabase.GetPara(12, lngSys)) <> 0
    If mblnCardHide Then txtCard.PasswordChar = "*"
    '就诊卡号码的长度
    mbytCardLen = Val(Split(gobjComLib.zlDatabase.GetPara(20, lngSys, , "7|7|7|7|7"), "|")(4))
    
    'IC卡对象
    On Error Resume Next
    Set mobjICCard = CreateObject("zlICCard.clsICCard")
    On Error GoTo 0
    
    Me.Show 1, frmParent
    ShowMe = mblnOK
    
    Screen.MousePointer = intMouse
    Exit Function
errH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If UCase(txtCard.Text) <> UCase(txtCard.Tag) Then
        MsgBox "当前卡号与病人的卡号不相符！", vbExclamation, gstrSysName
        Unload Me: Exit Sub '卡号不匹配，不准重试
    End If
    If txtPass.Text <> txtPass.Tag Then
        MsgBox "密码输入错误！", vbExclamation, gstrSysName
        txtPass.Text = "": mintCount = mintCount - 1
        If mintCount = 0 Then
            Unload Me '密码错误，可输入2次
        ElseIf txtPass.Enabled Then
            txtPass.SetFocus
        End If
        Exit Sub
    End If
    
    mblnOK = True
    Unload Me
End Sub

Private Sub cmdReadIC_Click()
    If Not mobjICCard Is Nothing Then
        txtCard.Text = mobjICCard.Read_Card(Me)
        If txtCard.Text <> "" Then
            txtPass.SetFocus
        Else
            txtCard.SetFocus
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mobjICCard = Nothing
End Sub

Private Sub txtCard_Change()
    txtPass.Enabled = txtCard.Text <> ""
    If Not txtPass.Enabled Then txtPass.Text = ""
End Sub

Private Sub txtCard_GotFocus()
    Call gobjComLib.zlControl.TxtSelAll(txtCard)
End Sub

Private Sub txtCard_KeyPress(KeyAscii As Integer)
    Static sngBegin As Single
    Dim sngNow As Single
    Dim blnCard As Boolean
    
    '是否刷卡完成
    blnCard = KeyAscii <> 8 And Len(txtCard.Text) = mbytCardLen - 1 And txtCard.SelLength <> Len(txtCard.Text)
    If blnCard Or KeyAscii = 13 Then
        If KeyAscii <> 13 Then
            txtCard.Text = txtCard.Text & Chr(KeyAscii)
            txtCard.SelStart = Len(txtCard.Text)
        End If
        KeyAscii = 0
        txtPass.SetFocus
    Else
        If InStr(":：;；?？" & Chr(22), Chr(KeyAscii)) > 0 Then
            KeyAscii = 0 '去除特殊符号，并且不允许粘贴
        Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        End If
           
        '安全刷卡检测
        If KeyAscii <> 0 And KeyAscii > 32 Then
            sngNow = Timer
            If txtCard.Text = "" Then
                sngBegin = sngNow
            ElseIf Format((sngNow - sngBegin) / (Len(txtCard.Text) + 1), "0.000") >= 0.04 Then '>0.007>=0.01
                txtCard.Text = Chr(KeyAscii)
                txtCard.SelStart = 1
                KeyAscii = 0
                sngBegin = sngNow
            End If
        End If
    End If
End Sub

Private Sub txtCard_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        glngTXTProc = GetWindowLong(txtCard.hwnd, GWL_WNDPROC)
        Call SetWindowLong(txtCard.hwnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txtCard_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SetWindowLong(txtCard.hwnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txtPass_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        glngTXTProc = GetWindowLong(txtPass.hwnd, GWL_WNDPROC)
        Call SetWindowLong(txtPass.hwnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txtPass_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SetWindowLong(txtPass.hwnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txtPass_GotFocus()
    Call gobjComLib.zlControl.TxtSelAll(txtPass)
End Sub

Private Sub txtPass_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call cmdOK_Click
    ElseIf KeyAscii = 22 Then
        KeyAscii = 0 '不允许粘贴
    End If
End Sub
