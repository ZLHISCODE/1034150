VERSION 5.00
Begin VB.Form frmIdentify 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "���������֤"
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
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdReadIC 
      Caption         =   "����"
      Height          =   405
      Left            =   4500
      TabIndex        =   1
      Top             =   1230
      Width           =   585
   End
   Begin VB.TextBox txtPass 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
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
         Name            =   "����"
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
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2505
      TabIndex        =   3
      Top             =   2865
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
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
         Caption         =   "���ˣ����������У�30��"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "ʣ���1000.00�����ν�1000.00"
         BeginProperty Font 
            Name            =   "����"
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
      Caption         =   "��  ��"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "��  ��"
      BeginProperty Font 
         Name            =   "����"
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

Private mobjICCard As Object 'IC������

Public Function ShowMe(frmParent As Object, ByVal lngSys As Long, ByVal lng����ID As Long, ByVal cur��� As Currency) As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim intMouse As Integer
    
    mblnOK = False
    mintCount = 3
    
    intMouse = Screen.MousePointer
    Screen.MousePointer = 0
    
    '��ȡ���￨��Ϣ
    On Error GoTo errH
    strSQL = "Select A.����,A.�Ա�,A.����,A.���￨��,A.����֤��,Nvl(B.Ԥ�����,0)-Nvl(B.�������,0) as ���" & _
        " From ������Ϣ A,������� B Where A.����ID=[1] And A.����ID=B.����ID(+) And B.����(+)=1"
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Name, lng����ID)
    Me.lblPati.Caption = "���ˣ�" & gobjComLib.zlCommFun.NVL(rsTmp!����) & _
        IIf(Not IsNull(rsTmp!�Ա�), "��" & rsTmp!�Ա�, "") & _
        IIf(Not IsNull(rsTmp!����), "��" & rsTmp!����, "")
    Me.lblMoney.Caption = "ʣ���" & Format(rsTmp!���, "0.00") & "�����ν�" & Format(cur���, "0.00")
    Me.txtCard.Tag = gobjComLib.zlCommFun.NVL(rsTmp!���￨��)
    Me.txtPass.Tag = gobjComLib.zlCommFun.NVL(rsTmp!����֤��)
    On Error GoTo 0
    '���û�о��￨��������֤
    If Me.txtCard.Tag = "" Then
        Screen.MousePointer = intMouse
        ShowMe = True: Exit Function
    End If
    
    '��ȡϵͳ������
    '���￨��������ʾ
    mblnCardHide = Val(gobjComLib.zlDatabase.GetPara(12, lngSys)) <> 0
    If mblnCardHide Then txtCard.PasswordChar = "*"
    '���￨����ĳ���
    mbytCardLen = Val(Split(gobjComLib.zlDatabase.GetPara(20, lngSys, , "7|7|7|7|7"), "|")(4))
    
    'IC������
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
        MsgBox "��ǰ�����벡�˵Ŀ��Ų������", vbExclamation, gstrSysName
        Unload Me: Exit Sub '���Ų�ƥ�䣬��׼����
    End If
    If txtPass.Text <> txtPass.Tag Then
        MsgBox "�����������", vbExclamation, gstrSysName
        txtPass.Text = "": mintCount = mintCount - 1
        If mintCount = 0 Then
            Unload Me '������󣬿�����2��
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
    
    '�Ƿ�ˢ�����
    blnCard = KeyAscii <> 8 And Len(txtCard.Text) = mbytCardLen - 1 And txtCard.SelLength <> Len(txtCard.Text)
    If blnCard Or KeyAscii = 13 Then
        If KeyAscii <> 13 Then
            txtCard.Text = txtCard.Text & Chr(KeyAscii)
            txtCard.SelStart = Len(txtCard.Text)
        End If
        KeyAscii = 0
        txtPass.SetFocus
    Else
        If InStr(":��;��?��" & Chr(22), Chr(KeyAscii)) > 0 Then
            KeyAscii = 0 'ȥ��������ţ����Ҳ�����ճ��
        Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        End If
           
        '��ȫˢ�����
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
        KeyAscii = 0 '������ճ��
    End If
End Sub
