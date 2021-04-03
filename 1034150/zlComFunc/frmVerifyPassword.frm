VERSION 5.00
Begin VB.Form frmVerifyPassword 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "密码验证"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4770
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmVerifyPassword.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   4770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picPati 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   4770
      TabIndex        =   5
      Top             =   0
      Width           =   4770
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   120
         Left            =   -135
         TabIndex        =   6
         Top             =   480
         Width           =   5100
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "姓名："
         Height          =   240
         Left            =   270
         TabIndex        =   12
         Top             =   195
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "性别："
         Height          =   240
         Left            =   2025
         TabIndex        =   11
         Top             =   195
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "年龄："
         Height          =   240
         Left            =   3345
         TabIndex        =   10
         Top             =   195
         Width           =   720
      End
      Begin VB.Label lbl姓名 
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   930
         TabIndex        =   9
         Top             =   195
         Width           =   960
      End
      Begin VB.Label lbl性别 
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   2685
         TabIndex        =   8
         Top             =   195
         Width           =   525
      End
      Begin VB.Label lbl年龄 
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   4005
         TabIndex        =   7
         Top             =   195
         Width           =   720
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   120
      Left            =   -105
      TabIndex        =   4
      Top             =   1575
      Width           =   5100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   400
      Left            =   2805
      TabIndex        =   2
      Top             =   1845
      Width           =   1450
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   400
      Left            =   1350
      TabIndex        =   1
      Top             =   1845
      Width           =   1450
   End
   Begin VB.TextBox txtPass 
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   2310
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   975
      Width           =   1950
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "请输入密码"
      Height          =   240
      Left            =   1020
      TabIndex        =   3
      Top             =   1035
      Width           =   1200
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   135
      Picture         =   "frmVerifyPassword.frx":058A
      Top             =   780
      Width           =   720
   End
End
Attribute VB_Name = "frmVerifyPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPass As String
Private mintTime As Integer
Private mblnOK As Boolean

Public Function ShowMe(frmParent As Object, ByVal strPass As String, Optional ByVal strName As String, Optional ByVal strSex As String, Optional ByVal strOld As String) As Boolean
    Dim objControl As Object
    
    mstrPass = strPass
    Load Me
    
    Me.lbl姓名.Caption = strName
    Me.lbl性别.Caption = strSex
    Me.lbl年龄.Caption = strOld
    If lbl姓名.Caption = "" And lbl性别.Caption = "" And lbl年龄.Caption = "" Then
        For Each objControl In Me.Controls
            If objControl Is picPati Then
                objControl.Visible = False
            Else
                objControl.Top = objControl.Top - picPati.Height
            End If
        Next
        Me.Height = Me.Height - picPati.Height
    End If
    
    Me.Show 1, frmParent
    ShowMe = mblnOK
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If txtPass.Text <> mstrPass Then
        If mintTime + 1 = 3 Then
            MsgBox "密码错误，你已经连续 3 次输入错误的密码，密码验证操作将中止！", vbExclamation, gstrSysName
            Unload Me
        Else
            MsgBox "密码输入错误，请重新输入！", vbExclamation, gstrSysName
        End If
        mintTime = mintTime + 1
        txtPass.Text = "": Exit Sub
    End If
    
    mblnOK = True
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdOK_Click
End Sub

Private Sub Form_Load()
    mblnOK = False
    mintTime = 0
End Sub

Private Sub txtPass_GotFocus()
    txtPass.SelStart = 0
    txtPass.SelLength = Len(txtPass.Text)
End Sub
